#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import io
import json
import re
import shutil
import subprocess
import sys
import zipfile
from copy import copy
from pathlib import Path
from tempfile import TemporaryDirectory

try:
    import numpy as np
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as ExcelImage
    from PIL import Image, ImageFilter
except ImportError as error:
    sys.exit(
        f"Missing Python dependency: {error.name}. Install with: "
        "python3 -m pip install openpyxl Pillow numpy"
    )


MIN_IMAGE_BYTES = 20 * 1024
MAX_IMAGE_BYTES = 200 * 1024
BLANK_HEADERS = {"联系方式", "船名", "许可照片", "许可开始", "许可截止", "许可类型", "换班类型"}
HEADER_ALIASES = {
    "sequence": ("序号", "编号", "no", "no."),
    "name": ("姓名", "船员姓名", "name"),
    "gender": ("性别", "sex", "gender"),
    "nationality": ("国籍", "nationality", "nation"),
    "passport": ("护照号码", "护照号", "证件号码", "passport no", "passport number", "passport"),
}


def _norm(value: object) -> str:
    return re.sub(r"[^a-z0-9一-龥]", "", str(value or "").lower())


def _sequence(path: Path) -> int | None:
    match = re.match(r"\s*0*(\d{1,3})(?:\D|$)", path.name)
    return int(match.group(1)) if match else None


def _numbered_files(folder: Path, suffixes: set[str]) -> dict[int, Path]:
    files: dict[int, Path] = {}
    for path in folder.iterdir():
        number = _sequence(path)
        if number and path.suffix.lower() in suffixes:
            if number in files:
                raise ValueError(f"Duplicate sequence {number}: {files[number].name}, {path.name}")
            files[number] = path
    return files


def _field_indexes(headers: list[object]) -> dict[str, int]:
    normalized = [_norm(value) for value in headers]
    indexes: dict[str, int] = {}
    for field, aliases in HEADER_ALIASES.items():
        for index, header in enumerate(normalized):
            if any(_norm(alias) == header or _norm(alias) in header for alias in aliases):
                indexes[field] = index
                break
    return indexes


def _crew_from_workbook(path: Path, sheet_name: str | None) -> list[dict[str, object]]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    candidates = [workbook[sheet_name]] if sheet_name else list(workbook.worksheets)
    selected = None
    header_row = 0
    indexes: dict[str, int] = {}
    for sheet in candidates:
        for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            found = _field_indexes(list(row))
            if {"name", "gender", "nationality", "passport"}.issubset(found):
                selected, header_row, indexes = sheet, row_number, found
                break
        if selected:
            break
    if selected is None:
        workbook.close()
        raise ValueError("Crew sheet header not found. Required: name, gender, nationality, passport number")
    crew = []
    next_sequence = 1
    for row in selected.iter_rows(min_row=header_row + 1, values_only=True):
        name = row[indexes["name"]] if indexes["name"] < len(row) else None
        passport = row[indexes["passport"]] if indexes["passport"] < len(row) else None
        if name in (None, "") and passport in (None, ""):
            continue
        sequence = row[indexes["sequence"]] if "sequence" in indexes and indexes["sequence"] < len(row) else next_sequence
        if sequence not in (None, ""):
            try:
                sequence = int(float(sequence))
            except (TypeError, ValueError):
                continue
        else:
            sequence = next_sequence
        nationality = str(row[indexes["nationality"]] or "")
        gender = str(row[indexes["gender"]] or "")
        crew.append({
            "sequence": sequence,
            "name": str(name or "").strip(),
            "gender": "女" if gender.strip().lower() in {"女", "f", "female", "2", "2-女"} else "男",
            "nationality": nationality.split("-", 1)[1] if "-" in nationality else nationality,
            "passport": str(passport or "").strip(),
        })
        next_sequence += 1
    workbook.close()
    if not crew:
        raise ValueError("No crew records found")
    return crew


def _crew_from_csv(path: Path) -> list[dict[str, object]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        headers = reader.fieldnames or []
        indexes = _field_indexes(headers)
        if not {"name", "gender", "nationality", "passport"}.issubset(indexes):
            raise ValueError("CSV headers must include name, gender, nationality and passport number")
        crew = []
        for position, row in enumerate(reader, start=1):
            values = [row.get(header, "") for header in headers]
            nationality = str(values[indexes["nationality"]] or "")
            gender = str(values[indexes["gender"]] or "")
            sequence = values[indexes["sequence"]] if "sequence" in indexes else position
            crew.append({
                "sequence": int(float(sequence)),
                "name": str(values[indexes["name"]] or "").strip(),
                "gender": "女" if gender.strip().lower() in {"女", "f", "female", "2", "2-女"} else "男",
                "nationality": nationality.split("-", 1)[1] if "-" in nationality else nationality,
                "passport": str(values[indexes["passport"]] or "").strip(),
            })
    return crew


def _read_crew(path: Path, sheet_name: str | None) -> list[dict[str, object]]:
    return _crew_from_csv(path) if path.suffix.lower() == ".csv" else _crew_from_workbook(path, sheet_name)


def _render_passport(pdf: Path, output: Path) -> None:
    result = subprocess.run(
        [
            "pdftoppm", "-jpeg", "-jpegopt", "quality=78,progressive=y,optimize=y",
            "-r", "100", "-f", "1", "-l", "1", "-singlefile", str(pdf), str(output.with_suffix("")),
        ], capture_output=True, text=True,
    )
    if result.returncode or not output.exists():
        raise RuntimeError(f"Passport render failed: {pdf.name}: {result.stderr.strip()}")


def _orient_passport(source: Path, folder: Path, sequence: int) -> Path:
    candidates: list[tuple[int, int, Path]] = []
    for angle in (0, 90, -90, 180):
        candidate = source
        if angle:
            candidate = folder / f"passport-{sequence:03d}-rotate-{angle}.jpg"
            subprocess.run(["magick", str(source), "-rotate", str(angle), str(candidate)], capture_output=True)
        result = subprocess.run(
            ["tesseract", str(candidate), "stdout", "-l", "eng", "--psm", "6"],
            capture_output=True, text=True,
        )
        lines = [re.sub(r"[^A-Z<]", "", line.upper()) for line in result.stdout.splitlines()]
        score = max((line.count("<") + (20 if re.match(r"P.?[A-Z]{3}", line) else 0) for line in lines), default=0)
        candidates.append((score, angle, candidate))
    score, angle, candidate = max(candidates, key=lambda item: item[0])
    return candidate if angle and score >= 25 else source


def _largest_run(flags: np.ndarray) -> tuple[int, int]:
    runs: list[tuple[int, int]] = []
    start = None
    for index, enabled in enumerate([*flags, False]):
        if enabled and start is None:
            start = index
        elif not enabled and start is not None:
            runs.append((start, index - 1))
            start = None
    return max(runs, key=lambda item: item[1] - item[0], default=(0, len(flags) - 1))


def _lower_document_crop(source: Path, color_only: bool) -> tuple[int, int, int, int]:
    with Image.open(source) as image:
        width, height = image.size
        small_width = 320
        small_height = round(height * small_width / width)
        pixels = np.asarray(image.convert("RGB").resize((small_width, small_height))).astype(float)
    saturation = np.ptp(pixels, axis=2)
    if color_only:
        mask, threshold = saturation > 15, 0.08
    else:
        edge = max(2, round(min(small_width, small_height) * 0.025))
        border = np.concatenate([
            pixels[:edge].reshape(-1, 3), pixels[-edge:].reshape(-1, 3),
            pixels[:, :edge].reshape(-1, 3), pixels[:, -edge:].reshape(-1, 3),
        ])
        background = np.median(border, axis=0)
        distance = np.max(np.abs(pixels - background), axis=2)
        mask = (distance > 20) & ((saturation > 7) | (pixels.mean(axis=2) < background.mean() - 14))
        threshold = 0.10
    expanded = np.asarray(Image.fromarray((mask * 255).astype("uint8")).filter(ImageFilter.MaxFilter(15))) > 0
    y0, y1 = _largest_run(expanded.mean(axis=1) > threshold)
    x0, x1 = _largest_run(expanded.mean(axis=0) > threshold)
    x_pad, y_pad = round((x1 - x0 + 1) * 0.04), round((y1 - y0 + 1) * 0.03)
    x0, x1 = max(0, x0 - x_pad), min(small_width - 1, x1 + x_pad)
    y0, y1 = max(0, y0 - y_pad), min(small_height - 1, y1 + y_pad)
    bio_top = round(y0 + (y1 - y0 + 1) * 0.46)
    return (
        round(x0 * width / small_width), round(bio_top * height / small_height),
        round((x1 + 1) * width / small_width), round((y1 + 1) * height / small_height),
    )


def _mrz_data_crop(source: Path, nationality: str) -> tuple[int, int, int, int] | None:
    result = subprocess.run(
        ["tesseract", str(source), "stdout", "-l", "eng", "--psm", "11", "tsv"],
        capture_output=True, text=True,
    )
    groups: dict[tuple[str, str, str], dict[str, list]] = {}
    for row in csv.DictReader(io.StringIO(result.stdout), delimiter="\t"):
        value = (row.get("text") or "").strip()
        if not value:
            continue
        key = (row["block_num"], row["par_num"], row["line_num"])
        group = groups.setdefault(key, {"text": [], "boxes": []})
        group["text"].append(value)
        group["boxes"].append(tuple(int(row[name]) for name in ("left", "top", "width", "height")))
    candidates = []
    for group in groups.values():
        if "".join(group["text"]).upper().count("<") < 5:
            continue
        boxes = group["boxes"]
        candidates.append((
            min(box[0] for box in boxes), min(box[1] for box in boxes),
            max(box[0] + box[2] for box in boxes), max(box[1] + box[3] for box in boxes),
        ))
    if not candidates:
        return None
    widest = max(candidate[2] - candidate[0] for candidate in candidates)
    lines = [candidate for candidate in candidates if candidate[2] - candidate[0] >= widest * 0.55]
    left, right = min(line[0] for line in lines), max(line[2] for line in lines)
    bottom = max(line[3] for line in lines)
    with Image.open(source) as image:
        width, height = image.size
    page_width = min(width, round((right - left) * 1.12))
    if page_width < width * 0.30:
        return None
    page_height = round(page_width / (1.42 if nationality == "埃及" else 1.55))
    center = (left + right) // 2
    x0 = max(0, center - page_width // 2)
    x1 = min(width, x0 + page_width)
    x0 = max(0, x1 - page_width)
    y1 = min(height, round(bottom + page_height * 0.06))
    return x0, max(0, y1 - page_height), x1, y1


def _crop_passport(source: Path, output: Path, nationality: str) -> None:
    box = None if nationality == "坦桑尼亚" else _mrz_data_crop(source, nationality)
    if box is None:
        box = _lower_document_crop(source, color_only=nationality == "坦桑尼亚")
    x0, y0, x1, y1 = box
    result = subprocess.run(
        [
            "magick", str(source), "-crop", f"{x1-x0}x{y1-y0}+{x0}+{y0}", "+repage",
            "-strip", "-interlace", "Plane", "-quality", "88", str(output),
        ], capture_output=True, text=True,
    )
    if result.returncode or not output.exists():
        raise RuntimeError(f"Passport crop failed: {source.name}: {result.stderr.strip()}")


def _bounded_image(source: Path, output: Path, normalize: bool = False) -> Path:
    if not normalize and MIN_IMAGE_BYTES <= source.stat().st_size <= MAX_IMAGE_BYTES:
        return source
    for dimension, quality in ((1400, 92), (1200, 88), (1000, 82), (900, 76), (800, 68), (700, 60), (600, 52)):
        result = subprocess.run(
            [
                "magick", str(source), "-auto-orient", "-thumbnail", f"{dimension}x{dimension}>",
                "-strip", "-interlace", "Plane", "-quality", str(quality), str(output),
            ], capture_output=True, text=True,
        )
        if result.returncode == 0 and output.exists() and MIN_IMAGE_BYTES <= output.stat().st_size <= MAX_IMAGE_BYTES:
            return output
    raise ValueError(f"Image cannot be compressed to 20-200 KB: {source.name}")


def _copy_row_style(sheet, source_row: int, target_row: int) -> None:
    sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height
    for column in range(1, sheet.max_column + 1):
        source, target = sheet.cell(source_row, column), sheet.cell(target_row, column)
        if source.has_style:
            target._style = copy(source._style)
        target.number_format = source.number_format
        target.protection = copy(source.protection)
        target.alignment = copy(source.alignment)


def _fit_image(path: Path, max_width: int, max_height: int) -> ExcelImage:
    image = ExcelImage(path)
    ratio = min(max_width / image.width, max_height / image.height)
    image.width, image.height = round(image.width * ratio), round(image.height * ratio)
    return image


def _find_target_sheet(workbook, requested: str | None):
    if requested:
        return workbook[requested]
    for sheet in workbook.worksheets:
        headers = {_norm(cell.value): cell.column for cell in sheet[1]}
        if {_norm("姓名"), _norm("人脸照片"), _norm("证件照片")}.issubset(headers):
            return sheet
    raise ValueError("Target template sheet not found")


def _target_columns(sheet) -> dict[str, int]:
    result = {}
    for cell in sheet[1]:
        header = _norm(cell.value)
        if header:
            result[header] = cell.column
    required = ["姓名", "性别", "国籍", "证件类型", "证件号码", "人脸照片", "证件照片"]
    missing = [name for name in required if _norm(name) not in result]
    if missing:
        raise ValueError(f"Template missing columns: {', '.join(missing)}")
    return result


def export(args: argparse.Namespace) -> dict[str, object]:
    crew = _read_crew(args.crew, args.crew_sheet)
    sequences = [int(item["sequence"]) for item in crew]
    if sequences != list(range(1, len(crew) + 1)):
        raise ValueError(f"Crew sequence must be continuous from 1: {sequences}")
    photos = _numbered_files(args.photos, {".jpg", ".jpeg", ".png"})
    passports = _numbered_files(args.passports, {".pdf"})
    expected = set(sequences)
    if set(photos) != expected or set(passports) != expected:
        raise ValueError(f"Sequence mismatch: crew={sorted(expected)}, photos={sorted(photos)}, passports={sorted(passports)}")

    workbook = load_workbook(args.template)
    sheet = _find_target_sheet(workbook, args.target_sheet)
    columns = _target_columns(sheet)
    sheet._images = []
    style_row = 2
    for row_number in range(2, 2 + len(crew)):
        if row_number != style_row:
            _copy_row_style(sheet, style_row, row_number)
        for column in range(1, sheet.max_column + 1):
            sheet.cell(row_number, column).value = None

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="ship-crew-info-") as temporary:
        temporary_path = Path(temporary)
        for item in crew:
            sequence = int(item["sequence"])
            row = sequence + 1
            sheet.cell(row, columns[_norm("姓名")], item["name"])
            sheet.cell(row, columns[_norm("性别")], item["gender"])
            sheet.cell(row, columns[_norm("国籍")], item["nationality"])
            sheet.cell(row, columns[_norm("证件类型")], args.document_type)
            sheet.cell(row, columns[_norm("证件号码")], item["passport"])

            face = _bounded_image(photos[sequence], temporary_path / f"face-{sequence:03d}.jpg", normalize=True)
            rendered = temporary_path / f"passport-{sequence:03d}.jpg"
            _render_passport(passports[sequence], rendered)
            oriented = _orient_passport(rendered, temporary_path, sequence)
            cropped = temporary_path / f"passport-cropped-{sequence:03d}.jpg"
            _crop_passport(oriented, cropped, str(item["nationality"]))
            document = _bounded_image(cropped, temporary_path / f"passport-bounded-{sequence:03d}.jpg")

            face_excel = _fit_image(face, args.image_width, args.image_height)
            face_excel.anchor = sheet.cell(row, columns[_norm("人脸照片")]).coordinate
            sheet.add_image(face_excel)
            document_excel = _fit_image(document, args.image_width, args.image_height)
            document_excel.anchor = sheet.cell(row, columns[_norm("证件照片")]).coordinate
            sheet.add_image(document_excel)
        workbook.save(args.output)

    blank_columns = [column for name, column in columns.items() if name in {_norm(value) for value in BLANK_HEADERS}]
    reloaded = load_workbook(args.output, data_only=False)
    target = _find_target_sheet(reloaded, args.target_sheet)
    if any(target.cell(row, column).value not in (None, "") for row in range(2, 2 + len(crew)) for column in blank_columns):
        raise ValueError("A required blank field contains data")
    formula_errors = [
        f"{worksheet.title}!{cell.coordinate}={cell.value}"
        for worksheet in reloaded.worksheets for row in worksheet.iter_rows() for cell in row
        if isinstance(cell.value, str) and cell.value.startswith(("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?"))
    ]
    if formula_errors:
        raise ValueError(f"Formula errors: {formula_errors}")
    with zipfile.ZipFile(args.output) as archive:
        media = [info for info in archive.infolist() if info.filename.startswith("xl/media/")]
    if len(media) != len(crew) * 2 or any(not MIN_IMAGE_BYTES <= info.file_size <= MAX_IMAGE_BYTES for info in media):
        raise ValueError("Embedded image count/size validation failed")
    return {
        "status": "success", "output": str(args.output.resolve()), "crew_count": len(crew),
        "image_count": len(media), "image_min_kb": round(min(info.file_size for info in media) / 1024, 1),
        "image_max_kb": round(max(info.file_size for info in media) / 1024, 1), "formula_errors": 0,
        "required_blank_fields": sorted(BLANK_HEADERS),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export crew data/photos/passports into the crew information Excel template")
    parser.add_argument("--template", type=Path, required=True, help="Target crew information .xlsx template")
    parser.add_argument("--crew", type=Path, required=True, help="Crew source .xlsx or .csv")
    parser.add_argument("--photos", type=Path, required=True, help="Numbered face photo directory")
    parser.add_argument("--passports", type=Path, required=True, help="Numbered passport PDF directory")
    parser.add_argument("--output", type=Path, required=True, help="Output .xlsx; do not point to the original template")
    parser.add_argument("--crew-sheet", help="Optional crew source sheet name")
    parser.add_argument("--target-sheet", help="Optional target template sheet name")
    parser.add_argument("--document-type", type=int, default=2, help="Passport document type code (default: 2)")
    parser.add_argument("--image-width", type=int, default=150)
    parser.add_argument("--image-height", type=int, default=220)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    for program in ("pdftoppm", "magick", "tesseract"):
        if not shutil.which(program):
            sys.exit(f"Missing system dependency: {program}. See references/INSTALL.md")
    for path in (args.template, args.crew, args.photos, args.passports):
        if not path.exists():
            sys.exit(f"Input not found: {path}")
    if args.output.resolve() == args.template.resolve():
        sys.exit("Refusing to overwrite the original template; choose another --output path")
    try:
        print(json.dumps(export(args), ensure_ascii=False, indent=2))
    except Exception as error:
        print(json.dumps({"status": "error", "error": str(error)}, ensure_ascii=False, indent=2), file=sys.stderr)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
