from __future__ import annotations

import re
import shutil
import subprocess
import zipfile
import csv
import io
from copy import copy
from pathlib import Path
from tempfile import TemporaryDirectory

import numpy as np
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image, ImageFilter


MIN_IMAGE_BYTES = 20 * 1024
MAX_IMAGE_BYTES = 200 * 1024


def _sequence(path: Path) -> int | None:
    match = re.match(r"\s*0*(\d{1,2})(?:\D|$)", path.name)
    return int(match.group(1)) if match else None


def _numbered_files(folder: Path, suffixes: set[str]) -> dict[int, Path]:
    files: dict[int, Path] = {}
    for path in folder.iterdir():
        number = _sequence(path)
        if number and path.suffix.lower() in suffixes:
            if number in files:
                raise ValueError(f"序号 {number} 有多个文件: {files[number].name}, {path.name}")
            files[number] = path
    return files


def _render_passport(pdf: Path, output: Path) -> None:
    if not shutil.which("pdftoppm"):
        raise RuntimeError("缺少 pdftoppm，无法生成护照证件照片")
    result = subprocess.run(
        [
            "pdftoppm", "-jpeg", "-jpegopt", "quality=78,progressive=y,optimize=y",
            "-r", "100", "-f", "1", "-l", "1", "-singlefile", str(pdf),
            str(output.with_suffix("")),
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode or not output.exists():
        raise RuntimeError(f"护照首页转换失败: {pdf.name}: {result.stderr.strip()}")


def _orient_passport(source: Path, folder: Path, sequence: int) -> Path:
    if not shutil.which("magick") or not shutil.which("tesseract"):
        return source
    candidates: list[tuple[int, int, Path]] = []
    for angle in (0, 90, -90, 180):
        candidate = source
        if angle:
            candidate = folder / f"passport-{sequence:02d}-rotate-{angle}.jpg"
            subprocess.run(
                ["magick", str(source), "-rotate", str(angle), str(candidate)],
                capture_output=True,
            )
        result = subprocess.run(
            ["tesseract", str(candidate), "stdout", "-l", "eng", "--psm", "6"],
            capture_output=True,
            text=True,
        )
        lines = [re.sub(r"[^A-Z<]", "", line.upper()) for line in result.stdout.splitlines()]
        score = max(
            (line.count("<") + (20 if re.match(r"P.?[A-Z]{3}", line) else 0) for line in lines),
            default=0,
        )
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
        mask = saturation > 15
        threshold = 0.08
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
    expanded = np.asarray(
        Image.fromarray((mask * 255).astype("uint8")).filter(ImageFilter.MaxFilter(15))
    ) > 0
    y0, y1 = _largest_run(expanded.mean(axis=1) > threshold)
    x0, x1 = _largest_run(expanded.mean(axis=0) > threshold)
    x_pad = round((x1 - x0 + 1) * 0.04)
    y_pad = round((y1 - y0 + 1) * 0.03)
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
        capture_output=True,
        text=True,
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
        text = "".join(group["text"]).upper()
        if text.count("<") < 5:
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
    y0 = max(0, y1 - page_height)
    return x0, y0, x1, y1


def _crop_passport(source: Path, output: Path, nationality: str) -> None:
    box = None if nationality == "坦桑尼亚" else _mrz_data_crop(source, nationality)
    if box is None:
        box = _lower_document_crop(source, color_only=nationality == "坦桑尼亚")
    x0, y0, x1, y1 = box
    result = subprocess.run(
        [
            "magick", str(source), "-crop", f"{x1 - x0}x{y1 - y0}+{x0}+{y0}", "+repage",
            "-strip", "-interlace", "Plane", "-quality", "88", str(output),
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode or not output.exists():
        raise RuntimeError(f"护照资料页裁切失败: {source.name}: {result.stderr.strip()}")


def _bounded_image(source: Path, output: Path, normalize: bool = False) -> Path:
    size = source.stat().st_size
    if not normalize and MIN_IMAGE_BYTES <= size <= MAX_IMAGE_BYTES:
        return source
    if not shutil.which("magick"):
        raise RuntimeError(f"图片大小不合规且缺少 ImageMagick: {source.name} ({size} bytes)")
    for dimension, quality in ((1200, 88), (1000, 82), (900, 76), (800, 68), (700, 60), (600, 52)):
        result = subprocess.run(
            [
                "magick", str(source), "-auto-orient", "-thumbnail", f"{dimension}x{dimension}>",
                "-strip", "-interlace", "Plane", "-quality", str(quality), str(output),
            ],
            capture_output=True,
            text=True,
        )
        if result.returncode == 0 and output.exists() and MIN_IMAGE_BYTES <= output.stat().st_size <= MAX_IMAGE_BYTES:
            return output
    raise ValueError(f"图片无法压缩到 20–200 KB: {source.name}")


def _copy_row_style(sheet, source_row: int, target_row: int) -> None:
    sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height
    for column in range(1, sheet.max_column + 1):
        source = sheet.cell(source_row, column)
        target = sheet.cell(target_row, column)
        if source.has_style:
            target._style = copy(source._style)
        target.number_format = source.number_format
        target.protection = copy(source.protection)
        target.alignment = copy(source.alignment)


def _fit_image(path: Path, max_width: int, max_height: int) -> ExcelImage:
    image = ExcelImage(path)
    ratio = min(max_width / image.width, max_height / image.height)
    image.width = round(image.width * ratio)
    image.height = round(image.height * ratio)
    return image


def export_crew_info(
    template: Path,
    crew_workbook: Path,
    photo_folder: Path,
    passport_folder: Path,
    output: Path,
) -> dict[str, object]:
    source_book = load_workbook(crew_workbook, data_only=True, read_only=True)
    source_sheet = source_book["船上非旅客人员清单"]
    crew = []
    for row in source_sheet.iter_rows(min_row=3, values_only=True):
        if not row[0]:
            continue
        nationality = str(row[4] or "")
        crew.append({
            "sequence": int(row[0]),
            "name": str(row[1] or ""),
            "gender": "女" if str(row[2] or "").startswith("2") else "男",
            "nationality": nationality.split("-", 1)[1] if "-" in nationality else nationality,
            "passport": str(row[8] or ""),
        })
    source_book.close()

    photos = _numbered_files(photo_folder, {".jpg", ".jpeg", ".png"})
    passports = _numbered_files(passport_folder, {".pdf"})
    expected = {item["sequence"] for item in crew}
    if set(photos) != expected or set(passports) != expected:
        raise ValueError(
            f"人员与图片序号不一致: 人员={sorted(expected)}, 照片={sorted(photos)}, 护照={sorted(passports)}"
        )

    workbook = load_workbook(template)
    sheet = workbook["Sheet1"]
    sheet._images = []
    for row_number in range(2, 2 + len(crew)):
        if row_number != 2:
            _copy_row_style(sheet, 2, row_number)
        for column in range(1, sheet.max_column + 1):
            sheet.cell(row_number, column).value = None

    output.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="ship-crew-info-") as temporary:
        temporary_path = Path(temporary)
        embedded_sizes: list[dict[str, object]] = []
        for item in crew:
            sequence = item["sequence"]
            row_number = sequence + 1
            sheet.cell(row_number, 1, item["name"])
            sheet.cell(row_number, 2, item["gender"])
            sheet.cell(row_number, 4, item["nationality"])
            sheet.cell(row_number, 6, 2)
            sheet.cell(row_number, 7, item["passport"])

            face_output = temporary_path / f"face-{sequence:02d}.jpg"
            face = _bounded_image(photos[sequence], face_output, normalize=True)
            passport_render = temporary_path / f"passport-{sequence:02d}.jpg"
            _render_passport(passports[sequence], passport_render)
            passport_oriented = _orient_passport(passport_render, temporary_path, sequence)
            passport_cropped = temporary_path / f"passport-cropped-{sequence:02d}.jpg"
            _crop_passport(passport_oriented, passport_cropped, item["nationality"])
            passport_image = _bounded_image(
                passport_cropped, temporary_path / f"passport-bounded-{sequence:02d}.jpg"
            )

            face_excel = _fit_image(face, 150, 220)
            face_excel.anchor = f"H{row_number}"
            sheet.add_image(face_excel)
            passport_excel = _fit_image(passport_image, 150, 220)
            passport_excel.anchor = f"I{row_number}"
            sheet.add_image(passport_excel)
            embedded_sizes.extend([
                {"sequence": sequence, "kind": "人脸照片", "bytes": face.stat().st_size},
                {"sequence": sequence, "kind": "证件照片", "bytes": passport_image.stat().st_size},
            ])

        workbook.save(output)

    with zipfile.ZipFile(output) as archive:
        media = [info for info in archive.infolist() if info.filename.startswith("xl/media/")]
        invalid = [info for info in media if not MIN_IMAGE_BYTES <= info.file_size <= MAX_IMAGE_BYTES]
    if len(media) != len(crew) * 2 or invalid:
        details = ", ".join(f"{item.filename}={item.file_size}" for item in invalid)
        raise ValueError(f"Excel 嵌入图片校验失败: 数量={len(media)}, 不合规={details}")

    return {
        "crew_count": len(crew),
        "image_count": len(media),
        "image_min_bytes": min(info.file_size for info in media),
        "image_max_bytes": max(info.file_size for info in media),
        "output": str(output),
    }
