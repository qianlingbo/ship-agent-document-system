from __future__ import annotations

import re
import hashlib
import shutil
import subprocess
from difflib import SequenceMatcher
from itertools import permutations
from pathlib import Path
from tempfile import TemporaryDirectory


def _ocr_first_page(path: Path) -> str:
    if not shutil.which("pdftoppm") or not shutil.which("tesseract"):
        return ""
    with TemporaryDirectory(prefix="ship-passport-") as temporary:
        prefix = str(Path(temporary) / "page")
        render = subprocess.run(
            ["pdftoppm", "-png", "-r", "200", "-f", "1", "-l", "1", str(path), prefix],
            capture_output=True,
        )
        images = sorted(Path(temporary).glob("page-*.png"))
        if render.returncode or not images:
            return ""
        orientations = [images[0]]
        if shutil.which("magick"):
            for angle in (90, -90, 180):
                rotated = Path(temporary) / f"page-{angle}.png"
                subprocess.run(["magick", str(images[0]), "-rotate", str(angle), str(rotated)], capture_output=True)
                if rotated.exists():
                    orientations.append(rotated)
        candidates: list[Path] = []
        for index, image in enumerate(orientations):
            candidates.append(image)
            if shutil.which("magick"):
                dimensions = subprocess.run(
                    ["magick", "identify", "-format", "%w %h", str(image)], capture_output=True, text=True,
                ).stdout.split()
                if len(dimensions) == 2:
                    width, height = map(int, dimensions)
                    bands = ((0.45, 0.3), (0.45, 0.22), (0.55, 0.22), (0.65, 0.22), (0.75, 0.22))
                    for band_index, (top, band_height) in enumerate(bands):
                        band = Path(temporary) / f"page-{index}-mrz-band-{band_index}.png"
                        geometry = f"{int(width * 0.8)}x{int(height * band_height)}+{int(width * 0.1)}+{int(height * top)}"
                        subprocess.run(
                            ["magick", str(image), "-crop", geometry, "+repage", "-resize", "200%",
                             "-colorspace", "Gray", "-contrast-stretch", "2%x2%", str(band)],
                            capture_output=True,
                        )
                        if band.exists():
                            candidates.append(band)
                cropped = Path(temporary) / f"page-{index}-bottom.png"
                subprocess.run(
                    ["magick", str(image), "-gravity", "South", "-crop", "100%x50%+0+0", "+repage",
                     "-resize", "150%", "-colorspace", "Gray", "-contrast-stretch", "2%x2%", str(cropped)],
                    capture_output=True,
                )
                if cropped.exists():
                    candidates.append(cropped)
        fallback = ""
        for image in candidates:
            result = subprocess.run(
                ["tesseract", str(image), "stdout", "-l", "eng", "--psm", "6"],
                capture_output=True,
                text=True,
            )
            text = result.stdout if result.returncode == 0 else ""
            fallback = fallback or text
            lines = [re.sub(r"[^A-Z<]", "", line.upper()) for line in text.splitlines()]
            if any(re.search(r"P(?:<|S<?)[A-Z]{3}.+<<", line) for line in lines):
                return text
        return fallback


def _mrz_name(text: str, expected_name: str) -> str:
    expected_tokens = re.findall(r"[A-Z]+", expected_name.upper())
    for raw in text.splitlines():
        line = raw.upper().translate(str.maketrans({"0": "O", "1": "I", "5": "S", "8": "B"}))
        line = re.sub(r"[^A-Z<]", "", line)
        match = re.search(r"P(?:<|S<?)[A-Z]{3}(.+)", line)
        if not match or "<<" not in match.group(1) or not expected_tokens:
            continue
        observed = re.sub(r"[^A-Z]", "", match.group(1))
        length = sum(map(len, expected_tokens))
        observed = observed[:length + 4]
        options = set(permutations(expected_tokens))
        best = max(options, key=lambda option: SequenceMatcher(None, "".join(option), observed).ratio())
        score = SequenceMatcher(None, "".join(best), observed).ratio()
        if score >= 0.72:
            return " ".join(best)
    return ""


def _sequence(path: Path) -> int | None:
    match = re.match(r"\s*0*(\d{1,2})(?:\D|$)", path.name)
    return int(match.group(1)) if match else None


def apply_passport_mrz_names(
    crew_rows: list[list[str]], files: list[Path], previous_checks: list[dict[str, str]] | None = None,
    manual_overrides: dict[str, str] | None = None,
) -> list[dict[str, str]]:
    passports = []
    for path in files:
        upper = str(path).upper()
        parent = path.parent.name.upper()
        if path.suffix.lower() != ".pdf" or "PASSPORT" not in upper:
            continue
        if any(token in upper for token in ("COMPETENCY", "FLAG COC")):
            continue
        # A common parent is named "CREW PASSPORT & SEAMAN BOOK"; only reject
        # files whose immediate folder is specifically the seaman-book set.
        if "SEAMAN" in parent and "PASSPORT" not in parent:
            continue
        passports.append(path)
    cache = {item.get("source", ""): item for item in previous_checks or []}
    overrides = {re.sub(r"[^A-Z0-9]", "", key.upper()): value for key, value in (manual_overrides or {}).items()}
    checks: list[dict[str, str]] = []
    for path in sorted(passports, key=lambda item: (_sequence(item) or 999, item.name)):
        sequence = _sequence(path)
        sequence_target = crew_rows[sequence - 1] if sequence and 1 <= sequence <= len(crew_rows) else None
        passport_key = re.sub(r"[^A-Z0-9]", "", sequence_target[8].upper()) if sequence_target else ""
        if sequence_target and passport_key in overrides:
            original = sequence_target[1]
            sequence_target[1] = overrides[passport_key]
            checks.append({
                "sequence": sequence_target[0], "passport_number": sequence_target[8],
                "crew_list_name": original, "mrz_name": sequence_target[1], "source": str(path),
                "verification": "manual_mrz",
                "source_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            })
            continue
        cached = cache.get(str(path))
        if cached and sequence_target and cached.get("crew_list_name") == sequence_target[1] and cached.get("mrz_name"):
            sequence_target[1] = cached["mrz_name"]
            cached = dict(cached)
            cached.setdefault("verification", "ocr_mrz")
            cached.setdefault("source_sha256", hashlib.sha256(path.read_bytes()).hexdigest())
            checks.append(cached)
            continue
        text = _ocr_first_page(path)
        normalized = re.sub(r"[^A-Z0-9]", "", text.upper())
        target = next(
            (row for row in crew_rows if row[8] and re.sub(r"[^A-Z0-9]", "", row[8].upper()) in normalized),
            None,
        )
        if target is None and sequence and 1 <= sequence <= len(crew_rows):
            target = crew_rows[sequence - 1]
        if target is None:
            checks.append({
                "sequence": str(sequence or ""), "passport_number": "", "crew_list_name": "",
                "mrz_name": "", "source": str(path), "verification": "needs_manual_review",
                "source_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            })
            continue
        original = target[1]
        mrz_name = _mrz_name(text, original)
        if not mrz_name:
            checks.append({
                "sequence": target[0], "passport_number": target[8], "crew_list_name": original,
                "mrz_name": "", "source": str(path), "verification": "needs_manual_review",
                "source_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            })
            continue
        target[1] = mrz_name
        checks.append({
            "sequence": target[0], "passport_number": target[8], "crew_list_name": original,
            "mrz_name": mrz_name, "source": str(path),
            "verification": "ocr_mrz",
            "source_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        })
    return checks
