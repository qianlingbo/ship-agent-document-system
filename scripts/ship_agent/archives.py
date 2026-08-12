from __future__ import annotations

import os
import shutil
import subprocess
import tarfile
import zipfile
from pathlib import Path


class ArchiveError(RuntimeError):
    pass


def _safe_target(root: Path, member: str) -> Path:
    destination = (root / member).resolve()
    if destination != root.resolve() and root.resolve() not in destination.parents:
        raise ArchiveError(f"拒绝解压越界路径: {member}")
    return destination


def extract_archive(path: str | Path, destination: str | Path) -> list[Path]:
    """Extract ZIP/TAR or an external-tool RAR into destination safely."""
    source = Path(path)
    target = Path(destination)
    target.mkdir(parents=True, exist_ok=True)
    suffix = source.suffix.lower()
    extracted: list[Path] = []
    if suffix == ".zip":
        with zipfile.ZipFile(source) as archive:
            for info in archive.infolist():
                _safe_target(target, info.filename)
            archive.extractall(target)
            extracted = [target / i.filename for i in archive.infolist() if not i.is_dir()]
        return extracted
    if suffix in {".tar", ".gz", ".tgz", ".bz2", ".xz"} or source.name.lower().endswith(".tar.gz"):
        with tarfile.open(source) as archive:
            for info in archive.getmembers():
                _safe_target(target, info.name)
                if info.issym() or info.islnk() or info.isdev():
                    raise ArchiveError(f"拒绝解压链接或设备文件: {info.name}")
            archive.extractall(target)
            extracted = [target / i.name for i in archive.getmembers() if i.isfile()]
        return extracted
    if suffix == ".rar":
        tool = next((shutil.which(x) for x in ("7z", "7zz", "unar", "bsdtar") if shutil.which(x)), None)
        if not tool:
            raise ArchiveError("检测到 RAR，但本机没有 7z/7zz/unar/bsdtar；请安装解压工具后重试")
        if os.path.basename(tool) in {"7z", "7zz"}:
            command = [tool, "x", "-y", f"-o{target}", str(source)]
        elif os.path.basename(tool) == "unar":
            command = [tool, "-f", "-o", str(target), str(source)]
        else:
            command = [tool, "-xf", str(source), "-C", str(target)]
        result = subprocess.run(command, capture_output=True, text=True)
        if result.returncode:
            raise ArchiveError(result.stderr.strip() or result.stdout.strip() or "RAR 解压失败")
        return [p for p in target.rglob("*") if p.is_file()]
    raise ArchiveError(f"不支持的压缩格式: {source.name}")
