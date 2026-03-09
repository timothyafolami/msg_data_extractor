from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


class ZipValidationError(ValueError):
    pass


def _should_skip_member(member_name: str) -> bool:
    parts = Path(member_name).parts
    name = Path(member_name).name
    return "__MACOSX" in parts or name.startswith("._") or name == ".DS_Store"


def _safe_member_path(target_dir: Path, member_name: str) -> Path:
    member_path = Path(member_name)
    if member_path.is_absolute():
        raise ZipValidationError(f"ZIP contains an absolute path: {member_name}")

    resolved_target_dir = target_dir.resolve()
    resolved = (resolved_target_dir / member_path).resolve()
    try:
        resolved.relative_to(resolved_target_dir)
    except ValueError:
        raise ZipValidationError(f"ZIP contains an unsafe path: {member_name}")
    return resolved


def extract_zip_bytes(zip_bytes: bytes, target_dir: Path) -> list[Path]:
    target_dir.mkdir(parents=True, exist_ok=True)
    extracted_files: list[Path] = []

    with ZipFile(BytesIO(zip_bytes)) as archive:
        for info in archive.infolist():
            if _should_skip_member(info.filename):
                continue
            output_path = _safe_member_path(target_dir, info.filename)
            if info.is_dir():
                output_path.mkdir(parents=True, exist_ok=True)
                continue

            output_path.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(info) as src, open(output_path, "wb") as dst:
                dst.write(src.read())
            extracted_files.append(output_path)

    return extracted_files


def create_zip_bytes(source_dir: Path) -> bytes:
    buffer = BytesIO()
    with ZipFile(buffer, mode="w", compression=ZIP_DEFLATED) as archive:
        for path in sorted(source_dir.rglob("*")):
            if path.is_file():
                archive.write(path, arcname=path.relative_to(source_dir).as_posix())
    return buffer.getvalue()
