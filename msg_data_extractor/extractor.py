import logging
import re
from pathlib import Path

from .ole2 import OLE2Reader
from .reporting import write_excel_log

log = logging.getLogger(__name__)

PROP_BODY_PLAIN = "__substg1.0_1000001F"
PROP_ATTACH_DATA = "__substg1.0_37010102"
PROP_ATTACH_LONG = "__substg1.0_3707001F"

IMAGE_MAGIC = {
    b"\xff\xd8\xff": ".jpg",
    b"\x89PNG": ".png",
    b"GIF8": ".gif",
    b"BM": ".bmp",
    b"RIFF": ".webp",
}


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def extract_name(body: str) -> str | None:
    body = body.replace("\r\n", "\n").replace("\r", "\n")

    match = re.search(
        r"(?i)(?<![A-Za-z])(?:Full\s)?Name[\t ]*[:\t][\t ]*([\w][^\t\r\n]*?)(?:\t|\r|\n|$)",
        body,
    )
    if match:
        prefix = body[max(0, match.start() - 10) : match.start()].lower()
        if not re.search(r"(first|last)", prefix):
            candidate = _clean(match.group(1))
            if candidate and not re.search(r"[@://]", candidate):
                return candidate

    first_match = re.search(r"(?i)First\s*Name[\t ]*[:\t][\t ]*([\w .'\-]+)", body)
    last_match = re.search(r"(?i)Last\s*Name[\t ]*[:\t][\t ]*([\w .'\-]+)", body)
    if first_match and last_match:
        first = _clean(first_match.group(1))
        last = _clean(last_match.group(1))
        if first and last:
            return f"{first} {last}"

    match = re.search(r"^([\w][\w .'\-]+ [\w .'\-]+)\s+just submitted", body, re.MULTILINE)
    if match:
        return _clean(match.group(1))

    match = re.search(r"(?i)\bName\s{2,}([\w][\w .'\-]+(?:\s+[\w .'\-]+)+)", body)
    if match:
        return _clean(match.group(1))

    return None


def _clean_contact_value(value: str) -> str:
    return value.strip().strip("<>").rstrip(",;.")


def extract_email(body: str) -> str | None:
    field_match = re.search(
        r"(?im)^Email[\t ]*[:\t ]+[\t ]*([^\s<]+@[A-Z0-9.-]+\.[A-Z]{2,})",
        body,
    )
    if field_match:
        return _clean_contact_value(field_match.group(1))

    generic_match = re.search(
        r"\b([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})\b",
        body,
        re.IGNORECASE,
    )
    if generic_match:
        return _clean_contact_value(generic_match.group(1))
    return None


def extract_phone(body: str) -> str | None:
    field_patterns = [
        r"(?im)^Mobile[\t ]*[:\t ]+[\t ]*([^\r\n]+)",
        r"(?im)^Phone(?:\s*Number)?[\t ]*[:\t ]+[\t ]*([^\r\n]+)",
    ]
    for pattern in field_patterns:
        field_match = re.search(pattern, body)
        if field_match:
            candidate = _clean_contact_value(field_match.group(1))
            if re.search(r"\d", candidate):
                return candidate

    generic_match = re.search(
        r"(\+?\d[\d()\- ]{7,}\d)",
        body,
    )
    if generic_match:
        return _clean_contact_value(generic_match.group(1))
    return None


def image_extension(data: bytes) -> str | None:
    for magic, ext in IMAGE_MAGIC.items():
        if data[: len(magic)] == magic:
            return ext
    return None


def filename_extension(filename: str | None) -> str:
    if filename:
        ext = Path(filename).suffix.lower()
        if ext in {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff"}:
            return ".jpg" if ext == ".jpeg" else ext
    return ""


def safe_name(full_name: str) -> str:
    name = re.sub(r"[^\w\s'\-]", "", full_name).strip()
    return re.sub(r"\s+", "_", name)


def unique_path(folder: Path, base: str, ext: str) -> Path:
    candidate = folder / f"{base}{ext}"
    counter = 2
    while candidate.exists():
        candidate = folder / f"{base}_{counter}{ext}"
        counter += 1
    return candidate


def _is_supported_msg_path(path: Path, root_dir: Path) -> bool:
    if not path.is_file() or path.suffix.lower() != ".msg":
        return False
    relative_parts = path.relative_to(root_dir).parts
    if "__MACOSX" in relative_parts:
        return False
    if path.name.startswith("._"):
        return False
    return True


def discover_msg_files(msg_dir: Path, recursive: bool) -> list[Path]:
    pattern = "**/*.msg" if recursive else "*.msg"
    return sorted(path for path in msg_dir.glob(pattern) if _is_supported_msg_path(path, msg_dir))


def process_msg(msg_path: Path, output_folder: Path) -> dict:
    result = {
        "msg_file": msg_path.name,
        "applicant_name": "",
        "email": "",
        "phone": "",
        "saved_files": [],
        "saved_paths": [],
        "error": "",
        "source_folder": "",
    }

    try:
        ole = OLE2Reader(str(msg_path))
    except Exception as exc:
        result["error"] = f"OLE2 parse error: {exc}"
        return result

    body_raw = ole.get_bytes(0, PROP_BODY_PLAIN)
    body = body_raw.decode("utf-16-le", errors="ignore") if body_raw else ""

    name = extract_name(body)
    if not name:
        result["error"] = "Name not found in body"
        name = msg_path.stem

    result["applicant_name"] = name
    result["email"] = extract_email(body) or ""
    result["phone"] = extract_phone(body) or ""
    safe = safe_name(name)

    image_attachments: list[tuple[bytes, str]] = []

    for att_idx in ole.attachment_storages():
        data = ole.get_bytes(att_idx, PROP_ATTACH_DATA)
        if not data:
            continue

        ext = image_extension(data)
        if ext is None:
            long_name = ole.get_string(att_idx, PROP_ATTACH_LONG)
            ext = filename_extension(long_name)

        if not ext:
            continue

        image_attachments.append((data, ext))

    total_images = len(image_attachments)

    for image_count, (data, ext) in enumerate(image_attachments, 1):
        base = safe if total_images == 1 else f"{safe}_{image_count}"
        out_path = unique_path(output_folder, base, ext)
        out_path.write_bytes(data)
        result["saved_files"].append(out_path.name)
        result["saved_paths"].append(str(out_path))

    if total_images == 0 and not result["error"]:
        result["error"] = "No image attachments found"

    return result


def run(
    msg_folder: str | Path,
    output_folder: str | Path,
    recursive: bool = True,
    preserve_structure: bool = True,
) -> dict:
    msg_dir = Path(msg_folder).expanduser().resolve()
    out_dir = Path(output_folder).expanduser().resolve()

    if not msg_dir.exists():
        raise FileNotFoundError(f"MSG folder does not exist: {msg_dir}")
    if not msg_dir.is_dir():
        raise NotADirectoryError(f"MSG folder is not a directory: {msg_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)

    msg_files = discover_msg_files(msg_dir, recursive=recursive)
    total = len(msg_files)
    if total == 0:
        log.warning("No .msg files found in: %s", msg_dir)
        return {"results": [], "log_path": None, "msg_dir": msg_dir, "out_dir": out_dir}

    log.info("Found %s .msg files in %s", total, msg_dir)
    log.info("Output folder: %s", out_dir)

    results = []
    for index, msg_path in enumerate(msg_files, 1):
        rel_parent = msg_path.parent.relative_to(msg_dir)
        source_folder = "." if str(rel_parent) == "." else rel_parent.as_posix()
        target_dir = out_dir / rel_parent if preserve_structure and str(rel_parent) != "." else out_dir
        target_dir.mkdir(parents=True, exist_ok=True)

        log.info("[%s/%s]  %s", index, total, msg_path.relative_to(msg_dir).as_posix())
        result = process_msg(msg_path, target_dir)
        result["source_folder"] = source_folder
        result["saved_paths"] = [Path(path).relative_to(out_dir).as_posix() for path in result["saved_paths"]]
        results.append(result)

        if result["saved_files"]:
            log.info("  Saved for %s -> %s", result["applicant_name"], ", ".join(result["saved_paths"]))
        else:
            log.warning("  Skipped: %s", result["error"])

    log_path = write_excel_log(results, out_dir)
    ok_count = sum(1 for result in results if result["saved_files"])

    log.info("%s", "=" * 60)
    log.info("Done. %s/%s files had photos extracted.", ok_count, total)
    log.info("Log saved: %s", log_path)

    return {
        "results": results,
        "log_path": log_path,
        "msg_dir": msg_dir,
        "out_dir": out_dir,
        "ok_count": ok_count,
        "total": total,
    }
