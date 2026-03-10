import argparse
import logging
from pathlib import Path

from .extractor import run


def build_parser(default_msg_folder: Path | None = None, default_output_folder: Path | None = None) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Extract photo attachments from Outlook .msg files and rename them by applicant name."
    )
    parser.add_argument(
        "msg_folder",
        nargs="?",
        default=str(default_msg_folder or (Path.cwd() / "MSG_Files")),
        help="Folder containing .msg files. Defaults to ./MSG_Files in the current directory.",
    )
    parser.add_argument(
        "-o",
        "--output-folder",
        default=str(default_output_folder or (Path.cwd() / "Extracted_Photos")),
        help="Folder where extracted photos and the Excel log will be written.",
    )
    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="Only process .msg files directly inside the input folder.",
    )
    parser.add_argument(
        "--flatten-output",
        action="store_true",
        help="Write all extracted photos into one output folder instead of preserving subfolders.",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        help="Number of worker threads to use. Defaults to a CPU-based value.",
    )
    return parser


def configure_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )


def main(
    argv: list[str] | None = None,
    default_msg_folder: Path | None = None,
    default_output_folder: Path | None = None,
) -> int:
    configure_logging()
    parser = build_parser(default_msg_folder, default_output_folder)
    args = parser.parse_args(argv)

    try:
        run(
            msg_folder=args.msg_folder,
            output_folder=args.output_folder,
            recursive=not args.no_recursive,
            preserve_structure=not args.flatten_output,
            workers=args.workers,
        )
    except (FileNotFoundError, NotADirectoryError) as exc:
        parser.error(str(exc))
    except Exception:
        logging.getLogger(__name__).exception("Unhandled extraction error")
        return 1
    return 0
