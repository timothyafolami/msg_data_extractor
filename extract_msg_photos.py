"""Compatibility wrapper for the installable msg_photo_extractor package."""

from pathlib import Path

from msg_data_extractor.cli import main


if __name__ == "__main__":
    script_dir = Path(__file__).resolve().parent
    raise SystemExit(
        main(
            default_msg_folder=script_dir / "MSG_Files",
            default_output_folder=script_dir / "Extracted_Photos",
        )
    )
