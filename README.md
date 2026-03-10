# MSG Data Extractor

Processes Outlook `.msg` files from a local folder tree, extracts image attachments, and writes applicant name, email, and phone into `extraction_log.xlsx`.

This project is intended for local batch processing on the client machine.

## Client Setup

On macOS or Linux, run the setup script once:

```bash
./setup_local.sh
```

That script creates `.venv`, upgrades `pip`, and installs the dependencies from `requirements.txt`.

On Windows, run:

```bat
setup_local.bat
```

## Client Processing

On macOS or Linux, run the processor with an input folder and optional output folder:

```bash
./process_folder.sh /path/to/client_folder
./process_folder.sh /path/to/client_folder /path/to/output_folder
```

On Windows, run:

```bat
process_folder.bat P:\path\to\client_folder
process_folder.bat P:\path\to\client_folder P:\path\to\output_folder
```

Behavior:

- recursively scans all `.msg` files under the input folder
- preserves subfolder structure in the output folder
- extracts photos
- writes `extraction_log.xlsx` with name, email, phone, saved files, and status

If no output folder is supplied, the script creates one next to the project folder using the input folder name plus `_extracted`.

## Direct Python Usage

The underlying CLI still works directly:

```bash
python extract_msg_photos.py /path/to/client_folder --output-folder /path/to/output
```
