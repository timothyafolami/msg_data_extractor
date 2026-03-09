# MSG Photo Extractor

Extracts photo attachments from Outlook `.msg` files and renames them using the applicant's name found in the email body.

## Install

```bash
pip install .
```

This installs the `msg-photo-extractor` command.

For a simple client-machine setup, install the web app dependencies with:

```bash
pip install -r requirements.txt
```

To run the ZIP upload/download web app:

```bash
pip install '.[web]'
streamlit run streamlit_app.py
```

## Usage

```bash
msg-photo-extractor /path/to/MSG_Files -o /path/to/output
```

By default, the tool scans folders recursively and preserves subfolder structure in the output.

## Options

- `--no-recursive`: only scan the top-level input folder.
- `--flatten-output`: save all photos into a single output folder.

## Web App Workflow

1. Zip the client folder that contains the `.msg` files and any subfolders.
2. Open the Streamlit app.
3. Upload the ZIP file.
4. Download the processed ZIP, which includes extracted photos and `extraction_log.xlsx`.

## Local script usage

The repository also keeps `extract_msg_photos.py` as a wrapper script for direct use during development.
