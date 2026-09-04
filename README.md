# navet-programtexter

Exports program texts from the Navet Expo GraphQL API to Markdown files that
can be reviewed and edited in [Obsidian](https://obsidian.md/).

Each program becomes one Markdown file in `program/`, containing its current
offer description(s) plus empty sections for a new text and comments.
Existing Markdown files are never overwritten, so editorial work is safe to
run the exporter again at any time.

This project is separate from `navet-info-display` and only handles the
text export workflow — it does not upload anything back to Expo.

## Setup

Create a virtual environment:

```
python -m venv .venv
```

Activate it (Windows):

```
.venv\Scripts\activate
```

Install the requirements:

```
pip install -r requirements.txt
```

Create a `.env` file in the project root with your API token:

```
API_TOKEN=your-token-here
```

## Usage

Run the exporter:

```
python export_program_texts.py
```

Markdown files are written to `program/`. The script logs each created or
skipped file and prints a summary when it finishes.