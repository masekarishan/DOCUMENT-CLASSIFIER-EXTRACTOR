# Travel Document Extractor

Extracts travel-accounting information from **PDFs and images** and writes a clean Excel workbook (and JSON) – with a simple web UI.

**Supported documents:** Air Ticket, Hotel Booking, Visa, Insurance, Train Ticket, Bus Ticket, Other.

## Features

- Reads text PDFs, scanned PDFs (OCR) and images (JPG, PNG, WEBP, TIF, BMP)
- Accepts a folder, a ZIP or a single file
- **Local mode** (no data leaves your PC) or **AI vision mode** (any OpenAI-compatible model)
- Web UI: upload files and see the JSON result in the browser
- Excel output with 6 columns: `Amount | Comission | TDS | CGST | SGST | Discription`
- Log file for every run, OCR debug text for troubleshooting

## Quick start (Windows)

1. Install **Python 3.10+** and **Tesseract OCR** (`C:\Program Files\Tesseract-OCR\`).
2. Double-click `setup_windows.bat` (installs the packages).
3. Put documents in the `input/` folder.
4. Double-click `run_windows.bat` → result in `output/travel_documents.xlsx`.


## Modes

| Mode | Description |
|------|-------------|
| `local` | PDF text + Tesseract OCR + rule-based parsing. Nothing is sent to an AI provider. |
| `ai` | Sends OCR text and the document image to an OpenAI-compatible vision model. Best for irregular layouts and handwriting. |
| `auto` | Uses AI if `AI_API_KEY` is set, otherwise local. |

For AI mode copy `.env.example` to `.env`:

```text
AI_API_KEY=your_key
AI_BASE_URL=https://api.openai.com/v1
AI_MODEL=gpt-4.1-mini
```

## Project structure

```text
Travel-Document-Extractor/
├── app.py                 # Flask web UI
├── travel_extractor.py    # extraction engine
├── templates/index.html   # UI page
├── requirements.txt
├── .env.example
├── setup_windows.bat / run_windows.bat / run_ui_windows.bat
├── input/                 # put documents here
├── output/                # results + extraction_log.txt
└── screenshots/           # execution images
```

## Troubleshooting

Check `output/extraction_log.txt`. Typical causes: Python not on PATH, packages not installed, Tesseract missing, empty `input/` folder, or an invalid AI key/model.

## Privacy

`--mode local` keeps everything on your computer. `--mode ai` sends document content to the configured provider – use only an approved provider. The `input/` and `output/` folders are git-ignored so personal travel data is never uploaded to GitHub.

## Author

Ishan Masekar
