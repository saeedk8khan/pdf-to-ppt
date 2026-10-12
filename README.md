# NotebookLM PDF → PowerPoint Converter

A local Streamlit app that converts PDFs exported from Google NotebookLM into PPTX files, using one full-page image per slide for visual preservation.

## What it preserves

- One PowerPoint slide per PDF page
- Page appearance, including text, images, charts, tables, and diagrams as rendered by the PDF
- Source aspect ratio (mixed page sizes are centered within the presentation canvas)
- No artificial 15-slide cap in the converter; 100+ page PDFs are supported, subject to available memory/disk and the PPTX file size

## Important limitation

The output is image-based: text, charts, and diagrams are not individually editable in PowerPoint. Exact visual fidelity and full editability are different goals. The PDF itself is the reference, and results can vary slightly due to PDF rendering and PowerPoint display.

## Windows installation

1. Install Python 3.10 or later from https://www.python.org/downloads/
   - During installation, enable **Add Python to PATH**.
2. Extract this ZIP file.
3. Open the extracted folder in File Explorer.
4. Double-click `run_windows.bat`.
5. Your browser should open the app. If it does not, open the local URL printed in the terminal (usually http://localhost:8501).
6. Upload the PDF downloaded from NotebookLM, choose DPI, and click **Convert PDF to PowerPoint**.
7. Click **Download PowerPoint (.pptx)**.

## Manual installation (Windows, macOS, Linux)

Open a terminal in this folder and run:

```bash
python -m venv .venv
```

Activate it:

Windows PowerShell:
```powershell
.\.venv\Scripts\Activate.ps1
```

macOS/Linux:
```bash
source .venv/bin/activate
```

Install packages and start:
```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
streamlit run app.py
```

## Recommended settings

- 180–220 DPI: good default for slide decks
- 240–300 DPI: sharper text/diagrams, larger output files and higher resource use
- For 100+ pages, start at 180 or 200 DPI if your computer has limited RAM or disk space

## Troubleshooting

- **The app does not start:** confirm Python is installed, then run the manual installation commands to see the error.
- **Very large PPTX:** lower the DPI; PNG images can make large files.
- **Upload fails:** increase the hosting platform's upload limit or run locally.
- **PowerPoint opens with a repair warning:** update `python-pptx`, rerun conversion, and test the output in desktop PowerPoint.
- **PDF is password-protected:** remove the password using a legitimate PDF viewer before uploading.
- **Slides have white margins:** the converter preserves page aspect ratio rather than stretching the PDF page. If the PDF has a different aspect ratio on a page, it is centered with margins instead of distorting it.

## Technical basis

The app uses PyMuPDF to render each PDF page to PNG and python-pptx to place each image on a blank slide. See:
- PyMuPDF: https://pymupdf.readthedocs.io/
- python-pptx: https://python-pptx.readthedocs.io/
- Streamlit: https://docs.streamlit.io/
