"""NotebookLM PDF to PowerPoint converter with editable-text reconstruction.

Text in text-based PDFs becomes native PowerPoint text boxes. Embedded raster images
are retained as images. Scanned/image-only PDFs can use OCR if Tesseract is installed.
No PDF-to-PPTX converter can guarantee both pixel-perfect fidelity and full editability.
"""
import io
import tempfile
from pathlib import Path

import fitz  # PyMuPDF
import streamlit as st
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

APP_TITLE = "NotebookLM PDF → Editable PowerPoint"
DEFAULT_DPI = 200

st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="centered")
st.title(APP_TITLE)
st.write(
    "Convert NotebookLM-exported PDFs into PowerPoint. Choose editable reconstruction "
    "to turn extractable text into editable text boxes instead of one full-slide image."
)
st.warning(
    "Editable reconstruction preserves text and approximate placement, not pixel-perfect layout. "
    "Photos and illustrations remain images; complex vector diagrams, decorative backgrounds, "
    "and some tables may need manual rebuilding."
)

uploaded = st.file_uploader("Upload your NotebookLM PDF", type=["pdf"])
mode = st.radio(
    "Conversion mode",
    ["Editable reconstruction (recommended)", "Hybrid: editable text + full-slide reference image", "Exact visual preservation (image slides)"],
    index=0,
    help="Hybrid keeps a reference image behind the editable text; the image can obscure text while editing. Use it for visual comparison, not as the cleanest editing experience."
)
use_ocr = st.checkbox("Use OCR when a page contains no extractable text (requires Tesseract installed)", value=True)
ocr_language = st.text_input("OCR language code", value="eng", help="Examples: eng, eng+msa. Language packs must be installed in Tesseract.")
dpi = st.slider("Reference-image resolution (DPI)", min_value=100, max_value=300, value=DEFAULT_DPI, step=25, disabled=(mode == "Editable reconstruction (recommended)"))
include_report = st.checkbox("Include conversion report (.txt)", value=True)


def page_size_inches(page):
    # Preserve page aspect ratio; keep both dimensions below PowerPoint's practical limit.
    max_dim = 56.0
    w = 13.333
    h = w * page.rect.height / page.rect.width
    scale = min(1.0, max_dim / max(w, h))
    return w * scale, h * scale


def pt_to_emu(value_pt, scale):
    return Inches(value_pt / 72.0 * scale)


def color_from_int(value):
    # PyMuPDF span colors are 0xRRGGBB.
    return RGBColor((value >> 16) & 255, (value >> 8) & 255, value & 255)


def add_text_span(slide, span, page_rect, slide_w_in, slide_h_in):
    text = span.get("text", "")
    if not text or not text.strip():
        return
    x0, y0, x1, y1 = span.get("bbox", (0, 0, 10, 10))
    sx = slide_w_in / page_rect.width
    sy = slide_h_in / page_rect.height
    left = max(0, x0 * sx)
    top = max(0, y0 * sy)
    width = max(0.03, min(slide_w_in - left, (x1 - x0) * sx + 0.015))
    height = max(0.08, min(slide_h_in - top, (y1 - y0) * sy + 0.025))
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = False
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    font = run.font
    # Use a safe common font if the PDF's embedded font isn't installed on the user's PC.
    pdf_font = span.get("font", "")
    font.name = "Arial" if not pdf_font else pdf_font.split(",")[0].replace("+", "")
    font.size = Pt(max(4, min(96, float(span.get("size", 12)) * sy * 72)))
    flags = int(span.get("flags", 0))
    font.italic = bool(flags & 2)
    font.bold = bool(flags & 16) or "bold" in pdf_font.lower()
    try:
        font.color.rgb = color_from_int(int(span.get("color", 0)))
    except Exception:
        font.color.rgb = RGBColor(0, 0, 0)


def extract_spans(page, ocr_enabled, language):
    data = page.get_text("dict")
    spans = []
    for block in data.get("blocks", []):
        if block.get("type") == 0:
            for line in block.get("lines", []):
                spans.extend(line.get("spans", []))
    source = "embedded text"
    if not any(s.get("text", "").strip() for s in spans) and ocr_enabled:
        try:
            textpage = page.get_textpage_ocr(language=language or "eng", dpi=200, full=True)
            ocr_data = page.get_text("dict", textpage=textpage)
            spans = []
            for block in ocr_data.get("blocks", []):
                if block.get("type") == 0:
                    for line in block.get("lines", []):
                        spans.extend(line.get("spans", []))
            source = "OCR" if any(s.get("text", "").strip() for s in spans) else "no text detected"
        except Exception as exc:
            source = f"OCR unavailable: {exc}"
    elif not any(s.get("text", "").strip() for s in spans):
        source = "no extractable text"
    return spans, source, data


def add_embedded_images(slide, page, page_rect, slide_w_in, slide_h_in):
    """Place embedded raster image blocks at approximate original positions."""
    count = 0
    data = page.get_text("dict")
    for block in data.get("blocks", []):
        if block.get("type") != 1 or not block.get("image"):
            continue
        bbox = block.get("bbox")
        if not bbox:
            continue
        x0, y0, x1, y1 = bbox
        # Do not reinsert a full-page raster: that would cover/defeat editable OCR text.
        area_ratio = max(0, x1 - x0) * max(0, y1 - y0) / max(1, page_rect.width * page_rect.height)
        if area_ratio >= 0.70:
            continue
        left = max(0, x0 / page_rect.width * slide_w_in)
        top = max(0, y0 / page_rect.height * slide_h_in)
        width = min(slide_w_in - left, max(0.01, (x1 - x0) / page_rect.width * slide_w_in))
        height = min(slide_h_in - top, max(0.01, (y1 - y0) / page_rect.height * slide_h_in))
        try:
            slide.shapes.add_picture(io.BytesIO(block["image"]), Inches(left), Inches(top), Inches(width), Inches(height))
            count += 1
        except Exception:
            pass
    return count

if uploaded:
    pdf_bytes = uploaded.getvalue()
    if not pdf_bytes.startswith(b"%PDF"):
        st.error("This file does not appear to be a valid PDF.")
        st.stop()
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as pdf:
            page_count = len(pdf)
            if not page_count:
                st.error("The PDF has no pages.")
                st.stop()
            st.write(f"**File:** {uploaded.name}  \n**Pages detected:** {page_count:,}")
            if st.button("Convert PDF to PowerPoint", type="primary", use_container_width=True):
                prs = Presentation()
                first = pdf[0]
                slide_w, slide_h = page_size_inches(first)
                prs.slide_width, prs.slide_height = Inches(slide_w), Inches(slide_h)
                blank = prs.slide_layouts[6]
                report = [
                    f"Source PDF: {uploaded.name}", f"Page count: {page_count}", f"Mode: {mode}",
                    "Text is reconstructed as editable PowerPoint text boxes where extraction/OCR succeeds.", ""
                ]
                progress = st.progress(0, text="Preparing slides…")
                status = st.empty()
                with tempfile.TemporaryDirectory(prefix="notebooklm_ppt_") as tmp:
                    out_path = Path(tmp) / "notebooklm_converted.pptx"
                    for i, page in enumerate(pdf):
                        status.write(f"Processing page {i + 1:,} of {page_count:,}…")
                        page_w, page_h = page.rect.width, page.rect.height
                        sw, sh = page_size_inches(page)
                        # If a PDF has mixed page sizes, this version keeps the deck size from page 1
                        # and scales content relative to each page's own dimensions.
                        slide = prs.slides.add_slide(blank)
                        spans, source, page_dict = extract_spans(page, use_ocr, ocr_language)
                        # Many NotebookLM PDFs store each slide as a single full-page image.
                        # If no usable text was extracted, render the page as a fallback so the
                        # resulting slide is never silently blank. Such fallback slides are visual
                        # images, not editable text.
                        has_text = any(s.get("text", "").strip() for s in spans)
                        full_page_image = any(
                            b.get("type") == 1 and b.get("bbox") and
                            ((b["bbox"][2] - b["bbox"][0]) * (b["bbox"][3] - b["bbox"][1]) /
                             max(1, page.rect.width * page.rect.height)) >= 0.70
                            for b in page_dict.get("blocks", [])
                        )
                        visual_fallback = (mode == "Editable reconstruction (recommended)" and
                                           (not has_text or full_page_image and source != "embedded text"))
                        if (mode == "Exact visual preservation (image slides)" or
                            mode.startswith("Hybrid:") or visual_fallback):
                            pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
                            slide.shapes.add_picture(io.BytesIO(pix.tobytes("png")), 0, 0, Inches(slide_w), Inches(slide_h))
                            if visual_fallback:
                                source = source + " — full-page image fallback (visual content preserved; not editable)"
                        editable_count = 0
                        if mode != "Exact visual preservation (image slides)" and not visual_fallback:
                            # Add actual embedded illustrations before text so text stays selectable.
                            # Large page-sized rasters are skipped by add_embedded_images.
                            if not mode.startswith("Hybrid:"):
                                image_count = add_embedded_images(slide, page, page.rect, slide_w, slide_h)
                            else:
                                image_count = 0
                            for span in spans:
                                if span.get("text", "").strip():
                                    add_text_span(slide, span, page.rect, slide_w, slide_h)
                                    editable_count += 1
                        else:
                            image_count = 1
                        report.append(f"Page {i+1}: text spans={editable_count}; image objects={image_count}; text source={source}")
                        progress.progress((i + 1) / page_count, text=f"Converted {i+1:,}/{page_count:,} pages")
                    status.write("Packaging PowerPoint…")
                    prs.core_properties.title = Path(uploaded.name).stem + " — Editable PDF conversion"
                    prs.core_properties.subject = "PDF to PowerPoint with editable text reconstruction"
                    prs.save(out_path)
                    pptx_bytes = out_path.read_bytes()
                st.session_state["converted_pptx"] = pptx_bytes
                st.session_state["converted_name"] = Path(uploaded.name).stem + ("_editable.pptx" if mode.startswith("Editable") else "_converted.pptx")
                st.session_state["converted_report"] = "\n".join(report)
                st.success(f"Conversion complete: {page_count:,} slides created.")
                st.info("Open PowerPoint and select text to verify editability. If a page is image-only, install Tesseract and enable OCR; OCR quality depends on image resolution and language packs.")
    except Exception as exc:
        st.error(f"Conversion failed: {exc}")
        st.exception(exc)

if "converted_pptx" in st.session_state:
    st.download_button(
        "Download PowerPoint (.pptx)", data=st.session_state["converted_pptx"],
        file_name=st.session_state["converted_name"],
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary", use_container_width=True
    )
    if include_report and "converted_report" in st.session_state:
        st.download_button("Download conversion report (.txt)", data=st.session_state["converted_report"].encode("utf-8"), file_name="conversion_report.txt", mime="text/plain", use_container_width=True)

st.divider()
st.caption("Conversion runs on the computer/server where you launch the app. For OCR fallback, install Tesseract separately; see README for Windows steps. Large PDFs require adequate RAM and temporary disk space.")
