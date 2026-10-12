import io
import os
import tempfile
from pathlib import Path

import fitz  # PyMuPDF
import streamlit as st
from pptx import Presentation
from pptx.util import Inches

APP_TITLE = "NotebookLM PDF → PowerPoint"
DEFAULT_DPI = 220

st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="centered")
st.title(APP_TITLE)
st.write(
    "Convert a PDF exported from Google NotebookLM into a PowerPoint deck. "
    "Each PDF page becomes a full-slide image, prioritizing visual fidelity."
)
st.info(
    "This is an exact-appearance converter, not an editable-slide converter. "
    "Text and diagrams remain part of each slide image. Higher DPI improves raster detail "
    "but increases PPTX size."
)

uploaded = st.file_uploader("Upload your NotebookLM PDF", type=["pdf"])
dpi = st.slider("Rendering resolution (DPI)", min_value=120, max_value=300, value=DEFAULT_DPI, step=20)
include_report = st.checkbox("Include a conversion report (.txt) in the download", value=True)

def slide_size_for_page(width_pt, height_pt):
    # Keep source aspect ratio and stay within PowerPoint's 56-inch dimension limit.
    max_dim = 56.0
    width_in = 13.333
    height_in = width_in * height_pt / width_pt
    scale = min(1.0, max_dim / max(width_in, height_in))
    return width_in * scale, height_in * scale

def fit_image(page_w_pt, page_h_pt, slide_w_in, slide_h_in):
    page_ratio = page_w_pt / page_h_pt
    slide_ratio = slide_w_in / slide_h_in
    if page_ratio >= slide_ratio:
        w = slide_w_in
        h = slide_w_in / page_ratio
    else:
        h = slide_h_in
        w = slide_h_in * page_ratio
    left = (slide_w_in - w) / 2
    top = (slide_h_in - h) / 2
    return Inches(left), Inches(top), Inches(w), Inches(h)

if uploaded:
    st.write(f"**File:** {uploaded.name}")
    pdf_bytes = uploaded.getvalue()

    if not pdf_bytes.startswith(b"%PDF"):
        st.error("This file does not appear to be a valid PDF.")
        st.stop()

    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as pdf:
            page_count = len(pdf)
            if page_count == 0:
                st.error("The PDF has no pages.")
                st.stop()

            st.write(f"**Pages detected:** {page_count:,}")
            st.write("Choose Convert to create a downloadable `.pptx` file.")

            if st.button("Convert PDF to PowerPoint", type="primary", use_container_width=True):
                progress = st.progress(0, text="Preparing presentation…")
                status = st.empty()
                prs = Presentation()
                first = pdf[0]
                slide_w, slide_h = slide_size_for_page(first.rect.width, first.rect.height)
                prs.slide_width = Inches(slide_w)
                prs.slide_height = Inches(slide_h)
                blank_layout = prs.slide_layouts[6]

                report_lines = [
                    f"Source PDF: {uploaded.name}",
                    f"Page count: {page_count}",
                    f"Render DPI: {dpi}",
                    "Conversion mode: full-page images for visual preservation",
                    f"PowerPoint slide size: {slide_w:.4f} x {slide_h:.4f} inches",
                    "",
                    "Per-page details:"
                ]

                with tempfile.TemporaryDirectory(prefix="notebooklm_ppt_") as temp_dir:
                    out_path = Path(temp_dir) / "notebooklm_converted.pptx"
                    for i, page in enumerate(pdf):
                        status.write(f"Rendering page {i + 1:,} of {page_count:,}…")
                        pix = page.get_pixmap(
                            matrix=fitz.Matrix(dpi / 72.0, dpi / 72.0),
                            alpha=False
                        )
                        image_bytes = pix.tobytes("png")
                        slide = prs.slides.add_slide(blank_layout)
                        left, top, width, height = fit_image(
                            page.rect.width, page.rect.height, slide_w, slide_h
                        )
                        slide.shapes.add_picture(io.BytesIO(image_bytes), left, top, width, height)
                        report_lines.append(
                            f"Page {i + 1}: PDF size {page.rect.width:.2f} x {page.rect.height:.2f} pt; "
                            f"render {pix.width} x {pix.height} px"
                        )
                        progress.progress((i + 1) / page_count, text=f"Converted {i + 1:,}/{page_count:,} pages")

                    status.write("Packaging PowerPoint…")
                    prs.core_properties.title = Path(uploaded.name).stem + " — NotebookLM PDF conversion"
                    prs.core_properties.subject = "Full-page visual conversion from PDF"
                    prs.save(out_path)
                    pptx_bytes = out_path.read_bytes()

                st.session_state["converted_pptx"] = pptx_bytes
                st.session_state["converted_name"] = Path(uploaded.name).stem + ".pptx"
                st.session_state["converted_report"] = "\n".join(report_lines)
                st.success(f"Conversion complete: {page_count:,} slides created.")
                st.caption("The output contains one slide for every PDF page. Review a few slides before teaching or distributing.")

    except Exception as exc:
        st.error(f"Conversion failed: {exc}")
        st.exception(exc)

if "converted_pptx" in st.session_state:
    st.download_button(
        "Download PowerPoint (.pptx)",
        data=st.session_state["converted_pptx"],
        file_name=st.session_state["converted_name"],
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True
    )
    if include_report and "converted_report" in st.session_state:
        st.download_button(
            "Download conversion report (.txt)",
            data=st.session_state["converted_report"].encode("utf-8"),
            file_name="conversion_report.txt",
            mime="text/plain",
            use_container_width=True
        )

st.divider()
st.caption(
    "Privacy: conversion runs on the computer/server where you launch this app. "
    "If deployed to a cloud host, the uploaded PDF is processed by that host. "
    "Large PDFs require sufficient RAM, temporary disk space, and browser upload limits."
)
