"""NotebookLM PDF -> editable PowerPoint (high-fidelity reconstruction).

NotebookLM PDFs are flat pictures. For every page this app:
  1. reads the text with RapidOCR (far more accurate than Tesseract on slide graphics),
  2. measures the real ink of each line (position, height, width, colour),
  3. works out font family (sans / serif / condensed) and bold by rendering candidates
     and comparing them with the picture,
  4. erases the original text from the picture (inpainting),
  5. places native PowerPoint text boxes on top.
Charts, icons and boxes stay in the background picture (they are not text).
"""
import io
import math
import re
from pathlib import Path

import numpy as np
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Inches, Pt

try:
    import pymupdf as fitz
except ImportError:  # older PyMuPDF
    import fitz

try:
    import cv2
except ImportError:
    cv2 = None
try:
    from rapidocr_onnxruntime import RapidOCR
except ImportError:
    RapidOCR = None
try:
    import wordninja
    _WN_VOCAB = wordninja.DEFAULT_LANGUAGE_MODEL._wordcost
except Exception:
    wordninja, _WN_VOCAB = None, {}

APP_TITLE = "NotebookLM PDF → Editable PowerPoint"

# ------------------------------------------------------------------ fonts
# family -> (PowerPoint font name, regular candidates, bold candidates)
FAMILIES = {
    "sans": ("Arial",
             ["C:/Windows/Fonts/arial.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
              "/Library/Fonts/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf"],
             ["C:/Windows/Fonts/arialbd.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
              "/Library/Fonts/Arial Bold.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]),
    "serif": ("Georgia",
              ["C:/Windows/Fonts/georgia.ttf", "C:/Windows/Fonts/times.ttf",
               "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
               "/Library/Fonts/Georgia.ttf", "/System/Library/Fonts/Supplemental/Georgia.ttf"],
              ["C:/Windows/Fonts/georgiab.ttf", "C:/Windows/Fonts/timesbd.ttf",
               "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf",
               "/Library/Fonts/Georgia Bold.ttf", "/System/Library/Fonts/Supplemental/Georgia Bold.ttf"]),
    "narrow": ("Arial Narrow",
               ["C:/Windows/Fonts/ARIALN.TTF", "/usr/share/fonts/truetype/liberation/LiberationSansNarrow-Regular.ttf",
                "/System/Library/Fonts/Supplemental/Arial Narrow.ttf"],
               ["C:/Windows/Fonts/ARIALNB.TTF", "/usr/share/fonts/truetype/liberation/LiberationSansNarrow-Bold.ttf",
                "/System/Library/Fonts/Supplemental/Arial Narrow Bold.ttf"]),
}
_font_cache = {}


def get_font(fam, bold, size):
    key = (fam, bold, size)
    if key not in _font_cache:
        f = None
        for p in FAMILIES[fam][2 if bold else 1]:
            if Path(p).is_file():
                try:
                    f = ImageFont.truetype(p, size)
                    break
                except Exception:
                    pass
        _font_cache[key] = f
    return _font_cache[key]


def available_families():
    return [f for f in FAMILIES if get_font(f, False, 20) and get_font(f, True, 20)]


# ------------------------------------------------------------- text fixing
_BULLETS = "\u00b7\u2022\u25cf\u25aa\u25e6\u2219\u2027\u25cb\u25a0"


def fix_text(t):
    t = t.strip()
    t = re.sub("^[" + _BULLETS + "\ue000-\uf8ff*]+\\s*", "\u2022 ", t)
    t = re.sub(r"(?<=[:;,])(?=[A-Za-z])", " ", t)  # "Rule:All" -> "Rule: All"
    if wordninja:  # "ofincome" -> "of income"
        def split(m):
            w = m.group(0)
            if len(w) < 8 or w.lower() in _WN_VOCAB or sum(c.isupper() for c in w) >= 2:
                return w
            parts = wordninja.split(w)
            if 2 <= len(parts) <= 4 and all(len(p) >= 2 and p.lower() in _WN_VOCAB for p in parts):
                out, pos = [], 0
                for p in parts:
                    out.append(w[pos:pos + len(p)])
                    pos += len(p)
                return " ".join(out)
            return w
        t = re.sub(r"[A-Za-z]{8,}", split, t)
    return re.sub(r"\s{2,}", " ", t)


# ------------------------------------------------------------ OCR / lines
_OCR = None


def get_ocr():
    global _OCR
    if _OCR is None:
        _OCR = RapidOCR()
    return _OCR


def ocr_lines(arr):
    res, _ = get_ocr()(arr)
    lines = []
    for box, text, conf in (res or []):
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        if not text.strip():
            continue
        lines.append({"box": (min(xs), min(ys), max(xs), max(ys)), "text": fix_text(text), "conf": float(conf)})
    return lines


def embedded_lines(page, k):
    out = []
    for b in page.get_text("dict").get("blocks", []):
        if b.get("type") != 0:
            continue
        for ln in b.get("lines", []):
            text = "".join(s["text"] for s in ln["spans"]).strip()
            if text:
                x0, y0, x1, y1 = ln["bbox"]
                out.append({"box": (x0 * k, y0 * k, x1 * k, y1 * k), "text": text, "conf": 1.0})
    return out


# ------------------------------------------------------------ ink analysis
def ring_pixels(arr, box, pad=4):
    x0, y0, x1, y1 = box
    h, w = arr.shape[:2]
    e = arr[max(0, y0 - pad):min(h, y1 + pad), max(0, x0 - pad):min(w, x1 + pad)]
    if e.size == 0:
        return np.zeros((1, 3), dtype=int)
    t = 3
    return np.concatenate([e[:t].reshape(-1, 3), e[-t:].reshape(-1, 3),
                           e[:, :t].reshape(-1, 3), e[:, -t:].reshape(-1, 3)]).astype(int)


def analyze_line(arr, box):
    """Tight ink box, text colour and ink mask of one OCR line."""
    h_img, w_img = arr.shape[:2]
    x0, y0, x1, y1 = box
    pad = max(3, int(0.12 * (y1 - y0)))
    xa, ya = max(0, int(x0) - pad), max(0, int(y0) - 2)  # vertical: stay inside the OCR box
    xb, yb = min(w_img, int(x1) + pad + 1), min(h_img, int(y1) + 3)
    crop = arr[ya:yb, xa:xb].astype(int)
    if crop.size == 0:
        return None
    bg = np.median(ring_pixels(arr, (xa, ya, xb, yb)), axis=0)
    dist = np.abs(crop - bg).sum(axis=2)
    if dist.max() < 60:
        return None
    mask = dist > max(60, 0.40 * dist.max())
    # keep only the text band that contains this line (drop ink of neighbouring lines)
    rows = mask.sum(axis=1)
    lo = hi = int(np.argmax(rows))
    while lo > 0 and rows[lo - 1] > 0:
        lo -= 1
    while hi < len(rows) - 1 and rows[hi + 1] > 0:
        hi += 1
    mask[:lo] = False
    mask[hi + 1:] = False
    ys, xs = np.where(mask)
    if len(ys) < 8:
        return None
    ix0, ix1, iy0, iy1 = xs.min(), xs.max() + 1, ys.min(), ys.max() + 1
    strong = dist >= np.quantile(dist[mask], 0.85)
    fg = tuple(int(v) for v in np.median(crop[strong & mask], axis=0))
    return {"ix0": xa + ix0, "ix1": xa + ix1, "iy0": ya + iy0, "iy1": ya + iy1,
            "mask": mask[iy0:iy1, ix0:ix1], "region": (xa, ya, mask), "fg": fg, "bg": tuple(int(v) for v in bg)}


DESC = set("gjpqy()[]{}/|@Q$")


def em_from_ink(text, ink_h):
    f = 0.72 + (0.21 if any(c in DESC for c in text) else 0.0)
    return ink_h / f


# --------------------------------------------------- font style analysis
def render_runs(runs, fam, em_px):
    """Render text with a local font; return the tight ink mask (or None)."""
    size = max(8, int(round(em_px)))
    fonts = {b: get_font(fam, b, size) for b in (False, True)}
    if any(fonts[b] is None for _, b in runs):
        return None
    total = sum(fonts[b].getlength(t) for t, b in runs)
    im = Image.new("L", (int(total + size), int(size * 1.7)), 255)
    d = ImageDraw.Draw(im)
    x, base = size * 0.3, int(size * 1.2)
    for t, b in runs:
        d.text((x, base), t, font=fonts[b], fill=0, anchor="ls")
        x += fonts[b].getlength(t)
    a = np.asarray(im) < 140
    ys, xs = np.where(a)
    if len(ys) == 0:
        return None
    return a[ys.min():ys.max() + 1, xs.min():xs.max() + 1]


def stroke_w(mask):
    """Average stroke thickness in pixels (2 x area / perimeter)."""
    mk = mask.astype(bool)
    p = np.pad(mk, 1)
    inner = p[1:-1, 1:-1] & p[:-2, 1:-1] & p[2:, 1:-1] & p[1:-1, :-2] & p[1:-1, 2:]
    return 2.0 * mk.sum() / max(1, (mk & ~inner).sum())


def _run_lengths(a):
    """Lengths of all horizontal ink runs of a boolean image."""
    d = np.diff(np.pad(a, ((0, 0), (1, 1))).astype(np.int8), axis=1)
    return (np.where(d == -1)[1] - np.where(d == 1)[1]).astype(float)


def serif_contrast(mask):
    """Thin-stroke / stem ratio: ~0.5 for serif fonts (hairlines), ~0.8+ for sans fonts."""
    h, v = _run_lengths(mask), _run_lengths(mask.T)
    if len(h) < 5 or len(v) < 5:
        return None
    return float(np.quantile(v, 0.20) / max(1.0, np.median(h)))


SERIF_BELOW = 0.66


def weight_ratio(mask, text, fam, em):
    """Target stroke thickness divided by that of the regular local font (1.0 = regular)."""
    r = render_runs([(text, False)], fam, em)
    if r is None or mask.size == 0:
        return None
    wr = stroke_w(r)
    return stroke_w(mask) / wr if wr > 0 else None


BOLD_RATIO = 1.40


def decide_style(items, blk, fams):
    pairs = [(serif_contrast(items[i]["mask"]), items[i]["mask"].shape[1]) for i in blk]
    pairs = [(c, w) for c, w in pairs if c is not None]
    contrast = sum(c * w for c, w in pairs) / sum(w for _, w in pairs) if pairs else 1.0
    fam = "serif" if ("serif" in fams and contrast < SERIF_BELOW) else "sans"
    if fam not in fams:
        fam = fams[0]
    ratios = [weight_ratio(items[i]["mask"], items[i]["text"], fam, items[i]["em"]) for i in blk]
    ratios = [r for r in ratios if r]
    bold = bool(ratios) and float(np.median(ratios)) >= BOLD_RATIO
    if fam == "sans" and "narrow" in fams:  # condensed typeface?
        rr = []
        for i in blk:
            r = render_runs([(items[i]["text"], bold)], "sans", items[i]["em"])
            if r is not None:
                rr.append(items[i]["mask"].shape[1] / r.shape[1])
        if rr and float(np.median(rr)) < 0.88:
            fam = "narrow"
    return fam, bold


def label_runs(it, fam):
    """Bold label before a colon ('Rule: text') on an otherwise regular line."""
    m = re.match(r"^((?:\u2022 )?[^:]{2,50}:)(\s?.{3,})$", it["text"])
    if not m:
        return None
    pre, rest = m.group(1), m.group(2)
    em, mask = it["em"], it["mask"]
    tot = render_runs([(pre, True), (rest, False)], fam, em)
    prb = render_runs([(pre, True)], fam, em)
    if tot is None or prb is None:
        return None
    cut = int(mask.shape[1] * prb.shape[1] / tot.shape[1])
    gap = int(0.25 * em)
    if cut < 4 or mask.shape[1] - cut - gap < 8:
        return None
    r_pre = weight_ratio(mask[:, :max(2, cut - int(0.1 * em))], pre, fam, em)
    r_rest = weight_ratio(mask[:, cut + gap:], rest, fam, em)
    if r_pre and r_rest and r_pre >= BOLD_RATIO and r_rest < BOLD_RATIO:
        return [(pre, True), (rest, False)]
    return None


# ---------------------------------------------------------------- layout
def group_blocks(items):
    order = sorted(range(len(items)), key=lambda i: (items[i]["iy0"], items[i]["ix0"]))
    blocks = []
    for i in order:
        it, h = items[i], items[i]["iy1"] - items[i]["iy0"]
        best, bd = None, 1e18
        for b in blocks:
            last = items[b[-1]]
            gap = it["iy0"] - last["iy1"]
            if not (-0.35 * h <= gap <= 1.1 * h):
                continue
            if not (0.8 <= it["em"] / last["em"] <= 1.25):
                continue
            dx = min(abs(it["ix0"] - last["ix0"]), abs((it["ix0"] + it["ix1"]) / 2 - (last["ix0"] + last["ix1"]) / 2))
            if dx > 0.6 * h:
                continue
            if gap + dx < bd:
                best, bd = b, gap + dx
        if best is None:
            blocks.append([i])
        else:
            best.append(i)
    return blocks


# ------------------------------------------------------- background clean
def clean_background(arr, items, extra_boxes):
    h, w = arr.shape[:2]
    mask = np.zeros((h, w), np.uint8)
    for it in items:
        xa, ya, m = it["region"]
        mask[ya:ya + m.shape[0], xa:xa + m.shape[1]][m] = 255
    for (x0, y0, x1, y1) in extra_boxes:
        mask[max(0, y0):min(h, y1), max(0, x0):min(w, x1)] = 255
    if not mask.any():
        return arr
    if cv2 is not None:
        mask = cv2.dilate(mask, np.ones((3, 3), np.uint8), iterations=2)
        return cv2.inpaint(arr, mask, 3, cv2.INPAINT_TELEA)
    out = arr.copy()
    for it in items:  # fallback: flat fill with surrounding colour
        xa, ya, m = it["region"]
        out[ya:ya + m.shape[0], xa:xa + m.shape[1]][m] = it["bg"]
    return out


# ---------------------------------------------------------- pptx writing
def add_textbox(slide, it, runs, fam, pt, spc_pt, scale_in, ox_in=0.0):
    """scale_in = inches per image pixel."""
    em_px = pt / 72.0 / scale_in
    baseline = it["iy1"] - (0.21 * em_px if any(c in DESC for c in it["text"]) else 0.0)
    centre = baseline - 0.3465 * em_px
    height_px = 1.3 * em_px
    left = it["ix0"] * scale_in
    top = (centre - height_px / 2) * scale_in
    width = (it["ix1"] - it["ix0"]) * scale_in + 0.06
    box = slide.shapes.add_textbox(Inches(max(0, left)), Inches(max(0, top)),
                                   Inches(width), Inches(height_px * scale_in))
    tf = box.text_frame
    tf.word_wrap = False
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    for text, bold in runs:
        r = p.add_run()
        r.text = text
        r.font.name = FAMILIES[fam][0]
        r.font.size = Pt(pt)
        r.font.bold = bold
        r.font.color.rgb = RGBColor(*it["fg"])
        if abs(spc_pt) >= 0.01:
            r._r.get_or_add_rPr().set("spc", str(int(round(spc_pt * 100))))


# -------------------------------------------------------------- converter
def convert_pdf(pdf_bytes, editable=True, dpi=120, remove_watermark=True, min_conf=0.5, progress=None):
    report, warnings = [], []
    k = dpi / 72.0
    fams = available_families()
    if editable and not fams:
        raise RuntimeError("No usable fonts found (need Arial/Liberation Sans).")
    with fitz.open(stream=pdf_bytes, filetype="pdf") as pdf:
        n = len(pdf)
        prs = Presentation()
        first = pdf[0].rect
        slide_w = 13.333
        slide_h = slide_w * first.height / first.width
        prs.slide_width, prs.slide_height = Inches(slide_w), Inches(slide_h)
        blank = prs.slide_layouts[6]

        for pi, page in enumerate(pdf):
            rect = page.rect
            pix = page.get_pixmap(matrix=fitz.Matrix(k, k), alpha=False)
            arr = np.array(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
            H, W = arr.shape[:2]
            scale_in = slide_w / W  # inches per pixel
            pt_per_px = scale_in * 72.0
            slide = prs.slides.add_slide(blank)
            items, extra, source = [], [], "image only"

            if editable:
                raw = embedded_lines(page, k)
                source = "embedded text"
                if len(raw) < 3:
                    if RapidOCR is None:
                        raise RuntimeError("RapidOCR is not installed. Run: pip install rapidocr-onnxruntime")
                    raw, source = ocr_lines(arr), "RapidOCR"
                for ln in raw:
                    x0, y0, x1, y1 = ln["box"]
                    if remove_watermark and "notebook" in ln["text"].lower() and y0 > 0.9 * H and x0 > 0.75 * W:
                        hh = y1 - y0
                        extra.append((int(x0 - 1.6 * hh), int(y0 - 3), int(x1 + 6), int(y1 + 3)))
                        continue
                    if ln["conf"] < min_conf:
                        continue
                    info = analyze_line(arr, ln["box"])
                    if info is None:
                        continue
                    it = {**info, "text": ln["text"], "conf": ln["conf"]}
                    it["em"] = em_from_ink(it["text"], it["iy1"] - it["iy0"])
                    items.append(it)
                    if ln["conf"] < 0.85:
                        warnings.append(f"Page {pi + 1}: check '{ln['text'][:50]}' (confidence {ln['conf']:.2f})")

            # ---- style per paragraph block: family, weight, one common size
            placed = []
            if items:
                for blk in group_blocks(items):
                    fam, bold = decide_style(items, blk, fams)
                    med = float(np.median([items[i]["em"] for i in blk]))
                    for i in blk:
                        it = items[i]
                        em = med if (len(blk) > 1 or abs(it["em"] / med - 1) < 0.12) else it["em"]
                        pt = max(6.0, round(em * pt_per_px * 2) / 2)
                        runs = [(it["text"], bold)]
                        if not bold:
                            runs = label_runs(it, fam) or runs
                        # letter-spacing so the width matches the picture exactly
                        nat = render_runs(runs, fam, pt / pt_per_px)
                        spc = 0.0
                        if nat is not None and len(it["text"]) > 1:
                            spc = (it["ix1"] - it["ix0"] - nat.shape[1]) / (len(it["text"]) - 1) * pt_per_px
                            spc = max(-4.0, min(4.0, spc))
                        placed.append((it, runs, fam, pt, spc))

            if editable:
                arr = clean_background(arr, items, extra)
            buf = io.BytesIO()
            Image.fromarray(arr).save(buf, "JPEG", quality=94, subsampling=0)
            buf.seek(0)
            slide.shapes.add_picture(buf, 0, 0, Inches(slide_w), Inches(slide_h))
            for it, runs, fam, pt, spc in placed:
                add_textbox(slide, it, runs, fam, pt, spc, scale_in)

            report.append(f"Page {pi + 1}: text boxes={len(placed)}; source={source}")
            if editable and not placed:
                warnings.append(f"Page {pi + 1}: no text found.")
            if progress:
                progress((pi + 1) / n, f"Converted {pi + 1}/{n} pages")

        out = io.BytesIO()
        prs.save(out)
    return out.getvalue(), report, warnings


# ---------------------------------------------------------------------- UI
if True:  # Streamlit script body
    st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="centered")
    st.title(APP_TITLE)
    st.write("Upload a NotebookLM PDF. Text becomes real, editable PowerPoint text; the rest of the "
             "slide stays as a clean background picture.")
    if RapidOCR is None:
        st.error("RapidOCR is not installed. Add `rapidocr-onnxruntime` to requirements.txt "
                 "(Python 3.10–3.12) and restart.")
    _fams = available_families()
    if "narrow" not in _fams:
        st.info("Arial Narrow was not found on this computer (it ships with Microsoft Office). "
                "Condensed slide text will be set in Arial with tighter letter-spacing instead.")
    if "serif" not in _fams:
        st.info("Georgia/Times was not found; serif text will be set in Arial.")
    uploaded = st.file_uploader("Upload your NotebookLM PDF", type=["pdf"])
    mode = st.radio("Conversion mode", ["Editable slides (recommended)", "Image-only slides (exact look)"], index=0)
    dpi = st.slider("Working resolution (DPI)", 90, 200, 120, 10,
                    help="Higher = sharper background and slightly better OCR, but slower and larger files.")
    remove_wm = st.checkbox("Remove the 'Gemini/NotebookLM' watermark", value=True)
    include_report = st.checkbox("Include conversion report (.txt)", value=True)

    if uploaded:
        pdf_bytes = uploaded.getvalue()
        if not pdf_bytes.startswith(b"%PDF"):
            st.error("This file does not appear to be a valid PDF.")
            st.stop()
        if st.button("Convert PDF to PowerPoint", type="primary", use_container_width=True):
            try:
                bar = st.progress(0, text="Starting… (first run loads the OCR model)")
                data, report, warns = convert_pdf(
                    pdf_bytes, mode.startswith("Editable"), dpi, remove_wm,
                    progress=lambda v, t: bar.progress(v, text=t))
                st.session_state["pptx"] = data
                st.session_state["name"] = Path(uploaded.name).stem + "_editable.pptx"
                st.session_state["report"] = "\n".join(report + [""] + warns)
                st.success(f"Conversion complete: {len(report)} slides.")
                if warns:
                    with st.expander(f"{len(warns)} item(s) worth a quick check"):
                        st.write("\n".join(warns))
            except Exception as exc:
                st.error(f"Conversion failed: {exc}")
                st.exception(exc)

    if "pptx" in st.session_state:
        st.download_button(
            "Download PowerPoint (.pptx)", data=st.session_state["pptx"], file_name=st.session_state["name"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary", use_container_width=True)
        if include_report:
            st.download_button("Download conversion report (.txt)", data=st.session_state["report"].encode("utf-8"),
                               file_name="conversion_report.txt", mime="text/plain", use_container_width=True)
