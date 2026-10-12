# Check if the page contains a full-page image
        full_page_image = any(
            b.get("type") == 1 and b.get("bbox") and
            ((b["bbox"][2] - b["bbox"][0]) * (b["bbox"][3] - b["bbox"][1]) /
             max(1, page.rect.width * page.rect.height)) >= 0.70
            for b in page_dict.get("blocks", [])
        )
        
        # If editable mode is selected but there's no extractable text or it's a full-page image,
        # fallback to rendering the slide as an image so it is never blank.
        visual_fallback = (
            mode.startswith("Editable reconstruction") and 
            (not has_text or full_page_image) and 
            source != "embedded text"
        )

        if (mode == "Exact visual preservation (image slides)" or
            mode.startswith("Hybrid:") or visual_fallback):
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
            slide.shapes.add_picture(io.BytesIO(pix.tobytes("png")), 0, 0, Inches(slide_w), Inches(slide_h))
            if visual_fallback:
                source = source + " — full-page image fallback (visual content preserved)"

        editable_count = 0
        if mode != "Exact visual preservation (image slides)" and not visual_fallback:
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
