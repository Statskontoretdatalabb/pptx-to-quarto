"""Editable text and placed artwork for slides with a matching PowerPoint PDF."""

from pathlib import Path
from html import escape
import re
import json
from pptx.text.text import _Run
import pymupdf as pdf


def box(shape, sx, sy):
    x, y = shape.left * sx, shape.top * sy
    return pdf.Rect(x, y, x + shape.width * sx, y + shape.height * sy)


def style(rect):
    return f"left:{rect.x0:.2f}px;top:{rect.y0:.2f}px;width:{rect.width:.2f}px"


def inline(paragraph):
    pieces = []
    for child in paragraph._p:
        if child.tag.endswith("}br"):
            pieces.append("<br>")
        elif child.tag.endswith("}r"):
            run = _Run(child, paragraph)
            value = escape(run.text)
            if run.font.bold:
                value = f"<strong>{value}</strong>"
            if run.font.italic:
                value = f"<em>{value}</em>"
            url = run.hyperlink.address
            if url:
                value = f'<a href="{escape(url, quote=True)}">{value}</a>'
            pieces.append(value)
    return "".join(pieces)


def write_reference(source, reference, output, fmt):
    deck = pdf.open(reference)
    if len(deck) != len(source.slides):
        raise ValueError("PDF page count differs from PowerPoint slide count")
    if abs(source.slide_width / source.slide_height - 16 / 9) > 0.01:
        raise ValueError("PDF layout mode requires a widescreen (16:9) presentation")
    sx, sy = 1280 / source.slide_width, 720 / source.slide_height
    css = [
        ".reveal .slides section.pdf-layout {box-sizing:border-box;height:720px;padding:0!important;overflow:hidden;text-align:left}",
        ".reveal .slides section.pdf-layout > h2,.pdf-layout .placed {position:absolute;margin:0}",
        ".pdf-layout .placed {z-index:1;box-sizing:border-box}",
        ".pdf-layout .placed img {display:block;width:100%;max-width:none;max-height:none;margin:0}",
        ".pdf-layout .placed p {margin:0 0 .30em;line-height:1.2}",
        ".pdf-layout .placed a {text-decoration:underline}",
        ".pdf-layout .item {padding-left:.95em;text-indent:-.95em}",
        ".pdf-layout .placed p.sub {margin-left:1.05em;font-size:.91em}",
        ".pdf-layout .ppt-bullet {color:var(--r-heading-color,currentColor);display:inline-block;width:.95em;text-indent:0}",
    ]
    body = []
    meta = [
        "format: " + json.dumps(fmt),
        "css: layout.css",
        "slide-level: 2",
        "width: 1280",
        "height: 720",
        "center: false",
        "margin: 0",
        "pagetitle: " + json.dumps(output.stem),
    ]
    for n, (slide, page) in enumerate(zip(source.slides, deck), 1):
        title_shape = slide.shapes.title
        title = (
            " ".join(title_shape.text.split())
            if title_shape is not None
            else f"Slide {n}"
        )
        scale = 1280 / page.rect.width
        if abs(page.rect.height * scale - 720) > 1:
            raise ValueError(f"PDF page {n} does not match the 16:9 slide canvas")
        lines = [
            line
            for block in page.get_text("dict", flags=0)["blocks"]
            for line in block.get("lines", [])
        ]
        text = " ".join("".join(s["text"] for s in line["spans"]) for line in lines)
        norm = lambda t: re.sub(r"\s+", "", t).casefold()
        if title_shape is not None and norm(title) not in norm(text):
            raise ValueError(f"PDF title differs on slide {n}: {title}")
        others = [s for s in slide.shapes if s != title_shape]
        notes = slide.notes_slide.notes_text_frame if slide.has_notes_slide else None
        if n == 1 and "rubrikbild" in slide.slide_layout.name.lower() and not notes:
            details = [
                p.text
                for s in others
                if s.has_text_frame
                for p in s.text_frame.paragraphs
                if p.text
            ]
            meta.append("title: " + json.dumps(title, ensure_ascii=False))
            meta.append(
                "subtitle: " + json.dumps("  \n".join(details), ensure_ascii=False)
            )
            css.append(
                ".reveal .slides #title-slide h1.title {font-size:68px;width:100%;max-width:100%}.reveal .slides #title-slide .subtitle {position:static;margin:32px 0 0;font-size:29px;font-weight:bold}"
            )
            continue
        diagram = any(s._element.tag.endswith(("}grpSp", "}cxnSp")) for s in others)
        pictures = [s for s in others if s._element.tag.endswith("}pic")]
        composite = diagram or (
            pictures
            and any(
                not s.is_placeholder and not s._element.tag.endswith("}pic")
                for s in others
            )
        )
        visuals = [
            s for s in others if s in pictures or (composite and not s.is_placeholder)
        ]
        closing = "avslutningsbild" in slide.slide_layout.name.lower()
        safe_title = re.sub(r"([\\*_`#{}\[\]<>])", r"\\\1", title)
        body.append(
            f"## {safe_title} {{.pdf-layout .page-{n}"
            + (" .closing" if closing else "")
            + "}"
        )
        if title_shape is not None:
            r = box(title_shape, sx, sy)
            heading = next(
                (
                    line["spans"][0]
                    for line in lines
                    if line["spans"]
                    and norm(title).startswith(norm(line["spans"][0]["text"]))
                    and line["spans"][0]["text"].strip()
                ),
                None,
            )
            if heading:
                x, y = heading["origin"]
                size = heading["size"] * scale
                width = r.width
                top = y * scale - size
                for picture in pictures:
                    image = box(picture, sx, sy)
                    if (
                        x * scale < image.x0
                        and image.y0 < top + 1.15 * size
                        and image.y1 > top
                    ):
                        width = min(width, image.x0 - x * scale - 20)
                if 0 < width < r.width:
                    # Reserve room for browser font substitution as well as the gutter.
                    advance = pdf.get_text_length(title, fontname="hebo", fontsize=1)
                    size = min(size, width / max(advance * 1.12, 1))
                css.append(
                    f".reveal .slides section.page-{n} > h2 {{left:{x * scale:.1f}px;top:{top:.1f}px;width:{width:.1f}px;font-size:{size:.1f}px;line-height:1.15}}"
                )
        groups = [visuals] if composite and visuals else [[s] for s in visuals]
        for idx, group in enumerate(groups, 1):
            bounds = pdf.Rect()
            for s in group:
                visible = box(s, sx, sy) & pdf.Rect(0, 0, 1280, 720)
                if not visible.is_empty:
                    bounds |= visible
            if bounds.is_empty:
                continue
            # Export the visible composition including PowerPoint crops, overlaps and shadows.
            clip = (bounds / scale) & page.rect
            asset = f"assets/slide-{n}-{idx}.png"
            page.get_pixmap(matrix=pdf.Matrix(2, 2), clip=clip).save(
                output.parent / asset
            )
            body.append(
                f'<div class="placed" style="{style(bounds)}"><img src="{asset}" alt="{escape(title, quote=True)}"></div>'
            )
        for shape in others:
            if shape in visuals or not shape.has_text_frame or not shape.text.strip():
                continue
            r = box(shape, sx, sy)
            tf = shape.text_frame
            r.x0 += tf.margin_left * sx
            r.x1 -= tf.margin_right * sx
            ps = [p for p in tf.paragraphs if p.text.strip()]
            # Use the rendered first baseline to account for placeholder insets.
            first = next(
                (
                    s
                    for line in lines
                    for s in line["spans"]
                    if len(s["text"].strip()) > 3
                    and norm(ps[0].text).startswith(norm(s["text"]))
                ),
                None,
            )
            size = first["size"] * scale if first else 29.33
            if first:
                r.y0 = first["origin"][1] * scale - size
            items = []
            for p in ps:
                content = inline(p)
                # Read PowerPoint's actual bullet or number at the start of the PDF line.
                marker = ""
                for line in lines:
                    spans = line["spans"]
                    start = spans[0]["text"].strip()
                    if (
                        re.fullmatch(r"[‒–•]|\d+\.", start)
                        and norm("".join(s["text"] for s in spans[1:]))
                        and norm(p.text).startswith(
                            norm("".join(s["text"] for s in spans[1:]))
                        )
                    ):
                        marker = start
                        break
                if p.level and not marker:
                    marker = "•"
                    for index, line in enumerate(lines[:-1]):
                        label = line["spans"][0]
                        following = lines[index + 1]["spans"][0]
                        if (
                            re.fullmatch(r"\d+\.", label["text"].strip())
                            and abs(label["origin"][1] - following["origin"][1]) < 2
                            and norm(p.text).startswith(norm(following["text"]))
                        ):
                            marker = label["text"].strip()
                            break
                klass = "item" if marker else ""
                if p.level:
                    klass += " sub"
                if marker:
                    content = f'<span class="ppt-bullet">{marker}</span>' + content
                items.append(f'<p class="{klass}">{content}</p>')
            body.append(
                f'<div class="placed" style="{style(r)};font-size:{size:.2f}px">\n'
                + "\n".join(items)
                + "\n</div>"
            )
        if notes and notes.text.strip():
            body.append("::: {.notes}\n\n" + notes.text + "\n\n:::")
    output.write_text(
        "---\n" + "\n".join(meta) + "\n---\n\n" + "\n\n".join(body) + "\n"
    )
    (output.parent / "layout.css").write_text("\n".join(css) + "\n")
