"""Convert PowerPoint slides into an editable Quarto Reveal.js presentation."""

import argparse
from pathlib import Path
import re
import shlex
import sys
import shutil
from tempfile import TemporaryDirectory
import yaml

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx2md.outputter import MarkdownFormatter
from pptx2md.parser import parse
from pptx2md.types import ConversionConfig, ElementType


def compact(text):
    return " ".join(text.split())


def warn(slide_number, message):
    print(f"Slide {slide_number}: {message}", file=sys.stderr)


def review_warnings(slide, number, slide_width, slide_height):
    """Flag layout losses, including content that upstream cannot represent."""
    reasons = set()
    pictures = 0

    def visit(shapes, top_level=True):
        nonlocal pictures
        for shape in shapes:
            if top_level and (
                shape.left < 0
                or shape.top < 0
                or shape.left + shape.width > slide_width
                or shape.top + shape.height > slide_height
            ):
                reasons.add(
                    "off-canvas objects are retained; check whether they belong in the draft"
                )
            if shape.rotation:
                reasons.add("object rotation is not preserved")
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                reasons.add(
                    "grouped diagram flattened; shapes and connections need manual reconstruction"
                )
                visit(shape.shapes, top_level=False)
            elif shape._element.tag.endswith("}pic"):
                pictures += 1
                if any(
                    getattr(shape, f"crop_{side}", 0)
                    for side in ("left", "right", "top", "bottom")
                ):
                    reasons.add(
                        "picture cropping is not applied; original images are extracted"
                    )
            elif shape.shape_type == MSO_SHAPE_TYPE.LINE:
                reasons.add("lines/connectors are not preserved")
            elif shape.has_chart:
                reasons.add("chart is unsupported and needs manual conversion")
            elif shape.has_table:
                reasons.add(
                    "table text retained; merged cells and table formatting need review"
                )
            elif not shape.has_text_frame:
                reasons.add(
                    f"unsupported object ({shape.shape_type}) needs manual conversion"
                )
            elif not shape.is_placeholder:
                reasons.add(
                    "free-positioned text/shapes are flattened; drawing geometry is not preserved"
                )

    visit(slide.shapes)
    if pictures > 1:
        reasons.add(
            f"{pictures} images extracted in reading order; arrange or select them manually"
        )
    if slide._element.get("show") == "0":
        reasons.add("source slide is hidden; it is included in this draft")
    for reason in sorted(reasons):
        warn(number, reason)


class DraftFormatter(MarkdownFormatter):
    """Reuse upstream inline formatting, without its title/order heuristics."""

    def emphasis(self, text, marker):
        # Upstream wraps whitespace inside markers, which can leave literal
        # underscores on slides when PowerPoint splits text into many runs.
        if not text.strip():
            return text
        leading = text[: len(text) - len(text.lstrip())]
        trailing = text[len(text.rstrip()) :]
        return f"{leading}{marker}{text.strip()}{marker}{trailing}"

    def get_strong(self, text):
        return self.emphasis(text, "__")

    def get_accent(self, text):
        return self.emphasis(text, "_")

    def output_draft(self, parsed, source, document_title, fmt):
        try:
            metadata = {"pagetitle": document_title, "format": fmt, "slide-level": 2}
            self.write(
                "---\n" + yaml.safe_dump(metadata, allow_unicode=True) + "---\n\n"
            )
            for number, (slide, original) in enumerate(
                zip(parsed.slides, source.slides, strict=True), 1
            ):
                title_shape = original.shapes.title
                source_title = (
                    compact(title_shape.text) if title_shape is not None else ""
                )
                self.put_title(self.get_escaped(source_title or f"Slide {number}"), 2)
                title_consumed = False
                previous = None
                for element in slide.elements:
                    kind = element.type
                    if previous == ElementType.ListItem and kind != previous:
                        self.put_list_footer()
                    if kind == ElementType.Title:
                        if (
                            source_title
                            and not title_consumed
                            and compact(element.content) == source_title
                        ):
                            title_consumed = True
                            continue
                        self.put_para(self.get_escaped(compact(element.content)))
                    elif kind == ElementType.Paragraph:
                        self.put_para(self.get_formatted_runs(element.content))
                    elif kind == ElementType.ListItem:
                        if previous != kind:
                            self.put_list_header()
                        self.put_list(
                            self.get_formatted_runs(element.content), element.level
                        )
                    elif kind == ElementType.Image:
                        self.put_image(element.path)
                    elif kind == ElementType.Table:
                        self.put_table(
                            [
                                [self.get_formatted_runs(cell) for cell in row]
                                for row in element.content
                            ]
                        )
                    previous = kind
                if previous == ElementType.ListItem:
                    self.put_list_footer()
                if slide.notes:
                    self.put_para("::: {.notes}")
                    for note in slide.notes:
                        # Pandoc accepts shorter closing fences too, so escape
                        # fence-like source text rather than extending our fence.
                        self.put_para(re.sub(r"(?m)^([ \t]*)(:{3,})", r"\1\\\2", note))
                    self.put_para(":::")
        finally:
            self.close()


def write_project(input_path, output_dir, reference, fmt):
    if input_path.suffix.lower() != ".pptx" or not input_path.is_file():
        raise ValueError(f"Expected an existing .pptx file: {input_path}")
    # Read directly: upstream load_pptx can write a repaired copy beside the input.
    source = Presentation(input_path)
    if not len(source.slides):
        raise ValueError("The presentation has no slides")
    output_path = output_dir / f"{input_path.stem}.qmd"
    assets = output_dir / "assets"
    assets.mkdir()
    if reference:
        from pdf_layout import write_reference

        write_reference(source, reference, output_path, fmt)
        return
    config = ConversionConfig(
        pptx_path=input_path,
        output_path=output_path,
        image_dir=assets,
        disable_color=True,
        min_block_size=0,
        try_multi_column=False,
    )
    for number, slide in enumerate(source.slides, 1):
        review_warnings(slide, number, source.slide_width, source.slide_height)
        # python-pptx's runs omit soft line breaks; keep adjacent words separate.
        # Only change the in-memory tree, never the source presentation.
        for br in list(
            slide._element.iter(
                "{http://schemas.openxmlformats.org/drawingml/2006/main}br"
            )
        ):
            run = OxmlElement("a:r")
            text = OxmlElement("a:t")
            text.text = " "
            run.append(text)
            br.getparent().replace(br, run)
    parsed = parse(config, source)
    DraftFormatter(config).output_draft(parsed, source, input_path.stem, fmt)


def convert(input_path, output_dir, reference=None, fmt=None, extension=None):
    if output_dir.exists():
        raise ValueError(f"Output directory already exists: {output_dir}")
    if extension:
        extension = extension.resolve()
        if output_dir.is_relative_to(extension):
            raise ValueError(
                "Choose an output directory outside the extension directory"
            )
        manifest = extension / "_extension.yml"
        if not manifest.is_file():
            raise ValueError(
                f"Expected a Quarto extension directory containing {manifest.name}"
            )
        formats = yaml.safe_load(manifest.read_text(encoding="utf-8"))["contributes"][
            "formats"
        ]
        if "revealjs" not in formats:
            raise ValueError("The extension must contribute a revealjs format")
        fmt = fmt or f"{extension.name}-revealjs"
    fmt = fmt or "revealjs"
    if fmt != "revealjs" and not re.fullmatch(
        r"[A-Za-z0-9][A-Za-z0-9_/-]*-revealjs", fmt
    ):
        raise ValueError("Use revealjs or an extension format ending in -revealjs")
    output_dir.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(dir=output_dir.parent) as temporary:
        project = Path(temporary) / "project"
        project.mkdir()
        write_project(input_path, project, reference, fmt)
        if extension:
            shutil.copytree(
                extension,
                project / "_extensions" / extension.name,
                ignore=shutil.ignore_patterns(".git", "__pycache__", ".quarto"),
            )
        project.rename(output_dir)
    print(f"Created {output_dir / (input_path.stem + '.qmd')}")
    print(f"  cd {shlex.quote(str(output_dir))}")
    print(f"  quarto render {shlex.quote(input_path.stem + '.qmd')}")


def main():
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("input", type=Path, help="source .pptx (read only)")
    parser.add_argument(
        "--output-dir",
        required=True,
        type=Path,
        help="new directory for .qmd and assets",
    )
    parser.add_argument(
        "--reference-pdf",
        type=Path,
        help="matching PowerPoint PDF export for positioned 16:9 slides",
    )
    parser.add_argument(
        "--extension-dir",
        type=Path,
        help="local Reveal.js extension directory containing _extension.yml",
    )
    parser.add_argument(
        "--format",
        dest="fmt",
        help="Quarto output format (default: revealjs, or the supplied extension)",
    )
    args = parser.parse_args()
    try:
        convert(
            args.input.resolve(),
            args.output_dir.resolve(),
            args.reference_pdf,
            args.fmt,
            args.extension_dir,
        )
    except Exception as error:
        parser.exit(1, f"Conversion failed: {error}\n")


if __name__ == "__main__":
    main()
