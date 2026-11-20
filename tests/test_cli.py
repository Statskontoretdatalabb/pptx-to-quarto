import subprocess
import tempfile
import unittest
from pathlib import Path

import pymupdf
import yaml
from pptx import Presentation


class ConversionTest(unittest.TestCase):
    def test_theme_selection(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "slides.pptx"
            presentation = Presentation()
            presentation.slide_width = 12192000
            presentation.slide_height = 6858000
            slide = presentation.slides.add_slide(presentation.slide_layouts[1])
            slide.shapes.title.text = "First slide"
            slide.placeholders[1].text = "An editable paragraph"
            presentation.save(source)
            reference = root / "slides.pdf"
            with pymupdf.open() as document:
                page = document.new_page(width=960, height=540)
                page.insert_text((50, 80), "First slide")
                page.insert_text((50, 160), "An editable paragraph")
                document.save(reference)
            extension = root / "my-theme"
            extension.mkdir()
            (extension / "_extension.yml").write_text(
                "contributes:\n  formats:\n    revealjs: {}\n"
            )
            for index, (flags, expected) in enumerate(
                [
                    ([], "revealjs"),
                    (["--reference-pdf", reference], "revealjs"),
                    (["--extension-dir", extension], "my-theme-revealjs"),
                    (["--format", "installed-revealjs"], "installed-revealjs"),
                    (
                        ["--reference-pdf", reference, "--extension-dir", extension],
                        "my-theme-revealjs",
                    ),
                ]
            ):
                with self.subTest(expected=expected, flags=flags):
                    output = root / str(index)
                    subprocess.run(
                        ["pptx-to-quarto", source, "--output-dir", output, *flags],
                        check=True,
                    )
                    qmd = (output / "slides.qmd").read_text()
                    self.assertEqual(
                        yaml.safe_load(qmd.split("---", 2)[1])["format"], expected
                    )
                    self.assertIn("An editable paragraph", qmd)
                    self.assertEqual(
                        (output / "_extensions/my-theme/_extension.yml").exists(),
                        "--extension-dir" in flags,
                    )
