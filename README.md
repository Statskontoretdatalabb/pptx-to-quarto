# pptx-to-quarto

Convert a PowerPoint presentation to an editable Quarto Reveal.js deck. The default output uses Quarto’s built-in theme.

With [uv](https://docs.astral.sh/uv/) and [Quarto](https://quarto.org/docs/get-started/) installed, run from a clone of this repository:

```sh
uv run pptx-to-quarto slides.pptx --output-dir deck
cd deck
quarto render slides.qmd
```

Open `slides.html` in a browser. Edit `slides.qmd` and render again. Choose a new output directory for each conversion.

**Preserve slide layout.** For a 16:9 deck, export a matching PDF from PowerPoint and add `--reference-pdf slides.pdf`. Headings, body text and links remain editable. The converter crops screenshots and complex illustrations from the PDF and places them at their original coordinates. Review the result for font differences and overlap. Without a PDF, the converter extracts Markdown, images and speaker notes in slide order.

**Use a custom theme.** Pass the directory containing the Quarto extension’s `_extension.yml`:

```sh
uv run pptx-to-quarto slides.pptx --output-dir branded \
  --reference-pdf slides.pdf --extension-dir /path/to/_extensions/statskontoret
```

This copies the extension into the deck and selects `statskontoret-revealjs`. To use a format already available in your Quarto project, pass `--format NAME-revealjs` and place the output within that project. For a built-in Reveal.js theme such as `serif`, set `format: {revealjs: {theme: serif}}` in the generated YAML.

Run `uv run pptx-to-quarto --help` for all options. The converter uses [pptx2md](https://github.com/ssine/pptx2md), [python-pptx](https://github.com/scanny/python-pptx), and [PyMuPDF](https://github.com/pymupdf/PyMuPDF). Licensed under [AGPL-3.0](LICENSE).
