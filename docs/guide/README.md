# KSC.Excel guides

| | Markdown | PDF |
|---|---|---|
| **English — complete guide** | [KSC.Excel-Guide.en.md](KSC.Excel-Guide.en.md) | [KSC.Excel-Guide.en.pdf](KSC.Excel-Guide.en.pdf) |
| **راهنمای کامل فارسی** | [KSC.Excel-Guide.fa.md](KSC.Excel-Guide.fa.md) | [KSC.Excel-Guide.fa.pdf](KSC.Excel-Guide.fa.pdf) |

Screenshots live in [`img/`](img/). Design & architecture (Persian): [../KSC.Excel-Design.md](../KSC.Excel-Design.md).

## Rebuilding the PDFs

The PDFs are generated from the Markdown files (Markdown → HTML → Chrome/Edge print):

```bash
pip install markdown
cd tools && npm install            # puppeteer-core (uses your installed Chrome/Edge)
python build-guide-html.py ../KSC.Excel-Guide.en.md ../_guide-en.html en
python build-guide-html.py ../KSC.Excel-Guide.fa.md ../_guide-fa.html fa
node print-pdf.js ../_guide-en.html ../KSC.Excel-Guide.en.pdf      # set CHROME_PATH if Chrome is not in the default location
node print-pdf.js ../_guide-fa.html ../KSC.Excel-Guide.fa.pdf
```

The `_guide-*.html` intermediates are git-ignored.
