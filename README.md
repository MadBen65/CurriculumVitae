# Curriculum Vitae Site

This repository contains my personal CV as a static GitHub Pages site.

The site is intentionally simple:

- [index.html](index.html) is the homepage and contains the full CV layout and content.
- GitHub Pages serves the site directly from the repository root.

## What The Site Shows

- Senior DevOps Engineer summary
- Core skills grouped as a compact tag-style list
- Professional experience
- Contact details with a clickable email address and LinkedIn profile

## Export To Word And PDF

Use [Convert-CV.ps1](Convert-CV.ps1) to convert the HTML master page into both `.docx` and `.pdf`.

### Default usage

```powershell
.\Convert-CV.ps1
```

This uses:

- Source HTML: `index.html`
- Output names: `Ben-Madle-Jordan-CV.docx` and `Ben-Madle-Jordan-CV.pdf`
- Quick Style Set: `Black & White (Classic)`
- Margins: `0.5` inches (narrow)

### Custom usage

```powershell
.\Convert-CV.ps1 -HtmlPath index.html -OutputBaseName "Ben-Madle-Jordan-CV" -OutputDirectory .
```
