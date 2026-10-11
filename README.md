# TLaFT Metadata Reuse Shiny App
R Shiny tool for clinical Title & Footnote metadata recycling, intelligent matching and compliant Excel export.

## Highlights
- Recursively scan legacy TLaFT files and build reusable metadata library
- Weighted cosine similarity to auto-match TOC entries
- Human-in-loop cart for review and in-place editing
- Auto formatted Excel output (border, zebra stripe, auto column width, remove TOC_No)
- Full audit trail preserved for regulatory QC

## Run
```r
shiny::runApp("app_folder")
