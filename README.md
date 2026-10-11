# TLaFT Metadata Reuse Shiny App
R Shiny tool for clinical Title & Footnote metadata recycling, intelligent matching and compliant Excel export.

## Highlights
- Recursively scan legacy TLaFT files and build reusable metadata library
- Weighted cosine similarity to auto-match TOC entries
- Human-in-loop cart for review and in-place editing
- Auto formatted Excel output (border, zebra stripe, auto column width, remove TOC_No)
- Full audit trail preserved for regulatory QC

## R packages
-	shiny (v1.13.0): Interactive web UI construction, drag panel resizing, custom JavaScript interaction and event-driven human-computer interaction;
-	stringdist (v0.9.17): Optimized cosine distance calculation for normalized title semantic similarity scoring;
-	openxlsx (v4.2.8.1): Regulatory-aligned Excel read/write, sheet styling, freeze panes, adaptive column width and bilingual CN/EN worksheet generation;
-	fs (v2.1.0): Recursive directory traversal, temporary file filtering, secure path parsing for scattered cross-study TLaFT resource recycling;
-	rhandsontable (v0.3.8): Editable interactive table for match review, locked protected core columns, right-click context menu batch operation;
-	Auxiliary supporting packages: dplyr (v1.2.1) and stringr (v1.6.0) for vectorized data cleansing, text normalization and structured metadata integration;
-	Additional dependent utilities: DT (0.34.0), readxl (1.5.0), shinyjs (2.1.1) for table rendering, raw Excel reading and frontend page interactive optimization.
  
