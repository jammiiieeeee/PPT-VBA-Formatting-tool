# PPT VBA Formatting Tool

A single-file VBA module for PowerPoint that inspects slides for issues and flattens presentations into clean, image-backed `.pptx` files — ideal for Articulate or LMS uploads where native shapes can cause rendering problems.

## Quick Start

1. Open your `.pptx` in PowerPoint
2. Press **Alt+F11** to open the VBA editor
3. **Insert → Module**, then paste the contents of `PPT VBA Tool.txt`
4. Press **Alt+F8**, choose a macro, and click **Run**

## Two Macros

| Macro | What it does |
|-------|-------------|
| `PPTTools_CheckSlides` | Scans every slide for issues — **read-only**, nothing is changed |
| `PPTTools_Flatten` | Flattens the presentation into a `_Flattened.pptx` copy |

**Recommended workflow:** run Check first → fix any reported issues → then Flatten.

## What CheckSlides Detects (11 checks)

- **Title/index errors** — duplicate titles, conflicting numbering, ambiguous shape selection
- **Missing notes** — slides with empty or absent speaker notes
- **Translation gaps** — shapes or notes containing text in a different language from the presentation default

## What Flatten Does

- Exports each slide as a high-resolution PNG (2×)
- Replaces all shapes with the PNG background, keeping only audio/video media
- Injects invisible native title placeholders so the slide sorter and accessibility tools still work
- Writes the result to `<filename>_Flattened.pptx` alongside the original

## Skip Guardrails

Slides are automatically skipped (not flattened, not renamed) when they match known non-content patterns — "Thank You" slides, Q&A slides, slides with ambiguous titles, etc. All skips are logged with a reason.

## Debug Log

Each run produces a timestamped log file beside the original presentation. Check runs open the log automatically; flatten runs write a brief summary.

## Requirements

- Microsoft PowerPoint (Windows)
- A saved `.pptx` file (the macros will exit with an error if the file is unsaved)

## User Manual

See `PPT_VBA_Tool_User_Manual.docx` for a detailed walkthrough with screenshots.
