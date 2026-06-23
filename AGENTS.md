# AGENTS.md — PPT VBA

> **Do not commit or push changes without explicit user request.**

## What this is
Single VBA PowerPoint macro in `slide_flatten_rename.txt`. Entry point: `Sub Unified_ProcessAndFlattenPresentation()`. Paste into VBA editor (Alt+F11) and run.

## Preconditions
- Operates on `ActivePresentation` — the presentation open and active when run
- **Must be saved before running** — script exits with error otherwise
- Creates `_Flattened.pptx` copy in the same directory (e.g., `Deck.pptx` → `Deck_Flattened.pptx`)

## Skip guardrails
Slides are skipped (no title extraction, no rename) when either condition is met:
1. **Text keywords** — any visible shape on the slide contains "thank you", "thanks", "Q&A", or "??" (lowercased, case-insensitive regex)
2. **Slide name** — the slide's `.Name` is not `"Slide" & slideNum` AND does not start with `"Slide"`. Catches renamed/imported slides.

## Title override (lines 122–148)
Before depth-based selection, a pre-scan checks all visible shapes for priority keywords (case-insensitive). First match wins immediately — shape's full text is used as the title as-is, bypassing regex and separator counting entirely.

**Trigger keywords:** "course objective", "learning outcomes", "table of content", "学习目标", "学习成果", "内容目录"

Override slides get `slideIndices = ""` so they are excluded from numerical index repair (Step 2, Case 2), but still participate in contiguous duplicate suffixing (Step 2, Case 1).

## Shape selection algorithm (lines 150–302)
**Scope:** Only visible shapes with text frames where `Top < SlideHeight × 0.5`.

**Selection:**
1. Split each shape's text into blocks (multi-line shapes evaluated per-line by `vbCr`)
2. Test each block against regex: `^(\d+(?:[\.\(/\\\-_\??]\d+)*)\s*(.*)`
3. Count separator characters (`. ( / \ - _ ?`) in the matched index portion
4. **Winner = shape with highest separator count** (deepest nesting, e.g. `1.2.3` beats `1.2`)
5. **Tie-breaker** (lines 242–254): if multiple shapes tie at max separator count, pick the **topmost visible shape among the tied shapes** (any shape, even non-text; falls back gracefully via `On Error Resume Next`)

Note: the separator character class allows **mixing** (e.g. `1.2/3` matches). The space between index and title is optional (`\s*`), so `1.2.3Title` also matches.

## Conflict resolution (lines 276–426)
Two independent branches, checked in order:

### 1. Contiguous duplicate titles (lines 287–342)
Same `slideTexts` value on consecutive slides → suffix each with `" (1/N)"`, `"(2/N)"` etc.
- Example: Slides 5–6 both have text `"Setup"` → becomes `"Setup (1/2)"`, `"Setup (2/2)"`
- Non-contiguous duplicates are flagged as non-fixable (manual attention needed)

### 2. Conflicting numerical indices (lines 344–426)
Same index on adjacent slides → cascade-renumber all subsequent slides at the same structural level (detected by matching the parent prefix).
- Example: Slides 3–4 both have index `"2.1"` → Slide 4 becomes `"2.2"`, Slide 5 (`"2.2"`) becomes `"2.3"`, etc.

## Export & flatten (lines 438–584)
- Exports each slide as PNG at 2× resolution into `PPT_Temp_Images_Internal\` (deleted after run)
- Opens `_Flattened.pptx` copy, deletes all shapes from each slide, inserts PNGs as full-slide backgrounds
- Injects invisible 1pt white-on-white title text into the native title placeholder (`ppLayoutTitleOnly`)

## Title vs. slide naming
Two separate name assignments serving different purposes:
- **Shape name** (line 535) — names the invisible title placeholder shape per-slide. Per-slide namespace, no collision risk.
- **Slide name** (line 572) — sets `sld.Name` for PowerPoint slide-sorter and `Slides("name")` lookup. Collision guardrail appends `_2`, `_3` etc. if a duplicate exists.

Both are intentional; see also tooltip/outline/accessibility binding via the native placeholder.

## Debug log
- Written to Desktop: `Articulate_Title_Update_Log_yyyymmdd_hhmmss.txt`
- Timestamped per run — never overwrites previous logs
- Contains per-slide inspection matrix: all shapes checked, all text blocks evaluated (matched/rejected/height-gated), tie-breaker winner shape, before/after titles

## Error handling (lines 628–638)
Emergency rollback on any error:
- Closes `_Flattened.pptx` without saving (`Saved = msoTrue` to suppress prompts)
- Deletes `PPT_Temp_Images_Internal\` temp folder
- Shows "System Failure" message box
