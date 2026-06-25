# AGENTS.md — PPT VBA

> **Do not commit or push changes without explicit user request.**

## What this is

Single VBA PowerPoint module in `slide_flatten_rename.txt`. Contains two workflows:

| Alt+F8 entry point | What it does |
|---|---|
| `PPTTools_CheckSlides` | Inspect all slides for issues — **no changes** to file |
| `PPTTools_Flatten` | Flatten presentation (assumes checks have passed) |

**Recommended workflow:** run `PPTTools_CheckSlides` → fix any problems → run `PPTTools_Flatten`.

Paste into VBA editor (Alt+F11 → Insert → Module), then Alt+F8 and double-click the macro name.

## Preconditions
- Operates on `ActivePresentation` — the presentation open and active when run
- **Must be saved before running** — both macros exit with error otherwise
- `PPTTools_Flatten` creates `_Flattened.pptx` copy in the same directory (e.g., `Deck.pptx` → `Deck_Flattened.pptx`)

## Skip guardrails
Slides are skipped (no title extraction, no flatten, no rename) when any condition is met. All skips are tagged with a `skipReason` string, logged per-slide and grouped in the **SKIPPED SLIDES SUMMARY**. First-matched reason wins (order below):

| # | Reason | Trigger |
|---|--------|---------|
| 1 | `"Keyword match"` | Any visible shape contains "thank you", "thanks", "Q&A", or "??" (case-insensitive) |
| 2 | `"Non-standard slide name"` | Slide `.Name` is not `"Slide" & slideNum` AND does not start with `"Slide"` |
| 3 | `"Multiple shapes tie at separator depth N"` | Two+ shapes have the same highest separator count (tie-breaker) |
| 4 | `"No valid title candidate found"` | No shape's text matches the regex in the top half of the slide |

## Title override
Before depth-based selection, a pre-scan checks all visible shapes for priority keywords (case-insensitive). First match wins immediately — shape's full text is used as the title as-is, bypassing regex and separator counting entirely.

**Trigger keywords:** "course objective", "learning outcomes", "table of content", "学习目标", "学习成果", "内容目录", "培训目标", "目录", "学习收获"

Override slides get `slideIndices = ""` so they are excluded from numerical index repair (Step 2, Case 2), but still participate in contiguous duplicate suffixing (Step 2, Case 1).

## Shape selection algorithm
**Scope:** Only visible shapes with text frames where `Top < SlideHeight × 0.5`.

**Selection:**
1. Split each shape's text into blocks (multi-line shapes evaluated per-line by `vbCr`)
2. Test each block against regex: `^(\d+(?:[\.\(/\\\-_\??]\d+)*)\s*(.*)`
3. Count separator characters (`. ( / \ - _ ?`) in the matched index portion
4. **Winner = shape with highest separator count** (deepest nesting, e.g. `1.2.3` beats `1.2`)
5. **Tie-breaker**: if multiple shapes tie at max separator count, the slide is **skipped entirely** (not flattened). The skip is tagged with reason `"Multiple shapes tie at separator depth N"`.
6. **No match**: if no shape's text passes the regex, the slide is **skipped entirely** (tagged `"No valid title candidate found"`).

Note: the separator character class allows **mixing** (e.g. `1.2/3` matches). The space between index and title is optional (`\s*`), so `1.2.3Title` also matches.

## PPTTools_CheckSlides — all 11 checks

### Title/index error analysis
| # | Check | What it detects |
|---|-------|-----------------|
| 1 | Keyword match skip | Guardrail — slide skipped from all checks |
| 2 | Non-standard slide name | Guardrail — slide skipped from all checks |
| 3 | Tie-breaker collision | Guardrail — slide skipped from all checks |
| 4 | No valid title candidate | Guardrail — slide skipped from all checks |
| 5 | Contiguous duplicate titles | Same title on consecutive slides — flagged fixable |
| 6 | Non-contiguous duplicate titles | Same title on non-consecutive slides — flagged non-fixable |
| 7 | Conflicting numerical indices | Same index on adjacent slides — flagged fixable |

### Notes checks
| # | Check | What it detects |
|---|-------|-----------------|
| 8 | Missing/empty notes | Slide has no `NotesSlide` or notes text is blank |
| 9 | Notes language mismatch | `DetectLanguage(notes)` ≠ reference language (per-slide title if available, else `pptLang`) — flagged with ±2 sentence context |

### Translation gap checks
| # | Check | What it detects |
|---|-------|-----------------|
| 10 | Shape text boxes — missed translation | Any visible shape whose language differs from the presentation-level language |
| 11 | Notes — missed translation | Any sentence in the notes whose language differs from the presentation-level language |

Language detection uses CJK Unicode range (`\u4E00`–`\u9FFF`) vs Latin letters ratio. Returns `"ZH"`, `"EN"`, or `"MIXED"`.

## PPTTools_Flatten
- Exports each slide as PNG at 2× resolution into `PPT_Temp_Images_Internal\` (deleted after run)
- Opens `_Flattened.pptx` copy, deletes all shapes from each slide, inserts PNGs as full-slide backgrounds
- Injects invisible 1pt white-on-white title text into the native title placeholder (`ppLayoutTitleOnly`)
- Does NOT run error analysis or notes checks (assumes user fixed issues found by `PPTTools_CheckSlides`)

## Title vs. slide naming
Two separate name assignments serving different purposes:
- **Shape name** — names the invisible title placeholder shape per-slide. Per-slide namespace, no collision risk.
- **Slide name** — sets `sld.Name` for PowerPoint slide-sorter and `Slides("name")` lookup. Collision guardrail appends `_2`, `_3` etc. if a duplicate exists.

Both are intentional; see also tooltip/outline/accessibility binding via the native placeholder.

## Debug log
- Written to Desktop: `Articulate_Title_Update_Log_yyyymmdd_hhmmss.txt`
- Timestamped per run — never overwrites previous logs
- **Check log** (`PPTTools_CheckSlides`): contains error analysis log, notes audit, translation gap check, skipped slides summary, and full per-slide debug matrix
- **Flatten log** (`PPTTools_Flatten`): brief — total slides, flattened count, skipped count

## Error handling
**Check workflow** (`RunAllChecks`): simple error trapping — no file changes to roll back.

**Flatten workflow** (`FlattenPresentation`): emergency rollback on any error:
- Closes `_Flattened.pptx` without saving (`Saved = msoTrue` to suppress prompts)
- Deletes `PPT_Temp_Images_Internal\` temp folder
- Shows "System Failure" message box