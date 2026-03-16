# PPTX Content Extraction & Template-Based Presentation Generator


---

## Overall Approach

The system transforms a **simple input presentation** (plain text grids, basic tables) into a **professionally designed output** by extracting content, analyzing visual templates, selecting the best structural match, and filling the template with extracted data.

```
┌──────────────┐     ┌──────────────┐     ┌──────────────┐     ┌──────────────┐
│   EXTRACT    │ ──▶ │   ANALYZE    │ ──▶ │    SCORE     │ ──▶ │   GENERATE   │
│              │     │              │     │              │     │              │
│ Parse input  │     │ Map template │     │ Match slides │     │ Fill slots   │
│ Detect grid  │     │ layouts &    │     │ Hungarian    │     │ Preserve all │
│ Classify     │     │ placeholder  │     │ algorithm    │     │ template     │
│ roles        │     │ zones        │     │ Dynamic wts  │     │ visuals      │
└──────────────┘     └──────────────┘     └──────────────┘     └──────────────┘
     Part 1               Part 2               Part 3               Part 4
```

### Design Philosophy

The input is **data**. The templates are **visual recipes**. The pipeline's job is to extract the data and render it inside the chosen recipe.

- **Input**: A plain PPTX with text arranged in grids — column headers, row labels, cell content. Minimal styling.
- **Templates**: Rich PPTX files with icons, colored header bars, card layouts, sidebars, accent dividers — containing `{placeholder}` text markers where content should go.
- **Output**: The template's visual design with the input's data filled into every placeholder. All shapes, colors, icons preserved exactly.

### The 4-Part Pipeline

**Part 1 — Content Extraction** (`src/extractor.py`)

Parses the input PPTX and builds a structured representation:

- Extracts every shape's text, position (normalized 0–1 coordinates), font metadata (name, size, bold, color, alignment)
- Combines multi-paragraph textboxes into single blocks (critical for preserving multi-line cell content like "Corporate decks\nPitch presentations\nInfographic design")
- Classifies each block's semantic role (`title`, `table_header`, `content_cell`, `row_label`, `footer`) using position, font size, bold state, shape width, and alignment
- **Grid Detection** (second pass): Groups blocks by y-position, separates left-side labels (x < 0.13) from content columns, identifies header row (bold, top area), and produces:
  ```json
  {
    "n_columns": 5,
    "n_data_rows": 2,
    "headers": ["Presentation Design", "Language Services", ...],
    "data_rows": [["Corporate decks\n...", ...], ["500+ projects/month\n...", ...]],
    "row_labels": ["Point 1", "Point 2"]
  }
  ```
- Generates 384-dimensional sentence embeddings (all-MiniLM-L6-v2) for semantic matching

**Part 2 — Template Analysis** (`src/template_analyzer.py`)

Analyzes each candidate template PPTX:

- Enumerates slide layouts and placeholders
- Maps text zones, shape zones, and table zones with normalized positions
- Extracts color scheme and font theme
- Builds capability profile: supported column counts, layout types, content capacity
- Handles corrupt files and Office temp files (`~$` prefix) gracefully

**Part 3 — Template Selection** (`src/scorer.py`)

See dedicated section below.

**Part 4 — Presentation Generation** (`src/generator.py`)

Opens the winning template and fills it via **slot-fill**:

- Scans every textbox and paragraph for `{placeholder}` patterns
- Replaces `{header_1}` → "Presentation Design", `{row1_col3}` → "Market research\nCompetitive analysis\n...", etc.
- Preserves all template formatting — font name, size, color, bold, alignment inherited from the template's placeholder text
- Preserves all decorative shapes — icons, colored bars, cards, sidebars untouched
- Handles images properly during cross-PPTX merging (binary extraction + re-insertion)

**Two output files**:
| File | Content |
|---|---|
| `final_presentation.pptx` | Best-ranked template filled with content |
| `all_templates_preview.pptx` | ALL templates filled with content, ranked, separated by divider slides showing template name and match score |

---

## How Scoring / Selection Works

The scoring system is **content-aware** and operates at the **individual slide level**.

### Step 1: Content Profiling

Before scoring, the system analyzes the input to understand what it contains:

- How many columns per slide? (detected via shape x-position clustering, table column count, and narrow-block pattern fallback)
- Does the input have tables? Metrics? Multi-column grids?
- What semantic roles are present? (titles, headers, body text, row labels)

This profile drives **dynamic weight computation** — the importance of each scoring signal adapts to the content:

| Content Characteristic | Effect on Weights |
|---|---|
| Multi-column grid (3+ cols) | `column_match` boosted to ~33% |
| Tables present | `table_fit` boosted to ~27% |
| Metric/stat cards detected | `zone_count_fit` boosted |
| Text-heavy body content | `semantic_role_fit` boosted |

Users can also override with custom weights via the interactive CLI.

### Step 2: Slide-Level Scoring (6 Signals)

Every (source slide, template slide) pair is independently scored on 6 signals:

| Signal | Weight (dynamic) | What It Measures | How |
|---|---|---|---|
| `column_match` | 15–33% | Column count match | Exact match = 1.0, ±1 = 0.4, ±2 = 0.15, else 0.0 |
| `zone_count_fit` | 10–20% | Enough text zones for content? | ratio of min/max block counts |
| `table_fit` | 5–27% | Table support match | 1.0 if both have/lack tables, 0.0 if source needs table and template lacks it |
| `spatial_overlap` | 15–20% | Position overlap of content zones | IoU (Intersection over Union) between normalized bounding boxes |
| `semantic_role_fit` | 10–20% | Right zone types at right positions | Title zones at top? Body zones in middle? Footer at bottom? |
| `layout_type_match` | 10% | Same layout classification | Classifies as grid_5, grid_3, table, body, title — exact match = 1.0 |

### Step 3: Hungarian Optimal Assignment

The system builds a full **score matrix** [n_source_slides × n_template_slides] and uses the **Hungarian algorithm** (`scipy.optimize.linear_sum_assignment`) to find the globally optimal 1-to-1 pairing of source slides to template slides.

Template score = average of optimally matched slide-pair scores.

### Step 4: Transparent Reporting

Every scoring decision is logged at slide level:

```
── template_header_sidebar (score=0.8732) ──
  Src 1 (metrics_5col) → Tmpl 1  score=0.873  [col=1.00 | zon=0.78 | tbl=1.00 | spa=0.72 | sem=0.85 | lay=1.00]

── template_minimal_grid (score=0.6521) ──
  Src 1 (metrics_5col) → Tmpl 1  score=0.652  [col=1.00 | zon=0.90 | tbl=1.00 | spa=0.30 | sem=0.50 | lay=0.10]
```

Full score matrices, per-slide breakdowns, and reasoning are saved to `output/scoring_report.json`.

---

## Why the Chosen Template Was Selected

The selection reasoning is generated automatically and saved in `output/scoring_report.json`. A typical explanation:

> Selected 'template_header_sidebar' (score=0.8732). Matched 1 slide pairs:
> Src slide 1 (metrics_5col) → Tmpl slide 1 (score=0.873).
> Strongest signal: column_match=1.00. Weakest signal: layout_type_match=0.10.
> Content has 5-column layouts → column_match weighted highest.
> Content has metric/stats → zone_count boosted.
> Runner-up: 'template_tall_cards' (score=0.8044, margin=0.0688).

The template wins because:
1. Its **column count matches** the input's grid structure (5 columns)
2. It has **enough text zones** for all the extracted content blocks
3. Its **spatial layout** (where text zones are positioned) overlaps well with where the input content sits
4. It provides the **right zone types** — header zones at the top, content zones in the middle, label zones on the side

---

## How to Run

### Prerequisites

```bash
Python 3.9+
pip install -r requirements.txt
```

### Option A: With Sample Files

```bash
# Generate sample input + 5 templates
python generate_samples.py

# Run pipeline (interactive — asks for input, weights, etc.)
python main.py

# Or fully automated
python main.py --auto
```

### Option B: With Your Own Files

```bash
# Place your input PPTX in input/
cp your_presentation.pptx input/

# Place your template PPTX files in templates/
# (Templates must contain {placeholder} text patterns — see Template Format below)
cp your_template_*.pptx templates/

# Run
python main.py
```

### CLI Options

```
python main.py [OPTIONS]

  -i, --input PATH      Path to input PPTX file
  -t, --templates DIR   Template directory (default: templates/)
  -o, --output DIR      Output directory (default: output/)
  -a, --auto            Non-interactive mode (all defaults, no prompts)
```

### Interactive Mode Features

| Feature | Options |
|---|---|
| Template sourcing | Use existing / Add your own / Regenerate samples |
| Input selection | Pick from `input/` folder or enter custom path |
| Scoring weights | Auto (content-derived) or Custom (manual per-signal) |

### Output Files

| File | Description |
|---|---|
| `extracted_content.json` | Full extraction: shapes, positions, fonts, grid structure, embeddings |
| `template_analysis.json` | Layout analysis of every template |
| `scoring_report.json` | Per-slide score matrices, rankings, reasoning |
| `selected_template.txt` | Name of the winning template |
| `final_presentation.pptx` | Best template filled with extracted content |
| `all_templates_preview.pptx` | All templates filled with content for visual comparison |
| `pipeline.log` | Full debug log (timestamped, per-shape detail) |

---

## Template Format

Templates are standard PPTX files with `{placeholder}` text markers:

| Placeholder | Content |
|---|---|
| `{title}` | Slide title |
| `{header_1}` ... `{header_5}` | Column headers |
| `{row1_label}`, `{row2_label}` | Row labels ("Point 1", "Point 2") |
| `{row1_col1}` ... `{row2_col5}` | Cell content |
| `{icon_1}` ... `{icon_5}` | Short text for icon circles |
| `{page}` | Page number |

Everything else in the template (shapes, colors, images, icons) is preserved as-is. The generator only replaces the placeholder text, inheriting whatever formatting the template applies to it.

---

## Libraries Used

| Library | Version | Purpose |
|---|---|---|
| `python-pptx` | ≥0.6.21 | PPTX parsing, shape manipulation, file generation |
| `sentence-transformers` | ≥2.2.0 | Semantic embeddings (all-MiniLM-L6-v2) for content role classification |
| `numpy` | ≥1.24.0 | Vector operations, score matrix computation |
| `scikit-learn` | ≥1.3.0 | Cosine similarity for spatial and semantic matching |
| `scipy` | ≥1.11.0 | Hungarian algorithm (`linear_sum_assignment`) for optimal slide matching |
| `Pillow` | ≥10.0.0 | Image handling during cross-PPTX merging |
| `torch` | ≥2.0.0 | Backend for sentence-transformers model |

---

## Key Features

- **Grid-aware extraction** — Detects column headers, data rows, and row labels from spatial position patterns, not just text content
- **Content-aware dynamic weights** — Scoring adapts to what the input actually contains (tables → table_fit boosted, multi-column → column_match boosted)
- **Slide-level transparency** — Every scoring decision visible per-slide in logs and JSON report
- **Hungarian optimal matching** — Globally optimal slide-to-slide pairing, not greedy
- **Lossless slot-fill generation** — Template visuals preserved perfectly; only placeholder text swapped
- **All-templates preview** — Single PPTX showing every template filled with content for side-by-side comparison
- **Robust file handling** — Skips Office temp files (`~$`), handles corrupt PPTX gracefully, proper image transfer across presentations
- **Interactive + automated modes** — Full CLI with `--auto` flag for scripted usage
- **Comprehensive logging** — Dual output (console INFO + file DEBUG) with full extraction and scoring detail

---

## Project Structure

```
ez-pptx-generator/
├── input/                          ← Input presentations
├── templates/                      ← Visual template files with {placeholders}
├── output/                         ← All generated outputs + logs
├── src/
│   ├── __init__.py
│   ├── logger.py                   ← Centralized dual logging
│   ├── utils.py                    ← EMU conversion, role classification, spatial helpers
│   ├── extractor.py                ← Part 1: Content + grid extraction
│   ├── template_analyzer.py        ← Part 2: Template layout analysis
│   ├── scorer.py                   ← Part 3: Dynamic scoring + Hungarian matching
│   └── generator.py                ← Part 4: Slot-fill generation + preview builder
├── generate_samples.py             ← Sample input + 5 template generator
├── main.py                         ← CLI entry point
├── requirements.txt
└── README.md
```
