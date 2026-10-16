# NovaRetail Capstone: Automated Executive QBR Deck

This project analyzes NovaRetail's three-sheet business dataset and populates the supplied five-slide PowerPoint template using Python and `python-pptx`.

## Project contents

- `generate_deck.py`: complete analysis and deck-generation workflow
- `novaretail_dataset.xlsx`: supplied source dataset
- `novaretail_template.pptx`: supplied fixed-layout template
- `icons/`: supplied icon assets
- `NovaRetail_QBR_Final.pptx`: generated executive presentation
- `requirements.txt`: Python dependencies

## Setup

```bash
python -m venv .venv
```

Activate the environment:

**Windows PowerShell**

```powershell
.venv\Scripts\Activate.ps1
```

**macOS/Linux**

```bash
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

## Generate the presentation

```bash
python generate_deck.py
```

The command creates `NovaRetail_QBR_Final.pptx` in the project folder.

## Execution logic

1. Load `Regional_Performance`, `Customer_Experience`, and `Operational_Risk` with `openpyxl`.
2. Index each record by quarter and region.
3. Calculate latest-quarter revenue, quarter-over-quarter growth, unit-weighted NPS, and region-level trends.
4. Open the supplied PowerPoint template.
5. Locate all writable areas by `shape.name`.
6. Replace KPI, headline, summary, insight, recommendation, and closing-statement text.
7. Remove dashed chart placeholders and instruction labels.
8. Add native PowerPoint line charts at the exact placeholder coordinates and dimensions.
9. Replace icon placeholders with the supplied theme-relevant artwork.
10. Save the completed five-slide presentation.

## Key dataset-driven narrative

- APAC is the strongest growth region and leads customer experience.
- North America remains the largest revenue market.
- Latin America shows declining revenue and worsening customer and operational indicators.
- APAC warehouse utilization and stockouts indicate a capacity constraint that could limit growth.

## Regeneration notes

- Keep the source filenames unchanged, or update the path constants at the top of `generate_deck.py`.
- Keep the template shape names unchanged because the script uses those names to locate content regions.
- If the dataset is refreshed with the same sheets and columns, rerun the script to regenerate the deck.
