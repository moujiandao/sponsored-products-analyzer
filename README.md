# Sponsored Products Analyzer

Analyzes Amazon Sponsored Products search term reports to surface actionable keyword insights.

## What it does

- Flags **negative keyword candidates** — terms with clicks but zero orders (wasted spend)
- Identifies **exact match candidates** — high-CVR terms worth bidding on directly
- Highlights **high ACoS terms** (40–60%+) that need bid review
- Surfaces **top performers** with ACoS under 20%
- Exports everything to a structured Excel workbook

## Usage

Place your search term report in `data/`, then run:

```bash
python parsers/search_terms.py
```

Output is written to `output/search_term_analysis.xlsx`.

## Configuration

Edit the thresholds at the top of `parsers/search_terms.py`:

| Variable | Default | Description |
|---|---|---|
| `MIN_CLICKS_FOR_NEGATIVE` | 8 | Clicks needed before flagging a zero-order term |
| `TARGET_ACOS` | 0.40 | ACoS threshold for underperforming terms |
| `MIN_CVR_FOR_EXACT` | 0.10 | Minimum CVR to suggest exact match |

## Input

Amazon Sponsored Products Search Term report (`.xlsx`), 60-day window recommended.
