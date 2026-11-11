# Suitcase Rex

A searchable, chronological web database of every track on Guided By Voices' four
*Suitcase* box sets (400 tracks), plus 7 *Briefcase* bonus tracks, built from the
`Suitcase_Rex.xlsx` spreadsheet.

Open `index.html` in a browser. It is a single self-contained page (data embedded),
so it also works when served by GitHub Pages from the repository root.

## Features

- Chronological view grouped by recording year; also box-set order, band name or title
- Year histogram, colored by Suitcase; click a bar to filter to that year (or "n.d." for undated)
- Full-text search across titles, band names, notes and credits
- Filter by Suitcase/Briefcase, songwriter, band (click a band name) and "dated only"

## Data

| File | Contents |
| --- | --- |
| `data/Suitcase_Rex.xlsx` | Source spreadsheet |
| `data/tracks.json` | Cleaned records plus Suitcase titles |
| `data/tracks.csv` | The same records as CSV |

Cleaning rules applied by `scripts/build.py`:

- Blank rows and the empty "remakes" heading are dropped.
- `Disc NN, track NN` is mapped to Suitcase (discs 1–4 → Suitcase 1, 5–8 → Suitcase 2, …)
  and a per-box disc number; the original sheet value is kept as `source_location`.
- Years like `1995?` and `85-'86?` are kept as approximate (`year_approx: true`);
  `?` and stray values are treated as undated.
- Songwriting credits are split on `-` into individual writers.
- The Excel escape `_x0092_` is turned back into an apostrophe.

## Rebuilding

After editing the spreadsheet or `src/index.template.html`:

```bash
pip install openpyxl
python3 scripts/build.py
```
