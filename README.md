# EMAS Action Plan — Word Document Generator

Turns the raw CSV export from Cority into one filled-in Word document per
action, using `Emas_Template_FINAL_placeholders.docx` as the template.

You do **not** need to know Python to run this. Follow the steps below.

---

## Running it (every time you have a fresh export)

1. **Open this project in GitHub Codespaces.**
   From the repository page on github.com, click the green **Code** button →
   **Codespaces** tab → **Create codespace on main**. This opens a full
   coding environment in your browser — nothing to install on your own
   computer.

2. **Replace the data file.**
   Drag the new Cority export into the `data/` folder and name it exactly:
   `EMAS_data_exported.csv`
   (overwrite the old one — right-click → Replace, or delete the old one
   first and drag in the new one).

3. **Open a terminal** in Codespaces (Terminal menu → New Terminal, or
   `` Ctrl+` ``).

4. **The first time only**, install the required packages:
   ```bash
   pip install docxtpl pandas
   ```

5. **Run the generator:**
   ```bash
   python generate_all.py
   ```

6. **Check the results**, printed directly in the terminal:
   - `Generated 111 document(s) in generated_docs/` — the filled Word docs,
     one per action, named by Action ID (e.g. `2025-CO2-01.docx`).
   - Any `FAILED` or `SKIPPED` rows are listed — see **Troubleshooting**
     below.

7. **Download the results.** In the Codespaces file explorer, right-click
   the `generated_docs` folder → **Download**, which gives you a zip of
   every generated document.

That's it — steps 2, 3, 5, 7 are the ones you repeat every time.

---

## What comes out

| File/folder | What it is |
|---|---|
| `generated_docs/` | One `.docx` per action, ready to use |
| `generation_log.csv` | One row per action: did it generate OK, get skipped, or fail — and why |
| `milestones_review.csv` | Milestones whose date wasn't in a standard format (e.g. "TBC", "Q1-2029") — worth a quick glance, not necessarily wrong |

---

## Troubleshooting

**A row shows `SKIPPED` in `generation_log.csv`.**
That row had no value in the `EMAS Action ID` column, so there was nothing
to name the file after. Check that row in the CSV — likely a blank/test row.

**A row shows `FAILED`.**
The error message is in the `reason` column of `generation_log.csv`. Most
likely cause: someone typed a stray `{` or `}` character into a text field
in Cority (e.g. inside a description), which can look like broken template
syntax. Find that row's text and remove/escape the stray character.

**A generated document has a blank or odd-looking field.**
Check the same row in the CSV directly — the script only formats what's in
the source data. If the CSV itself is blank or unusual there, that's a
Cority data-entry issue, not a script issue.

**Some milestone dates look like "TBC" or "Q1-2029" instead of a normal date.**
This is expected — see `milestones_review.csv`. The source data isn't
always entered with consistent date formats by every contributor. The
document is still generated; only the date format is non-standard.

---

## If the Cority export format ever changes

This script depends on the CSV having specific column names, and on the
milestone fields following the `Title,Date,Responsible,Percent%` pattern
inside each `Milestone 1`–`4` column (a quirk of how Cority currently
flattens repeating milestones into single cells).

If Cority changes its export structure, or if it becomes possible to export
milestones as their own rows instead of jammed into one cell, the parsing
logic in `generate_all.py` (functions `split_milestone_units` and
`parse_unit`) will need to be revisited. **Worth checking with whoever
administers Cority whether a cleaner export option exists** — that would be
a better long-term fix than parsing around the current format.

## If the Word template needs a new field

1. Open `Emas_Template_FINAL_placeholders.docx`.
2. Add `{{ Your_Field_Name }}` wherever the value should go (letters,
   numbers, underscores only — no spaces or punctuation in the name).
3. In `generate_all.py`, add a matching line inside the `context = {...}`
   block, e.g.:
   ```python
   'Your_Field_Name': clean(row.get('The CSV Column Name')),
   ```
4. Re-run and check the output before trusting it in bulk.

---

## Project structure

```
emas-doc-generator/
├── generate_all.py                          # the script — run this
├── Emas_Template_FINAL_placeholders.docx    # the Word template
├── data/
│   └── EMAS_data_exported.csv               # drop the fresh export here
├── generated_docs/                          # output appears here
├── milestones_review.csv                    # generated on each run
├── generation_log.csv                       # generated on each run
└── README.md                                # this file
```
