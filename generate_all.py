"""
EMAS Action Plan document generator.

Reads the CSV export, parses the flattened milestone fields into clean
records, resolves DG/PG, and generates one populated Word document per
action using Emas_Template_FINAL_placeholders.docx.

Setup (in Codespaces):
  pip install docxtpl pandas

Usage:
  python generate_all.py

Files expected:
  - data/EMAS_data_exported.csv   (drop the fresh Cority export here)
  - Emas_Template_FINAL_placeholders.docx  (do not rename or move)

Output:
  - generated_docs/<ActionID>.docx  (one per action)
  - milestones_review.csv           (flagged freeform-date milestones, for a quick spot-check)
  - generation_log.csv              (one row per action: status + any issues)
"""

import re
import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate

# ---- CONFIG -----------------------------------------------------------------
DATA_PATH = "data/EMAS_data_exported.csv"
TEMPLATE_PATH = "Emas_Template_FINAL_placeholders.docx"
OUTPUT_DIR = Path("generated_docs")
REVIEW_OUTPUT = "milestones_review.csv"
LOG_OUTPUT = "generation_log.csv"
# --------------------------------------------------------------------------------

STRICT_DATE_PATTERNS = [r'\d{1,2}/\d{1,2}/\d{4}', r'\d{4}', r'TBD', r'TBC']
STRICT_DATE_RE = '(' + '|'.join(STRICT_DATE_PATTERNS) + ')'


# ---------------------------------------------------------------------------
# Milestone parsing 
# ---------------------------------------------------------------------------

def split_milestone_units(raw: str):
    if not isinstance(raw, str) or not raw.strip():
        return []
    parts = re.split(r'(?<=%)(?=[A-Za-z])', raw)
    return [p.strip() for p in parts if p.strip()]


def parse_unit(unit: str) -> dict:
    m = re.search(r',?\s*(\d*)\s*%\s*$', unit)
    percent = (m.group(1) + '%') if m else ''
    remainder = unit[:m.start()].rstrip() if m else unit

    if re.search(r',\s*,\s*$', remainder):
        title = re.sub(r',\s*,\s*$', '', remainder).strip()
        return {'title': title, 'date': '', 'responsible': '', 'percent': percent, 'review_needed': False}

    dm = re.search(r',\s*' + STRICT_DATE_RE + r'\s*,', remainder)
    if not dm:
        dm = re.search(r',\s*' + STRICT_DATE_RE + r'\s*$', remainder)

    if dm:
        title = remainder[:dm.start()].strip().rstrip(',').strip()
        date = dm.group(1)
        responsible = remainder[dm.end():].strip().strip(',').strip()
        return {'title': title, 'date': date, 'responsible': responsible, 'percent': percent, 'review_needed': False}

    segments = remainder.rsplit(',', 2)
    if len(segments) == 3:
        title, date, responsible = segments
    elif len(segments) == 2:
        title, date = segments
        responsible = ''
    else:
        title, date, responsible = remainder, '', ''

    return {
        'title': title.strip(), 'date': date.strip(), 'responsible': responsible.strip(),
        'percent': percent, 'review_needed': True,
    }


def parse_action_milestones(row: pd.Series) -> list:
    milestones = []
    for col in ['Milestone 1', 'Milestone 2', 'Milestone 3', 'Milestone 4']:
        for unit in split_milestone_units(row.get(col, '')):
            milestones.append(parse_unit(unit))
    return milestones


def resolve_dg_pg(value: str) -> dict:
    value = (value or '').strip()
    if value.startswith('DG '):
        return {'DGUnit': value, 'PG': ''}
    return {'DGUnit': '', 'PG': value}


def clean(value) -> str:
    """Normalize a raw CSV value: NaN -> '', strip whitespace."""
    if pd.isna(value):
        return ''
    return str(value).strip()


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

def main():
    df = pd.read_csv(DATA_PATH, sep=';', quotechar='"', dtype=str, engine='python', encoding='utf-8-sig')
    OUTPUT_DIR.mkdir(exist_ok=True)

    review_rows = []
    log_rows = []
    generated = 0

    for idx, row in df.iterrows():
        action_id = clean(row.get('EMAS Action ID'))

        if not action_id:
            log_rows.append({'row': idx, 'ActionID': '', 'status': 'SKIPPED', 'reason': 'missing EMAS Action ID'})
            continue

        milestones = parse_action_milestones(row)
        for m in milestones:
            if m['review_needed']:
                review_rows.append({'ActionID': action_id, **m})

        dgpg = resolve_dg_pg(row.get('DG/PG', ''))

        context = {
            'EMAS_Action_ID': action_id,
            'EMAS_Action_Title': clean(row.get('EMAS Action Title')),
            'Project_Manager': clean(row.get('Project Manager')),
            'DGUnit': dgpg['DGUnit'],
            'PG': dgpg['PG'],
            'Action_Description_Objective_and_Scope': clean(row.get('Action Description, Objective, and Scope')),
            'milestones': [
                {'title': m['title'], 'date': m['date'], 'responsible': m['responsible'], 'percent': m['percent']}
                for m in milestones
            ],
            'Human_Resources': clean(row.get('Human Resources')),
            'Budgetary_Resources': clean(row.get('Budgetary Resources')),
            'Expected_Completion_Date': clean(row.get('Expected Completion Date')),
            'Status_Update': clean(row.get('Status Update')),
            'Last_Updated': clean(row.get('Last Updated')),
        }

        try:
            doc = DocxTemplate(TEMPLATE_PATH)
            doc.render(context)
            safe_id = re.sub(r'[^\w\-. ]', '_', action_id)
            out_path = OUTPUT_DIR / f"{safe_id}.docx"
            doc.save(out_path)
            generated += 1
            log_rows.append({'row': idx, 'ActionID': action_id, 'status': 'OK', 'reason': ''})
        except Exception as e:
            log_rows.append({'row': idx, 'ActionID': action_id, 'status': 'FAILED', 'reason': str(e)})

    pd.DataFrame(review_rows).to_csv(REVIEW_OUTPUT, index=False)
    pd.DataFrame(log_rows).to_csv(LOG_OUTPUT, index=False)

    print(f"Generated {generated} document(s) in {OUTPUT_DIR}/")
    print(f"{len(review_rows)} milestone record(s) flagged for a spot-check -> {REVIEW_OUTPUT}")
    failed = [r for r in log_rows if r['status'] == 'FAILED']
    skipped = [r for r in log_rows if r['status'] == 'SKIPPED']
    if failed:
        print(f"\n{len(failed)} row(s) FAILED to generate -- see {LOG_OUTPUT}:")
        for r in failed:
            print(f"  row {r['row']} ({r['ActionID']}): {r['reason']}")
    if skipped:
        print(f"{len(skipped)} row(s) skipped (no Action ID) -- see {LOG_OUTPUT}")


if __name__ == "__main__":
    main()
