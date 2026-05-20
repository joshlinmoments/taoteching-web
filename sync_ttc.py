#!/usr/bin/env python3
"""Sync Tao Te Ching Google Sheet data to index.html.
Usage: python3 sync_ttc.py [--dry-run]
"""

import csv, io, re, sys, urllib.request
from pathlib import Path

SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1vo8Kt3NYk_zOw4Jny7FyMTiUW7OTGib8bobqs9Pz-wg"
    "/export?format=csv&gid=2140898367"
)
HTML_PATH = Path(__file__).parent / "index.html"

COL_ENG_TITLE  = 1
COL_ENGLISH    = 2
COL_ZH_NAME    = 6
COL_CHINESE    = 7
COL_COMMENTARY = 8

# Chinese ordinals for chapters 1-81, longest-first to prevent prefix clashes
# e.g. "第二十" must be checked before "第二"
_NUMS = [
    "一","二","三","四","五","六","七","八","九","十",
    "十一","十二","十三","十四","十五","十六","十七","十八","十九","二十",
    "二十一","二十二","二十三","二十四","二十五","二十六","二十七","二十八","二十九","三十",
    "三十一","三十二","三十三","三十四","三十五","三十六","三十七","三十八","三十九","四十",
    "四十一","四十二","四十三","四十四","四十五","四十六","四十七","四十八","四十九","五十",
    "五十一","五十二","五十三","五十四","五十五","五十六","五十七","五十八","五十九","六十",
    "六十一","六十二","六十三","六十四","六十五","六十六","六十七","六十八","六十九","七十",
    "七十一","七十二","七十三","七十四","七十五","七十六","七十七","七十八","七十九","八十","八十一",
]
ORDINAL_MAP = sorted(
    {f"第{n}": i + 1 for i, n in enumerate(_NUMS)}.items(),
    key=lambda x: -len(x[0])
)


def zh_name_to_ch(zh_name):
    for suffix, num in ORDINAL_MAP:
        if zh_name.endswith(suffix):
            return num
    return None


def strip_ordinal(zh_name):
    for suffix, _ in ORDINAL_MAP:
        if zh_name.endswith(suffix):
            return zh_name[: -len(suffix)]
    return zh_name


# CJK Radicals Supplement → standard CJK equivalents
# These variant forms can appear in spreadsheet data and look subtly different
_CJK_RADICAL_FIX = str.maketrans({
    '⺠': '民',  # ⺠ → 民
    '⺾': '筵',  # ⽾ → 竵 (竹)
    '⽎': '民',  # 民 kangxi radical form
})

def normalize_cjk(text):
    return text.translate(_CJK_RADICAL_FIX)

def to_br(text):
    return (normalize_cjk(text)
            .replace('\r\n', '\n').replace('\r', '\n')
            .replace('\n', '<br>').replace('`', "'"))


def fetch_csv():
    print("  Connecting to Google Sheets...")
    with urllib.request.urlopen(SHEET_URL, timeout=20) as r:
        return r.read().decode('utf-8')


def parse_sheet(csv_text):
    reader = csv.reader(io.StringIO(csv_text))
    rows = list(reader)

    chapter_rows = {}  # ch_num -> list of rows
    pending_ch = None  # track current chapter for rows with empty 中文篇名

    for row in rows[1:]:  # skip header
        if len(row) <= COL_ZH_NAME:
            continue

        zh_name = row[COL_ZH_NAME].strip()

        if zh_name:
            ch_num = zh_name_to_ch(zh_name)
            if not ch_num:
                continue  # intro / non-chapter row
            pending_ch = ch_num
        else:
            ch_num = pending_ch  # continuation of split chapter

        if ch_num is None:
            continue

        chapter_rows.setdefault(ch_num, []).append(row)

    result = {}
    for ch_num, rlist in chapter_rows.items():

        def first(col):
            return next((r[col] for r in rlist if len(r) > col and r[col].strip()), '')

        def concat(col, sep='\n'):
            parts = [r[col] for r in rlist if len(r) > col and r[col].strip()]
            return sep.join(parts)

        zh_name_first = first(COL_ZH_NAME)
        result[ch_num] = {
            'classical':   strip_ordinal(zh_name_first) if zh_name_first else '',
            'title':       to_br(first(COL_ENG_TITLE)),
            'english':     to_br(first(COL_ENGLISH)),
            'chinese':     to_br(concat(COL_CHINESE, '\n')),
            'commentary':  to_br(concat(COL_COMMENTARY, '\n\n')),
        }

    return result


def update_field(html, ch_num, field, new_val):
    pat = re.compile(
        r'(num:\s*' + str(ch_num) + r'\b.*?' + re.escape(field) + r':\s*`)([^`]*)(`)',
        re.DOTALL
    )
    def repl(m): return m.group(1) + new_val + m.group(3)
    new_html, n = pat.subn(repl, html)
    return new_html, n


def main():
    dry_run = '--dry-run' in sys.argv

    csv_text = fetch_csv()
    sheet = parse_sheet(csv_text)
    print(f"  Parsed {len(sheet)} chapters from spreadsheet")

    html = HTML_PATH.read_text(encoding='utf-8')
    changes = []

    for ch_num in sorted(sheet):
        data = sheet[ch_num]
        for field in ('title', 'classical', 'english', 'chinese', 'commentary'):
            val = data.get(field, '')
            if not val:
                continue
            new_html, n = update_field(html, ch_num, field, val)
            if n == 0:
                print(f"  WARNING: ch{ch_num}.{field} — pattern not found in HTML, skipped")
            elif new_html != html:
                changes.append(f"ch{ch_num}.{field}")
                html = new_html

    if not changes:
        print("Already up to date — nothing changed.")
        return 0

    print(f"\n  Updated fields: {', '.join(changes)}")

    if dry_run:
        print("[dry-run] File not written; git not touched.")
        return 0

    HTML_PATH.write_text(html, encoding='utf-8')
    print(f"  Saved {HTML_PATH.name}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
