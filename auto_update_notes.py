#!/usr/bin/env python3
"""
Tao Te Ching Book Club Notes Auto-Updater
Runs every Thursday & Friday at noon (via launchd).
- Downloads latest .docx from Google Drive (Chinese & English folders)
- Replaces names with role labels
- Injects formatted HTML into index.html via update_field()
- Git commits and pushes to deploy GitHub Pages
"""

import io, json, os, re, subprocess, sys, tempfile
from datetime import datetime
from pathlib import Path

# ── Config ──────────────────────────────────────────────────────────────────
REPO_DIR       = Path(__file__).parent
TOKEN_PATH     = Path.home() / ".google-mcp/tokens/josh.json"
CREDS_PATH     = Path.home() / ".google-mcp/credentials.json"
GH_TOKEN       = os.environ.get("GH_TOKEN", "")

ZH_FOLDER_ID   = "1SDW9Baa-D1cU6FMy5Ufo9KRQX_INi7bI"
EN_FOLDER_ID   = "1SDhO0VzK9eYIaTI6JUzwTvVMJCF49dWy"

# ── Google Auth ──────────────────────────────────────────────────────────────
def build_drive_service():
    import warnings
    warnings.filterwarnings("ignore")
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    t = json.loads(TOKEN_PATH.read_text())
    c = json.loads(CREDS_PATH.read_text())["installed"]
    creds = Credentials(
        token=None,
        refresh_token=t["refresh_token"],
        client_id=t.get("client_id") or c["client_id"],
        client_secret=t.get("client_secret") or c["client_secret"],
        token_uri="https://oauth2.googleapis.com/token",
    )
    creds.refresh(Request())
    return build("drive", "v3", credentials=creds)

# ── Drive helpers ────────────────────────────────────────────────────────────
def latest_docx(service, folder_id):
    """Return (file_id, name, modifiedTime) of the most recently modified .docx."""
    q = (f"'{folder_id}' in parents and trashed=false "
         f"and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'")
    res = service.files().list(
        q=q, orderBy="modifiedTime desc", pageSize=5,
        fields="files(id,name,modifiedTime)"
    ).execute()
    files = res.get("files", [])
    if not files:
        return None
    return files[0]

def download_docx(service, file_id):
    from googleapiclient.http import MediaIoBaseDownload
    buf = io.BytesIO()
    req = service.files().get_media(fileId=file_id)
    dl  = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    return buf

# ── docx → plain text ────────────────────────────────────────────────────────
def docx_to_lines(buf):
    """Extract paragraphs from .docx preserving heading styles."""
    import docx
    doc = docx.Document(buf)
    lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            lines.append("")
            continue
        style = para.style.name if para.style else ""
        is_bold = all(run.bold for run in para.runs if run.text.strip())
        # Prefix headings with ## so the HTML formatter recognises them
        if "Heading 1" in style or "Title" in style:
            lines.append(f"## {text}")
        elif "Heading 2" in style or "Heading 3" in style or is_bold:
            lines.append(f"### {text}")
        else:
            lines.append(text)
    return lines

# ── Name replacement ─────────────────────────────────────────────────────────
# Matches 2-4 CJK characters (names) followed by ： or :
_ZH_NAME_COLON = re.compile(
    r'^([一-鿿]{2,4})\s*[：:]',
)
# English name pattern: "FirstName LastName:" or "FirstName:"
_EN_NAME_COLON = re.compile(
    r'^([A-Z][a-z]+(?: [A-Z][a-z]+)?)\s*[：:]',
)

# Heuristic: track first speaker per paragraph block = 點傳師/Master
# subsequent distinct speakers = 同修/Participant
# Also replaces inline name mentions in body lines.
_INLINE_NAME_RE = re.compile(r'[一-鿿]{2,4}')

def replace_names_zh(lines):
    seen = {}   # name -> assigned role
    out  = []
    for line in lines:
        m = _ZH_NAME_COLON.match(line)
        if m:
            name = m.group(1)
            if name not in seen:
                role = "點傳師" if not seen else "同修"
                seen[name] = role
            label = seen[name]
            rest  = line[m.end():]
            out.append(f"{label}：{rest.lstrip()}")
        else:
            # Replace inline occurrences of known speaker names
            processed = line
            for name, role in sorted(seen.items(), key=lambda x: -len(x[0])):
                processed = processed.replace(name, role)
            out.append(processed)
    return out

def replace_names_en(lines):
    seen = {}
    out  = []
    for line in lines:
        m = _EN_NAME_COLON.match(line)
        if m:
            name = m.group(1)
            if name not in seen:
                role = "Master" if not seen else "Participant"
                seen[name] = role
            label = seen[name]
            rest  = line[m.end():]
            out.append(f"{label}: {rest.lstrip()}")
        else:
            out.append(line)
    return out

# ── HTML formatter (mirrors sync_ttc.py logic) ───────────────────────────────
_ZH_H_EXACT = frozenset({
    '概覽','概述','核心概念或理論','核心概念與理論','核心概念',
    '提出的重要問題','重要提問','重要問題',
    '重點摘要與學習目標','重點摘要','學習目標',
    '相關問答','問答','補充資源','補充資源和後續閱讀',
    '可行的後續步驟/作業','行動步驟/作業','行動步驟',
    '後續步驟','作業','總結','知道・道學讀書會',
})
_ZH_TOPIC_RE  = re.compile(r'^主題[一二三四五六七八九十百\d]+[：:]')
_ZH_Q_RE      = re.compile(r'^(點傳師|講師)[：:]|^問[（(][^）)]*[）)][：:]?|^問\s*[：:]')
_ZH_A_RE      = re.compile(r'^(同修[甲乙丙丁]?|前賢[甲乙丙丁]?|學員|線上學員)[：:]|^答[（(][^）)]*[）)][：:]?|^答\s*[：:]')
_EN_H_EXACT   = frozenset({
    'Overview','Key Takeaways','Discussion Highlights',
    'Key Concepts or Theories','Important Questions Raised',
    'Key Points','Key Insights','The Discussion','Conclusion',
})
_EN_CLASS_RE  = re.compile(r'^Class \d+[：:]', re.I)
_EN_TOPIC_RE  = re.compile(r'^Topic \d+[：:]', re.I)
_EN_Q_RE      = re.compile(r'^(Question|Q)\s*[：:]', re.I)
_EN_A_RE      = re.compile(r'^(Answer|A)\s*[：:]', re.I)
_EN_MASTER_RE = re.compile(r'^Master\b[^：:]*[：:]')
_EN_PART_RE   = re.compile(r'^Participant\s*[A-Z]?\s*[：:]')

def _kind(s):
    if not s:                             return 'blank', s
    if s.startswith('## '):              return 'h',     s[3:]
    if s.startswith('### '):             return 'sh',    s[4:]
    if s in _ZH_H_EXACT:                return 'h',     s
    if _ZH_TOPIC_RE.match(s):           return 'h',     s
    if s in _EN_H_EXACT:                return 'h',     s
    if _EN_CLASS_RE.match(s):           return 'h',     s
    if _EN_TOPIC_RE.match(s):           return 'h',     s
    if _ZH_Q_RE.match(s):              return 'q',     s
    if _ZH_A_RE.match(s):              return 'a',     s
    if _EN_Q_RE.match(s):              return 'q',     s
    if _EN_A_RE.match(s):              return 'a',     s
    if _EN_MASTER_RE.match(s):         return 'q',     s
    if _EN_PART_RE.match(s):           return 'a',     s
    if s.startswith('- ') or s.startswith('• '):
        return 'li', s[2:]
    return 'body', s

def lines_to_html(lines):
    out, buf, in_list = [], [], False

    def flush():
        nonlocal in_list
        if in_list:
            out.append('</ul>')
            in_list = False
        if buf:
            out.append('<div class="disc-p">' + '<br>'.join(buf) + '</div>')
            buf.clear()

    for raw in lines:
        s = raw.strip().replace('`', "'")
        kind, content = _kind(s)
        if kind == 'blank':     flush()
        elif kind == 'h':       flush(); out.append(f'<div class="disc-h">{content}</div>')
        elif kind == 'sh':      flush(); out.append(f'<div class="disc-sh">{content}</div>')
        elif kind == 'q':       flush(); out.append(f'<div class="disc-q">{content}</div>')
        elif kind == 'a':       flush(); out.append(f'<div class="disc-a">{content}</div>')
        elif kind == 'li':
            if buf: flush()
            if not in_list:
                out.append('<ul class="disc-ul">')
                in_list = True
            out.append(f'<li>{content}</li>')
        else:
            if in_list: flush()
            buf.append(content)
    flush()
    return ''.join(out)

# ── Chapter detection ────────────────────────────────────────────────────────
_ZH_NUMS = [
    "一","二","三","四","五","六","七","八","九","十",
    "十一","十二","十三","十四","十五","十六","十七","十八","十九","二十",
    "二十一","二十二","二十三","二十四","二十五","二十六","二十七","二十八","二十九","三十",
    "三十一","三十二","三十三","三十四","三十五","三十六","三十七","三十八","三十九","四十",
    "四十一","四十二","四十三","四十四","四十五","四十六","四十七","四十八","四十九","五十",
    "五十一","五十二","五十三","五十四","五十五","五十六","五十七","五十八","五十九","六十",
    "六十一","六十二","六十三","六十四","六十五","六十六","六十七","六十八","六十九","七十",
    "七十一","七十二","七十三","七十四","七十五","七十六","七十七","七十八","七十九","八十","八十一",
]
_ORDINAL = sorted({f"第{n}": i+1 for i,n in enumerate(_ZH_NUMS)}.items(), key=lambda x: -len(x[0]))

def detect_chapters(filename, lines):
    """Return a list of chapter numbers detected from filename or document content."""
    chapters = []

    # Check filename for Chinese ordinals
    for prefix, num in _ORDINAL:
        if prefix in filename and num not in chapters:
            chapters.append(num)

    # Check filename for patterns like "Chapter6_7", "ch6-7", "Chapter6&7"
    for m in re.finditer(r'(?:ch|chapter|第)\s*(\d+)(?:[_&\-](\d+))?', filename, re.I):
        num = int(m.group(1))
        if 1 <= num <= 81 and num not in chapters:
            chapters.append(num)
        if m.group(2):
            num2 = int(m.group(2))
            if 1 <= num2 <= 81 and num2 not in chapters:
                chapters.append(num2)

    if chapters:
        return sorted(chapters)

    # Scan full document content
    for line in lines:
        for prefix, num in _ORDINAL:
            if prefix in line and num not in chapters:
                chapters.append(num)
        for m in re.finditer(r'(?:chapter|ch)\s*(\d+)', line, re.I):
            num = int(m.group(1))
            if 1 <= num <= 81 and num not in chapters:
                chapters.append(num)

    return sorted(chapters) if chapters else []

# ── index.html patching ──────────────────────────────────────────────────────
def patch_html(html, ch_num, field, new_val):
    pat = re.compile(
        r'(num:\s*' + str(ch_num) + r'\b.*?' + re.escape(field) + r':\s*`)([^`]*)(`)',
        re.DOTALL
    )
    new_html, n = pat.subn(lambda m: m.group(1) + new_val + m.group(3), html)
    return new_html, n

# ── Git push ─────────────────────────────────────────────────────────────────
def git_push(message):
    env = os.environ.copy()
    token = GH_TOKEN
    if token:
        # Configure credential helper via env
        remote_url = f"https://{token}@github.com/joshlinmoments/taoteching-web.git"
        subprocess.run(["git", "remote", "set-url", "origin", remote_url],
                       cwd=REPO_DIR, check=True, capture_output=True)

    subprocess.run(["git", "add", "index.html"], cwd=REPO_DIR, check=True)
    result = subprocess.run(
        ["git", "diff", "--cached", "--quiet"], cwd=REPO_DIR
    )
    if result.returncode == 0:
        print("Nothing to commit.")
        return False

    subprocess.run(
        ["git", "commit", "-m", message + "\n\nCo-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"],
        cwd=REPO_DIR, check=True
    )
    subprocess.run(["git", "push", "origin", "main"], cwd=REPO_DIR, check=True)
    return True

# ── Network retry ────────────────────────────────────────────────────────────
def build_drive_service_with_retry(log, max_wait_minutes=480, interval_minutes=5):
    """Retry Google auth until network is available, up to max_wait_minutes."""
    import time
    attempts = max_wait_minutes // interval_minutes
    for i in range(attempts):
        try:
            return build_drive_service()
        except Exception as e:
            if i == 0:
                log(f"  Network unavailable ({e.__class__.__name__}), will retry every {interval_minutes} min (up to {max_wait_minutes} min)...")
            else:
                log(f"  Still waiting for network... ({(i+1)*interval_minutes}/{max_wait_minutes} min)")
            time.sleep(interval_minutes * 60)
    raise RuntimeError(f"Network unavailable after {max_wait_minutes} minutes, giving up.")

# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    log = lambda msg: print(f"[{datetime.now():%Y-%m-%d %H:%M}] {msg}")

    log("=== Tao Te Ching Notes Auto-Update ===")
    service = build_drive_service_with_retry(log)

    html_path = REPO_DIR / "index.html"
    html      = html_path.read_text(encoding="utf-8")
    changes   = []

    tasks = [
        (ZH_FOLDER_ID, "zh", "discussion"),
        (EN_FOLDER_ID, "en", "en_discussion"),
    ]

    for folder_id, lang, field in tasks:
        log(f"Checking {lang} folder...")
        meta = latest_docx(service, folder_id)
        if not meta:
            log(f"  No .docx found in {lang} folder, skipping.")
            continue

        log(f"  Found: {meta['name']} (modified {meta['modifiedTime'][:10]})")
        buf   = download_docx(service, meta["id"])
        lines = docx_to_lines(buf)

        chs = detect_chapters(meta["name"], lines)
        if not chs:
            log(f"  Cannot detect chapter number from '{meta['name']}', skipping.")
            continue
        log(f"  Detected chapter(s): {chs}")

        if lang == "zh":
            lines = replace_names_zh(lines)
        else:
            lines = replace_names_en(lines)

        html_content = lines_to_html(lines)

        for ch in chs:
            new_html, n = patch_html(html, ch, field, html_content)
            if n == 0:
                log(f"  WARNING: pattern for ch{ch}.{field} not found in HTML.")
                continue
            if new_html == html:
                log(f"  ch{ch}.{field} unchanged, skipping.")
                continue
            html = new_html
            changes.append(f"ch{ch}.{field}")
            log(f"  Updated ch{ch}.{field}")

    if not changes:
        log("No updates needed.")
        return

    html_path.write_text(html, encoding="utf-8")
    log(f"Saved index.html ({', '.join(changes)})")

    pushed = git_push(f"Auto-update notes: {', '.join(changes)}")
    if pushed:
        log("Pushed to GitHub Pages. Site will update in ~1 min.")
        log("https://joshlinmoments.github.io/taoteching-web/")


if __name__ == "__main__":
    main()
