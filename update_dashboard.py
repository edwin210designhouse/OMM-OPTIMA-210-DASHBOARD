"""
OMM Optima Container Dashboard — Auto-update script
Double-click this file to update the dashboard from Excel and push to GitHub.

Steps it runs automatically:
  1. Reads container schedule + unit matrix from Excel
  2. Injects fresh data into OMM_Optima_Dashboard.html
  3. Pushes to GitHub → live site updates in ~60 seconds
"""

import os
import re
import json
import subprocess
from datetime import datetime
import openpyxl

# ── CONFIG ──────────────────────────────────────────────────────────────────
SITE_DIR       = r"C:\Users\Edwin\OneDrive - Studio Snaidero Chicago\OMM_Optima Website"
EXCEL_FILE     = os.path.join(SITE_DIR, "Output_Optima_Container Matrix & Schedule .xlsx")
TEMPLATE_FILE  = os.path.join(SITE_DIR, "OMM_Optima_Template.html")   # design lives here — never auto-edited
HTML_FILE      = os.path.join(SITE_DIR, "OMM_Optima_Dashboard.html")  # output — always regenerated from template

AUTO_GIT = True   # Set to False to skip the git push

SH_CTN  = "OMM_Overall Container Matrix"
SH_UNIT = "OMM_Unit Matrix"

GITHUB_TOKEN  = ""   # ← paste your GitHub Personal Access Token here (between the quotes)
GITHUB_REMOTE = "https://github.com/edwin210designhouse/OMM-OPTIMA-210-DASHBOARD.git"
# ────────────────────────────────────────────────────────────────────────────

def fmt_date(val):
    if val is None:
        return "—"
    if isinstance(val, datetime):
        months = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()
        return f"{months[val.month - 1]} {val.day}, {val.year}"
    s = str(val).strip()
    return s if s else "—"

def read_containers(ws):
    containers = []
    for row in ws.iter_rows(min_row=11, values_only=True):
        if len(row) < 32:
            continue
        num = row[4]
        if not isinstance(num, int):
            continue
        containers.append({
            "num":          num,
            "status":       str(row[3]).strip() if row[3] else "PROJ",
            "week":         str(row[6]).strip()  if row[6] else "--",
            "floors":       str(row[7]).strip()  if row[7] else "--",
            "contCode":     str(row[10]).strip() if row[10] else "--",
            "contSeal":     str(row[11]).strip() if row[11] else "--",
            "loadDate":     fmt_date(row[13]),
            "shipDate":     fmt_date(row[16]),
            "vessel1":      str(row[17]).strip() if row[17] else "TBD",
            "vessel2":      str(row[19]).strip() if row[19] else "TBD",
            "etaLAX":       fmt_date(row[21]),
            "etaPhoenix":   fmt_date(row[27]),
            "deliveryDate": fmt_date(row[31]),
            "installDate":  fmt_date(row[39]) if len(row) > 39 else "--",
        })
    return sorted(containers, key=lambda x: x["num"])

COMP_COLS = [
    ("Foyer",     7,  8),   ("Den",       9,  10),  ("Hall",     11, 12),
    ("Utility",   13, 14),  ("Closet",    15, 16),  ("Bev Ctr",  17, 18),
    ("Laundry",   19, 20),  ("Kitchen",   21, 22),  ("Dining",   23, 24),
    ("Great Rm",  25, 26),  ("Prim Bath", 27, 28),  ("Bath 2",   29, 30),
    ("Bath 3",    31, 32),  ("Bath 4",    33, 34),  ("Prim Bed", 35, 36),
    ("Bed 2",     37, 38),  ("Bed 3",     39, 40),  ("Bed 4",    41, 42),
]

def read_units(ws):
    units = []
    for row in ws.iter_rows(min_row=11, values_only=True):
        unit_num = row[3]
        if not isinstance(unit_num, int):
            continue
        status   = row[6]
        in_scope = (status != "NOT IN SCOPE")
        comps = []
        if in_scope:
            for name, si, ci in COMP_COLS:
                if si >= len(row) or ci >= len(row):
                    continue
                if row[si] == "YES" and isinstance(row[ci], (int, float)):
                    comps.append({"name": name, "ctn": int(row[ci])})
        ctns = sorted(set(c["ctn"] for c in comps))
        units.append({
            "floor":      row[1],
            "unit":       unit_num,
            "inScope":    in_scope,
            "combo":      row[4] if row[4] and row[4] != "NO" else None,
            "type":       row[5] or None,
            "status":     status,
            "components": comps,
            "ctns":       ctns,
        })
    return units

def remote_url():
    """Build the push URL — embeds token if provided."""
    if GITHUB_TOKEN:
        return GITHUB_REMOTE.replace("https://", f"https://{GITHUB_TOKEN}@")
    return GITHUB_REMOTE

def git_setup():
    """Initialize or repair git repo inside the OneDrive folder."""
    def r(cmd):
        return subprocess.run(cmd, cwd=SITE_DIR, capture_output=True, text=True)
    r(["git", "init"])
    r(["git", "config", "user.email", "flores.edwin9271@gmail.com"])
    r(["git", "config", "user.name", "Edwin"])
    r(["git", "branch", "-M", "main"])
    subprocess.run(["git", "remote", "remove", "origin"], cwd=SITE_DIR, capture_output=True)
    r(["git", "remote", "add", "origin", remote_url()])

def git_push(commit_msg):
    """Add, commit, and push. Auto-repairs git if it's broken."""
    def r(cmd):
        return subprocess.run(cmd, cwd=SITE_DIR, capture_output=True, text=True)

    # Remove stale lock file if present
    lock = os.path.join(SITE_DIR, ".git", "index.lock")
    if os.path.exists(lock):
        try:
            os.remove(lock)
        except Exception:
            pass

    r(["git", "add", "-A"])
    rc = r(["git", "commit", "-m", commit_msg])
    if "nothing to commit" in (rc.stdout + rc.stderr):
        r(["git", "commit", "--allow-empty", "-m", commit_msg])

    rp = r(["git", "push", remote_url(), "main", "--force"])
    if rp.returncode == 0:
        return True, "pushed"

    # Push failed — re-init git and retry once
    print("  Push failed, re-initializing git...")
    git_setup()
    r(["git", "add", "-A"])
    r(["git", "commit", "--allow-empty", "-m", commit_msg])
    rp2 = r(["git", "push", remote_url(), "main", "--force"])
    if rp2.returncode == 0:
        return True, "pushed (after reinit)"
    return False, rp2.stderr.strip()

def main():
    print("=" * 54)
    print("  OMM Optima Dashboard -- Auto-update")
    print("=" * 54)

    # ── Read Excel ──────────────────────────────────────────────
    print(f"\nReading: {os.path.basename(EXCEL_FILE)}")
    if not os.path.exists(EXCEL_FILE):
        print(f"\nERROR: Excel file not found at:\n  {EXCEL_FILE}")
        input("\nPress Enter to close...")
        return

    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    except Exception as e:
        print(f"\nERROR opening Excel: {e}")
        print("Make sure Excel is fully closed, then try again.")
        input("\nPress Enter to close...")
        return

    if SH_CTN not in wb.sheetnames:
        print(f"\nERROR: Sheet '{SH_CTN}' not found.")
        print(f"Available sheets: {wb.sheetnames}")
        input("\nPress Enter to close...")
        return

    containers = read_containers(wb[SH_CTN])
    units      = read_units(wb[SH_UNIT])
    in_scope   = sum(1 for u in units if u["inScope"])
    print(f"  {len(containers)} containers | {in_scope} units in scope")

    if not containers:
        print("\nERROR: No containers found. Check sheet name.")
        input("\nPress Enter to close...")
        return

    # ── Build timestamp ─────────────────────────────────────────
    now = datetime.now()
    months = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()
    updated = f"{months[now.month - 1]} {now.day}, {now.year} at {now.strftime('%I:%M %p')}"

    # ── Build data block ────────────────────────────────────────
    data_block = (
        "// DATA_START\n"
        f"const CONTAINERS = {json.dumps(containers, ensure_ascii=False)};\n"
        f"const UNITS = {json.dumps(units, ensure_ascii=False)};\n"
        f'const UPDATED = "{updated}";\n'
        "// DATA_END"
    )

    # ── Read TEMPLATE, inject data, write to dashboard ──────────
    print(f"\nUpdating dashboard HTML...")
    if not os.path.exists(TEMPLATE_FILE):
        print(f"\nERROR: Template file not found at:\n  {TEMPLATE_FILE}")
        print("Make sure OMM_Optima_Template.html is in the folder.")
        input("\nPress Enter to close...")
        return

    try:
        with open(TEMPLATE_FILE, "rb") as f:
            raw = f.read().replace(b'\x00', b'')
        idx = raw.find(b'<!DOCTYPE html>')
        if idx > 0:
            raw = raw[idx:]
        html = raw.decode("utf-8", errors="replace")
    except Exception as e:
        print(f"\nERROR reading template: {e}")
        input("\nPress Enter to close...")
        return

    new_html, count = re.subn(
        r"// DATA_START.*?// DATA_END",
        lambda m: data_block,
        html,
        flags=re.DOTALL
    )

    if count == 0:
        print("\nERROR: DATA_START / DATA_END markers not found in HTML.")
        print("The HTML file may be corrupted. Contact support.")
        input("\nPress Enter to close...")
        return

    try:
        with open(HTML_FILE, "w", encoding="utf-8") as f:
            f.write(new_html)
        opens  = new_html.count('<script')
        closes = new_html.count('</script>')
        print(f"  HTML updated ({opens} open / {closes} close script tags).")
    except Exception as e:
        print(f"\nERROR writing HTML: {e}")
        input("\nPress Enter to close...")
        return

    # ── Git push ─────────────────────────────────────────────────
    if AUTO_GIT:
        print("\nPushing to GitHub...")
        commit_msg = f"Auto-update {now.strftime('%Y-%m-%d %H:%M')} -- {len(containers)} containers"
        ok, msg = git_push(commit_msg)
        if ok:
            print(f"  OK   ({msg})")
            print(f"\nDone! Site will update within ~60 seconds.")
            print(f"\n  https://edwin210designhouse.github.io/OMM-OPTIMA-210-DASHBOARD/")
        else:
            print(f"\nERROR pushing to GitHub: {msg}")
            print("Try closing all apps and running again.")
    else:
        print(f"\nDone! (Git push skipped — AUTO_GIT = False)")

    input("\nPress Enter to close...")

if __name__ == "__main__":
    main()
