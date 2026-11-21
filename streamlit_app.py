import csv
import html as _html
import io
import re

import streamlit as st

from diff_engine import compare_workbooks

st.set_page_config(page_title="ExcelDiff", page_icon="📊", layout="wide", initial_sidebar_state="collapsed")

MAX_ROWS_SHOWN = 500          # cap rendered diffs; CSV download always has everything
ACCEPTED = ["xlsx", "xlsm"]


# ── Helpers ──────────────────────────────────────────────────────────────────
def md(s: str):
    """Render raw HTML. Newlines + indentation are collapsed so Streamlit's
    markdown parser never sees blank lines or 4-space code blocks."""
    st.markdown(re.sub(r"\n\s*", "", s.strip()), unsafe_allow_html=True)


def esc(v) -> str:
    return _html.escape(str(v))


def icon(d, size=20, color="currentColor"):
    return (f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" '
            f'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="{d}"/></svg>')


ICONS = {
    "zap": "M13 2L3 14h9l-1 8 10-12h-9l1-8z",
    "eye": "M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8zM12 9a3 3 0 1 0 0 6 3 3 0 0 0 0-6z",
    "shield": "M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z",
    "download": "M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3",
    "arrow": "M5 12h14M12 5l7 7-7 7",
    "grid": "M3 3h7v7H3zM14 3h7v7h-7zM14 14h7v7h-7zM3 14h7v7H3z",
    "check": "M20 6L9 17l-5-5",
    "x": "M18 6L6 18M6 6l12 12",
}

SERIF = "'DM Serif Display',Georgia,serif"


# ── Global styling ───────────────────────────────────────────────────────────
GLOBAL_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@400;500;600;700&display=swap');
:root{color-scheme:light}
header[data-testid="stHeader"],[data-testid="stToolbar"],[data-testid="stDecoration"],[data-testid="stStatusWidget"],footer,#MainMenu{display:none !important}
.stApp{font-family:'DM Sans',system-ui,sans-serif;color:#0f172a}
[data-testid="stMainBlockContainer"],.block-container{padding:0 !important;max-width:100% !important}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"]{gap:0 !important}
[data-testid="stMarkdownContainer"] p{margin:0}
.xd a{text-decoration:none}
@keyframes xdspin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
.xd-spin{animation:xdspin 1s linear infinite}

/* nav */
.xd-nav{background:#fff;border-bottom:1px solid #e8edf5}
.xd-nav-in{max-width:72rem;margin:0 auto;padding:0 24px;height:56px;display:flex;align-items:center;justify-content:space-between}
.xd-nav a.link{color:#64748b;font-size:14px;margin-right:24px}
.xd-nav a.link:hover{color:#2563eb}
.xd-cta{background:#2563eb;color:#fff !important;border-radius:8px;font-size:14px;font-weight:600;padding:7px 18px;display:inline-block}
.xd-cta:hover{background:#1d4ed8}
.xd-hero-btn{background:linear-gradient(135deg,#2563eb,#0ea5e9);color:#fff !important;border-radius:12px;padding:15px 36px;font-size:16px;font-weight:700;display:inline-flex;align-items:center;gap:10px;box-shadow:0 8px 30px rgba(37,99,235,.35)}
.xd-hero-btn:hover{opacity:.9}
.xd-card{border:1px solid #e8edf5;border-radius:14px;padding:28px;background:#fff;transition:box-shadow .2s,transform .2s}
.xd-card:hover{box-shadow:0 8px 32px rgba(37,99,235,.1);transform:translateY(-3px)}

/* results table */
table.xd-table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:13px;margin:0}
table.xd-table th{text-align:left;padding:11px 16px !important;color:#94a3b8;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;background:#f8fafc !important;border:none !important;border-bottom:2px solid #e8edf5 !important}
table.xd-table td{padding:11px 16px !important;vertical-align:middle;border:none !important;border-bottom:1px solid #f1f5f9 !important;background:transparent;color:#0f172a;overflow:hidden;text-overflow:ellipsis}
table.xd-table tr.fx td{background:#fdfaff !important}
table.xd-table tr:hover td{background:#fafbff !important}
table.xd-table tr.fx:hover td{background:#f7f0ff !important}
table.xd-table tr:last-child td{border-bottom:none !important}
.xd-mono{font-family:ui-monospace,Consolas,monospace}

/* streamlit widgets */
[data-testid="stFileUploader"] label p{font-weight:700;font-size:16px;color:#0f172a}
[data-testid="stFileUploaderDropzone"]{border:2px dashed #cbd5e1;border-radius:16px;background:#fafbfc;padding:32px 20px;transition:all .2s}
[data-testid="stFileUploaderDropzone"]:hover{border-color:#2563eb;background:#eff6ff}
[data-testid="stFileUploaderDropzone"] button{border-radius:8px;border:1px solid #cbd5e1;background:#fff;color:#0f172a;font-weight:600}
[data-testid="stBaseButton-primary"]{background:#2563eb;border:none;border-radius:10px;color:#fff;font-weight:600;padding:.55rem 1.1rem}
[data-testid="stBaseButton-primary"]:hover{background:#1d4ed8;color:#fff}
[data-testid="stBaseButton-secondary"]{background:#fff;border:1px solid #e2e8f0;border-radius:10px;color:#64748b;font-weight:600;padding:.55rem 1.1rem}
[data-testid="stBaseButton-secondary"]:hover{background:#f9fafb;border-color:#cbd5e1;color:#0f172a}
.st-key-compare_btn button{background:linear-gradient(135deg,#2563eb,#0ea5e9);box-shadow:0 6px 20px rgba(37,99,235,.3);font-size:15px;font-weight:700;padding:.75rem 1.5rem}
.st-key-compare_btn button:hover{background:linear-gradient(135deg,#1d4ed8,#0284c7);color:#fff}
.st-key-compare_btn button:disabled{background:#e2e8f0;color:#94a3b8;box-shadow:none}
.st-key-reset_btn button,.st-key-new_btn button,.st-key-download_btn button{font-size:14px}
.st-key-new_btn,.st-key-download_btn{margin-top:14px}
[data-testid="stBaseButton-pills"]{border-radius:8px;border:1px solid #e2e8f0;background:#fff;color:#64748b;font-weight:600;font-size:13px}
[data-testid="stBaseButton-pillsActive"]{border-radius:8px;border:1px solid #2563eb;background:#eff6ff;color:#1d4ed8;font-weight:600;font-size:13px}
</style>
"""


def page_bg(bg):
    st.markdown(f"<style>.stApp{{background:{bg}}}</style>", unsafe_allow_html=True)


# ── Navigation ───────────────────────────────────────────────────────────────
def render_nav(landing: bool):
    links = ""
    if landing:
        links = ('<a class="link" href="#features" target="_self">Features</a>'
                 '<a class="link" href="#pricing" target="_self">Pricing</a>')
    md(f"""
    <div class="xd xd-nav"><div class="xd-nav-in">
      <a href="?page=landing" target="_self" style="display:flex;align-items:center;gap:8px">
        <span style="background:linear-gradient(135deg,#2563eb,#0ea5e9);border-radius:8px;width:28px;height:28px;display:inline-flex;align-items:center;justify-content:center;color:#fff">{icon(ICONS['grid'], 18)}</span>
        <span style="font-family:{SERIF};color:#0f172a;font-size:18px;letter-spacing:-.3px">Excel<span style="color:#2563eb">Diff</span></span>
      </a>
      <div style="display:flex;align-items:center">{links}<a class="xd-cta" href="?page=upload" target="_self">Get Started</a></div>
    </div></div>
    """)


# ── Landing page ─────────────────────────────────────────────────────────────
def render_landing():
    features = [
        ("zap", "Instant Comparison", "Results appear in seconds. Upload two files, click compare, done."),
        ("eye", "Visual Diff Table", "Color-coded changes across sheets, cells, rows, and formula differences."),
        ("shield", "Secure Processing", "Files are processed in memory and are never saved by the app."),
        ("download", "Export Reports", "Download a full CSV report of every change for audits, reviews, or archives."),
    ]
    steps = [
        ("01", "Upload Two Files", "Add your old and new .xlsx files using the upload boxes."),
        ("02", "Click Compare", "Hit the compare button and let ExcelDiff scan every cell across all sheets."),
        ("03", "Review Changes", "Browse the color-highlighted diff table and download your change report."),
    ]
    plans = [
        ("Free", "&#36;0", "forever", ["5 comparisons/mo", "Up to 10 sheets", "CSV export"], "Start Free", False),
        ("Pro", "&#36;12", "/month", ["Unlimited comparisons", "Formula diff", "History log", "Priority support"], "Start Trial", True),
        ("Enterprise", "Custom", "", ["SSO &amp; access control", "API access", "SLA &amp; compliance", "Dedicated support"], "Contact Us", False),
    ]

    feat_html = "".join(f"""
      <div class="xd-card">
        <div style="width:42px;height:42px;border-radius:10px;background:#dbeafe;display:flex;align-items:center;justify-content:center;color:#2563eb;margin-bottom:16px">{icon(ICONS[i])}</div>
        <div style="font-weight:700;color:#0f172a;margin-bottom:8px">{t}</div>
        <div style="color:#64748b;font-size:14px;line-height:1.6">{d}</div>
      </div>""" for i, t, d in features)

    steps_html = "".join(f"""
      <div style="text-align:center">
        <div style="font-family:{SERIF};font-size:48px;color:#dbeafe;line-height:1;margin-bottom:12px">{n}</div>
        <div style="font-weight:700;color:#0f172a;font-size:17px;margin-bottom:8px">{t}</div>
        <div style="color:#64748b;font-size:14px;line-height:1.6">{d}</div>
      </div>""" for n, t, d in steps)

    plans_html = ""
    for name, price, period, feats, cta, hi in plans:
        border = "2px solid #2563eb" if hi else "1px solid #e8edf5"
        bg = "#eff6ff" if hi else "#fff"
        shadow = "0 8px 32px rgba(37,99,235,.12)" if hi else "none"
        badge = ('<div style="position:absolute;top:-12px;left:50%;transform:translateX(-50%);background:#2563eb;color:#fff;'
                 'border-radius:999px;padding:3px 14px;font-size:12px;font-weight:700">Most Popular</div>') if hi else ""
        li = "".join(f'<div style="display:flex;align-items:center;gap:8px;color:#475569;font-size:14px;margin-bottom:10px">'
                     f'<span style="color:#2563eb;display:flex">{icon(ICONS["check"], 16)}</span>{f}</div>' for f in feats)
        btn_bg, btn_fg = ("#2563eb", "#fff") if hi else ("#f1f5f9", "#0f172a")
        plans_html += f"""
      <div style="border:{border};border-radius:16px;padding:32px;background:{bg};box-shadow:{shadow};position:relative">
        {badge}
        <div style="font-weight:700;color:#0f172a;font-size:16px;margin-bottom:8px">{name}</div>
        <div style="display:flex;align-items:baseline;gap:4px;margin-bottom:20px"><span style="font-family:{SERIF};font-size:36px;color:#0f172a">{price}</span><span style="color:#94a3b8;font-size:14px">{period}</span></div>
        <div style="margin-bottom:24px">{li}</div>
        <a href="?page=upload" target="_self" style="display:block;text-align:center;padding:11px 0;border-radius:10px;background:{btn_bg};color:{btn_fg};font-weight:600;font-size:14px">{cta}</a>
      </div>"""

    preview = [
        ("Sales Q1", "B4", "12,400", "14,800", "modified", False),
        ("Inventory", "A9", "—", "SKU-512", "added", False),
        ("Summary", "B2", "=SUM(B3:B10)", "=SUM(B3:B12)", "modified", True),
    ]
    prev_stats = "".join(f'<div style="background:{bg};border-radius:10px;padding:14px 16px"><div style="color:{c};font-weight:800;font-size:22px">{v}</div><div style="color:#64748b;font-size:11px;margin-top:2px">{l}</div></div>'
                         for l, v, bg, c in [("Sheets Changed", "1", "#dbeafe", "#1d4ed8"), ("Cells Changed", "3", "#fef9c3", "#a16207"),
                                              ("Rows Added", "1", "#dcfce7", "#15803d"), ("Formula Diffs", "1", "#f3e8ff", "#7c3aed")])
    prev_rows = ""
    for sh, cl, ov, nv, t, fx in preview:
        ff = "ui-monospace,Consolas,monospace" if fx else "inherit"
        o_col, o_bg = ("#6b7280", "transparent") if t == "added" else ("#b91c1c", "#fef2f2")
        n_col, n_bg = ("#15803d", "#f0fdf4") if t == "added" else ("#a16207", "#fefce8")
        if fx:
            tag = '<span style="background:#f3e8ff;color:#7c3aed;border-radius:999px;padding:2px 8px;font-size:11px;font-weight:600">formula</span>'
        else:
            tb, tc = ("#dcfce7", "#15803d") if t == "added" else ("#fef9c3", "#a16207")
            tag = f'<span style="background:{tb};color:{tc};border-radius:999px;padding:2px 8px;font-size:11px;font-weight:600">{t}</span>'
        prev_rows += f"""
          <div style="display:grid;grid-template-columns:1fr 60px 1fr 1fr 80px;gap:8px;padding:8px 0;border-bottom:1px solid #f1f5f9;font-size:12px;align-items:center;text-align:left">
            <span style="color:#475569">{sh}</span>
            <span style="font-family:ui-monospace,Consolas,monospace;background:#f1f5f9;border-radius:4px;padding:2px 6px;text-align:center">{cl}</span>
            <span style="color:{o_col};background:{o_bg};border-radius:4px;padding:2px 8px;font-family:{ff}">{ov}</span>
            <span style="color:{n_col};background:{n_bg};border-radius:4px;padding:2px 8px;font-family:{ff}">{nv}</span>
            <span>{tag}</span>
          </div>"""

    md(f"""
    <div class="xd">
      <section style="background:linear-gradient(160deg,#f0f6ff 0%,#fff 60%);padding:80px 24px 96px;text-align:center">
        <div style="max-width:56rem;margin:0 auto">
          <div style="display:inline-flex;align-items:center;gap:8px;background:#dbeafe;border-radius:999px;padding:5px 14px;margin-bottom:28px;color:#1d4ed8">{icon(ICONS['zap'], 16)}<span style="font-size:13px;font-weight:600">No signup required · Free to use</span></div>
          <div style="font-family:{SERIF};font-size:clamp(36px,6vw,68px);color:#0f172a;line-height:1.1;letter-spacing:-1px;margin-bottom:22px">Compare Excel Files<br><span style="background:linear-gradient(90deg,#2563eb,#0ea5e9);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;color:transparent">in Seconds</span></div>
          <div style="color:#475569;font-size:18px;line-height:1.7;max-width:520px;margin:0 auto 40px">Detect changed cells, rows, sheets, and formulas instantly. No accounts needed.</div>
          <a class="xd-hero-btn" href="?page=upload" target="_self">Upload Files {icon(ICONS['arrow'])}</a>
          <div style="border-radius:16px;border:1px solid #dde4f0;background:#fff;box-shadow:0 20px 60px rgba(0,0,0,.08);overflow:hidden;max-width:820px;margin:64px auto 0">
            <div style="background:#f8fafc;border-bottom:1px solid #e2e8f0;padding:10px 16px;display:flex;gap:6px;text-align:left">
              <span style="width:11px;height:11px;border-radius:50%;background:#ff5f57;display:inline-block"></span>
              <span style="width:11px;height:11px;border-radius:50%;background:#febc2e;display:inline-block"></span>
              <span style="width:11px;height:11px;border-radius:50%;background:#28c840;display:inline-block"></span>
            </div>
            <div style="padding:24px 28px">
              <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px;text-align:left">{prev_stats}</div>
              {prev_rows}
            </div>
          </div>
        </div>
      </section>
      <section id="features" style="padding:80px 24px;background:#fff">
        <div style="max-width:64rem;margin:0 auto">
          <div style="text-align:center;margin-bottom:52px">
            <div style="font-family:{SERIF};font-size:clamp(26px,4vw,40px);color:#0f172a;margin-bottom:12px">Everything you need to audit changes</div>
            <div style="color:#64748b;font-size:16px">Purpose-built for finance, ops, and data teams who live in spreadsheets.</div>
          </div>
          <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:24px">{feat_html}</div>
        </div>
      </section>
      <section style="padding:80px 24px;background:#f8fafc">
        <div style="max-width:56rem;margin:0 auto">
          <div style="text-align:center;margin-bottom:52px"><div style="font-family:{SERIF};font-size:clamp(26px,4vw,40px);color:#0f172a">Three steps. Done.</div></div>
          <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:32px">{steps_html}</div>
        </div>
      </section>
      <section id="pricing" style="padding:80px 24px;background:#fff">
        <div style="max-width:64rem;margin:0 auto">
          <div style="text-align:center;margin-bottom:52px">
            <div style="font-family:{SERIF};font-size:clamp(26px,4vw,40px);color:#0f172a;margin-bottom:12px">Simple, transparent pricing</div>
            <div style="color:#64748b;font-size:16px">Start free. Upgrade when you need more.</div>
          </div>
          <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:24px;align-items:start">{plans_html}</div>
        </div>
      </section>
      <footer style="background:#0f172a;padding:40px 24px">
        <div style="max-width:64rem;margin:0 auto;display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:16px">
          <span style="font-family:{SERIF};color:#fff;font-size:18px">Excel<span style="color:#60a5fa">Diff</span></span>
          <span style="color:#64748b;font-size:13px">© 2026 ExcelDiff. Built for spreadsheet people.</span>
          <span style="color:#64748b;font-size:13px">Privacy · Terms · Contact</span>
        </div>
      </footer>
    </div>
    """)


# ── Upload page ──────────────────────────────────────────────────────────────
def render_upload():
    page_bg("#f8fafc")
    ss = st.session_state
    _, mid, _ = st.columns([1, 2.6, 1])
    with mid:
        md(f"""
        <div class="xd" style="text-align:center;padding:48px 0 32px">
          <div style="font-family:{SERIF};font-size:clamp(24px,4vw,38px);color:#0f172a;margin-bottom:10px">Upload Your Files</div>
          <div style="color:#64748b;font-size:16px">Drop your old and new Excel files below to detect every change.</div>
        </div>
        """)
        ca, car, cb = st.columns([10, 1, 10], vertical_alignment="center")
        n = ss.up_n
        old = ca.file_uploader("Old File", type=ACCEPTED, key=f"old_{n}")
        with car:
            md(f'<div style="color:#cbd5e1;display:flex;justify-content:center">{icon(ICONS["arrow"])}</div>')
        new = cb.file_uploader("New File", type=ACCEPTED, key=f"new_{n}")

        if not (old and new):
            md('<div style="text-align:center;margin:12px 0 20px;color:#94a3b8;font-size:14px">📂 Upload both files to enable comparison</div>')

        if ss.get("error"):
            md(f'<div class="xd" style="background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:12px 16px;margin:8px 0 16px;display:flex;align-items:center;gap:10px;color:#b91c1c;font-size:14px">{icon(ICONS["x"])}<span>{esc(ss.error)}</span></div>')

        _, b1, b2, _ = st.columns([1, 2, 1.3, 1])
        go = b1.button("⚡ Compare Files", key="compare_btn", type="primary", use_container_width=True, disabled=not (old and new))
        if b2.button("↻ Reset", key="reset_btn", use_container_width=True):
            ss.up_n += 1
            ss.error = None
            st.rerun()

        if go:
            ss.error = None
            ph = st.empty()
            with ph.container():
                md(f"""
                <div class="xd" style="margin-top:32px;text-align:center"><div style="display:inline-flex;flex-direction:column;align-items:center;gap:16px;background:#fff;border-radius:16px;padding:32px 48px;border:1px solid #e8edf5;box-shadow:0 8px 32px rgba(0,0,0,.07)">
                  <svg class="xd-spin" width="52" height="52" viewBox="0 0 52 52" fill="none"><circle cx="26" cy="26" r="22" stroke="#e2e8f0" stroke-width="4"/><path d="M26 4a22 22 0 0 1 22 22" stroke="#2563eb" stroke-width="4" stroke-linecap="round"/></svg>
                  <div><div style="font-weight:700;color:#0f172a;margin-bottom:4px">Analyzing files…</div><div style="color:#94a3b8;font-size:13px">Scanning sheets, cells, and formulas</div></div>
                </div></div>
                """)
            try:
                ss.result = compare_workbooks(old.getvalue(), new.getvalue())
                ss.page = "results"
            except ValueError as e:
                ss.error = str(e)
            except Exception:
                ss.error = "Could not read these files. Please make sure both are valid .xlsx files."
            st.rerun()


# ── Results page ─────────────────────────────────────────────────────────────
FILTERS = {"All Changes": "all", "Modified": "modified", "Added": "added", "Removed": "removed", "Formula Diffs": "formula"}
TYPE_STYLE = {
    "added": ("#f0fdf4", "#15803d", "#bbf7d0"),
    "removed": ("#fef2f2", "#b91c1c", "#fecaca"),
    "modified": ("#fefce8", "#a16207", "#fef08a"),
}


def pill(text, bg, fg, border, cap=True):
    return (f'<span style="background:{bg};color:{fg};border:1px solid {border};border-radius:999px;padding:3px 10px;'
            f'font-size:12px;font-weight:600;white-space:nowrap;{"text-transform:capitalize" if cap else ""}">{esc(text)}</span>')


def chip(text, bg, fg, mono=False, italic=False):
    style = f"background:{bg};color:{fg};border-radius:6px;padding:3px 10px;"
    if mono:
        style += "font-family:ui-monospace,Consolas,monospace;"
    if italic:
        style += "font-style:italic;"
    return f'<span style="{style}">{esc(text)}</span>'


def build_csv(diffs):
    buf = io.StringIO()
    w = csv.writer(buf, quoting=csv.QUOTE_ALL)
    w.writerow(["Sheet", "Cell", "Old Value", "New Value", "Old Formula", "New Formula", "Type"])
    for d in diffs:
        w.writerow([d["sheet"], d["cell"], d["oldVal"], d["newVal"], d["oldFormula"], d["newFormula"], d["type"]])
    return buf.getvalue().encode("utf-8-sig")


def table_html(diffs, mode):
    cell_chip = lambda c: chip(c, "#f1f5f9", "#475569", mono=True)
    rows = []
    for d in diffs:
        sheet = f'<td style="font-weight:600">{esc(d["sheet"])}</td>'
        cell = f'<td>{cell_chip(d["cell"])}</td>'
        if d["hasValueDiff"] and mode != "formula":
            bg, fg, bd = TYPE_STYLE[d["type"]]
            old = chip(d["oldVal"], "#fef2f2" if d["type"] == "removed" else "#f1f5f9", "#b91c1c" if d["type"] == "removed" else "#475569") if d["oldVal"] else '<span style="color:#cbd5e1">—</span>'
            new = chip(d["newVal"], bg if d["type"] != "removed" else "#f1f5f9", fg if d["type"] != "removed" else "#475569") if d["newVal"] else '<span style="color:#cbd5e1">—</span>'
            rows.append(f'<tr>{sheet}{cell}<td>{old}</td><td>{new}</td><td>{pill(d["type"], bg, fg, bd)}</td></tr>')
        if d["hasFormulaDiff"]:
            old = chip(d["oldFormula"], "#fef2f2", "#b91c1c", mono=True) if d["oldFormula"] else chip("no formula", "transparent", "#94a3b8", italic=True)
            new = chip(d["newFormula"], "#f0fdf4", "#15803d", mono=True) if d["newFormula"] else chip("no formula", "transparent", "#94a3b8", italic=True)
            rows.append(f'<tr class="fx">{sheet}{cell}<td>{old}</td><td>{new}</td><td>{pill("formula", "#f3e8ff", "#7c3aed", "#e9d5ff", cap=False)}</td></tr>')
    return ('<table class="xd-table"><colgroup><col style="width:18%"><col style="width:10%"><col style="width:28%"><col style="width:28%"><col style="width:16%"></colgroup>'
            '<thead><tr><th>Sheet</th><th>Cell</th><th>Old Value</th><th>New Value</th><th>Type</th></tr></thead>'
            f'<tbody>{"".join(rows)}</tbody></table>')


def render_results():
    page_bg("#f8fafc")
    ss = st.session_state
    summary, diffs = ss.result

    _, mid, _ = st.columns([1, 4, 1])
    with mid:
        h, bn, bd = st.columns([5, 1.7, 2.1])
        with h:
            md(f"""
            <div class="xd" style="padding:36px 0 24px">
              <div style="font-family:{SERIF};font-size:clamp(22px,4vw,34px);color:#0f172a;margin-bottom:4px">Comparison Results</div>
              <div style="color:#64748b;font-size:14px">Showing all detected differences between the two files.</div>
            </div>
            """)
        if bn.button("↻ New Compare", key="new_btn", use_container_width=True):
            ss.page = "upload"
            ss.result = None
            ss.up_n += 1
            st.rerun()
        bd.download_button("⬇ Download Report", build_csv(diffs), "exceldiff-report.csv", "text/csv",
                           key="download_btn", type="primary", use_container_width=True)

        cards = [
            ("Sheets Changed", summary["sheetsChanged"], "#dbeafe", "#1d4ed8"),
            ("Cells Changed", summary["cellsChanged"], "#fef9c3", "#a16207"),
            ("Rows Added", summary["rowsAdded"], "#dcfce7", "#15803d"),
            ("Rows Removed", summary["rowsRemoved"], "#fee2e2", "#b91c1c"),
            ("Formula Diffs", summary["formulaDiffs"], "#f3e8ff", "#7c3aed"),
            ("Value Diffs", summary["valueDiffs"], "#fff7ed", "#c2410c"),
        ]
        cards_html = "".join(f"""
          <div style="background:#fff;border:1px solid #e8edf5;border-radius:14px;padding:20px 22px;box-shadow:0 2px 8px rgba(0,0,0,.04)">
            <div style="font-family:{SERIF};font-size:36px;color:{c};line-height:1">{v}</div>
            <div style="color:#64748b;font-size:13px;margin-top:6px">{l}</div>
            <div style="height:4px;border-radius:999px;background:{bg};margin-top:12px"></div>
          </div>""" for l, v, bg, c in cards)
        md(f'<div class="xd" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:16px;margin-bottom:24px">{cards_html}</div>')

        notes = ""
        if summary["sheetsAdded"]:
            notes += f'<div style="background:#eff6ff;border:1px solid #bfdbfe;color:#1d4ed8;border-radius:10px;padding:10px 14px;font-size:13px;margin-bottom:8px">Sheets only in the new file: {esc(", ".join(summary["sheetsAdded"]))}</div>'
        if summary["sheetsRemoved"]:
            notes += f'<div style="background:#fffbeb;border:1px solid #fde68a;color:#a16207;border-radius:10px;padding:10px 14px;font-size:13px;margin-bottom:8px">Sheets only in the old file: {esc(", ".join(summary["sheetsRemoved"]))}</div>'
        if notes:
            md(f'<div class="xd" style="margin-bottom:12px">{notes}</div>')

        fl, fr = st.columns([5, 1], vertical_alignment="center")
        with fl:
            label = st.pills("Filter", list(FILTERS), selection_mode="single", default="All Changes",
                             key="filter", label_visibility="collapsed")
        mode = FILTERS.get(label or "All Changes", "all")
        if mode == "all":
            filtered = diffs
        elif mode == "formula":
            filtered = [d for d in diffs if d["hasFormulaDiff"]]
        else:
            filtered = [d for d in diffs if d["type"] == mode]
        with fr:
            md(f'<div style="text-align:right;color:#94a3b8;font-size:13px">{len(filtered)} results</div>')

        shown = filtered[:MAX_ROWS_SHOWN]
        if shown:
            md(f'<div class="xd" style="background:#fff;border:1px solid #e8edf5;border-radius:16px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,.05);margin-top:6px">{table_html(shown, mode)}</div>')
        else:
            md('<div class="xd" style="background:#fff;border:1px solid #e8edf5;border-radius:16px;padding:48px;text-align:center;color:#94a3b8">No results for this filter.</div>')
        if len(filtered) > MAX_ROWS_SHOWN:
            md(f'<div style="margin:12px 0 40px;text-align:center;color:#94a3b8;font-size:13px">Showing the first {MAX_ROWS_SHOWN:,} of {len(filtered):,} changes. Download the report for the full list.</div>')
        else:
            md('<div style="height:40px"></div>')


# ── Router ───────────────────────────────────────────────────────────────────
ss = st.session_state
ss.setdefault("page", "landing")
ss.setdefault("result", None)
ss.setdefault("up_n", 0)
ss.setdefault("error", None)

if "page" in st.query_params:           # links in the nav / hero use ?page=...
    requested = st.query_params["page"]
    ss.page = requested if requested in ("landing", "upload") else "landing"
    st.query_params.clear()

if ss.page == "results" and not ss.result:
    ss.page = "upload"

st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
render_nav(landing=ss.page == "landing")

if ss.page == "landing":
    page_bg("#ffffff")
    render_landing()
elif ss.page == "upload":
    render_upload()
else:
    render_results()
