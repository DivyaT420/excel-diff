"""Core Excel diff logic (no Flask / no web framework). Used by the Streamlit app."""
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

MAX_CELLS_PER_SHEET = 200_000


def _cell(ws_val, ws_fml, r, c):
    v = ws_val.cell(row=r, column=c).value
    f = ws_fml.cell(row=r, column=c).value
    is_fml = isinstance(f, str) and f.startswith("=")
    return ("" if v is None else str(v)), (f if is_fml else "")


def _diff_sheet(name, ov, of, nv, nf):
    # Read dimensions BEFORE touching cells (ws.cell() can inflate max_row/col)
    max_row = max(ov.max_row or 0, nv.max_row or 0)
    max_col = max(ov.max_column or 0, nv.max_column or 0)
    if max_row * max_col > MAX_CELLS_PER_SHEET:
        raise ValueError(
            f"Sheet '{name}' has {max_row * max_col:,} cells, above the "
            f"{MAX_CELLS_PER_SHEET:,} limit."
        )
    diffs = []
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            old_val, old_fml = _cell(ov, of, r, c)
            new_val, new_fml = _cell(nv, nf, r, c)
            val_changed = old_val != new_val
            fml_changed = old_fml != new_fml
            if not (val_changed or fml_changed):
                continue
            if old_val == "" and old_fml == "":
                kind = "added"
            elif new_val == "" and new_fml == "":
                kind = "removed"
            else:
                kind = "modified"
            diffs.append({
                "sheet": name, "cell": f"{get_column_letter(c)}{r}", "type": kind,
                "hasValueDiff": val_changed, "hasFormulaDiff": fml_changed,
                "oldVal": old_val, "newVal": new_val,
                "oldFormula": old_fml, "newFormula": new_fml,
            })
    return diffs


def compare_workbooks(old_bytes: bytes, new_bytes: bytes):
    wo_v = load_workbook(BytesIO(old_bytes), data_only=True)
    wo_f = load_workbook(BytesIO(old_bytes), data_only=False)
    wn_v = load_workbook(BytesIO(new_bytes), data_only=True)
    wn_f = load_workbook(BytesIO(new_bytes), data_only=False)

    old_names, new_names = wo_v.sheetnames, wn_v.sheetnames
    primary = old_names[0] if old_names else None
    old_rows = max(0, (wo_v[primary].max_row or 1) - 1) if primary else 0
    new_rows = max(0, (wn_v[primary].max_row or 1) - 1) if primary in new_names else 0

    all_diffs, sheets_changed = [], 0
    for name in old_names:
        if name not in new_names:
            continue
        d = _diff_sheet(name, wo_v[name], wo_f[name], wn_v[name], wn_f[name])
        sheets_changed += 1 if d else 0
        all_diffs.extend(d)

    summary = {
        "sheetsChanged": sheets_changed,
        "sheetsAdded": [s for s in new_names if s not in old_names],
        "sheetsRemoved": [s for s in old_names if s not in new_names],
        "cellsChanged": len(all_diffs),
        "rowsAdded": max(0, new_rows - old_rows),
        "rowsRemoved": max(0, old_rows - new_rows),
        "formulaDiffs": sum(d["hasFormulaDiff"] for d in all_diffs),
        "valueDiffs": sum(d["hasValueDiff"] for d in all_diffs),
    }
    return summary, all_diffs
