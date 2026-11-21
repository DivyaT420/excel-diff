import pandas as pd
import streamlit as st
from diff_engine import compare_workbooks

st.set_page_config(page_title="ExcelDiff", page_icon="📊", layout="wide")

st.title("📊 ExcelDiff")
st.caption("Compare two Excel files: values and formulas, across all sheets.")

c1, c2 = st.columns(2)
old_file = c1.file_uploader("Old file", type=["xlsx", "xlsm"], key="old")
new_file = c2.file_uploader("New file", type=["xlsx", "xlsm"], key="new")

if st.button("Compare files", type="primary", disabled=not (old_file and new_file)):
    with st.spinner("Comparing..."):
        try:
            st.session_state["result"] = compare_workbooks(old_file.getvalue(), new_file.getvalue())
        except Exception as e:
            st.session_state.pop("result", None)
            st.error(f"Could not compare these files: {e}")

if "result" in st.session_state:
    summary, diffs = st.session_state["result"]

    m = st.columns(6)
    m[0].metric("Sheets changed", summary["sheetsChanged"])
    m[1].metric("Cells changed", summary["cellsChanged"])
    m[2].metric("Rows added", summary["rowsAdded"])
    m[3].metric("Rows removed", summary["rowsRemoved"])
    m[4].metric("Formula diffs", summary["formulaDiffs"])
    m[5].metric("Value diffs", summary["valueDiffs"])
    if summary["sheetsAdded"]:
        st.info("Sheets only in new file: " + ", ".join(summary["sheetsAdded"]))
    if summary["sheetsRemoved"]:
        st.warning("Sheets only in old file: " + ", ".join(summary["sheetsRemoved"]))

    choice = st.radio("Show", ["All", "Modified", "Added", "Removed", "Formula diffs"], horizontal=True)

    # Flat table: one row for a value change, one row for a formula change
    rows = []
    for d in diffs:
        if choice == "Formula diffs" and not d["hasFormulaDiff"]:
            continue
        if choice in ("Modified", "Added", "Removed") and d["type"] != choice.lower():
            continue
        if d["hasValueDiff"] and choice != "Formula diffs":
            rows.append([d["sheet"], d["cell"], d["oldVal"] or "—", d["newVal"] or "—", d["type"].capitalize()])
        if d["hasFormulaDiff"]:
            rows.append([d["sheet"], d["cell"], d["oldFormula"] or "no formula", d["newFormula"] or "no formula", "Formula"])

    df = pd.DataFrame(rows, columns=["Sheet", "Cell", "Old Value", "New Value", "Type"])
    colors = {"Added": "#dcfce7", "Removed": "#fee2e2", "Modified": "#fef9c3", "Formula": "#f3e8ff"}
    styled = df.style.apply(lambda r: [f"background-color: {colors.get(r['Type'], '')}" if c == "Type" else "" for c in df.columns], axis=1)

    st.caption(f"{len(df)} rows")
    st.dataframe(styled, use_container_width=True, hide_index=True)
    st.download_button("Download report (CSV)", df.to_csv(index=False).encode(), "exceldiff-report.csv", "text/csv")
