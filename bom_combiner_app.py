"""
BOM Combiner — Streamlit Web Version
FabTools | Tool 3
Combines multiple BOM Excel files into one ready-to-use Combined BOM file.
"""

import io
from copy import copy
from typing import Any, Dict, List, Optional, Tuple

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="BOM Combiner — FabTools",
    page_icon="📋",
    layout="centered",
)

# ─────────────────────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Source Sans 3', sans-serif; }

.hero-banner {
    background: linear-gradient(135deg, #0C447C 0%, #185FA5 55%, #378ADD 100%);
    border-radius: 12px; padding: 36px 36px 28px; color: white; margin-bottom: 32px;
}
.hero-banner h1 { font-size: 28px; font-weight: 700; margin: 0 0 8px 0; line-height: 1.2; }
.hero-banner p  { font-size: 15px; color: rgba(255,255,255,0.82); margin: 0 0 20px 0; line-height: 1.6; }
.hero-steps     { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.hero-step      { display: flex; align-items: center; gap: 8px; font-size: 14px; font-weight: 600; color: rgba(255,255,255,0.95); }
.step-num {
    width: 24px; height: 24px; border-radius: 50%;
    background: rgba(255,255,255,0.2); border: 1.5px solid rgba(255,255,255,0.4);
    display: flex; align-items: center; justify-content: center;
    font-size: 12px; font-weight: 700; text-align: center; line-height: 24px;
}
.hero-arrow { color: rgba(255,255,255,0.4); font-size: 16px; }
.trust-badge {
    display: inline-flex; align-items: center; gap: 8px;
    background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.2);
    border-radius: 8px; padding: 7px 14px; font-size: 13px;
    color: rgba(255,255,255,0.85); margin-top: 16px;
}
.section-header {
    font-size: 13px; font-weight: 600; letter-spacing: 1.2px;
    text-transform: uppercase; color: #185FA5; margin: 0 0 12px 0;
    padding-top: 8px; border-top: 2px solid #E6F1FB;
}
.info-box {
    background: #E6F1FB; border-left: 3px solid #185FA5;
    border-radius: 0 8px 8px 0; padding: 10px 14px;
    font-size: 13px; color: #0C447C; line-height: 1.55; margin-bottom: 12px;
}
.warn-box {
    background: #FFF8E1; border-left: 3px solid #F59E0B;
    border-radius: 0 8px 8px 0; padding: 10px 14px;
    font-size: 13px; color: #92400E; line-height: 1.55; margin-bottom: 12px;
}
.assembly-row {
    display: flex; align-items: center; gap: 12px;
    background: #F8F9FC; border: 1px solid #DDE1EA;
    border-radius: 8px; padding: 10px 14px; margin-bottom: 8px;
}
.assembly-name { font-weight: 600; font-size: 14px; color: #1A1F2E; flex: 1; }
.assembly-source { font-size: 11px; color: #8E96A8; }
.stats-row {
    display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 16px;
}
.stat-card {
    background: #F8F9FC; border: 1px solid #DDE1EA;
    border-radius: 8px; padding: 14px; text-align: center;
}
.stat-num   { font-size: 28px; font-weight: 700; color: #185FA5; }
.stat-label { font-size: 12px; color: #8E96A8; margin-top: 2px; }
.footer-bar {
    text-align: center; font-size: 12px; color: #888;
    margin-top: 40px; padding-top: 20px; border-top: 1px solid #EFF1F5;
}
.conflict-box {
    background: #FFF3CD; border: 1px solid #F59E0B;
    border-radius: 8px; padding: 16px; margin-bottom: 12px;
}
.conflict-title { font-weight: 700; color: #92400E; margin-bottom: 8px; font-size: 14px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def read_bom_sheets(file_bytes: bytes, filename: str) -> List[Dict]:
    """Read all sheets from a BOM file. Returns list of sheet info dicts."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheets = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # E2 = unit quantity
        raw_qty = ws["E2"].value
        try:
            qty = int(float(str(raw_qty))) if raw_qty is not None else 1
        except (ValueError, TypeError):
            qty = 1
        sheets.append({
            "sheet_name":    sheet_name,
            "source_file":   filename,
            "original_qty":  qty,
            "file_bytes":    file_bytes,
        })
    return sheets


def copy_sheet_to_workbook(src_wb_bytes: bytes, src_sheet_name: str,
                            dst_wb, dst_sheet_name: str, new_qty: int) -> None:
    """Copy a sheet from source workbook bytes into destination workbook with updated E2."""
    src_wb = load_workbook(io.BytesIO(src_wb_bytes), data_only=False)
    src_ws = src_wb[src_sheet_name]

    dst_ws = dst_wb.create_sheet(dst_sheet_name)

    # Copy column widths and row heights
    for col, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col].width = dim.width
    for row, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row].height = dim.height

    # Copy all cells with style
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = dst_ws.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                new_cell.font        = copy(cell.font)
                new_cell.border      = copy(cell.border)
                new_cell.fill        = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection  = copy(cell.protection)
                new_cell.alignment   = copy(cell.alignment)

    # Copy merged cells
    for merged in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged))

    # Update E2 with new quantity (keep style)
    dst_ws["E2"].value = new_qty

    # Update row 1 title to reflect new sheet name if it differs
    if src_sheet_name != dst_sheet_name:
        cell_a1 = dst_ws["A1"]
        if cell_a1.value and src_sheet_name in str(cell_a1.value):
            cell_a1.value = str(cell_a1.value).replace(src_sheet_name, dst_sheet_name)


def build_combined_bom(assemblies: List[Dict]) -> bytes:
    """
    Build combined BOM workbook.
    assemblies: list of {sheet_name, final_name, source_file, file_bytes, qty}
    """
    dst_wb = Workbook()
    # Remove default empty sheet
    dst_wb.remove(dst_wb.active)

    for asm in assemblies:
        copy_sheet_to_workbook(
            src_wb_bytes  = asm["file_bytes"],
            src_sheet_name= asm["sheet_name"],
            dst_wb        = dst_wb,
            dst_sheet_name= asm["final_name"],
            new_qty       = asm["qty"],
        )

    buf = io.BytesIO()
    dst_wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────────────────────

if "assemblies" not in st.session_state:
    st.session_state.assemblies   = []   # list of sheet dicts
if "conflicts_resolved" not in st.session_state:
    st.session_state.conflicts_resolved = False
if "final_names" not in st.session_state:
    st.session_state.final_names  = {}   # index → final sheet name
if "result_bytes" not in st.session_state:
    st.session_state.result_bytes = None


# ─────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero-banner">
  <h1>📋 BOM Combiner</h1>
  <p>Upload multiple BOM Excel files, review and adjust assembly quantities,
     resolve any name conflicts, and download a single Combined BOM file
     ready for use with the BOM Production Automation tool.</p>
  <div class="hero-steps">
    <div class="hero-step"><span class="step-num">1</span> Upload BOM files</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">2</span> Review quantities</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">3</span> Resolve conflicts</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">4</span> Download</div>
  </div>
  <div class="trust-badge">🔒 Your files are processed in memory and never stored on any server.</div>
</div>
""", unsafe_allow_html=True)

# ── STEP 1: Upload ───────────────────────────────────────────
st.markdown('<div class="section-header">Step 1 — Upload BOM files</div>', unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
  Upload one or more BOM Excel files (.xlsx). Each sheet in each file is treated as
  one <strong>assembly</strong>. The unit quantity is read from cell <strong>E2</strong>
  of each sheet — you can review and change it in the next step.
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Drop BOM files here",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploaded_files:
    # Re-read files whenever upload changes
    all_sheets = []
    for f in uploaded_files:
        file_bytes = f.read()
        sheets = read_bom_sheets(file_bytes, f.name)
        all_sheets.extend(sheets)
    st.session_state.assemblies = all_sheets
    st.session_state.conflicts_resolved = False
    st.session_state.result_bytes = None

assemblies = st.session_state.assemblies

if not assemblies:
    st.caption("Upload at least one BOM file to continue.")
    st.stop()

# ── STEP 2: Review & edit quantities ─────────────────────────
st.markdown('<div class="section-header">Step 2 — Review assembly quantities</div>', unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
  Below are all assemblies found across your files. The <strong>Quantity</strong> column
  shows the value from cell E2 of each sheet. You can edit it directly —
  column E formulas will recalculate automatically in Excel.
</div>
""", unsafe_allow_html=True)

# Build quantity editor
qty_values = {}
total_qty = 0

# Header row
col_name, col_source, col_qty = st.columns([3, 3, 1])
col_name.markdown("**Assembly name**")
col_source.markdown("**Source file**")
col_qty.markdown("**Qty**")
st.divider()

for i, asm in enumerate(assemblies):
    col_name, col_source, col_qty = st.columns([3, 3, 1])
    with col_name:
        st.markdown(f"📄 **{asm['sheet_name']}**")
    with col_source:
        st.caption(asm["source_file"])
    with col_qty:
        qty = st.number_input(
            label=f"qty_{i}",
            min_value=0,
            max_value=9999,
            value=asm["original_qty"],
            step=1,
            key=f"qty_{i}",
            label_visibility="collapsed",
        )
        qty_values[i] = qty
        total_qty += qty

st.divider()
col_total1, col_total2, col_total3 = st.columns([3, 3, 1])
col_total1.markdown("**Total assemblies**")
col_total2.markdown(f"**{len(assemblies)} sheets from {len(uploaded_files)} file(s)**")
col_total3.markdown(f"**{total_qty}**")

# ── STEP 3: Resolve conflicts ─────────────────────────────────
st.markdown('<div class="section-header">Step 3 — Resolve name conflicts</div>', unsafe_allow_html=True)

# Find duplicate sheet names
from collections import defaultdict
name_groups: Dict[str, List[int]] = defaultdict(list)
for i, asm in enumerate(assemblies):
    name_groups[asm["sheet_name"]].append(i)

conflicts = {name: indices for name, indices in name_groups.items() if len(indices) > 1}

# Build final_names from session state or defaults
if "rename_choices" not in st.session_state:
    st.session_state.rename_choices = {}

if not conflicts:
    st.success("✅ No conflicts — all assembly names are unique.")
    # Set final names = original names
    final_names = {i: asm["sheet_name"] for i, asm in enumerate(assemblies)}
    conflicts_ok = True
else:
    st.markdown(f"""
    <div class="warn-box">
      ⚠️ <strong>{len(conflicts)} name conflict(s) found.</strong>
      The same assembly name appears in multiple files.
      For each conflict, choose what to do.
    </div>
    """, unsafe_allow_html=True)

    final_names = {i: asm["sheet_name"] for i, asm in enumerate(assemblies)}
    conflicts_ok = True

    for name, indices in conflicts.items():
        st.markdown(f'<div class="conflict-box"><div class="conflict-title">⚠️ Conflict: "{name}" appears {len(indices)} times</div>', unsafe_allow_html=True)

        for conflict_num, idx in enumerate(indices):
            asm = assemblies[idx]
            col_info, col_action = st.columns([3, 2])

            with col_info:
                st.markdown(f"**Occurrence {conflict_num + 1}:** `{asm['source_file']}`")

            with col_action:
                choice = st.radio(
                    label=f"action_{name}_{idx}",
                    options=["Keep this name", "Rename"],
                    index=0,
                    key=f"action_{name}_{idx}",
                    horizontal=True,
                    label_visibility="collapsed",
                )

                if choice == "Rename":
                    # Suggest a name with suffix
                    suggested = f"{name}_{conflict_num + 1}"
                    new_name = st.text_input(
                        label=f"New name for occurrence {conflict_num + 1}",
                        value=suggested,
                        key=f"newname_{name}_{idx}",
                        placeholder="Enter new assembly name...",
                    )
                    final_names[idx] = new_name.strip() if new_name.strip() else suggested
                else:
                    final_names[idx] = name

        st.markdown("</div>", unsafe_allow_html=True)

    # Check if after user choices there are still duplicates
    chosen_names = list(final_names.values())
    still_dupes = [n for n in set(chosen_names) if chosen_names.count(n) > 1]
    if still_dupes:
        st.error(f"❌ Still duplicate names: {', '.join(still_dupes)}. Please rename to unique names.")
        conflicts_ok = False


# ── STEP 4: Generate ─────────────────────────────────────────
st.markdown('<div class="section-header">Step 4 — Generate Combined BOM</div>', unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
  The output is a single BOM Excel file with all assemblies as separate sheets.
  Quantities in E2 are updated to your values — column E formulas will
  recalculate automatically when you open the file in Excel.
  <br><br>
  This file is <strong>ready to use directly with Tool 1 — BOM Production Automation.</strong>
</div>
""", unsafe_allow_html=True)

# Summary before generate
st.markdown(f"""
<div class="stats-row">
  <div class="stat-card">
    <div class="stat-num">{len(uploaded_files)}</div>
    <div class="stat-label">Files uploaded</div>
  </div>
  <div class="stat-card">
    <div class="stat-num">{len(assemblies)}</div>
    <div class="stat-label">Assemblies</div>
  </div>
  <div class="stat-card">
    <div class="stat-num">{total_qty}</div>
    <div class="stat-label">Total units</div>
  </div>
</div>
""", unsafe_allow_html=True)

generate_ready = conflicts_ok and len(assemblies) > 0
generate_clicked = st.button(
    "⚡  Generate Combined BOM",
    type="primary",
    use_container_width=True,
    disabled=not generate_ready,
)

if generate_clicked and generate_ready:
    # Build final assembly list
    final_assemblies = []
    for i, asm in enumerate(assemblies):
        final_assemblies.append({
            "sheet_name":  asm["sheet_name"],
            "final_name":  final_names[i],
            "source_file": asm["source_file"],
            "file_bytes":  asm["file_bytes"],
            "qty":         qty_values[i],
        })

    with st.spinner("Building Combined BOM..."):
        try:
            result_bytes = build_combined_bom(final_assemblies)
            st.session_state.result_bytes = result_bytes
            st.success(f"✅ Combined BOM ready! {len(final_assemblies)} assemblies combined.")
        except Exception as e:
            st.error(f"❌ Error: {e}")

# Download button
if st.session_state.result_bytes:
    st.download_button(
        label="📥 Download Combined BOM.xlsx",
        data=st.session_state.result_bytes,
        file_name="Combined BOM.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # Preview table
    with st.expander("📋 Assembly summary"):
        col_h1, col_h2, col_h3, col_h4 = st.columns([2, 2, 2, 1])
        col_h1.markdown("**Original name**")
        col_h2.markdown("**Final name**")
        col_h3.markdown("**Source file**")
        col_h4.markdown("**Qty**")
        st.divider()
        for i, asm in enumerate(assemblies):
            c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
            original = asm["sheet_name"]
            final    = final_names[i]
            renamed  = original != final
            c1.markdown(f"`{original}`")
            c2.markdown(f"`{final}` {'✏️' if renamed else ''}")
            c3.caption(asm["source_file"])
            c4.markdown(f"**{qty_values[i]}**")

# Footer
st.markdown("""
<div class="footer-bar">
  <a href="https://oleksandr-horiachyi.github.io/BOM/" target="_blank">← Back to FabTools</a>
  &nbsp;·&nbsp; Built with Python &amp; Streamlit &nbsp;·&nbsp;
  <a href="https://github.com/oleksandr-horiachyi/BOM" target="_blank">GitHub</a>
</div>
""", unsafe_allow_html=True)
