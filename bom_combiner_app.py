"""
BOM Combiner — Streamlit Web Version
FabTools | Tool 3
"""

import io
from collections import defaultdict
from copy import copy
from typing import Dict, List

import streamlit as st
from openpyxl import Workbook, load_workbook
from streamlit_sortables import sort_items

st.set_page_config(page_title="BOM Combiner — FabTools", page_icon="📋", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Source Sans 3', sans-serif; }
.hero-banner {
    background: linear-gradient(135deg, #0C447C 0%, #185FA5 55%, #378ADD 100%);
    border-radius: 12px; padding: 36px 36px 28px; color: white; margin-bottom: 32px;
}
.hero-banner h1 { font-size: 28px; font-weight: 700; margin: 0 0 8px 0; }
.hero-banner p  { font-size: 15px; color: rgba(255,255,255,0.82); margin: 0 0 20px 0; line-height: 1.6; }
.hero-steps { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.hero-step  { display: flex; align-items: center; gap: 8px; font-size: 14px; font-weight: 600; color: rgba(255,255,255,0.95); }
.step-num {
    width: 24px; height: 24px; border-radius: 50%;
    background: rgba(255,255,255,0.2); border: 1.5px solid rgba(255,255,255,0.4);
    font-size: 12px; font-weight: 700; display: flex; align-items: center; justify-content: center;
}
.hero-arrow { color: rgba(255,255,255,0.4); font-size: 16px; }
.trust-badge {
    display: inline-flex; align-items: center; gap: 8px;
    background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.2);
    border-radius: 8px; padding: 7px 14px; font-size: 13px; color: rgba(255,255,255,0.85); margin-top: 16px;
}
.section-header {
    font-size: 13px; font-weight: 600; letter-spacing: 1.2px; text-transform: uppercase;
    color: #185FA5; margin: 0 0 12px 0; padding-top: 8px; border-top: 2px solid #E6F1FB;
}
.info-box {
    background: #E6F1FB; border-left: 3px solid #185FA5; border-radius: 0 8px 8px 0;
    padding: 10px 14px; font-size: 13px; color: #0C447C; line-height: 1.55; margin-bottom: 12px;
}
.warn-box {
    background: #FFF8E1; border-left: 3px solid #F59E0B; border-radius: 0 8px 8px 0;
    padding: 10px 14px; font-size: 13px; color: #92400E; line-height: 1.55; margin-bottom: 12px;
}
.conflict-box { background: #FFF3CD; border: 1px solid #F59E0B; border-radius: 8px; padding: 16px; margin-bottom: 12px; }
.conflict-title { font-weight: 700; color: #92400E; margin-bottom: 10px; font-size: 14px; }
.renamed-tag { font-size: 10px; background: #FFF3CD; color: #92400E; padding: 1px 6px; border-radius: 4px; margin-left: 6px; }
.qty-row {
    display: flex; align-items: center; gap: 12px;
    background: #F8F9FC; border: 1px solid #DDE1EA;
    border-radius: 8px; padding: 10px 16px; margin-bottom: 6px;
}
.qty-name { font-size: 14px; font-weight: 600; color: #1A1F2E; flex: 1; }
.qty-src  { font-size: 11px; color: #8E96A8; }
.stats-row { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 16px; }
.stat-card { background: #F8F9FC; border: 1px solid #DDE1EA; border-radius: 8px; padding: 14px; text-align: center; }
.stat-num  { font-size: 28px; font-weight: 700; color: #185FA5; }
.stat-label{ font-size: 12px; color: #8E96A8; margin-top: 2px; }
.footer-bar{ text-align: center; font-size: 12px; color: #888; margin-top: 40px; padding-top: 20px; border-top: 1px solid #EFF1F5; }
.drag-hint { font-size: 12px; color: #8E96A8; margin-bottom: 8px; }
</style>
""", unsafe_allow_html=True)

# ── HELPERS ──────────────────────────────────────────────────

def read_bom_sheets(file_bytes: bytes, filename: str) -> List[Dict]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    result = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        raw = ws["E2"].value
        try:
            qty = int(float(str(raw))) if raw is not None else 1
        except:
            qty = 1
        result.append({"sheet_name": sheet_name, "source_file": filename,
                        "original_qty": qty, "file_bytes": file_bytes})
    return result


def copy_sheet_to_wb(src_bytes, src_name, dst_wb, dst_name, new_qty):
    src_wb = load_workbook(io.BytesIO(src_bytes), data_only=False)
    src_ws = src_wb[src_name]
    dst_ws = dst_wb.create_sheet(dst_name)
    for col, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col].width = dim.width
    for row, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row].height = dim.height
    for row in src_ws.iter_rows():
        for cell in row:
            nc = dst_ws.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                nc.font = copy(cell.font); nc.border = copy(cell.border)
                nc.fill = copy(cell.fill); nc.number_format = cell.number_format
                nc.protection = copy(cell.protection); nc.alignment = copy(cell.alignment)
    for merged in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged))
    dst_ws["E2"].value = new_qty
    if src_name != dst_name:
        a1 = dst_ws["A1"]
        if a1.value and src_name in str(a1.value):
            a1.value = str(a1.value).replace(src_name, dst_name)


def build_combined_bom(assemblies: List[Dict]) -> bytes:
    dst_wb = Workbook()
    dst_wb.remove(dst_wb.active)
    for asm in assemblies:
        copy_sheet_to_wb(asm["file_bytes"], asm["sheet_name"],
                         dst_wb, asm["final_name"], asm["qty"])
    buf = io.BytesIO()
    dst_wb.save(buf)
    return buf.getvalue()


def make_label(idx: int, assemblies, final_names) -> str:
    """Create a unique drag-and-drop label for each assembly."""
    asm   = assemblies[idx]
    fname = final_names[idx]
    renamed = fname != asm["sheet_name"]
    tag   = " ✏️" if renamed else ""
    return f"{fname}{tag}  |  {asm['source_file']}"


def parse_label(label: str) -> str:
    """Extract final name from drag label."""
    return label.split("  |  ")[0].replace(" ✏️", "").strip()


# ── SESSION STATE ─────────────────────────────────────────────

for key, val in [("assemblies", []), ("result_bytes", None), ("prev_files", [])]:
    if key not in st.session_state:
        st.session_state[key] = val

# ── HERO ─────────────────────────────────────────────────────

st.markdown("""
<div class="hero-banner">
  <h1>📋 BOM Combiner</h1>
  <p>Upload multiple BOM Excel files, review quantities, drag to reorder assemblies,
     resolve name conflicts, and download a single Combined BOM ready for Tool 1.</p>
  <div class="hero-steps">
    <div class="hero-step"><span class="step-num">1</span> Upload</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">2</span> Resolve conflicts</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">3</span> Drag to reorder & set qty</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">4</span> Download</div>
  </div>
  <div class="trust-badge">🔒 Your files are processed in memory and never stored on any server.</div>
</div>
""", unsafe_allow_html=True)

# ── STEP 1: UPLOAD ────────────────────────────────────────────

st.markdown('<div class="section-header">Step 1 — Upload BOM files</div>', unsafe_allow_html=True)
st.markdown("""
<div class="info-box">
  Upload one or more BOM Excel files (.xlsx). Each sheet = one assembly.
  Unit quantity is read from cell <strong>E2</strong> of each sheet.
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Drop BOM files here", type=["xlsx"],
    accept_multiple_files=True, label_visibility="collapsed",
)

if uploaded_files:
    file_names = [f.name for f in uploaded_files]
    if file_names != st.session_state.prev_files:
        sheets = []
        for f in uploaded_files:
            sheets.extend(read_bom_sheets(f.read(), f.name))
        st.session_state.assemblies   = sheets
        st.session_state.result_bytes = None
        st.session_state.prev_files   = file_names

assemblies = st.session_state.assemblies
if not assemblies:
    st.caption("Upload at least one BOM file to continue.")
    st.stop()

n = len(assemblies)

# ── STEP 2: CONFLICTS ─────────────────────────────────────────

st.markdown('<div class="section-header">Step 2 — Resolve name conflicts</div>', unsafe_allow_html=True)

name_groups: Dict[str, List[int]] = defaultdict(list)
for i, asm in enumerate(assemblies):
    name_groups[asm["sheet_name"]].append(i)
conflicts = {name: idxs for name, idxs in name_groups.items() if len(idxs) > 1}

final_names: Dict[int, str] = {i: asm["sheet_name"] for i, asm in enumerate(assemblies)}
conflicts_ok = True

if not conflicts:
    st.success("✅ No conflicts — all assembly names are unique.")
else:
    st.markdown(f"""
    <div class="warn-box">⚠️ <strong>{len(conflicts)} name conflict(s) found.</strong>
    The same assembly name appears in multiple files.</div>
    """, unsafe_allow_html=True)

    for name, indices in conflicts.items():
        st.markdown(f'<div class="conflict-box"><div class="conflict-title">⚠️ Conflict: "{name}" — {len(indices)} occurrences</div></div>', unsafe_allow_html=True)
        for conflict_num, idx in enumerate(indices):
            asm = assemblies[idx]
            col_info, col_action = st.columns([3, 2])
            with col_info:
                st.markdown(f"**Occurrence {conflict_num + 1}:** `{asm['source_file']}`")
            with col_action:
                choice = st.radio(
                    label=f"c_{name}_{idx}",
                    options=["Keep this name", "Rename"],
                    index=0 if conflict_num == 0 else 1,
                    key=f"choice_{name}_{idx}",
                    horizontal=True, label_visibility="collapsed",
                )
            if choice == "Rename":
                suggested = f"{name}_{conflict_num + 1}"
                new_name = st.text_input(
                    f"New name — occurrence {conflict_num + 1}",
                    value=suggested, key=f"newname_{name}_{idx}",
                )
                final_names[idx] = new_name.strip() or suggested
            else:
                final_names[idx] = name

    chosen = list(final_names.values())
    still_dupes = [nm for nm in set(chosen) if chosen.count(nm) > 1]
    if still_dupes:
        st.error(f"❌ Still duplicate names: {', '.join(still_dupes)}. Please use unique names.")
        conflicts_ok = False

# ── STEP 3: DRAG TO REORDER + QTY ────────────────────────────

st.markdown('<div class="section-header">Step 3 — Reorder assemblies & set quantities</div>', unsafe_allow_html=True)

# ── 3A: DRAG AND DROP ────────────────────────────────────────
st.markdown("""
<div class="info-box">
  <strong>Drag and drop</strong> the assemblies below to set the sheet order in the output file.
  Then set the quantity for each assembly in the table below.
</div>
""", unsafe_allow_html=True)

st.markdown('<p class="drag-hint">☰ Drag items to reorder</p>', unsafe_allow_html=True)

# Build labels for sortable (must be unique strings)
labels = [make_label(i, assemblies, final_names) for i in range(n)]

sorted_labels = sort_items(labels, direction="vertical", key="sortable_assemblies")

# Map sorted labels back to assembly indices
label_to_idx = {make_label(i, assemblies, final_names): i for i in range(n)}
sorted_order = [label_to_idx.get(lbl, i) for i, lbl in enumerate(sorted_labels)]

# ── 3B: QTY TABLE ────────────────────────────────────────────
st.markdown("**Quantities:**")

# Header
ch1, ch2, ch3 = st.columns([1, 4, 1])
ch1.markdown("**#**")
ch2.markdown("**Assembly**")
ch3.markdown("**Qty**")
st.divider()

qty_values: Dict[int, int] = {}
for pos, idx in enumerate(sorted_order):
    asm   = assemblies[idx]
    fname = final_names[idx]
    renamed = fname != asm["sheet_name"]

    c1, c2, c3 = st.columns([1, 4, 1])
    with c1:
        st.markdown(f"**{pos + 1}.**")
    with c2:
        tag = ' <span class="renamed-tag">renamed</span>' if renamed else ""
        st.markdown(
            f'<div class="qty-row"><div>'
            f'<div class="qty-name">{fname}{tag}</div>'
            f'<div class="qty-src">{asm["source_file"]}'
            + (f' · original: {asm["sheet_name"]}' if renamed else "")
            + f'</div></div></div>',
            unsafe_allow_html=True,
        )
    with c3:
        qty_values[idx] = st.number_input(
            label=f"qty_{idx}",
            min_value=0, max_value=9999,
            value=asm["original_qty"], step=1,
            key=f"qty_{idx}",
            label_visibility="collapsed",
        )

st.divider()
total_qty = sum(qty_values.values())
ct1, ct2, ct3 = st.columns([5, 2, 1])
ct1.markdown(f"**{n} assemblies** from **{len(uploaded_files)} file(s)**")
ct2.markdown("**Total units:**")
ct3.markdown(f"**{total_qty}**")

# ── STEP 4: GENERATE ─────────────────────────────────────────

st.markdown('<div class="section-header">Step 4 — Generate Combined BOM</div>', unsafe_allow_html=True)
st.markdown("""
<div class="info-box">
  All assemblies are combined into one BOM file in the order above.
  <strong>Ready to use directly with Tool 1 — BOM Production Automation.</strong>
</div>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="stats-row">
  <div class="stat-card"><div class="stat-num">{len(uploaded_files)}</div><div class="stat-label">Files uploaded</div></div>
  <div class="stat-card"><div class="stat-num">{n}</div><div class="stat-label">Assemblies</div></div>
  <div class="stat-card"><div class="stat-num">{total_qty}</div><div class="stat-label">Total units</div></div>
</div>
""", unsafe_allow_html=True)

if st.button("⚡  Generate Combined BOM", type="primary",
             use_container_width=True, disabled=not conflicts_ok):
    final_assemblies = [
        {"sheet_name": assemblies[idx]["sheet_name"],
         "final_name": final_names[idx],
         "source_file": assemblies[idx]["source_file"],
         "file_bytes":  assemblies[idx]["file_bytes"],
         "qty":         qty_values[idx]}
        for idx in sorted_order
    ]
    with st.spinner("Building Combined BOM..."):
        try:
            st.session_state.result_bytes = build_combined_bom(final_assemblies)
            st.success(f"✅ Done! {n} assemblies combined.")
        except Exception as e:
            st.error(f"❌ Error: {e}")

if st.session_state.result_bytes:
    st.download_button(
        label="📥 Download Combined BOM.xlsx",
        data=st.session_state.result_bytes,
        file_name="Combined BOM.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    with st.expander("📋 Final sheet order"):
        for pos, idx in enumerate(sorted_order, 1):
            renamed = final_names[idx] != assemblies[idx]["sheet_name"]
            tag = " ✏️" if renamed else ""
            st.markdown(
                f"**{pos}.** `{final_names[idx]}`{tag}"
                f" &nbsp; qty **{qty_values[idx]}**"
                f" &nbsp; *{assemblies[idx]['source_file']}*"
            )

st.markdown("""
<div class="footer-bar">
  <a href="https://oleksandr-horiachyi.github.io/BOM/" target="_blank">← Back to FabTools</a>
  &nbsp;·&nbsp; Built with Python &amp; Streamlit &nbsp;·&nbsp;
  <a href="https://github.com/oleksandr-horiachyi/BOM" target="_blank">GitHub</a>
</div>
""", unsafe_allow_html=True)
