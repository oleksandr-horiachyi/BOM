"""
BOM Combiner — Streamlit Web Version
FabTools | Tool 3
Drag-and-drop reordering + inline name/qty editing — all in sync.
"""

import io
from copy import copy
from typing import Dict, List

import streamlit as st
from openpyxl import Workbook, load_workbook
from streamlit_sortables import sort_items

st.set_page_config(page_title="BOM Combiner — FabTools", page_icon="📋", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Source Sans 3', sans-serif; }
.hero-banner {
    background: linear-gradient(135deg, #0C447C 0%, #185FA5 55%, #378ADD 100%);
    border-radius: 12px; padding: 32px 36px 24px; color: white; margin-bottom: 28px;
}
.hero-banner h1 { font-size: 26px; font-weight: 700; margin: 0 0 8px 0; }
.hero-banner p  { font-size: 14px; color: rgba(255,255,255,0.82); margin: 0 0 16px 0; line-height: 1.6; }
.hero-steps { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.hero-step  { display: flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 600; color: rgba(255,255,255,0.95); }
.step-num {
    width: 22px; height: 22px; border-radius: 50%;
    background: rgba(255,255,255,0.2); border: 1.5px solid rgba(255,255,255,0.4);
    font-size: 11px; font-weight: 700; display: flex; align-items: center; justify-content: center;
}
.hero-arrow { color: rgba(255,255,255,0.4); font-size: 14px; }
.trust-badge {
    display: inline-flex; align-items: center; gap: 8px;
    background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.2);
    border-radius: 8px; padding: 6px 12px; font-size: 12px; color: rgba(255,255,255,0.85); margin-top: 12px;
}
.section-header {
    font-size: 12px; font-weight: 600; letter-spacing: 1.2px; text-transform: uppercase;
    color: #185FA5; margin: 0 0 10px 0; padding-top: 8px; border-top: 2px solid #E6F1FB;
}
.info-box {
    background: #E6F1FB; border-left: 3px solid #185FA5; border-radius: 0 8px 8px 0;
    padding: 9px 12px; font-size: 13px; color: #0C447C; line-height: 1.5; margin-bottom: 10px;
}
.drag-section {
    background: #F8F9FC; border: 1px solid #DDE1EA;
    border-radius: 10px; padding: 16px; margin-bottom: 16px;
}
.drag-hint { font-size: 12px; color: #8E96A8; margin-bottom: 8px; }
.conflict-badge {
    display: inline-block; background: #FEE2E2; color: #B91C1C;
    font-size: 10px; font-weight: 700; padding: 1px 6px; border-radius: 4px;
}
.renamed-badge {
    display: inline-block; background: #FEF3C7; color: #92400E;
    font-size: 10px; font-weight: 700; padding: 1px 6px; border-radius: 4px;
}
.stats-row { display: flex; gap: 16px; margin-bottom: 14px; flex-wrap: wrap; }
.stat-card { background: #F8F9FC; border: 1px solid #DDE1EA; border-radius: 8px; padding: 12px 20px; text-align: center; }
.stat-num  { font-size: 24px; font-weight: 700; color: #185FA5; }
.stat-label{ font-size: 11px; color: #8E96A8; margin-top: 2px; }
.footer-bar{ text-align: center; font-size: 12px; color: #888; margin-top: 40px; padding-top: 20px; border-top: 1px solid #EFF1F5; }
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
        result.append({"original_name": sheet_name, "source_file": filename,
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


def build_combined_bom(rows: List[Dict]) -> bytes:
    dst_wb = Workbook()
    dst_wb.remove(dst_wb.active)
    for row in rows:
        copy_sheet_to_wb(row["file_bytes"], row["original_name"],
                         dst_wb, row["final_name"], row["qty"])
    buf = io.BytesIO()
    dst_wb.save(buf)
    return buf.getvalue()


def make_drag_label(idx: int) -> str:
    """Unique drag label that encodes the original index."""
    name = st.session_state.names[idx]
    src  = st.session_state.sheets[idx]["source_file"]
    deleted = st.session_state.deleted[idx]
    prefix = "🚫 " if deleted else ""
    return f"{prefix}{name}  —  {src}  [{idx}]"


def parse_idx_from_label(label: str) -> int:
    """Extract original index from drag label."""
    return int(label.rsplit("[", 1)[-1].rstrip("]"))


def init_state(sheets: List[Dict]):
    st.session_state.sheets  = sheets
    st.session_state.order   = list(range(len(sheets)))
    st.session_state.names   = {i: s["original_name"] for i, s in enumerate(sheets)}
    st.session_state.qtys    = {i: s["original_qty"]  for i, s in enumerate(sheets)}
    st.session_state.deleted = {i: False for i in range(len(sheets))}
    st.session_state.result  = None


def delete_row(idx):
    st.session_state.deleted[idx] = True


def restore_row(idx):
    st.session_state.deleted[idx] = False


# ── SESSION STATE ─────────────────────────────────────────────

for key, val in [("sheets", []), ("order", []), ("names", {}),
                 ("qtys", {}), ("deleted", {}), ("result", None), ("prev_files", [])]:
    if key not in st.session_state:
        st.session_state[key] = val

# ── HERO ─────────────────────────────────────────────────────

st.markdown("""
<div class="hero-banner">
  <h1>📋 BOM Combiner</h1>
  <p>Upload multiple BOM files — then drag to reorder, rename, set quantities,
     and remove duplicates. Everything in one place.</p>
  <div class="hero-steps">
    <div class="hero-step"><span class="step-num">1</span> Upload</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">2</span> Drag to reorder</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">3</span> Edit names & qty</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">4</span> Generate</div>
  </div>
  <div class="trust-badge">🔒 Files processed in memory — never stored on any server.</div>
</div>
""", unsafe_allow_html=True)

# ── STEP 1: UPLOAD ────────────────────────────────────────────

st.markdown('<div class="section-header">Step 1 — Upload BOM files</div>', unsafe_allow_html=True)
st.markdown("""
<div class="info-box">
  Upload one or more BOM Excel files (.xlsx). Each sheet = one assembly.
  Quantity is read from cell <strong>E2</strong> of each sheet.
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
        init_state(sheets)
        st.session_state.prev_files = file_names

if not st.session_state.sheets:
    st.caption("Upload at least one BOM file to continue.")
    st.stop()

sheets = st.session_state.sheets
n = len(sheets)

# ── STEP 2A: DRAG TO REORDER ──────────────────────────────────

st.markdown('<div class="section-header">Step 2 — Drag to reorder, then edit names & quantities</div>', unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
  <strong>① Drag</strong> the items below to set the sheet order in the output file.<br>
  <strong>② Edit</strong> names and quantities in the table that appears below the list.
</div>
""", unsafe_allow_html=True)

# Build current drag labels from current order
current_labels = [make_drag_label(idx) for idx in st.session_state.order]

st.markdown('<p class="drag-hint">☰ &nbsp;Drag to reorder — deleted rows shown with 🚫</p>', unsafe_allow_html=True)

sorted_labels = sort_items(current_labels, direction="vertical", key="drag_order")

# Parse the new order from sorted labels and save to session state
new_order = [parse_idx_from_label(lbl) for lbl in sorted_labels]
if new_order != st.session_state.order:
    st.session_state.order = new_order
    st.rerun()

order = st.session_state.order

# ── STEP 2B: EDIT TABLE (synced with drag order) ──────────────

st.markdown("---")

# Column headers
h_pos, h_name, h_src, h_qty, h_del = st.columns([0.5, 4, 2.5, 1, 0.8])
h_pos.markdown('<span style="font-size:11px;color:#8E96A8;font-weight:700">#</span>', unsafe_allow_html=True)
h_name.markdown('<span style="font-size:11px;color:#8E96A8;font-weight:700">FINAL NAME</span>', unsafe_allow_html=True)
h_src.markdown('<span style="font-size:11px;color:#8E96A8;font-weight:700">SOURCE FILE</span>', unsafe_allow_html=True)
h_qty.markdown('<span style="font-size:11px;color:#8E96A8;font-weight:700">QTY</span>', unsafe_allow_html=True)
h_del.markdown('<span style="font-size:11px;color:#8E96A8;font-weight:700">DEL</span>', unsafe_allow_html=True)

# Detect conflicts among active rows
active_order = [idx for idx in order if not st.session_state.deleted[idx]]
active_name_count: Dict[str, int] = {}
for idx in active_order:
    nm = st.session_state.names[idx]
    active_name_count[nm] = active_name_count.get(nm, 0) + 1
conflict_names = {nm for nm, cnt in active_name_count.items() if cnt > 1}

display_pos = 1
for idx in order:
    asm     = sheets[idx]
    deleted = st.session_state.deleted[idx]
    cur_name= st.session_state.names[idx]
    orig    = asm["original_name"]
    renamed = cur_name != orig
    conflict= (cur_name in conflict_names) and not deleted

    c_pos, c_name, c_src, c_qty, c_del = st.columns([0.5, 4, 2.5, 1, 0.8])

    with c_pos:
        if deleted:
            st.markdown('<span style="color:#ccc">—</span>', unsafe_allow_html=True)
        else:
            st.markdown(f"**{display_pos}.**")
            display_pos += 1

    with c_name:
        if deleted:
            st.markdown(f'<span style="color:#bbb;text-decoration:line-through">{cur_name}</span>',
                        unsafe_allow_html=True)
        else:
            new_name = st.text_input(
                label=f"name_{idx}",
                value=cur_name,
                key=f"name_input_{idx}",
                label_visibility="collapsed",
                help=f"Original: {orig}",
            )
            if new_name.strip() and new_name.strip() != cur_name:
                st.session_state.names[idx] = new_name.strip()
                st.rerun()

            if conflict:
                st.markdown('<span class="conflict-badge">⚠️ DUPLICATE — rename or delete</span>',
                            unsafe_allow_html=True)
            elif renamed:
                st.markdown(f'<span class="renamed-badge">✏️ renamed from "{orig}"</span>',
                            unsafe_allow_html=True)

    with c_src:
        color = "#bbb" if deleted else "#8E96A8"
        st.markdown(f'<span style="font-size:12px;color:{color}">{asm["source_file"]}</span>',
                    unsafe_allow_html=True)

    with c_qty:
        if deleted:
            st.markdown(f'<span style="color:#bbb">—</span>', unsafe_allow_html=True)
        else:
            new_qty = st.number_input(
                label=f"qty_{idx}",
                min_value=0, max_value=9999,
                value=st.session_state.qtys[idx],
                step=1, key=f"qty_input_{idx}",
                label_visibility="collapsed",
            )
            st.session_state.qtys[idx] = new_qty

    with c_del:
        if not deleted:
            if st.button("🗑", key=f"del_{idx}", help="Remove from output"):
                delete_row(idx)
                st.rerun()
        else:
            if st.button("↩", key=f"restore_{idx}", help="Restore"):
                restore_row(idx)
                st.rerun()

st.divider()

# Totals
total_qty     = sum(st.session_state.qtys[idx] for idx in active_order)
deleted_count = sum(1 for idx in order if st.session_state.deleted[idx])

c1, c2, c3 = st.columns([5, 2, 1])
info = f"**{len(active_order)} active** assemblies from **{len(uploaded_files)} file(s)**"
if deleted_count:
    info += f" &nbsp;·&nbsp; <span style='color:#aaa'>{deleted_count} excluded</span>"
c1.markdown(info, unsafe_allow_html=True)
c2.markdown("**Total units:**")
c3.markdown(f"**{total_qty}**")

# ── STEP 3: GENERATE ─────────────────────────────────────────

st.markdown('<div class="section-header">Step 3 — Generate Combined BOM</div>', unsafe_allow_html=True)

# Final conflict check
final_name_count: Dict[str, int] = {}
for idx in active_order:
    nm = st.session_state.names[idx]
    final_name_count[nm] = final_name_count.get(nm, 0) + 1
final_conflicts = {nm for nm, cnt in final_name_count.items() if cnt > 1}

if final_conflicts:
    st.error(f"❌ Duplicate names: **{', '.join(final_conflicts)}** — rename or delete before generating.")
    can_generate = False
elif len(active_order) == 0:
    st.warning("⚠️ No assemblies — restore or upload files.")
    can_generate = False
else:
    st.success(f"✅ {len(active_order)} assemblies ready — no conflicts.")
    can_generate = True

st.markdown(f"""
<div class="stats-row">
  <div class="stat-card"><div class="stat-num">{len(uploaded_files)}</div><div class="stat-label">Files</div></div>
  <div class="stat-card"><div class="stat-num">{len(active_order)}</div><div class="stat-label">Assemblies</div></div>
  <div class="stat-card"><div class="stat-num">{deleted_count}</div><div class="stat-label">Excluded</div></div>
  <div class="stat-card"><div class="stat-num">{total_qty}</div><div class="stat-label">Total units</div></div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
  All active assemblies are combined into one BOM file in the order shown above.
  <strong>Ready to use directly with Tool 1 — BOM Production Automation.</strong>
</div>
""", unsafe_allow_html=True)

if st.button("⚡  Generate Combined BOM", type="primary",
             use_container_width=True, disabled=not can_generate):
    final_rows = [
        {"original_name": sheets[idx]["original_name"],
         "final_name":    st.session_state.names[idx],
         "file_bytes":    sheets[idx]["file_bytes"],
         "qty":           st.session_state.qtys[idx]}
        for idx in active_order
    ]
    with st.spinner("Building Combined BOM..."):
        try:
            st.session_state.result = build_combined_bom(final_rows)
            st.success(f"✅ Done! {len(final_rows)} assemblies combined.")
        except Exception as e:
            st.error(f"❌ Error: {e}")

if st.session_state.result:
    st.download_button(
        label="📥 Download Combined BOM.xlsx",
        data=st.session_state.result,
        file_name="Combined BOM.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    with st.expander("📋 Final sheet list"):
        for pos, idx in enumerate(active_order, 1):
            renamed = st.session_state.names[idx] != sheets[idx]["original_name"]
            tag = f" ✏️ *(was {sheets[idx]['original_name']})*" if renamed else ""
            st.markdown(
                f"**{pos}.** `{st.session_state.names[idx]}`{tag}"
                f" — qty **{st.session_state.qtys[idx]}**"
                f" — *{sheets[idx]['source_file']}*"
            )

st.markdown("""
<div class="footer-bar">
  <a href="https://oleksandr-horiachyi.github.io/BOM/" target="_blank">← Back to FabTools</a>
  &nbsp;·&nbsp; Built with Python &amp; Streamlit &nbsp;·&nbsp;
  <a href="https://github.com/oleksandr-horiachyi/BOM" target="_blank">GitHub</a>
</div>
""", unsafe_allow_html=True)
