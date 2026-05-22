"""
BOM Combiner — Streamlit Web Version
FabTools | Tool 3
Combines multiple BOM Excel files into one ready-to-use Combined BOM file.
"""

import io
from collections import defaultdict
from copy import copy
from typing import Dict, List

import streamlit as st
from openpyxl import Workbook, load_workbook

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="BOM Combiner — FabTools",
    page_icon="📋",
    layout="centered",
)

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
.hero-steps     { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.hero-step      { display: flex; align-items: center; gap: 8px; font-size: 14px; font-weight: 600; color: rgba(255,255,255,0.95); }
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
.stats-row { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 16px; }
.stat-card { background: #F8F9FC; border: 1px solid #DDE1EA; border-radius: 8px; padding: 14px; text-align: center; }
.stat-num   { font-size: 28px; font-weight: 700; color: #185FA5; }
.stat-label { font-size: 12px; color: #8E96A8; margin-top: 2px; }
.footer-bar { text-align: center; font-size: 12px; color: #888; margin-top: 40px; padding-top: 20px; border-top: 1px solid #EFF1F5; }
.conflict-box { background: #FFF3CD; border: 1px solid #F59E0B; border-radius: 8px; padding: 16px; margin-bottom: 12px; }
.conflict-title { font-weight: 700; color: #92400E; margin-bottom: 8px; font-size: 14px; }
.renamed-tag { font-size: 11px; background: #FFF3CD; color: #92400E; padding: 1px 7px; border-radius: 4px; margin-left: 6px; }
.order-preview {
    background: #F0F9F5; border: 1px solid #A8D5BF; border-radius: 8px;
    padding: 14px 18px; margin-bottom: 16px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def read_bom_sheets(file_bytes: bytes, filename: str) -> List[Dict]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheets = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        raw = ws["E2"].value
        try:
            qty = int(float(str(raw))) if raw is not None else 1
        except:
            qty = 1
        sheets.append({
            "sheet_name":   sheet_name,
            "source_file":  filename,
            "original_qty": qty,
            "file_bytes":   file_bytes,
        })
    return sheets


def copy_sheet_to_workbook(src_wb_bytes, src_sheet_name, dst_wb, dst_sheet_name, new_qty):
    src_wb = load_workbook(io.BytesIO(src_wb_bytes), data_only=False)
    src_ws = src_wb[src_sheet_name]
    dst_ws = dst_wb.create_sheet(dst_sheet_name)

    for col, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col].width = dim.width
    for row, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row].height = dim.height

    for row in src_ws.iter_rows():
        for cell in row:
            nc = dst_ws.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                nc.font          = copy(cell.font)
                nc.border        = copy(cell.border)
                nc.fill          = copy(cell.fill)
                nc.number_format = cell.number_format
                nc.protection    = copy(cell.protection)
                nc.alignment     = copy(cell.alignment)

    for merged in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged))

    dst_ws["E2"].value = new_qty

    # Update title in A1 if it contains the old name
    if src_sheet_name != dst_sheet_name:
        a1 = dst_ws["A1"]
        if a1.value and src_sheet_name in str(a1.value):
            a1.value = str(a1.value).replace(src_sheet_name, dst_sheet_name)


def build_combined_bom(assemblies: List[Dict]) -> bytes:
    dst_wb = Workbook()
    dst_wb.remove(dst_wb.active)
    for asm in assemblies:
        copy_sheet_to_workbook(
            asm["file_bytes"], asm["sheet_name"],
            dst_wb, asm["final_name"], asm["qty"]
        )
    buf = io.BytesIO()
    dst_wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────

if "assemblies"    not in st.session_state: st.session_state.assemblies    = []
if "result_bytes"  not in st.session_state: st.session_state.result_bytes  = None
if "prev_files"    not in st.session_state: st.session_state.prev_files    = []

# ─────────────────────────────────────────────────────────────
# HERO
# ─────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero-banner">
  <h1>📋 BOM Combiner</h1>
  <p>Upload multiple BOM Excel files, review and adjust quantities,
     set the sheet order, resolve name conflicts, and download a
     single Combined BOM ready for Tool 1.</p>
  <div class="hero-steps">
    <div class="hero-step"><span class="step-num">1</span> Upload</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">2</span> Resolve conflicts</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">3</span> Set order & qty</div>
    <span class="hero-arrow">→</span>
    <div class="hero-step"><span class="step-num">4</span> Download</div>
  </div>
  <div class="trust-badge">🔒 Your files are processed in memory and never stored on any server.</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# STEP 1 — UPLOAD
# ─────────────────────────────────────────────────────────────

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
        all_sheets = []
        for f in uploaded_files:
            fb = f.read()
            all_sheets.extend(read_bom_sheets(fb, f.name))
        st.session_state.assemblies   = all_sheets
        st.session_state.result_bytes = None
        st.session_state.prev_files   = file_names

assemblies = st.session_state.assemblies

if not assemblies:
    st.caption("Upload at least one BOM file to continue.")
    st.stop()

n = len(assemblies)

# ─────────────────────────────────────────────────────────────
# STEP 2 — RESOLVE CONFLICTS
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-header">Step 2 — Resolve name conflicts</div>', unsafe_allow_html=True)

name_groups: Dict[str, List[int]] = defaultdict(list)
for i, asm in enumerate(assemblies):
    name_groups[asm["sheet_name"]].append(i)
conflicts = {name: idxs for name, idxs in name_groups.items() if len(idxs) > 1}

# final_names: index → current final name
final_names: Dict[int, str] = {i: asm["sheet_name"] for i, asm in enumerate(assemblies)}

conflicts_ok = True

if not conflicts:
    st.success("✅ No conflicts — all assembly names are unique.")
else:
    st.markdown(f"""
    <div class="warn-box">
      ⚠️ <strong>{len(conflicts)} name conflict(s) found.</strong>
      The same assembly name appears in multiple files. Choose what to do for each.
    </div>
    """, unsafe_allow_html=True)

    for name, indices in conflicts.items():
        with st.container():
            st.markdown(f'<div class="conflict-box"><div class="conflict-title">⚠️ Conflict: "{name}" appears {len(indices)} time(s)</div></div>', unsafe_allow_html=True)

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
                        horizontal=True,
                        label_visibility="collapsed",
                    )
                if choice == "Rename":
                    suggested = f"{name}_{conflict_num + 1}"
                    new_name = st.text_input(
                        label=f"New name — occurrence {conflict_num + 1}",
                        value=suggested,
                        key=f"newname_{name}_{idx}",
                    )
                    final_names[idx] = new_name.strip() or suggested
                else:
                    final_names[idx] = name

    # Check for remaining duplicates
    chosen = list(final_names.values())
    still_dupes = [nm for nm in set(chosen) if chosen.count(nm) > 1]
    if still_dupes:
        st.error(f"❌ Still duplicate names after renaming: {', '.join(still_dupes)}. Please use unique names.")
        conflicts_ok = False

# ─────────────────────────────────────────────────────────────
# STEP 3 — ORDER & QTY TABLE
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-header">Step 3 — Set sheet order & quantities</div>', unsafe_allow_html=True)
st.markdown("""
<div class="info-box">
  Set the <strong>order</strong> (position in the output file) and <strong>quantity</strong>
  for each assembly. The <em>Final name</em> column already reflects any renames from Step 2.
  Sheets with the same order number will be sorted alphabetically.
</div>
""", unsafe_allow_html=True)

# Table header
c_ord, c_final, c_orig, c_src, c_qty = st.columns([1, 3, 2, 2, 1])
c_ord.markdown("**Order**")
c_final.markdown("**Final name**")
c_orig.markdown("**Original**")
c_src.markdown("**Source file**")
c_qty.markdown("**Qty**")
st.divider()

qty_values  = {}
order_values = {}

for i, asm in enumerate(assemblies):
    c_ord, c_final, c_orig, c_src, c_qty = st.columns([1, 3, 2, 2, 1])

    with c_ord:
        order_values[i] = st.number_input(
            label=f"order_{i}", min_value=1, max_value=n,
            value=i + 1, step=1, key=f"ord_{i}",
            label_visibility="collapsed",
        )

    with c_final:
        fname = final_names[i]
        renamed = fname != asm["sheet_name"]
        tag = ' <span class="renamed-tag">renamed</span>' if renamed else ""
        st.markdown(f"**{fname}**{tag}", unsafe_allow_html=True)

    with c_orig:
        st.caption(asm["sheet_name"])

    with c_src:
        st.caption(asm["source_file"])

    with c_qty:
        qty_values[i] = st.number_input(
            label=f"qty_{i}", min_value=0, max_value=9999,
            value=asm["original_qty"], step=1, key=f"qty_{i}",
            label_visibility="collapsed",
        )

st.divider()

# Totals row
total_qty = sum(qty_values.values())
ct1, ct2, ct3, ct4, ct5 = st.columns([1, 3, 2, 2, 1])
ct2.markdown("**Total assemblies**")
ct3.markdown(f"**{n} sheets**")
ct4.markdown(f"**{len(uploaded_files)} file(s)**")
ct5.markdown(f"**{total_qty}**")

# Live order preview
sorted_indices = sorted(range(n), key=lambda i: (order_values[i], final_names[i]))
st.markdown("**Sheet order preview:**")
preview_parts = [f"`{i+1}.` {final_names[idx]}  ×{qty_values[idx]}" for i, idx in enumerate(sorted_indices)]
st.markdown("&nbsp;&nbsp;→&nbsp;&nbsp;".join(preview_parts))

# ─────────────────────────────────────────────────────────────
# STEP 4 — GENERATE
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-header">Step 4 — Generate Combined BOM</div>', unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
  The output is a single BOM Excel file with all assemblies as separate sheets
  in the order you set above. <strong>Ready to use directly with Tool 1.</strong>
</div>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="stats-row">
  <div class="stat-card"><div class="stat-num">{len(uploaded_files)}</div><div class="stat-label">Files uploaded</div></div>
  <div class="stat-card"><div class="stat-num">{n}</div><div class="stat-label">Assemblies</div></div>
  <div class="stat-card"><div class="stat-num">{total_qty}</div><div class="stat-label">Total units</div></div>
</div>
""", unsafe_allow_html=True)

generate_ready = conflicts_ok and n > 0
if st.button("⚡  Generate Combined BOM", type="primary", use_container_width=True, disabled=not generate_ready):
    # Build sorted assembly list
    final_assemblies = [
        {
            "sheet_name": assemblies[idx]["sheet_name"],
            "final_name": final_names[idx],
            "source_file": assemblies[idx]["source_file"],
            "file_bytes":  assemblies[idx]["file_bytes"],
            "qty":         qty_values[idx],
        }
        for idx in sorted_indices
    ]
    with st.spinner("Building Combined BOM..."):
        try:
            st.session_state.result_bytes = build_combined_bom(final_assemblies)
            st.success(f"✅ Combined BOM ready! {n} assemblies in the correct order.")
        except Exception as e:
            st.error(f"❌ Error: {e}")

# Download
if st.session_state.result_bytes:
    st.download_button(
        label="📥 Download Combined BOM.xlsx",
        data=st.session_state.result_bytes,
        file_name="Combined BOM.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    with st.expander("📋 Final assembly list"):
        for pos, idx in enumerate(sorted_indices, 1):
            renamed = final_names[idx] != assemblies[idx]["sheet_name"]
            tag = " ✏️ renamed" if renamed else ""
            st.markdown(
                f"**{pos}.** `{final_names[idx]}`{tag} — "
                f"qty **{qty_values[idx]}** — "
                f"*{assemblies[idx]['source_file']}*"
            )

st.markdown("""
<div class="footer-bar">
  <a href="https://oleksandr-horiachyi.github.io/BOM/" target="_blank">← Back to FabTools</a>
  &nbsp;·&nbsp; Built with Python &amp; Streamlit &nbsp;·&nbsp;
  <a href="https://github.com/oleksandr-horiachyi/BOM" target="_blank">GitHub</a>
</div>
""", unsafe_allow_html=True)
