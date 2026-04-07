"""
Draw Maker – Streamlit UI
Run with:  streamlit run app.py"""

import io
import streamlit as st

from draw_maker import (
    read_nominations,
    generate_draw_excel,
    POWERS_OF_2,
    next_power_of_2,
    valid_group_sizes,
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Draw Maker",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
        .main-header {
            background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
            color: white;
            padding: 22px 30px;
            border-radius: 10px;
            margin-bottom: 24px;
        }
        .main-header h1 { margin: 0; font-size: 2rem; }
        .main-header p  { margin: 6px 0 0 0; opacity: 0.85; font-size: 0.95rem; }
        .step-label {
            background: #EDF4FB;
            border-left: 5px solid #ED7D31;
            padding: 8px 14px;
            border-radius: 0 6px 6px 0;
            font-weight: 700;
            font-size: 1.05rem;
            color: #1F4E79;
            margin: 24px 0 12px 0;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Hero header ───────────────────────────────────────────────────────────────
st.markdown(
    """
    <div class="main-header">
        <h1>Draw Maker</h1>
        <p>Upload the Self Nomination file → Review analysis → Configure draw → Download Excel</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Category metadata ─────────────────────────────────────────────────────────
CATEGORIES = {
    "MS":    "Men's Singles",
    "MD":    "Men's Doubles",
    "Mixed": "Mixed Doubles",
    "WS":    "Women's Singles",
    "WD":    "Women's Doubles",
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 – Upload
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="step-label">Step 1 — Upload Self Nomination File</div>',
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "Choose the Self Nomination Excel file (.xlsx)",
    type=["xlsx"],
    help="Upload the Microsoft Forms export of the Self Nomination survey.",
)

if uploaded is None:
    st.info("👆 Please upload the **Self Nomination Excel file** to continue.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# Parse & cache nominations
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="📋 Analysing nominations…")
def load_nominations(file_bytes: bytes):
    return read_nominations(io.BytesIO(file_bytes))


file_bytes = uploaded.read()
data, emp_lookup = load_nominations(file_bytes)

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 – Nomination Analysis
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="step-label">Step 2 — Nomination Analysis</div>',
    unsafe_allow_html=True,
)

total = sum(len(v) for v in data.values())
st.success(f"File loaded! **{total} total entries** found across all categories.")

metrics_cols = st.columns(5)
for i, (cat, label) in enumerate(CATEGORIES.items()):
    n = len(data[cat])
    suggested = next_power_of_2(n) if n > 0 else "–"
    with metrics_cols[i]:
        st.metric(
            label=label,
            value=f"{n} {'entry' if n == 1 else 'entries'}",
            delta=f"Suggest draw: {suggested}" if n > 0 else "No entries",
        )

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 – Configure Draw Settings
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="step-label">Step 3 — Configure Draw Settings</div>',
    unsafe_allow_html=True,
)
st.caption(
    "💡 **Draw Size** must be a power of 2 (4 · 8 · 16 · 32 · 64 · 128 · 256).  "
    "**Group Size** must be even and divide the Draw Size exactly.  \n"
    "Choosing a **smaller draw** than the number of entries → extra players are "
    "**waitlisted** (listed in a separate shaded section at the bottom of the sheet).  "
    "Choosing a **larger draw** → empty slots become **BYE** entries."
)

draw_configs = {}

for cat, label in CATEGORIES.items():
    players = data[cat]
    n = len(players)
    if n == 0:
        continue

    with st.container(border=True):
        st.markdown(f"**{label}** — {n} {'entry' if n == 1 else 'entries'}")

        col_ds, col_gs, col_status = st.columns([2, 2, 3])

        suggested_size = next_power_of_2(n)
        # Offer draw sizes from 4 to twice the suggested (capped at 256)
        max_offered = min(max(suggested_size * 2, 128), 256)
        avail_sizes = [p for p in POWERS_OF_2 if p <= max_offered]
        default_ds_idx = (
            avail_sizes.index(suggested_size)
            if suggested_size in avail_sizes
            else len(avail_sizes) - 1
        )

        with col_ds:
            draw_size = st.selectbox(
                "Draw Size",
                options=avail_sizes,
                index=default_ds_idx,
                key=f"ds_{cat}",
                help="Total bracket slots including BYEs. Must be a power of 2.",
                format_func=lambda x, s=suggested_size: (
                    f"{x}  ✅ (suggested)" if x == s else str(x)
                ),
            )

        gs_options = valid_group_sizes(draw_size)
        default_gs = 16 if 16 in gs_options else gs_options[-1]

        with col_gs:
            group_size = st.selectbox(
                "Group Size",
                options=gs_options,
                index=gs_options.index(default_gs),
                key=f"gs_{cat}",
                help="Players per round-robin group. Must be even and divide Draw Size.",
            )

        byes       = max(0, draw_size - n)
        waitlisted = max(0, n - draw_size)
        num_groups = draw_size // group_size

        with col_status:
            if waitlisted > 0:
                st.warning(
                    f"⚠️ **{num_groups} group(s) of {group_size}**  \n"
                    f"**{waitlisted}** player(s) will be **waitlisted** "
                    f"(shown separately in Excel — coordinator decides)."
                )
            elif byes > 0:
                st.info(
                    f"✅ **{num_groups} group(s) of {group_size}**  \n"
                    f"**{byes}** BYE slot(s) will be added to fill the bracket."
                )
            else:
                st.success(
                    f"✅ **{num_groups} group(s) of {group_size}**  \n"
                    f"Perfect fit — no BYEs and no waitlisted players!"
                )

        draw_configs[cat] = {
            "entries":    players,
            "draw_size":  draw_size,
            "group_size": group_size,
        }

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 – Review Summary
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="step-label">Step 4 — Review & Generate Draw</div>',
    unsafe_allow_html=True,
)

if not draw_configs:
    st.warning("No categories with entries were detected. Please check the file.")
    st.stop()

# Show a compact review table
review_rows = []
for cat, label in CATEGORIES.items():
    if cat not in draw_configs:
        continue
    cfg = draw_configs[cat]
    n = len(cfg["entries"])
    ds = cfg["draw_size"]
    gs = cfg["group_size"]
    byes = max(0, ds - n)
    wl   = max(0, n - ds)
    review_rows.append({
        "Category":   label,
        "Entries":    n,
        "Draw Size":  ds,
        "Groups":     ds // gs,
        "Group Size": gs,
        "BYEs":       byes,
        "Waitlisted": wl,
    })

st.dataframe(review_rows, use_container_width=True, hide_index=True)

st.divider()

col_btn, col_hint = st.columns([1, 3])

with col_btn:
    generate_clicked = st.button(
        "🎲  Make a Draw",
        type="primary",
        use_container_width=True,
    )

with col_hint:
    st.caption(
        "Each click produces a **different random arrangement**. "
        "Click again to regenerate if you're not satisfied with the draw."
    )

if generate_clicked:
    with st.spinner("Generating draw…"):
        excel_bytes = generate_draw_excel(draw_configs)
    st.session_state["draw_bytes"] = excel_bytes
    st.session_state["draw_ready"] = True
    st.success("✅ Draw generated successfully!")

if st.session_state.get("draw_ready"):
    st.download_button(
        label="📥  Download Draw Excel",
        data=st.session_state["draw_bytes"],
        file_name="TT Cybage Internal 2026 Draws.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
