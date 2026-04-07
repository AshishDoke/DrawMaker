"""
Draw Maker – Streamlit UI
Run with:  streamlit run app.py"""

import io
import streamlit as st

from draw_maker import (
    read_nominations,
    read_draw_file,
    generate_draw_excel,
    generate_seeded_draw_excel,
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
        <p>Create tournament draws from nominations &nbsp;|&nbsp; Apply seeding to an existing draw</p>
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


# ── Shared helper: draw-config widgets ───────────────────────────────────────
def build_draw_config_ui(cat, label, n, key_prefix):
    col_ds, col_gs, col_status = st.columns([2, 2, 3])

    suggested_size = next_power_of_2(n)
    max_offered    = min(max(suggested_size * 2, 128), 256)
    avail_sizes    = [p for p in POWERS_OF_2 if p <= max_offered]
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
            key=f"{key_prefix}_ds_{cat}",
            help="Total bracket slots including BYEs. Must be a power of 2.",
            format_func=lambda x, s=suggested_size: (
                f"{x}  \u2705 (suggested)" if x == s else str(x)
            ),
        )

    gs_options = valid_group_sizes(draw_size)
    default_gs = 16 if 16 in gs_options else gs_options[-1]

    with col_gs:
        group_size = st.selectbox(
            "Group Size",
            options=gs_options,
            index=gs_options.index(default_gs),
            key=f"{key_prefix}_gs_{cat}",
            help="Players per round-robin group. Must be even and divide Draw Size.",
        )

    byes_n       = max(0, draw_size - n)
    waitlisted_n = max(0, n - draw_size)
    num_groups   = draw_size // group_size

    with col_status:
        if waitlisted_n > 0:
            st.warning(
                f"\u26a0\ufe0f **{num_groups} group(s) of {group_size}**  \n"
                f"**{waitlisted_n}** player(s) will be **waitlisted** "
                f"(shown separately in Excel)."
            )
        elif byes_n > 0:
            st.info(
                f"\u2705 **{num_groups} group(s) of {group_size}**  \n"
                f"**{byes_n}** BYE slot(s) will be added to fill the bracket."
            )
        else:
            st.success(
                f"\u2705 **{num_groups} group(s) of {group_size}**  \n"
                f"Perfect fit \u2014 no BYEs and no waitlisted players!"
            )

    return draw_size, group_size


# ── Cached parsers ────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="\U0001f4cb Analysing nominations\u2026")
def load_nominations(file_bytes: bytes):
    return read_nominations(io.BytesIO(file_bytes))


@st.cache_data(show_spinner="\U0001f4ca Reading draw file\u2026")
def load_draw_file(file_bytes: bytes):
    return read_draw_file(io.BytesIO(file_bytes))


# ── Tabs ──────────────────────────────────────────────────────────────────────


# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 – CREATE FROM NOMINATIONS
# ═════════════════════════════════════════════════════════════════════════════
def render_tab1():
    # Step 1 – Upload ─────────────────────────────────────────────────────────
    st.markdown(
        '<div class="step-label">Step 1 \u2014 Upload Self Nomination File</div>',
        unsafe_allow_html=True,
    )
    uploaded = st.file_uploader(
        "Choose the Self Nomination Excel file (.xlsx)",
        type=["xlsx"],
        key="nom_upload",
        help="Upload the Microsoft Forms export of the Self Nomination survey.",
    )

    if uploaded is None:
        st.info("\U0001f446 Please upload the **Self Nomination Excel file** to continue.")
        return

    file_bytes      = uploaded.read()
    data, _emp_lkup = load_nominations(file_bytes)

    # Step 2 – Analysis ───────────────────────────────────────────────────────
    st.markdown(
        '<div class="step-label">Step 2 \u2014 Nomination Analysis</div>',
        unsafe_allow_html=True,
    )
    total = sum(len(v) for v in data.values())
    st.success(f"File loaded! **{total} total entries** found across all categories.")

    metrics_cols = st.columns(5)
    for i, (cat, label) in enumerate(CATEGORIES.items()):
        n         = len(data[cat])
        suggested = next_power_of_2(n) if n > 0 else "\u2013"
        with metrics_cols[i]:
            st.metric(
                label=label,
                value=f"{n} {'entry' if n == 1 else 'entries'}",
                delta=f"Suggest draw: {suggested}" if n > 0 else "No entries",
            )

    # Step 3 – Configure ──────────────────────────────────────────────────────
    st.markdown(
        '<div class="step-label">Step 3 \u2014 Configure Draw Settings</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "\U0001f4a1 **Draw Size** must be a power of 2 (4 \u00b7 8 \u00b7 16 \u00b7 32 \u00b7 64 \u00b7 128 \u00b7 256).  "
        "**Group Size** must be even and divide the Draw Size exactly.  \n"
        "Smaller draw than entries \u2192 extra players **waitlisted**.  "
        "Larger draw \u2192 empty slots become **BYE** entries."
    )

    nom_draw_configs = {}
    for cat, label in CATEGORIES.items():
        players = data[cat]
        n = len(players)
        if n == 0:
            continue
        with st.container(border=True):
            st.markdown(f"**{label}** \u2014 {n} {'entry' if n == 1 else 'entries'}")
            draw_size, group_size = build_draw_config_ui(cat, label, n, "nom")
            nom_draw_configs[cat] = {
                "entries":    players,
                "draw_size":  draw_size,
                "group_size": group_size,
            }

    # Step 4 – Review & Generate ──────────────────────────────────────────────
    st.markdown(
        '<div class="step-label">Step 4 \u2014 Review &amp; Generate Draw</div>',
        unsafe_allow_html=True,
    )

    if not nom_draw_configs:
        st.warning("No categories with entries found. Please check the file.")
        return

    nom_review = []
    for cat, label in CATEGORIES.items():
        if cat not in nom_draw_configs:
            continue
        cfg = nom_draw_configs[cat]
        n, ds, gs = len(cfg["entries"]), cfg["draw_size"], cfg["group_size"]
        nom_review.append({
            "Category":   label,
            "Entries":    n,
            "Draw Size":  ds,
            "Groups":     ds // gs,
            "Group Size": gs,
            "BYEs":       max(0, ds - n),
            "Waitlisted": max(0, n - ds),
        })
    st.dataframe(nom_review, use_container_width=True, hide_index=True)
    st.divider()

    c1, c2 = st.columns([1, 3])
    with c1:
        nom_clicked = st.button(
            "\U0001f3b2  Make a Draw",
            type="primary",
            use_container_width=True,
            key="nom_generate",
        )
    with c2:
        st.caption(
            "Each click produces a **different random arrangement**. "
            "Click again to regenerate."
        )

    if nom_clicked:
        with st.spinner("Generating draw\u2026"):
            excel_bytes = generate_draw_excel(nom_draw_configs)
        st.session_state["nom_draw_bytes"] = excel_bytes
        st.session_state["nom_draw_ready"] = True
        st.success("\u2705 Draw generated successfully!")

    if st.session_state.get("nom_draw_ready"):
        st.download_button(
            label="\U0001f4e5  Download Draw Excel",
            data=st.session_state["nom_draw_bytes"],
            file_name="Draw.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="nom_download",
        )


# ═════════════════════════════════════════════════════════════════════════════
# TAB 2 – APPLY SEEDING
# ═════════════════════════════════════════════════════════════════════════════
def render_tab2():
    st.markdown(
        """
        **How it works:**
        1. Generate a draw using **Tab 1** \u2192 open in Excel \u2192 **bold the names** of seeded players
        2. Save and upload that file here
        3. The app detects bold names as seeds and places them at the **top and bottom**
           of each group in a balanced pattern

        > \U0001f4cc **Seed order** = top-to-bottom order of bold names in the sheet.
        > Seed 1 \u2192 Group 1 top, Seed 2 \u2192 Group 2 top \u2026 Seed N+1 \u2192 Group N bottom (reversed for balance).
        """
    )

    # Step 1 – Upload draw ────────────────────────────────────────────────────
    st.markdown(
        '<div class="step-label">Step 1 \u2014 Upload Draw Excel</div>',
        unsafe_allow_html=True,
    )
    seed_uploaded = st.file_uploader(
        "Upload the draw Excel (bold player names = seeded)",
        type=["xlsx"],
        key="seed_upload",
        help="Upload a draw Excel previously generated by this app. Bold player names = seeds.",
    )

    if seed_uploaded is None:
        st.info("\U0001f446 Upload a draw Excel file (generated by Tab 1) to continue.")
        return

    seed_file_bytes = seed_uploaded.read()
    draw_data       = load_draw_file(seed_file_bytes)

    # Step 2 – Detected seeds ─────────────────────────────────────────────────
    st.markdown(
        '<div class="step-label">Step 2 \u2014 Detected Players &amp; Seeds</div>',
        unsafe_allow_html=True,
    )

    has_any = any(draw_data[c]["players"] for c in CATEGORIES)
    if not has_any:
        st.error(
            "No player data found. Make sure you upload a draw Excel generated by this app "
            "(sheets named MS, MD, Mixed, WS, WD)."
        )
        return

    seed_metrics_cols = st.columns(5)
    for i, (cat, label) in enumerate(CATEGORIES.items()):
        info = draw_data[cat]
        with seed_metrics_cols[i]:
            st.metric(
                label=label,
                value=f"{len(info['players'])} player(s)",
                delta=(
                    f"\U0001f3c6 {len(info['seeded'])} seed(s) detected"
                    if info["seeded"] else "No seeds"
                ),
            )

    for cat, label in CATEGORIES.items():
        seeded = draw_data[cat]["seeded"]
        if not seeded:
            continue
        with st.expander(f"\U0001f3c6 {label} \u2014 {len(seeded)} seed(s)", expanded=False):
            for idx, s in enumerate(seeded, 1):
                st.markdown(f"**Seed {idx}:** {s['label']}")

    st.caption(
        "\U0001f4a1 Seeds will be placed at the **top and bottom** of each group "
        "(Seed 1\u2026N across groups left\u2192right, then Seed N+1\u2026 right\u2192left). "
        "Non-seeded slots are randomly reshuffled on every click. "
        "Draw size and group size are kept the same as the uploaded file."
    )

    # Build configs directly from detected structure
    seed_draw_configs = {}
    for cat in CATEGORIES:
        info    = draw_data[cat]
        players = info["players"]
        seeded  = info["seeded"]
        ds      = info.get("draw_size", 0)
        gs      = info.get("group_size", 0)
        if not players or not ds or not gs:
            continue
        seed_draw_configs[cat] = {
            "players":    players,
            "seeded":     seeded,
            "draw_size":  ds,
            "group_size": gs,
        }

    # Step 3 – Review & Generate ──────────────────────────────────────────────
    st.markdown(
        '<div class="step-label">Step 3 \u2014 Review &amp; Generate Seeded Draw</div>',
        unsafe_allow_html=True,
    )

    if not seed_draw_configs:
        st.warning("No player data found in any category sheet.")
        return

    seed_review = []
    for cat, label in CATEGORIES.items():
        if cat not in seed_draw_configs:
            continue
        cfg     = seed_draw_configs[cat]
        n       = len(cfg["players"])
        n_seeds = len(cfg["seeded"])
        ds      = cfg["draw_size"]
        gs      = cfg["group_size"]
        seed_review.append({
            "Category":   label,
            "Players":    n,
            "Seeds":      n_seeds,
            "Draw Size":  ds,
            "Groups":     ds // gs,
            "Group Size": gs,
        })
    st.dataframe(seed_review, use_container_width=True, hide_index=True)
    st.divider()

    s1, s2 = st.columns([1, 3])
    with s1:
        seed_clicked = st.button(
            "\U0001f3b2  Make Seeded Draw",
            type="primary",
            use_container_width=True,
            key="seed_generate",
        )
    with s2:
        st.caption(
            "Seeds are fixed at top/bottom positions. "
            "Non-seeded slots are **randomly** reshuffled on every click."
        )

    if seed_clicked:
        with st.spinner("Generating seeded draw\u2026"):
            seed_excel_bytes = generate_seeded_draw_excel(seed_draw_configs)
        st.session_state["seed_draw_bytes"] = seed_excel_bytes
        st.session_state["seed_draw_ready"] = True
        st.success(
            "\u2705 Seeded draw generated! "
            "Seeded players are shown in **bold dark blue** in the Excel."
        )

    if st.session_state.get("seed_draw_ready"):
        st.download_button(
            label="\U0001f4e5  Download Seeded Draw Excel",
            data=st.session_state["seed_draw_bytes"],
            file_name="Draw_Seeded.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="seed_download",
        )


# -- Tabs --
tab1, tab2 = st.tabs(["\U0001f4cb  Create from Nominations", "\U0001f3c6  Apply Seeding"])

with tab1:
    render_tab1()

with tab2:
    render_tab2()
