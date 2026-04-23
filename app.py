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
from ttclash_draw_maker import (
    read_ttclash_nominations,
    read_ttclash_draw_file,
    generate_ttclash_draw_excel,
    generate_ttclash_seeded_draw_excel,
    TTCLASH_CATEGORIES,
    COMPANY_DRAW_CATEGORIES,
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


@st.cache_data(show_spinner="\U0001f4cb Analysing TTClash nominations\u2026")
def load_ttclash_nominations(file_bytes: bytes, _cache_ver: int = 3):
    try:
        return read_ttclash_nominations(io.BytesIO(file_bytes))
    except ValueError as exc:
        msg = str(exc)
        if msg.startswith("FRAMESET_WORKBOOK:"):
            companion = msg.split(":", 1)[1]
            raise ValueError(companion) from None
        raise


@st.cache_data(show_spinner="\U0001f4ca Reading TTClash draw file\u2026")
def load_ttclash_draw_file(file_bytes: bytes):
    return read_ttclash_draw_file(io.BytesIO(file_bytes))


# ── Tabs ──────────────────────────────────────────────────────────────────────


# ═════════════════════════════════════════════════════════════════════════════
# TT CLASH – TAB 1 (CREATE FROM NOMINATIONS)
# ═════════════════════════════════════════════════════════════════════════════
def render_ttclash_tab1():
    # Step 1 – Upload \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    st.markdown(
        '<div class="step-label">Step 1 \u2014 Upload TT Clash Nomination File</div>',
        unsafe_allow_html=True,
    )
    uploaded = st.file_uploader(
        "Choose the TT Clash Nomination file\u00a0(.xls / .xlsx / .htm)",
        type=["xls", "xlsx", "htm", "html"],
        key="tc_nom_upload",
        help=(
            "Upload the TT Clash registration export (.xls) file. "
            "If your .xls file is a multi-frame workbook, upload the "
            "sheet001.htm file from its companion folder instead."
        ),
    )

    if uploaded is None:
        st.info("\U0001f446 Please upload the **TT Clash Nomination file** to continue.")
        return

    file_bytes = uploaded.read()
    try:
        categories, player_info, has_city = load_ttclash_nominations(file_bytes)
        # Store for the Players browser tab
        st.session_state["tc_player_info"] = player_info
        st.session_state["tc_categories"]  = categories
        st.session_state["tc_has_city"]    = has_city
    except ValueError as exc:
        companion = str(exc)
        # companion looks like  "Registrants Details - ..._files/sheet001.htm"
        folder = companion.split("/")[0] if "/" in companion else companion
        st.error(
            "\u26a0\ufe0f **Multi-frame workbook detected** — this .xls file is a "
            "frameset wrapper and cannot be read directly.\n\n"
            "**How to fix:** \n"
            f"1. Find the companion folder **`{folder}`** next to your .xls file.\n"
            "2. Open that folder and upload **`sheet001.htm`** here instead.\n"
            "3. *Alternatively*, open the .xls in Excel → **File ▸ Save As** → "
            "choose **Web Page (*.htm)** with **\u2018Selection: Sheet\u2019** or "
            "**\u2018HTML Only\u2019** → save and upload that file."
        )
        return

    # Step 2 – Analysis \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    st.markdown(
        '<div class="step-label">Step 2 \u2014 Nomination Analysis</div>',
        unsafe_allow_html=True,
    )
    total = sum(len(v) for v in categories.values())
    st.success(f"File loaded! **{total} total entries** found across all categories.")

    metrics_cols = st.columns(6)
    for i, (cat, label) in enumerate(TTCLASH_CATEGORIES.items()):
        n         = len(categories[cat])
        suggested = next_power_of_2(n) if n > 0 else "\u2013"
        with metrics_cols[i]:
            st.metric(
                label=label,
                value=f"{n} {'entry' if n == 1 else 'entries'}",
                delta=f"Suggest draw: {suggested}" if n > 0 else "No entries",
            )

    # Step 3 – Configure \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    st.markdown(
        '<div class="step-label">Step 3 \u2014 Configure Draw Settings</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "\U0001f4a1 **Draw Size** must be a power of 2 (4\u00b78\u00b716\u00b732\u00b764\u00b7128\u00b7256).  "
        "**Group Size** must be even and divide the Draw Size exactly.  \n"
        "Smaller draw than entries \u2192 extra players **waitlisted**.  "
        "Larger draw \u2192 empty slots become **BYE** entries."
    )

    tc_draw_configs = {}
    for cat, label in TTCLASH_CATEGORIES.items():
        players = categories[cat]
        n = len(players)
        if n == 0:
            continue
        with st.container(border=True):
            st.markdown(f"**{label}** \u2014 {n} {'entry' if n == 1 else 'entries'}")
            draw_size, group_size = build_draw_config_ui(cat, label, n, "tc")
            draw_by_city    = False
            draw_by_company = False
            if cat in COMPANY_DRAW_CATEGORIES:
                col_by_ci, col_by_co = st.columns(2)
                with col_by_ci:
                    draw_by_city = st.checkbox(
                        "\U0001f3d9\ufe0f Group by City \u2014 "
                        "try to keep same-city players in different groups",
                        key=f"tc_byci_{cat}",
                        disabled=not has_city,
                        help=(
                            "Players from the same city will be spread across groups where possible. "
                            "If both City and Company are selected, City takes priority."
                        ) if has_city else (
                            "\u26a0\ufe0f No City column found in the uploaded file. "
                            "Upload a registration file that includes a 'City' column to enable this option."
                        ),
                    )
                with col_by_co:
                    draw_by_company = st.checkbox(
                        "\U0001f3e2 Group by Company / Organisation \u2014 "
                        "try to keep same-company players in different groups",
                        key=f"tc_byco_{cat}",
                        help=(
                            "Players from the same company/organisation will be spread across groups "
                            "where possible. If City is also selected and city data is available, "
                            "City takes priority."
                        ),
                    )
            tc_draw_configs[cat] = {
                "entries":         players,
                "draw_size":       draw_size,
                "group_size":      group_size,
                "draw_by_city":    draw_by_city,
                "draw_by_company": draw_by_company,
            }

    # Step 4 – Review & Generate \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    st.markdown(
        '<div class="step-label">Step 4 \u2014 Review &amp; Generate Draw</div>',
        unsafe_allow_html=True,
    )

    if not tc_draw_configs:
        st.warning("No categories with entries found. Please check the file.")
        return

    tc_review = []
    for cat, label in TTCLASH_CATEGORIES.items():
        if cat not in tc_draw_configs:
            continue
        cfg = tc_draw_configs[cat]
        n, ds, gs = len(cfg["entries"]), cfg["draw_size"], cfg["group_size"]
        if cfg.get("draw_by_city") and has_city:
            grouping = "\U0001f3d9\ufe0f City"
        elif cfg.get("draw_by_company"):
            grouping = "\U0001f3e2 Company"
        else:
            grouping = "\u2014"
        tc_review.append({
            "Category":   label,
            "Entries":    n,
            "Draw Size":  ds,
            "Groups":     ds // gs,
            "Group Size": gs,
            "BYEs":       max(0, ds - n),
            "Waitlisted": max(0, n - ds),
            "Grouping":   grouping,
        })
    st.dataframe(tc_review, use_container_width=True, hide_index=True)

    include_players = st.checkbox(
        "\U0001f4cb Include Player Info sheet (name, mobile, company, gender, t-shirt, email) "
        "\u2014 adds a clickable '...' next to each player name in the draw",
        value=True,
        key="tc_include_players",
        help=(
            "Adds a 'Players' sheet with full registration details. "
            "In each draw sheet, clicking '...' beside a player name jumps to their full record."
        ),
    )

    st.divider()
    c1, c2 = st.columns([1, 3])
    with c1:
        tc_clicked = st.button(
            "\U0001f3b2  Make a Draw",
            type="primary",
            use_container_width=True,
            key="tc_generate",
        )
    with c2:
        st.caption(
            "Each click produces a **different random arrangement**. "
            "Click again to regenerate."
        )

    if tc_clicked:
        with st.spinner("Generating draw\u2026"):
            excel_bytes = generate_ttclash_draw_excel(
                tc_draw_configs,
                player_info=player_info if include_players else None,
            )
        st.session_state["tc_draw_bytes"] = excel_bytes
        st.session_state["tc_draw_ready"] = True
        st.success("\u2705 Draw generated successfully!")

    if st.session_state.get("tc_draw_ready"):
        st.download_button(
            label="\U0001f4e5  Download TT Clash Draw Excel",
            data=st.session_state["tc_draw_bytes"],
            file_name="TTClash_Draw.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="tc_download",
        )


# ═════════════════════════════════════════════════════════════════════════════
# TT CLASH – TAB 2 (APPLY SEEDING)
# ═════════════════════════════════════════════════════════════════════════════
def render_ttclash_tab2():
    st.markdown(
        """
        **How it works for TT Clash seeding:**
        1. Generate a TT Clash draw using **Tab 1** \u2192 open in Excel \u2192 **bold the names** of seeded players
        2. Save and upload that file here
        3. The app detects bold names as seeds and places them at the **top and bottom**
           of each group in a balanced pattern

        > \U0001f4cc **Seed order** = top-to-bottom order of bold names in the sheet.
        > Seed 1 \u2192 Group 1 top, Seed 2 \u2192 Group 2 top \u2026 Seed N+1 \u2192 Group N bottom (reversed for balance).
        """
    )

    # Step 1 – Upload draw \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    st.markdown(
        '<div class="step-label">Step 1 \u2014 Upload TT Clash Draw Excel</div>',
        unsafe_allow_html=True,
    )
    seed_uploaded = st.file_uploader(
        "Upload the TT Clash draw Excel (bold player names = seeded)",
        type=["xlsx"],
        key="tc_seed_upload",
        help="Upload a TT Clash draw Excel generated by this app. Bold player names = seeds.",
    )

    if seed_uploaded is None:
        st.info("\U0001f446 Upload a TT Clash draw Excel file (generated by Tab 1) to continue.")
        return

    seed_file_bytes = seed_uploaded.read()
    draw_data       = load_ttclash_draw_file(seed_file_bytes)

    # Step 2 – Detected seeds \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    st.markdown(
        '<div class="step-label">Step 2 \u2014 Detected Players &amp; Seeds</div>',
        unsafe_allow_html=True,
    )

    has_any = any(draw_data[c]["players"] for c in TTCLASH_CATEGORIES)
    if not has_any:
        st.error(
            "No player data found. Make sure you upload a TT Clash draw Excel generated by this app "
            "(sheets named MS, WS, 35MS, MD, WD, Mixed)."
        )
        return

    tc_seed_cols = st.columns(6)
    for i, (cat, label) in enumerate(TTCLASH_CATEGORIES.items()):
        info = draw_data[cat]
        with tc_seed_cols[i]:
            st.metric(
                label=label,
                value=f"{len(info['players'])} player(s)",
                delta=(
                    f"\U0001f3c6 {len(info['seeded'])} seed(s) detected"
                    if info["seeded"] else "No seeds"
                ),
            )

    for cat, label in TTCLASH_CATEGORIES.items():
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

    tc_seed_draw_configs = {}
    for cat in TTCLASH_CATEGORIES:
        info    = draw_data[cat]
        players = info["players"]
        seeded  = info["seeded"]
        ds      = info.get("draw_size", 0)
        gs      = info.get("group_size", 0)
        if not players or not ds or not gs:
            continue
        tc_seed_draw_configs[cat] = {
            "players":    players,
            "seeded":     seeded,
            "draw_size":  ds,
            "group_size": gs,
        }

    # Step 3 – Review & Generate \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    st.markdown(
        '<div class="step-label">Step 3 \u2014 Review &amp; Generate Seeded Draw</div>',
        unsafe_allow_html=True,
    )

    if not tc_seed_draw_configs:
        st.warning("No player data found in any category sheet.")
        return

    tc_seed_review = []
    for cat, label in TTCLASH_CATEGORIES.items():
        if cat not in tc_seed_draw_configs:
            continue
        cfg     = tc_seed_draw_configs[cat]
        n       = len(cfg["players"])
        n_seeds = len(cfg["seeded"])
        ds      = cfg["draw_size"]
        gs      = cfg["group_size"]
        tc_seed_review.append({
            "Category":   label,
            "Players":    n,
            "Seeds":      n_seeds,
            "Draw Size":  ds,
            "Groups":     ds // gs,
            "Group Size": gs,
        })
    st.dataframe(tc_seed_review, use_container_width=True, hide_index=True)
    st.divider()

    ts1, ts2 = st.columns([1, 3])
    with ts1:
        tc_seed_clicked = st.button(
            "\U0001f3b2  Make Seeded Draw",
            type="primary",
            use_container_width=True,
            key="tc_seed_generate",
        )
    with ts2:
        st.caption(
            "Seeds are fixed at top/bottom positions. "
            "Non-seeded slots are **randomly** reshuffled on every click."
        )

    if tc_seed_clicked:
        with st.spinner("Generating seeded draw\u2026"):
            tc_seed_bytes = generate_ttclash_seeded_draw_excel(tc_seed_draw_configs)
        st.session_state["tc_seed_draw_bytes"] = tc_seed_bytes
        st.session_state["tc_seed_draw_ready"] = True
        st.success(
            "\u2705 Seeded draw generated! "
            "Seeded players are shown in **bold dark blue** in the Excel."
        )

    if st.session_state.get("tc_seed_draw_ready"):
        st.download_button(
            label="\U0001f4e5  Download TT Clash Seeded Draw Excel",
            data=st.session_state["tc_seed_draw_bytes"],
            file_name="TTClash_Draw_Seeded.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="tc_seed_download",
        )


# ═════════════════════════════════════════════════════════════════════════════
# TT CLASH – PLAYERS BROWSER TAB
# ═════════════════════════════════════════════════════════════════════════════
def render_ttclash_players_tab():
    player_info = st.session_state.get("tc_player_info")
    categories  = st.session_state.get("tc_categories")
    has_city    = st.session_state.get("tc_has_city", False)

    if not player_info:
        st.info(
            "👆 Go to **Create from Nominations** tab, select **⚡ TT Clash** "
            "and upload the nomination file first."
        )
        return

    st.markdown(
        '<div class="step-label">Player Directory</div>',
        unsafe_allow_html=True,
    )

    # Build category membership map  email → list of category labels
    cat_map: dict = {}
    if categories:
        for cat, players in categories.items():
            for p in players:
                e = (p.get("email") or "").lower()
                if e:
                    cat_map.setdefault(e, []).append(TTCLASH_CATEGORIES.get(cat, cat))

    import pandas as pd

    # Server-side search input (filters rows on Enter / focus-out)
    search = st.text_input(
        "🔍 Search player",
        placeholder="Type name, company, city, email… then press Enter",
        key="tc_player_search",
    )
    q = search.strip().lower()

    # Build rows, applying server-side filter when a query is present
    rows = []
    for p in player_info:
        if q and not any(
            q in (p.get(f) or "").lower()
            for f in ("name", "company", "city", "email", "mobile")
        ):
            continue
        email = (p.get("email") or "").lower()
        cats_joined = ", ".join(cat_map.get(email, ["-"]))
        row = {
            "#":          p.get("regno",   ""),
            "Name":       p.get("name",    ""),
            "Company":    p.get("company", ""),
            "Gender":     p.get("gender",  ""),
            "Mobile":     p.get("mobile",  ""),
            "Email":      email,
            "T-Shirt":    p.get("tshirt",  ""),
            "Categories": cats_joined,
        }
        if has_city:
            row["City"] = p.get("city", "")
        rows.append(row)

    col_order = ["#", "Name", "Company"]
    if has_city:
        col_order.append("City")
    col_order += ["Gender", "Mobile", "Email", "T-Shirt", "Categories"]

    df = pd.DataFrame(rows)[col_order]
    st.caption(
        f"Showing **{len(rows)}** of **{len(player_info)}** registrants — "
        "you can also use the **🔍 icon** in the table toolbar (top-right on hover) to filter in real-time"
    )
    st.dataframe(df, use_container_width=True, hide_index=True)


# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 – CREATE FROM NOMINATIONS
# ═════════════════════════════════════════════════════════════════════════════
def render_tab1():    # ── Tournament format selector ───────────────────────────────────────────
    _tt_type = st.radio(
        "Select Tournament Format",
        ["\U0001f3e2 TT Internal Cybage", "\u26a1 TT Clash"],
        horizontal=True,
        key="tab1_tournament_type",
    )
    st.divider()
    if _tt_type == "\u26a1 TT Clash":
        render_ttclash_tab1()
        return
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
    # ── Tournament format selector ───────────────────────────────────────────
    _tt_type = st.radio(
        "Select Draw Type to Seed",
        ["\U0001f3e2 TT Internal Cybage", "\u26a1 TT Clash"],
        horizontal=True,
        key="tab2_tournament_type",
    )
    st.divider()
    if _tt_type == "\u26a1 TT Clash":
        render_ttclash_tab2()
        return

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
tab1, tab2, tab3 = st.tabs([
    "📋  Create from Nominations",
    "🏆  Apply Seeding",
    "👥  Players",
])

with tab1:
    render_tab1()

with tab2:
    render_tab2()

with tab3:
    render_ttclash_players_tab()
