"""Sixes Match Analysis — Coach Dashboard v0.3.2 (Coding Template V1.2).

Four coach-facing tabs: Overview, Offensive Patterns, Offensive Tempo, Infractions.
Workbook QC always runs; analyst-only details stay in analysis.py's CSV output.
"""
from __future__ import annotations

from pathlib import Path
from html import escape
import tempfile

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from analysis import analyze_match, load_infraction_log, load_match_workbook
from dashboard_data import (
    KPI_FIELDS, TYPES, computed_final_score, efficiency_by_type, end_reasons,
    infraction_by_team, infraction_changes, overview_kpis, possession_types,
    score_consistent, tooltip, tempo_medians, tempo_summary_table,
)

st.set_page_config(page_title="Sixes Match Analysis", page_icon="🥍", layout="wide")
st.title("🥍 Sixes Match Analysis")
# Cosmetic theme only; all calculations and plots outside Overview/Tempo stay unchanged.
st.markdown("""
<style>
/* Help coaches discover the four interactive navigation tabs. */
.stTabs [data-baseweb="tab-list"] {
    gap: 0.5rem;
    padding: 0.3rem 0 0.55rem;
}
.stTabs button[data-baseweb="tab"] {
    border-radius: 0.65rem 0.65rem 0.3rem 0.3rem;
    background-color: rgba(105, 119, 140, 0.08);
    padding: 0.65rem 1.05rem !important;
    font-weight: 600 !important;
    transition: background-color 120ms ease;
}
.stTabs button[data-baseweb="tab"] p {
    font-size: 1.03rem !important;
    font-weight: 600 !important;
}
.stTabs button[data-baseweb="tab"]:hover {
    background-color: rgba(71, 127, 180, 0.12);
}
.stTabs button[data-baseweb="tab"][aria-selected="true"] {
    background-color: rgba(71, 127, 180, 0.15);
    color: #356C9E;
    box-shadow: inset 0 -3px 0 #477FB4;
}
.stTabs button[data-baseweb="tab"]:focus-visible {
    outline: 2px solid #477FB4;
    outline-offset: 2px;
}
/* Compact, readable side-by-side home/away numbers inside each KPI card. */
.coach-kpi-heading {
    font-size: 0.91rem;
    font-weight: 650;
    line-height: 1.25;
    padding-top: 0.2rem;
    min-height: 2.3rem;
}
.coach-team-kpi {
    padding: 0.55rem 0.15rem 0.6rem;
    border-radius: 0.55rem;
    background-color: var(--kpi-bg);
    text-align: center;
    min-width: 0;
}
.coach-team-kpi .coach-team-name {
    color: var(--kpi-accent);
    font-weight: 700;
    font-size: 0.82rem;
    letter-spacing: 0.01em;
    white-space: nowrap;
}
.coach-team-kpi .coach-kpi-value {
    font-size: clamp(1.05rem, 1.45vw, 1.55rem);
    font-weight: 700;
    color: var(--text-color, inherit);
    margin-top: 0.15rem;
    white-space: nowrap;
    font-variant-numeric: tabular-nums;
}
@media (max-width: 900px) {
    .stTabs button[data-baseweb="tab"] { padding: 0.65rem 0.72rem !important; }
    .stTabs button[data-baseweb="tab"] p { font-size: 0.96rem !important; }
}
</style>
""", unsafe_allow_html=True)
st.caption("Single Match · Coach Dashboard v0.3.2 · Possession-based performance analysis")


# ------------------------------- File loading -------------------------------
@st.cache_data(show_spinner=False)
def load_uploaded(file_bytes: bytes):
    """Upload is not saved as a persistent file on the cloud server."""
    with tempfile.TemporaryDirectory() as temp_dir:
        path = Path(temp_dir) / "match.xlsx"
        path.write_bytes(file_bytes)
        return (*load_match_workbook(path), load_infraction_log(path))


@st.cache_data(show_spinner=False)
def load_local(path: str, mtime_ns: int, size: int):
    """mtime/size are cache keys; reloading edited local workbook stays reliable."""
    source = Path(path)
    return (*load_match_workbook(source), load_infraction_log(source))


def local_workbooks() -> list[Path]:
    base = Path(__file__).resolve().parent
    roots = [Path.cwd(), Path.cwd() / "data", Path.cwd() / "Match_Records", base, base / "data", base.parent / "Match_Records"]
    paths: dict[str, Path] = {}
    for root in roots:
        if root.is_dir():
            for file in root.glob("*.xlsx"):
                if not file.name.startswith("~$"):
                    paths[str(file.resolve())] = file.resolve()
    return sorted(paths.values(), key=lambda p: str(p))


def show_definition(label: str):
    """Bilingual definition and one formula; no interpretation paragraph."""
    with st.popover("ⓘ"):
        en, zh, formula = tooltip(label).split("\n\n")
        st.markdown(f"**{label}**")
        st.write(en)
        st.caption(zh)
        st.code(formula, language=None)


def section_heading(label: str):
    a, b = st.columns([12, 1], vertical_alignment="center")
    with a:
        st.subheader(label)
    with b:
        show_definition(label)


def fmt_number(value, is_percent=False):
    if value is None or pd.isna(value):
        return "—"
    return f"{float(value):.1f}%" if is_percent else str(int(value))


def kpi_team_panel(team: str, value: str, position: int):
    """Render an accessible pastel team panel; never embed unsanitized workbook text."""
    if position == 0:
        color, bg = "#356A9B", "rgba(71, 127, 180, 0.12)"
    else:
        color, bg = "#986426", "rgba(206, 150, 89, 0.14)"
    st.markdown(
        f'<div class="coach-team-kpi" style="--kpi-accent: {color}; --kpi-bg: {bg};">'
        f'<div class="coach-team-name">{escape(str(team))}</div>'
        f'<div class="coach-kpi-value">{escape(str(value))}</div></div>',
        unsafe_allow_html=True,
    )


def show_tempo_bar(times: pd.DataFrame, value_col: str, metric_name: str,
                   teams: list[str], palette: dict):
    """Coach-facing median comparison; mean and valid n remain available on hover."""
    section_heading(metric_name)
    data = tempo_medians(times, value_col, teams)
    if data.empty:
        st.info("No valid observation times are available for this metric.")
        return

    fig = px.bar(
        data, x="Possession_Type_Label", y="Median_s", color="Team",
        barmode="group", text="Bar_Label", color_discrete_map=palette,
        category_orders={"Possession_Type_Label": ["Transition", "Settled Offense"],
                         "Team": teams},
        custom_data=["Mean_s", "Valid_n"],
        labels={"Possession_Type_Label": "Possession Type",
                "Median_s": "Median time (seconds)"},
    )
    fig.update_traces(
        textposition="outside", cliponaxis=False,
        hovertemplate=(
            "%{fullData.name} · %{x}<br>"
            "Median: %{y:.1f} s<br>"
            "Mean: %{customdata[0]:.1f} s<br>"
            "Valid n: %{customdata[1]:.0f}<extra></extra>"
        ),
    )
    top = float(data["Median_s"].max())
    fig.update_layout(
        height=355,
        margin=dict(t=28, b=10),
        legend_title_text="Team",
        yaxis_range=[0, max(5.0, top * 1.35)],
    )
    st.plotly_chart(fig, use_container_width=True)
    st.caption(
        "Bars show median seconds; labels show median and valid n. "
        "Hover for mean. Other/uncategorized possessions are not included "
        "in the type comparison."
    )


# ------------------------------- Sidebar ------------------------------------
with st.sidebar:
    st.header("Match setup")
    source = st.radio("Data source", ["Upload Excel", "Local Excel"], horizontal=True)
    loaded = None
    source_name = ""
    if source == "Upload Excel":
        upload = st.file_uploader("Coding Template V1.2 (.xlsx)", type=["xlsx"])
        if upload is not None:
            source_name = upload.name
            try:
                loaded = load_uploaded(upload.getvalue())
            except Exception as exc:
                st.error(f"Cannot read workbook. Please contact the match analyst. ({exc})")
    else:
        files = local_workbooks()
        if files:
            selected = st.selectbox("Excel file", files, format_func=lambda f: f.name)
            source_name = selected.name
            try:
                stat = selected.stat()
                loaded = load_local(str(selected), stat.st_mtime_ns, stat.st_size)
            except Exception as exc:
                st.error(f"Cannot read workbook. ({exc})")
        else:
            st.info("No local Excel found. Use Upload Excel, or place the file in ./data.")
    st.divider()
    st.caption("Requires Match_Info, Match_Record, Infraction_Log, and Score_Check.")
    st.caption("Missing broadcast observations are not imputed.")

if loaded is None:
    st.info("Upload a completed Coding Template V1.2 workbook to view the match analysis.")
    st.stop()

record, meta, score_qc, audit, events = loaded
teams = [meta.team_home, meta.team_away]
palette = {teams[0]: "#477FB4", teams[1]: "#CE9659"}
package = analyze_match(record, meta, score_qc=score_qc, score_audit=audit, infractions=events)

# Important: a coach view should never present unsafe numbers as though validated.
errors = package["qc"].loc[package["qc"]["Severity"] == "error"]
if not errors.empty:
    st.error(f"The uploaded match contains {len(errors)} validation error(s). Please ask the match analyst to check the file before presenting results.")
    st.stop()
if not score_consistent(record, meta, audit):
    st.error("Score validation did not pass. Please ask the match analyst to check the scoring records before viewing this report.")
    st.stop()

# Score matches and all critical QC checks pass. Legitimate warning/info records
# remain in the analyst-side qc.csv and do not overwhelm the coach interface.
with st.sidebar:
    st.markdown(f"**Match:** {meta.match_id}")
    st.caption(f"Home: {teams[0]} · Away: {teams[1]}")

home_score, away_score = computed_final_score(record, *teams)
team_overview = overview_kpis(package["kpis"], teams)

# -------------------------------- Four tabs ---------------------------------
overview_tab, offense_tab, tempo_tab, infr_tab = st.tabs([
    "Overview", "Offensive Patterns", "Offensive Tempo", "Infractions",
])

with overview_tab:
    st.subheader("Match Information")
    a, b, c = st.columns([1, 1.2, 1])
    with a:
        st.caption("HOME")
        st.header(teams[0])
    with b:
        st.caption("Official Final Score")
        st.header(f"{home_score}  –  {away_score}")
    with c:
        st.caption("AWAY")
        st.header(teams[1])
    st.caption(f"Match ID: {meta.match_id} · Reconstructed final score matches the official result. · KPIs use observed coded possessions.")
    st.divider()

    st.subheader("Core Performance Indicators")
    metric_cols = st.columns(5, gap="small")
    lookup = team_overview.set_index("Team")
    for column, (label, _) in zip(metric_cols, KPI_FIELDS):
        percent = label != "Possessions"
        hval = lookup.at[teams[0], label]
        aval = lookup.at[teams[1], label]
        with column:
            with st.container(border=True):
                title_col, help_col = st.columns([5, 1], gap="small", vertical_alignment="center")
                with title_col:
                    st.markdown(f'<div class="coach-kpi-heading">{escape(label)}</div>',
                                unsafe_allow_html=True)
                with help_col:
                    show_definition(label)
                left, right = st.columns(2, gap="small")
                with left:
                    kpi_team_panel(teams[0], fmt_number(hval, percent), 0)
                with right:
                    kpi_team_panel(teams[1], fmt_number(aval, percent), 1)

    st.subheader("Offensive Performance Comparison")
    bars = team_overview.melt(id_vars="Team", value_vars=[label for label, _ in KPI_FIELDS[1:]],
                              var_name="Metric", value_name="Percent")
    fig = px.bar(bars, x="Metric", y="Percent", color="Team", barmode="group",
                 text="Percent", color_discrete_map=palette,
                 category_orders={"Metric": [m[0] for m in KPI_FIELDS[1:]]},
                 labels={"Percent": "Percent (%)"})
    fig.update_traces(texttemplate="%{y:.1f}%", textposition="outside",
                      hovertemplate="%{fullData.name} · %{x}<br>%{y:.1f}%<extra></extra>")
    fig.update_layout(yaxis_range=[0, 109], legend_title_text="Team",
                      margin=dict(t=12, b=12), height=400)
    st.plotly_chart(fig, use_container_width=True)

with offense_tab:
    # 1: Type distribution — denominator includes Other.
    section_heading("Possession Type Distribution")
    ptypes = possession_types(record, teams)
    ptypes["Type"] = ptypes["Possession_Type"].replace({"Settled_Offense": "Settled Offense"})
    fig_ptype = px.bar(ptypes, x="Type", y="Count", color="Team", barmode="group",
                       text="Count", color_discrete_map=palette,
                       category_orders={"Type": ["Transition", "Settled Offense", "Other"]},
                       custom_data=["Share_pct"], labels={"Count": "Possessions"})
    fig_ptype.update_traces(textposition="outside", hovertemplate=(
        "%{fullData.name} · %{x}<br>Count: %{y}<br>Share: %{customdata[0]:.1f}% of team possessions<extra></extra>"))
    fig_ptype.update_layout(height=350, margin=dict(t=15, b=15))
    st.plotly_chart(fig_ptype, use_container_width=True)
    st.caption("Share denominator includes Other. Types are coded at the first offensive threat, based on defensive organization.")

    # 2: Goals per possession by type; ratio labels discourage n=3 over-reading.
    section_heading("Offensive Efficiency by Possession Type")
    eff = efficiency_by_type(record, teams)
    eff["Type"] = eff["Possession_Type"].replace({"Settled_Offense": "Settled Offense"})
    fig_eff = px.bar(eff, x="Type", y="Efficiency_pct", color="Team", barmode="group",
                     color_discrete_map=palette, text="Ratio",
                     category_orders={"Type": ["Transition", "Settled Offense"]},
                     custom_data=["Goals", "Possessions"],
                     labels={"Efficiency_pct": "Goals per possession (%)"})
    fig_eff.update_traces(textposition="outside", hovertemplate=(
        "%{fullData.name} · %{x}<br>Goal / Possession: %{customdata[0]} / %{customdata[1]}<br>"
        "Efficiency: %{y:.1f}%<extra></extra>"))
    fig_eff.update_layout(height=340, yaxis_range=[0, 110], margin=dict(t=15, b=15))
    st.plotly_chart(fig_eff, use_container_width=True)
    st.caption("Bar labels show Goals / Possessions. Results describe the full possession, not necessarily the phase when the goal occurred.")

    # 3: End reason counts, not aggregated as a 100% stacked chart.
    section_heading("Possession End Reasons")
    ends = end_reasons(record, teams)
    combined = ends.groupby("End_Reason")["Count"].sum().sort_values(ascending=False)
    ordering = combined.index.to_list()
    fig_ends = px.bar(ends, y="End_Reason", x="Count", color="Team", orientation="h",
                      barmode="group", color_discrete_map=palette, text="Count",
                      category_orders={"End_Reason": ordering[::-1]},
                      custom_data=["Share_pct"], labels={"End_Reason": "End Reason"})
    fig_ends.update_traces(textposition="outside", hovertemplate=(
        "%{fullData.name} · %{y}<br>Count: %{x}<br>Share: %{customdata[0]:.1f}% of team possessions<extra></extra>"))
    fig_ends.update_layout(height=425, margin=dict(t=10, b=10), legend_title_text="Team")
    st.plotly_chart(fig_ends, use_container_width=True)

with tempo_tab:
    st.subheader("Tempo Summary")
    tm = package["time_by_team"].copy()
    summary_left, summary_right = st.columns(2, gap="large")
    with summary_left:
        section_heading("Possession Duration")
        st.dataframe(tempo_summary_table(tm, "Possession Duration", teams),
                     hide_index=True, use_container_width=True)
    with summary_right:
        section_heading("Time to First Shot")
        st.dataframe(tempo_summary_table(tm, "Time to First Shot", teams),
                     hide_index=True, use_container_width=True)
    st.caption("Valid Possessions (n) excludes unobservable timings only for the relevant measure. "
               "Values use the official game clock, not shot-clock usage or active-ball time. "
               "Missing observations are never counted as zero.")
    times = package["possession_times"].copy()
    show_tempo_bar(times, "Time_to_First_Shot_s", "Time to First Shot", teams, palette)
    show_tempo_bar(times, "Possession_Duration_s", "Possession Duration", teams, palette)
    st.caption("No-shot possessions do not enter Time to First Shot. Unobservable broadcast timings are excluded only from the relevant time metric.")

with infr_tab:
    section_heading("Infraction Type Distribution")
    infra = infraction_by_team(package["infractions"], teams)
    global_max = int(infra["Count"].max()) if not infra.empty else 0
    columns = st.columns(2)
    for column, team in zip(columns, teams):
        with column:
            st.markdown(f"**{team} Infraction Distribution**")
            part = infra.loc[infra["Team"] == team].copy()
            st.caption(f"Total: {int(part['Count'].sum())} recorded events")
            if part.empty:
                st.info("No recorded infractions.")
            else:
                part = part.sort_values("Count", ascending=True)
                fig_infra = px.bar(part, y="Infraction_Type", x="Count", orientation="h",
                                   text="Count", color_discrete_sequence=[palette[team]],
                                   labels={"Infraction_Type": "Infraction Type"})
                fig_infra.update_traces(textposition="outside", hovertemplate=(
                    "%{y}<br>Count: %{x}<extra></extra>"))
                fig_infra.update_layout(xaxis_range=[0, max(1, global_max) * 1.2],
                                        showlegend=False, height=max(275, 60 * len(part) + 100),
                                        margin=dict(t=15, b=15))
                st.plotly_chart(fig_infra, use_container_width=True)
    st.caption("Only observed infraction types appear. Both charts use the same count scale. Loose_Ball_Push is tracked separately and does not count as Defensive Fouls Drawn.")

    section_heading("Possession Change by Infraction")
    changes = infraction_changes(package["infractions"], teams)
    if changes["Total"].sum() == 0:
        st.info("No infraction events are recorded for this match.")
    else:
        fig_change = go.Figure()
        for name, color in [("Yes", "#477FB4"), ("No", "#CE9659")]:
            values = changes[f"{name}_pct"].fillna(0).to_numpy()
            counts = changes[name].astype(int).to_numpy()
            fig_change.add_trace(go.Bar(
                name=name, y=changes["Team"], x=values, orientation="h",
                marker_color=color, text=[str(n) if n else "" for n in counts],
                textposition="inside", customdata=counts,
                hovertemplate=(f"{name} · %{{y}}<br>Count: %{{customdata}}"
                               "<br>Share: %{x:.1f}%<extra></extra>"),
            ))
        fig_change.update_layout(barmode="stack", xaxis_range=[0, 100],
                                 xaxis_title="Share of team infractions (%)", yaxis_title="Team",
                                 height=260, margin=dict(t=10, b=15), legend_title_text="Possession Change")
        st.plotly_chart(fig_change, use_container_width=True)
        st.caption("Percentages use each team's recorded infractions as the denominator. ‘No’ does not mean an infraction had no effect on play.")
        st.dataframe(changes[["Team", "Yes", "No", "Total"]], hide_index=True, use_container_width=True)

st.divider()
st.caption("Coach view · Based on observed coded possessions · Analyst-side QC is preserved in analysis.py (qc.csv).")
