"""Sixes Lacrosse Match Analysis Engine v0.3.1

Coding Template V1.2: Match_Info, Match_Record, Infraction_Log, Score_Check.
Scores are rebuilt from initial Home/Away scores and per-possession Goal records.
All score differentials refer to the possessing team at possession START.
Original technical-tactical KPI formulas are retained; violation-based losses and fouls
are read from Infraction_Log, and two clock-based metrics are added.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable
import argparse
import json
import math

import pandas as pd


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

REQUIRED_COLUMNS = [
    "Possession_ID", "Quarter", "Game_Clock", "First_Shot_Clock", "End_Clock",
    "Team", "Start_Type", "Start_Zone", "Possession_Type",
    "Pass", "Shot", "Shot_on_Goal", "Goal", "TimeOut", "End_Reason",
]

COUNT_COLUMNS = ["Pass", "Shot", "Shot_on_Goal", "Goal", "TimeOut"]
INFRACTION_COLUMNS = [
    "Infraction_ID", "Possession_ID", "Game_Clock", "Infraction_Team",
    "Infraction_Type", "Possession_Change",
]
VALID_END_REASONS = {
    "Goal", "GK_Save_Loss", "Miss_Loss", "Ball_Loss",
    "Infraction_Loss", "Period_End", "Other",
}
# Violation/procedural events are not automatically "defensive fouls drawn".
# Only an actual foul by the non-possessing team is counted as drawn.
FOUL_TYPES = {
    "Holding", "Pushing", "Interference", "Charging", "Warding",
    "Illegal_Pick", "Withholding_Ball", "Slashing", "Cross_Check", "Tripping",
    "Illegal_Body_Check", "Unnecessary_Roughness", "Unsportsmanlike_Conduct",
}
VIOLATION_TYPES = {
    "Shot_Clock_Violation", "Ten_Second_Violation", "Over_And_Back",
    "Crease_Violation", "Goalkeeper_Five_Second", "Illegal_Procedure",
    "Delay_of_Game", "Offside",
}
# Loose-ball push is recorded as a foul event but excluded from defensive
# fouls drawn; it occurs while the ball is loose, rather than a defense
# against an established possession.
KNOWN_INFRACTION_TYPES = FOUL_TYPES | VIOLATION_TYPES | {"Loose_Ball_Push", "Other", "Unknown"}

VALID_POSSESSION_TYPES = {"Transition", "Settled_Offense", "Other", "Unknown"}
VALID_START_ZONES = {"Defensive_Zone", "Middle_Zone", "Attacking_Zone", "Unknown"}

# Aliases let the coding sheet evolve without breaking old files.
# Values on the right are the canonical labels used by the analysis engine.
CATEGORY_ALIASES = {
    "Start_Type": {
        "GK_Possession": "Gk_Possession",
        "GK Possession": "Gk_Possession",
        "Live-ball_Gain": "Live_ball_Gain",
        "Live Ball Gain": "Live_ball_Gain",
        "Dead-ball_Restart": "Dead_ball_Restart",
        "Dead Ball Restart": "Dead_ball_Restart",
        "Goal Restart": "Goal_Restart",
        "Goal-Restart": "Goal_Restart",
    },
    "Possession_Type": {
        "Settled": "Settled_Offense",
        "Settled Offense": "Settled_Offense",
    },
    "End_Reason": {
        "Gk_Save_Loss": "GK_Save_Loss",
        "GK Save Loss": "GK_Save_Loss",
        "Miss-Loss": "Miss_Loss",  # Accept older workbook spelling without rewriting the file.
        "Miss Loss": "Miss_Loss",
        "Ball Loss": "Ball_Loss",
        "Infraction Loss": "Infraction_Loss",
        "Period End": "Period_End",
        # Older files may still be encountered during transition to V1.2.
        "Offensive_Foul": "Infraction_Loss",
        "Shot-Clock_Violation": "Infraction_Loss",
        "10sec_Violation": "Infraction_Loss",
    },
}

CONTROLLABLE_LOSS_REASONS = {"Ball_Loss", "Infraction_Loss"}


@dataclass(frozen=True)
class MatchMeta:
    match_id: str = "match_1"
    focus_team: str = "TPE"
    opponent: str | None = None
    competition: str | None = None
    stage: str | None = None
    note: str | None = None
    team_home: str | None = None
    team_away: str | None = None
    first_possession_id: int = 1
    initial_score_home: int = 0
    initial_score_away: int = 0
    official_final_score_home: int | None = None
    official_final_score_away: int | None = None


# ---------------------------------------------------------------------------
# Loading / cleaning
# ---------------------------------------------------------------------------

def load_match_excel(path: str | Path, sheet_name: str = "Match_Record") -> pd.DataFrame:
    """Load only the coding columns from one match workbook.

    Extra helper/dropdown columns on the right side of the worksheet are ignored.
    Completely blank rows are removed.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)

    df = pd.read_excel(path, sheet_name=sheet_name, usecols=lambda c: c in REQUIRED_COLUMNS)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    df = df[REQUIRED_COLUMNS].copy()
    df = df.dropna(how="all")

    # Remove accidental non-data rows that may remain below the table.
    df = df[df["Possession_ID"].notna()].copy()

    # Clean text columns.
    text_cols = [
        "Quarter",
        "Team",
        "Start_Type",
        "Start_Zone",
        "Possession_Type",
        "End_Reason",
    ]
    for col in text_cols:
        df[col] = df[col].astype("string").str.strip()
        df[col] = df[col].replace({"<NA>": pd.NA, "nan": pd.NA, "": pd.NA})

    # Apply known label aliases.
    for col, mapping in CATEGORY_ALIASES.items():
        df[col] = df[col].replace(mapping)

    # Numeric fields.
    df["Possession_ID"] = pd.to_numeric(df["Possession_ID"], errors="coerce")
    for col in COUNT_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Keep missing count cells visible to QC rather than silently converting them to zero.
    df = df.sort_values("Possession_ID", kind="stable").reset_index(drop=True)
    for col in ["Game_Clock", "First_Shot_Clock", "End_Clock"]:
        df[col + "_Display"] = df[col].apply(format_game_clock)
    return df


INFO_COLUMNS = [
    "Match_ID", "Team_Home", "Team_Away", "First_Possession_ID",
    "Initial_Score_Home", "Initial_Score_Away",
    "Official_Final_Score_Home", "Official_Final_Score_Away",
]
CHECK_COLUMNS = ["Before_Possession_ID", "Score_Home", "Score_Away", "Check_Reason"]
CHECK_REASONS = {"Quarter_Start", "Video_Gap", "Manual_Check"}


def _nonnegative_integer(value: Any, label: str) -> int:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number) or not math.isfinite(float(number)) or number < 0 or float(number) != int(number):
        raise ValueError(f"{label} must be a non-negative whole number; got {value!r}")
    return int(number)


def load_match_info(path: str | Path) -> MatchMeta:
    """Read the single Match_Info row. Home/Away reflect official World Lacrosse labels."""
    try:
        info = pd.read_excel(path, sheet_name="Match_Info", dtype=object)
    except ValueError as exc:
        raise ValueError("The workbook must contain a Match_Info worksheet.") from exc
    missing = [col for col in INFO_COLUMNS if col not in info.columns]
    if missing:
        raise ValueError(f"Match_Info missing columns: {missing}")
    info = info.dropna(how="all", subset=INFO_COLUMNS)
    if len(info) != 1:
        raise ValueError(f"Match_Info requires exactly one populated match row; found {len(info)}")
    r = info.iloc[0]
    home, away = str(r["Team_Home"]).strip(), str(r["Team_Away"]).strip()
    if not home or not away or home.lower() == "nan" or away.lower() == "nan" or home == away:
        raise ValueError("Team_Home and Team_Away must be two distinct team codes.")
    match_id = str(r["Match_ID"]).strip()
    if not match_id or match_id.lower() == "nan":
        raise ValueError("Match_Info.Match_ID cannot be blank")
    first_id = _nonnegative_integer(r["First_Possession_ID"], "First_Possession_ID")
    if first_id < 1:
        raise ValueError("First_Possession_ID must be at least 1")
    return MatchMeta(
        match_id=match_id,
        focus_team=home,
        opponent=away,
        team_home=home,
        team_away=away,
        first_possession_id=first_id,
        initial_score_home=_nonnegative_integer(r["Initial_Score_Home"], "Initial_Score_Home"),
        initial_score_away=_nonnegative_integer(r["Initial_Score_Away"], "Initial_Score_Away"),
        official_final_score_home=_nonnegative_integer(r["Official_Final_Score_Home"], "Official_Final_Score_Home"),
        official_final_score_away=_nonnegative_integer(r["Official_Final_Score_Away"], "Official_Final_Score_Away"),
    )


def load_score_check(path: str | Path) -> pd.DataFrame:
    """An empty Score_Check with just four headings is valid."""
    try:
        checks = pd.read_excel(path, sheet_name="Score_Check", dtype=object)
    except ValueError as exc:
        raise ValueError("The workbook must contain a Score_Check worksheet (it may be empty).") from exc
    missing = [col for col in CHECK_COLUMNS if col not in checks.columns]
    if missing:
        raise ValueError(f"Score_Check missing columns: {missing}")
    checks = checks[CHECK_COLUMNS].dropna(how="all").copy()
    for idx in checks.index:
        excel_row = int(idx) + 2
        checks.at[idx, "Before_Possession_ID"] = _nonnegative_integer(
            checks.at[idx, "Before_Possession_ID"], f"Score_Check row {excel_row} Before_Possession_ID")
        for col in ["Score_Home", "Score_Away"]:
            checks.at[idx, col] = _nonnegative_integer(checks.at[idx, col], f"Score_Check row {excel_row} {col}")
        reason = str(checks.at[idx, "Check_Reason"]).strip()
        if reason not in CHECK_REASONS:
            raise ValueError(f"Score_Check row {excel_row} Check_Reason must be one of {sorted(CHECK_REASONS)}")
        checks.at[idx, "Check_Reason"] = reason
    checks = checks.reset_index(drop=True)
    if checks["Before_Possession_ID"].duplicated().any():
        raise ValueError("Score_Check has duplicate Before_Possession_ID values")
    return checks


def _score_qc_issue(df: pd.DataFrame, row_idx: int, issue: str, severity: str = "warning") -> pd.DataFrame:
    mask = pd.Series(False, index=df.index)
    mask.iloc[row_idx] = True
    return _issue_rows(df, mask, issue, severity)


def rebuild_possession_scores(
    record: pd.DataFrame,
    meta: MatchMeta,
    checks: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Build possession-start scores; check official checkpoints and final score.

    A mismatching Video_Gap checkpoint explicitly resynchronizes the NEXT observed
    possession's starting score, without manufacturing an unseen possession/goal.
    Other mismatching checkpoints warn but do not silently modify the score.
    """
    df = record.copy().reset_index(drop=True)
    if df.empty:
        raise ValueError("Match_Record has no possession rows")
    if df["Possession_ID"].isna().any() or df["Possession_ID"].duplicated().any():
        raise ValueError("Match_Record Possession_ID must be nonmissing and unique for score reconstruction")
    ids = df["Possession_ID"].tolist()
    if any(float(pid) != int(pid) or int(pid) < 1 for pid in ids):
        raise ValueError("Possession_ID must contain positive whole numbers")
    if int(ids[0]) != meta.first_possession_id:
        raise ValueError(
            f"Match_Info First_Possession_ID={meta.first_possession_id}, "
            f"but Match_Record starts at {int(ids[0])}")
    valid_teams = {meta.team_home, meta.team_away}
    unexpected = sorted(set(df["Team"].dropna().tolist()) - valid_teams)
    if unexpected or df["Team"].isna().any():
        raise ValueError(f"Match_Record.Team must be {sorted(valid_teams)}; unexpected: {unexpected}")

    for idx, goal in enumerate(df["Goal"].tolist()):
        _nonnegative_integer(goal, f"Match_Record Possession_ID={int(ids[idx])} Goal")

    present_ids = {int(x) for x in ids}
    absent = sorted(set(checks["Before_Possession_ID"].tolist()) - present_ids)
    if absent:
        raise ValueError(f"Score_Check references nonexistent Before_Possession_ID: {absent}")
    check_by_id = {int(r["Before_Possession_ID"]): r for _, r in checks.iterrows()}

    home = meta.initial_score_home
    away = meta.initial_score_away
    timeline = []
    audit_rows = []
    warnings = []
    for i, row in df.iterrows():
        possession_id = int(row["Possession_ID"])
        sync_source = "Goal_Records"
        if possession_id in check_by_id:
            ch = check_by_id[possession_id]
            official_h, official_a = int(ch["Score_Home"]), int(ch["Score_Away"])
            agrees = (home == official_h and away == official_a)
            is_gap = ch["Check_Reason"] == "Video_Gap"
            resynced = bool(not agrees and is_gap)
            audit_rows.append({
                "Before_Possession_ID": possession_id,
                "Check_Reason": ch["Check_Reason"],
                "Computed_Score_Home": home,
                "Computed_Score_Away": away,
                "Official_Score_Home": official_h,
                "Official_Score_Away": official_a,
                "Matches": bool(agrees),
                "Resynced": resynced,
            })
            if not agrees:
                description = (
                    f"Score_Check {ch['Check_Reason']} mismatch: computed "
                    f"{meta.team_home} {home}:{away} {meta.team_away}, "
                    f"official {official_h}:{official_a}"
                )
                if is_gap:
                    description += "; resynced from verified Video_Gap"
                else:
                    description += "; NOT resynced (review scoring / footage)"
                warnings.append(_score_qc_issue(df, i, description))
                if is_gap:
                    home, away = official_h, official_a
                    sync_source = "Video_Gap_Resync"

        before_h, before_a = home, away
        diff = (home - away) if row["Team"] == meta.team_home else (away - home)
        match_status = "Leading" if diff > 0 else "Trailing" if diff < 0 else "Tied"
        timeline.append({
            "Score_Home_Before": before_h,
            "Score_Away_Before": before_a,
            "Score_Differential": diff,
            "Match_Status": match_status,
            "Score_Source": sync_source,
        })
        if row["Team"] == meta.team_home:
            home += int(row["Goal"])
        else:
            away += int(row["Goal"])

    for col in ["Score_Home_Before", "Score_Away_Before", "Score_Differential", "Match_Status", "Score_Source"]:
        df[col] = [entry[col] for entry in timeline]
    final_expected = (meta.official_final_score_home, meta.official_final_score_away)
    if (home, away) != final_expected:
        warnings.append(_score_qc_issue(
            df, len(df) - 1,
            f"Official final score mismatch: computed {meta.team_home} {home}:{away} {meta.team_away}, "
            f"official {final_expected[0]}:{final_expected[1]}; review missing footage or Goal counts",
        ))
    score_issues = pd.concat(warnings, ignore_index=True) if warnings else pd.DataFrame(
        columns=["Severity", "Issue", "Possession_ID"])
    audit = pd.DataFrame(audit_rows, columns=[
        "Before_Possession_ID", "Check_Reason", "Computed_Score_Home", "Computed_Score_Away",
        "Official_Score_Home", "Official_Score_Away", "Matches", "Resynced",
    ])
    audit.attrs["computed_final_score"] = (home, away)
    return df, score_issues, audit


def load_match_workbook(path: str | Path) -> tuple[pd.DataFrame, MatchMeta, pd.DataFrame, pd.DataFrame]:
    """Return (coded records with derived scores, metadata, score QC, checkpoint audit)."""
    meta = load_match_info(path)
    df = load_match_excel(path, sheet_name="Match_Record")
    checks = load_score_check(path)
    scored, score_qc, audit = rebuild_possession_scores(df, meta, checks)
    return scored, meta, score_qc, audit


def format_game_clock(value: Any) -> str | None:
    """Read MM:SS, or the special Excel time-of-day representation of MM:SS.

    In the user's coding file, typing 07:29 into Excel stores a time of 07:29 AM
    (fraction of a day). Here its *hour* is the game-clock minutes and the
    *minute* is the game-clock seconds. Do NOT multiply this by 60 twice.
    """
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        if not math.isfinite(float(value)) or value < 0:
            return None
        total_seconds = int(round(float(value) * 86400))
        return f"{(total_seconds // 3600) % 24:02d}:{(total_seconds % 3600) // 60:02d}"
    if hasattr(value, "hour") and hasattr(value, "minute"):
        return f"{int(value.hour):02d}:{int(value.minute):02d}"
    if hasattr(value, "total_seconds"):
        total = int(round(value.total_seconds()))
        mm, ss = divmod(total, 60)
        return f"{mm:02d}:{ss:02d}"
    string = str(value).strip()
    if not string:
        return None
    if string.count(":") == 2 and string.endswith(":00"):
        string = string[:-3]  # e.g. Excel time "07:29:00" -> 07:29
    if ":" not in string:
        return string
    try:
        minutes, seconds = map(int, string.split(":"))
        return f"{minutes:02d}:{seconds:02d}"
    except (ValueError, TypeError):
        return string


def clock_seconds(value: Any) -> float:
    """Return seconds remaining in quarter (or NaN if not parseable)."""
    text = format_game_clock(value)
    if text is None:
        return float("nan")
    try:
        minutes, seconds = map(int, text.split(":"))
    except (TypeError, ValueError):
        return float("nan")
    if minutes < 0 or minutes > 8 or seconds < 0 or seconds > 59 or minutes * 60 + seconds > 480:
        return float("nan")
    return float(minutes * 60 + seconds)


def add_time_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate full-possession and first-shot elapsed *game-clock* seconds.

    Missing/unobservable times remain NaN, never zero. In a game with an
    unknown opening-possession start, that partial possession is excluded from
    elapsed-time summaries while kept for non-time KPIs. This is game-clock
    elapsed time, not pure live-ball or shot-clock usage.
    """
    work = df.copy()
    for col in ["Game_Clock", "First_Shot_Clock", "End_Clock"]:
        if col + "_Display" not in work.columns:
            work[col + "_Display"] = work[col].apply(format_game_clock)
        work[col + "_Seconds"] = work[col].apply(clock_seconds)
    opening_unverified = (
        (work["Possession_ID"] == work.iloc[0]["Possession_ID"])
        & work["Start_Type"].isin(["Unknown"])
    )
    durations = []
    first_shots = []
    duration_quality = []
    first_shot_quality = []
    for i, row in work.iterrows():
        start, end, first = (row[c] for c in [
            "Game_Clock_Seconds", "End_Clock_Seconds", "First_Shot_Clock_Seconds"])
        valid_start = pd.notna(start)
        valid_end = pd.notna(end)
        valid_first = pd.notna(first)
        if opening_unverified.loc[i]:
            duration_quality.append("Unverified_Opening_Start")
            durations.append(float("nan"))
        elif not valid_end:
            duration_quality.append("Unobservable_End" if pd.isna(row["End_Clock"]) else "Invalid_End_Clock")
            durations.append(float("nan"))
        elif not valid_start or end > start:
            duration_quality.append("Invalid_Clock_Order")
            durations.append(float("nan"))
        else:
            duration_quality.append("Complete")
            durations.append(float(start-end))

        if row["Shot"] == 0:
            first_shot_quality.append("No_Shot")
            first_shots.append(float("nan"))
        elif opening_unverified.loc[i]:
            first_shot_quality.append("Unverified_Opening_Start")
            first_shots.append(float("nan"))
        elif not valid_first:
            first_shot_quality.append("Unobservable_First_Shot" if pd.isna(row["First_Shot_Clock"]) else "Invalid_First_Shot_Clock")
            first_shots.append(float("nan"))
        elif not valid_start or first > start or (valid_end and first < end):
            first_shot_quality.append("Invalid_Clock_Order")
            first_shots.append(float("nan"))
        else:
            first_shot_quality.append("Complete")
            first_shots.append(float(start-first))
    work["Possession_Duration_s"] = durations
    work["Time_to_First_Shot_s"] = first_shots
    work["Duration_Quality"] = duration_quality
    work["First_Shot_Quality"] = first_shot_quality
    return work


def load_infraction_log(path: str | Path) -> pd.DataFrame:
    """Read event-level records. An empty log with headings is valid."""
    try:
        raw = pd.read_excel(path, sheet_name="Infraction_Log", dtype=object)
    except ValueError as exc:
        raise ValueError("The workbook must contain an Infraction_Log worksheet.") from exc
    missing = [col for col in INFRACTION_COLUMNS if col not in raw.columns]
    if missing:
        raise ValueError(f"Infraction_Log missing columns: {missing}")
    log = raw[INFRACTION_COLUMNS].dropna(how="all").copy()
    # Keep incomplete event rows: QC must flag a missing ID, not silently delete the event.
    for col in ["Infraction_ID", "Possession_ID"]:
        log[col] = pd.to_numeric(log[col], errors="coerce")
    for col in ["Infraction_Team", "Infraction_Type", "Possession_Change"]:
        log[col] = log[col].astype("string").str.strip()
    log["Possession_Change"] = log["Possession_Change"].str.capitalize()
    log["Game_Clock_Display"] = log["Game_Clock"].apply(format_game_clock)
    log["Game_Clock_Seconds"] = log["Game_Clock"].apply(clock_seconds)
    return log.reset_index(drop=True)


def attach_infraction_counts(df: pd.DataFrame, log: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Join events to possessions and infer fouls drawn without manual duplicate coding.

    A defensive foul drawn is counted only if a KNOWN defensive foul (FOUL_TYPES) is
    committed by the team opposing the possession team. Violations, Other,
    Unknown and generic procedural incidents are not silently counted as fouls.
    """
    work = df.copy()
    events = log.copy()
    teams_by_id = work.set_index("Possession_ID")["Team"].to_dict()
    events["Possession_Team"] = events["Possession_ID"].map(teams_by_id)
    events["Is_Offensive_Infraction"] = events["Infraction_Team"] == events["Possession_Team"]
    events["Is_Defensive_Foul"] = (
        events["Possession_Team"].notna()
        & (events["Infraction_Team"] != events["Possession_Team"])
        & events["Infraction_Type"].isin(FOUL_TYPES)
    )
    count_drawn = events.groupby("Possession_ID")["Is_Defensive_Foul"].sum().to_dict()
    count_events = events.groupby("Possession_ID").size().to_dict()
    count_changes = (events[events["Possession_Change"] == "Yes"]
                     .groupby("Possession_ID").size().to_dict())
    work["Defensive_Foul"] = work["Possession_ID"].map(count_drawn).fillna(0).astype(int)
    work["Infraction_Count"] = work["Possession_ID"].map(count_events).fillna(0).astype(int)
    work["Possession_Change_Events"] = work["Possession_ID"].map(count_changes).fillna(0).astype(int)
    return work, events


def infraction_quality_check(df: pd.DataFrame, events: pd.DataFrame, meta: MatchMeta) -> pd.DataFrame:
    """Check event identifiers, possession links, labels, change/end consistency, timing."""
    findings = []
    p_by_id = df.set_index("Possession_ID").to_dict("index")
    event_counts_yes = {}
    for _, e in events.iterrows():
        event_id, pid = e["Infraction_ID"], e["Possession_ID"]
        def add(issue: str, severity: str = "error"):
            findings.append({"Severity": severity, "Issue": issue,
                             "Infraction_ID": event_id, "Possession_ID": pid})
        if pd.isna(event_id) or float(event_id) != int(event_id) or event_id < 1:
            add("Infraction_ID must be a positive integer")
        if pd.isna(pid) or float(pid) != int(pid) or pid not in p_by_id:
            add("Infraction_Log references unknown / invalid Possession_ID")
            continue
        record = p_by_id[pid]
        if e["Infraction_Team"] not in {meta.team_home, meta.team_away}:
            add("Infraction_Team must match Home/Away team")
        if pd.isna(e["Infraction_Type"]) or not str(e["Infraction_Type"]).strip():
            add("Infraction_Type is missing")
        elif e["Infraction_Type"] not in KNOWN_INFRACTION_TYPES:
            add(f"Unrecognized Infraction_Type: {e['Infraction_Type']}", "warning")
        if e["Possession_Change"] not in {"Yes", "No"}:
            add("Possession_Change must be Yes or No")
        if e["Possession_Change"] == "Yes":
            event_counts_yes[pid] = event_counts_yes.get(pid, 0) + 1
            if record["End_Reason"] != "Infraction_Loss":
                add("Possession_Change=Yes but End_Reason is not Infraction_Loss")
            if e["Infraction_Team"] != record["Team"]:
                add("Possession_Change=Yes by non-possessing team; review decision", "warning")
        tick = e["Game_Clock_Seconds"]
        if pd.isna(tick):
            add("Infraction event Game_Clock unavailable or invalid", "info")
        else:
            start, end = record["Game_Clock_Seconds"], record["End_Clock_Seconds"]
            if pd.notna(start) and tick > start:
                add("Infraction occurs before possession start")
            if pd.notna(end) and tick < end:
                add("Infraction occurs after possession end")
    for duplicated_id in events.loc[events["Infraction_ID"].duplicated(keep=False), "Infraction_ID"].dropna().unique():
        findings.append({"Severity": "error", "Issue": "Duplicate Infraction_ID",
                         "Infraction_ID": duplicated_id, "Possession_ID": pd.NA})
    for _, record in df[df["End_Reason"] == "Infraction_Loss"].iterrows():
        pid = record["Possession_ID"]
        if event_counts_yes.get(pid, 0) != 1:
            findings.append({"Severity": "error", "Issue":
                             "Infraction_Loss must have exactly one matching Possession_Change=Yes event",
                             "Infraction_ID": pd.NA, "Possession_ID": pid})
    return pd.DataFrame(findings, columns=["Severity", "Issue", "Infraction_ID", "Possession_ID"])


def time_quality_check(df: pd.DataFrame) -> pd.DataFrame:
    """Clock QC: explicit informational notes for known unobservable TV footage."""
    issues = []
    for column in ["Game_Clock", "End_Clock", "First_Shot_Clock"]:
        invalid = df[column].notna() & df[column + "_Seconds"].isna()
        issues.append(_issue_rows(df, invalid, f"Invalid {column} value", "error"))
    issues.append(_issue_rows(df, df["Game_Clock"].isna(), "Missing possession start clock", "error"))
    issues.append(_issue_rows(df, df["Duration_Quality"] == "Invalid_Clock_Order",
                              "End_Clock is later than Game_Clock", "error"))
    issues.append(_issue_rows(df, df["First_Shot_Quality"] == "Invalid_Clock_Order",
                              "First shot is outside possession time range", "error"))
    issues.append(_issue_rows(df, (df["Shot"] == 0) & df["First_Shot_Clock"].notna(),
                              "First_Shot_Clock given for possession without a shot", "error"))
    issues.append(_issue_rows(df, df["Duration_Quality"] == "Unobservable_End",
                              "End_Clock not observable in broadcast (duration excluded)", "info"))
    issues.append(_issue_rows(df, df["First_Shot_Quality"] == "Unobservable_First_Shot",
                              "First_Shot_Clock not observable despite Shot>0 (time-to-shot excluded)", "info"))
    issues.append(_issue_rows(df, df["Duration_Quality"] == "Unverified_Opening_Start",
                              "Opening possession start unverified (duration excluded)", "info"))
    valid_long = (df["Duration_Quality"] == "Complete") & (df["Possession_Duration_s"] > 30)
    issues.append(_issue_rows(df, valid_long,
                              "Duration >30s: check shot-clock reset / game interruptions (not necessarily an error)", "info"))
    have = [x for x in issues if not x.empty]
    return pd.concat(have, ignore_index=True) if have else pd.DataFrame(columns=["Severity", "Issue", "Possession_ID"])


def time_summary(df: pd.DataFrame, by_type: bool = False) -> pd.DataFrame:
    """Describe time ONLY on available, verified intervals; report denominators."""
    by = ["Team", "Possession_Type"] if by_type else ["Team"]
    rows = []
    for keys, g in df.groupby(by, dropna=False):
        if not isinstance(keys, tuple):
            keys = (keys,)
        duration = g["Possession_Duration_s"].dropna()
        first_shot = g["Time_to_First_Shot_s"].dropna()
        row = dict(zip(by, keys))
        row.update({
            "Possessions": len(g),
            "With_Shot": int((g["Shot"] > 0).sum()),
            "Duration_Valid_n": int(len(duration)),
            "Duration_Missing_n": int(len(g) - len(duration)),
            "Mean_Possession_Duration_s": float(duration.mean()),
            "Median_Possession_Duration_s": float(duration.median()),
            "First_Shot_Valid_n": int(len(first_shot)),
            "First_Shot_Missing_Among_Shot_n": int((g["Shot"] > 0).sum() - len(first_shot)),
            "Mean_Time_to_First_Shot_s": float(first_shot.mean()),
            "Median_Time_to_First_Shot_s": float(first_shot.median()),
        })
        rows.append(row)
    return pd.DataFrame(rows)


def infraction_summary(events: pd.DataFrame, by: list[str]) -> pd.DataFrame:
    if events.empty:
        return pd.DataFrame(columns=by + ["Count"])
    return (events.groupby(by, dropna=False).size().rename("Count")
            .reset_index().sort_values("Count", ascending=False).reset_index(drop=True))


# ---------------------------------------------------------------------------
# Quality control
# ---------------------------------------------------------------------------

def _issue_rows(df: pd.DataFrame, mask: pd.Series, issue: str, severity: str = "warning") -> pd.DataFrame:
    cols = [
        "Possession_ID",
        "Quarter",
        "Game_Clock_Display",
        "Team",
        "Start_Type",
        "Possession_Type",
        "Shot",
        "Shot_on_Goal",
        "Goal",
        "End_Reason",
    ]
    out = df.loc[mask, cols].copy()
    if out.empty:
        return out.assign(Severity=pd.Series(dtype="string"), Issue=pd.Series(dtype="string"))
    out.insert(0, "Severity", severity)
    out.insert(1, "Issue", issue)
    return out


def quality_check(df: pd.DataFrame) -> pd.DataFrame:
    """Return a row-level QC table. Empty output means no flagged issues.

    Rules are deliberately conservative: warnings do not automatically mean the
    coding is wrong; they identify possessions that deserve a quick analyst review.
    """
    issues: list[pd.DataFrame] = []

    # Required values.
    for col in ["Possession_ID", "Quarter", "Team", "End_Reason"]:
        issues.append(_issue_rows(df, df[col].isna(), f"Missing {col}", "error"))

    for col in COUNT_COLUMNS:
        issues.append(_issue_rows(df, df[col].isna(), f"Missing numeric count: {col}", "error"))
        issues.append(_issue_rows(df, df[col].fillna(0) < 0, f"Negative count: {col}", "error"))

    # Invalid end reason names must not silently propagate into KPI groups.
    issues.append(_issue_rows(df, df["End_Reason"].notna() & ~df["End_Reason"].isin(VALID_END_REASONS),
                              "Unrecognized End_Reason", "error"))

    # Duplicate IDs.
    dup = df["Possession_ID"].duplicated(keep=False)
    issues.append(_issue_rows(df, dup, "Duplicate Possession_ID", "error"))

    # Shot hierarchy: Goal <= SOG <= Shot.
    hierarchy_bad = (
        (df["Goal"].fillna(0) > df["Shot_on_Goal"].fillna(0))
        | (df["Shot_on_Goal"].fillna(0) > df["Shot"].fillna(0))
    )
    issues.append(_issue_rows(df, hierarchy_bad, "Shot hierarchy violated: Goal ≤ SOG ≤ Shot", "error"))

    # Goal / end reason consistency.
    issues.append(_issue_rows(
        df,
        (df["Goal"].fillna(0) > 0) & (df["End_Reason"] != "Goal"),
        "Goal recorded but End_Reason is not Goal",
        "error",
    ))
    issues.append(_issue_rows(
        df,
        (df["End_Reason"] == "Goal") & (df["Goal"].fillna(0) == 0),
        "End_Reason is Goal but Goal count is 0",
        "error",
    ))

    # Unknown/other tactical classifications are legal but useful to monitor.
    issues.append(_issue_rows(
        df,
        df["Start_Type"].isin(["Other", "Unknown"]) | df["Start_Type"].isna(),
        "Start_Type is Other/Unknown/missing",
        "warning",
    ))
    issues.append(_issue_rows(
        df,
        df["Possession_Type"].isin(["Other", "Unknown"]) | df["Possession_Type"].isna(),
        "Possession_Type is Other/Unknown/missing",
        "warning",
    ))
    issues.append(_issue_rows(
        df,
        df["Start_Zone"].isin(["Unknown"]) | df["Start_Zone"].isna(),
        "Start_Zone is Unknown/missing",
        "warning",
    ))

    # Possession sequence should normally alternate teams inside a quarter.
    # Same-team starts are flagged only when the previous possession did not end the period.
    prev_team = df["Team"].shift(1)
    prev_quarter = df["Quarter"].shift(1)
    prev_end = df["End_Reason"].shift(1)
    same_team_unexpected = (
        (df["Team"] == prev_team)
        & (df["Quarter"] == prev_quarter)
        & (prev_end != "Period_End")
    )
    issues.append(_issue_rows(
        df,
        same_team_unexpected.fillna(False),
        "Consecutive possessions by same team within quarter; verify possession boundary",
        "warning",
    ))

    # Goal should normally be followed by opponent possession / goal restart,
    # except where footage is missing or a period boundary intervenes.
    next_team = df["Team"].shift(-1)
    next_start = df["Start_Type"].shift(-1)
    next_quarter = df["Quarter"].shift(-1)
    goal_followup_bad = (
        (df["End_Reason"] == "Goal")
        & (df["Quarter"] == next_quarter)
        & ((next_team == df["Team"]) | (~next_start.isin(["Goal_Restart", "Unknown", "Other"])))
    )
    issues.append(_issue_rows(
        df,
        goal_followup_bad.fillna(False),
        "Goal follow-up sequence unusual; verify next possession",
        "warning",
    ))

    nonempty = [x for x in issues if not x.empty]
    if not nonempty:
        return pd.DataFrame(columns=["Severity", "Issue", "Possession_ID"])
    return pd.concat(nonempty, ignore_index=True).sort_values(
        ["Severity", "Possession_ID"],
        key=lambda s: s.map({"error": 0, "warning": 1}).fillna(s) if s.name == "Severity" else s,
    ).reset_index(drop=True)


# ---------------------------------------------------------------------------
# KPI engine
# ---------------------------------------------------------------------------

def safe_div(numerator: float | int, denominator: float | int) -> float:
    if denominator is None or denominator == 0 or pd.isna(denominator):
        return float("nan")
    return float(numerator) / float(denominator)


def team_kpis(team_df: pd.DataFrame) -> dict[str, float | int]:
    n_poss = len(team_df)
    shots = float(team_df["Shot"].fillna(0).sum())
    sog = float(team_df["Shot_on_Goal"].fillna(0).sum())
    goals = float(team_df["Goal"].fillna(0).sum())
    passes = float(team_df["Pass"].fillna(0).sum())
    fouls_drawn = float(team_df["Defensive_Foul"].fillna(0).sum()) if "Defensive_Foul" in team_df else 0.0
    timeouts = float(team_df["TimeOut"].fillna(0).sum())

    valid_type = team_df[team_df["Possession_Type"].isin(["Transition", "Settled_Offense"])]
    transition = valid_type[valid_type["Possession_Type"] == "Transition"]
    settled = valid_type[valid_type["Possession_Type"] == "Settled_Offense"]

    controllable_losses = int(team_df["End_Reason"].isin(CONTROLLABLE_LOSS_REASONS).sum())

    return {
        "Possessions": int(n_poss),
        "Passes": int(passes),
        "Shots": int(shots),
        "SOG": int(sog),
        "Goals": int(goals),
        "Defensive_Fouls_Drawn": int(fouls_drawn),
        "Timeouts": int(timeouts),
        "Passes_per_Possession": safe_div(passes, n_poss),
        "Offensive_Efficiency": safe_div(goals, n_poss),
        "Shot_Generation_Rate": safe_div(int((team_df["Shot"].fillna(0) > 0).sum()), n_poss),
        "Shot_Accuracy": safe_div(sog, shots),
        "Shooting_Efficiency": safe_div(goals, shots),
        "SOG_Conversion": safe_div(goals, sog),
        "Transition_Frequency": safe_div(len(transition), len(valid_type)),
        "Transition_Efficiency": safe_div(transition["Goal"].fillna(0).sum(), len(transition)),
        "Settled_Efficiency": safe_div(settled["Goal"].fillna(0).sum(), len(settled)),
        "Controllable_Loss_Rate": safe_div(controllable_losses, n_poss),
        "Defensive_Fouls_Drawn_per_Possession": safe_div(fouls_drawn, n_poss),
        "Timeouts_per_Possession": safe_div(timeouts, n_poss),
        "Defensive_Fouls_Drawn_per_100_Possessions": safe_div(fouls_drawn * 100, n_poss),
        "Timeouts_per_100_Possessions": safe_div(timeouts * 100, n_poss),
    }


def core_kpi_table(df: pd.DataFrame, teams: Iterable[str] | None = None) -> pd.DataFrame:
    if teams is None:
        teams = [x for x in df["Team"].dropna().unique().tolist()]
    rows = []
    for team in teams:
        row = {"Team": team}
        row.update(team_kpis(df[df["Team"] == team]))
        rows.append(row)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Tactical breakdowns
# ---------------------------------------------------------------------------

def _group_summary(group: pd.DataFrame) -> pd.Series:
    n = len(group)
    shots = group["Shot"].fillna(0).sum()
    sog = group["Shot_on_Goal"].fillna(0).sum()
    goals = group["Goal"].fillna(0).sum()
    losses = group["End_Reason"].isin(CONTROLLABLE_LOSS_REASONS).sum()
    return pd.Series({
        "Possessions": n,
        "Share": float("nan"),  # filled within each team below
        "Passes_per_Possession": safe_div(group["Pass"].fillna(0).sum(), n),
        "Shot_Generation_Rate": safe_div((group["Shot"].fillna(0) > 0).sum(), n),
        "Shot_Accuracy": safe_div(sog, shots),
        "Shooting_Efficiency": safe_div(goals, shots),
        "Goals_per_Possession": safe_div(goals, n),
        "Controllable_Loss_Rate": safe_div(losses, n),
        "Defensive_Fouls_Drawn_per_Possession": safe_div(group["Defensive_Foul"].fillna(0).sum() if "Defensive_Foul" in group else 0, n),
    })


def breakdown(df: pd.DataFrame, by: str, include_unknown: bool = True) -> pd.DataFrame:
    if by not in df.columns:
        raise KeyError(by)

    work = df.copy()
    if not include_unknown:
        work = work[~work[by].isin(["Other", "Unknown"]) & work[by].notna()]

    out = (
        work.groupby(["Team", by], dropna=False, observed=False)
        .apply(_group_summary, include_groups=False)
        .reset_index()
    )

    if not out.empty:
        team_totals = out.groupby("Team")["Possessions"].transform("sum")
        out["Share"] = out["Possessions"] / team_totals
    return out


def end_reason_summary(df: pd.DataFrame) -> pd.DataFrame:
    out = (
        df.groupby(["Team", "End_Reason"], dropna=False)
        .size()
        .rename("Count")
        .reset_index()
    )
    totals = out.groupby("Team")["Count"].transform("sum")
    out["Share"] = out["Count"] / totals
    return out.sort_values(["Team", "Count"], ascending=[True, False]).reset_index(drop=True)


def quarter_summary(df: pd.DataFrame) -> pd.DataFrame:
    out = (
        df.groupby(["Team", "Quarter"], dropna=False, observed=False)
        .apply(_group_summary, include_groups=False)
        .reset_index()
    )
    return out


def possession_funnel(df: pd.DataFrame) -> pd.DataFrame:
    """Coach-facing possession -> shot -> SOG -> goal funnel by team.

    Shot/SOG/Goal stages use possessions containing at least one event, rather
    than total event counts, so the funnel remains a possession conversion view.
    """
    rows = []
    for team, g in df.groupby("Team", dropna=False):
        n = len(g)
        shot_poss = int((g["Shot"].fillna(0) > 0).sum())
        sog_poss = int((g["Shot_on_Goal"].fillna(0) > 0).sum())
        goal_poss = int((g["Goal"].fillna(0) > 0).sum())
        rows.append({
            "Team": team,
            "Possessions": n,
            "Possessions_with_Shot": shot_poss,
            "Possessions_with_SOG": sog_poss,
            "Possessions_with_Goal": goal_poss,
            "Shot_Possession_Rate": safe_div(shot_poss, n),
            "SOG_Possession_Rate": safe_div(sog_poss, n),
            "Goal_Possession_Rate": safe_div(goal_poss, n),
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Video-review lists
# ---------------------------------------------------------------------------

def _review_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "Possession_ID",
        "Quarter",
        "Game_Clock_Display",
        "First_Shot_Clock_Display",
        "End_Clock_Display",
        "Possession_Duration_s",
        "Time_to_First_Shot_s",
        "Team",
        "Score_Differential",
        "Start_Type",
        "Start_Zone",
        "Possession_Type",
        "Pass",
        "Shot",
        "Shot_on_Goal",
        "Goal",
        "TimeOut",
        "End_Reason",
    ]
    return df[cols].copy()


def review_lists(df: pd.DataFrame, focus_team: str = "TPE",
                 events: pd.DataFrame | None = None) -> dict[str, pd.DataFrame]:
    team = df[df["Team"] == focus_team]
    if events is None:
        events = pd.DataFrame(columns=["Possession_ID", "Infraction_Type", "Possession_Team"])
    def from_infraction(event_type: str) -> pd.DataFrame:
        pids = events.loc[
            (events["Possession_Team"] == focus_team)
            & (events["Infraction_Type"] == event_type), "Possession_ID"].unique()
        return team[team["Possession_ID"].isin(pids)]
    lists = {
        "shot_clock_violations": from_infraction("Shot_Clock_Violation"),
        "ten_second_violations": from_infraction("Ten_Second_Violation"),
        "transition_without_shot": team[
            (team["Possession_Type"] == "Transition") & (team["Shot"].fillna(0) == 0)
        ],
        "settled_without_shot": team[
            (team["Possession_Type"] == "Settled_Offense") & (team["Shot"].fillna(0) == 0)
        ],
        "live_ball_gain_goals": team[
            (team["Start_Type"] == "Live_ball_Gain") & (team["Goal"].fillna(0) > 0)
        ],
        "infraction_losses": team[team["End_Reason"] == "Infraction_Loss"],
        "timeouts": team[team["TimeOut"].fillna(0) > 0],
        "miss_losses": team[team["End_Reason"] == "Miss_Loss"],
    }
    return {name: _review_columns(x) for name, x in lists.items()}


# ---------------------------------------------------------------------------
# Match package / export
# ---------------------------------------------------------------------------

def analyze_match(
    df: pd.DataFrame, meta: MatchMeta,
    score_qc: pd.DataFrame | None = None,
    score_audit: pd.DataFrame | None = None,
    infractions: pd.DataFrame | None = None,
) -> dict[str, Any]:
    """Analyze a V1.2 workbook's records, scores, infraction events, and clocks."""
    if infractions is None:
        infractions = pd.DataFrame(columns=INFRACTION_COLUMNS)
    work = add_time_metrics(df)
    work, enriched_events = attach_infraction_counts(work, infractions)
    iqc = infraction_quality_check(work, enriched_events, meta)
    qc = pd.concat([quality_check(work), time_quality_check(work), iqc], ignore_index=True)
    if score_qc is not None and not score_qc.empty:
        qc = pd.concat([qc, score_qc], ignore_index=True)
    return {
        "meta": meta,
        "qc": qc,
        "score_audit": score_audit if score_audit is not None else pd.DataFrame(),
        "possession_scores": work[[
            "Possession_ID", "Quarter", "Game_Clock_Display", "Team", "Goal",
            "Score_Home_Before", "Score_Away_Before", "Score_Differential",
            "Match_Status", "Score_Source",
        ]].copy(),
        "possession_times": work[[
            "Possession_ID", "Quarter", "Team", "Possession_Type", "Shot",
            "Game_Clock_Display", "First_Shot_Clock_Display", "End_Clock_Display",
            "Possession_Duration_s", "Time_to_First_Shot_s",
            "Duration_Quality", "First_Shot_Quality",
        ]].copy(),
        "time_by_team": time_summary(work),
        "time_by_possession_type": time_summary(work, by_type=True),
        "infractions": enriched_events,
        "infraction_by_type": infraction_summary(enriched_events, ["Infraction_Team", "Infraction_Type"]),
        "infraction_by_team": infraction_summary(enriched_events, ["Infraction_Team", "Possession_Change"]),
        "kpis": core_kpi_table(work),
        "funnel": possession_funnel(work),
        "by_possession_type": breakdown(work, "Possession_Type"),
        "by_start_type": breakdown(work, "Start_Type"),
        "by_start_zone": breakdown(work, "Start_Zone"),
        "by_quarter": quarter_summary(work),
        "end_reasons": end_reason_summary(work),
        "review_lists": review_lists(work, meta.focus_team, enriched_events),
    }


def export_analysis(package: dict[str, Any], output_dir: str | Path) -> Path:
    """Export flat CSV files for auditing / poster work / future Streamlit use."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    tables = {
        "qc": package["qc"],
        "score_checks": package["score_audit"],
        "possession_scores": package["possession_scores"],
        "possession_times": package["possession_times"],
        "time_by_team": package["time_by_team"],
        "time_by_possession_type": package["time_by_possession_type"],
        "infractions": package["infractions"],
        "infraction_by_type": package["infraction_by_type"],
        "infraction_by_team": package["infraction_by_team"],
        "core_kpis": package["kpis"],
        "possession_funnel": package["funnel"],
        "by_possession_type": package["by_possession_type"],
        "by_start_type": package["by_start_type"],
        "by_start_zone": package["by_start_zone"],
        "by_quarter": package["by_quarter"],
        "end_reasons": package["end_reasons"],
    }
    for name, table in tables.items():
        table.to_csv(output_dir / f"{name}.csv", index=False, encoding="utf-8-sig")

    review_dir = output_dir / "review_lists"
    review_dir.mkdir(exist_ok=True)
    for name, table in package["review_lists"].items():
        table.to_csv(review_dir / f"{name}.csv", index=False, encoding="utf-8-sig")

    meta = package["meta"]
    with open(output_dir / "meta.json", "w", encoding="utf-8") as f:
        json.dump(meta.__dict__, f, ensure_ascii=False, indent=2)

    return output_dir


def percent_columns(table: pd.DataFrame) -> pd.DataFrame:
    """Presentation helper: convert 0-1 ratio fields to percentage points.

    Important: count-normalized metrics such as Passes_per_Possession and
    *_per_100_Possessions are intentionally NOT converted here.
    """
    out = table.copy()
    explicit_ratio_cols = {
        "Offensive_Efficiency",
        "Shot_Generation_Rate",
        "Shot_Accuracy",
        "Shooting_Efficiency",
        "SOG_Conversion",
        "Transition_Frequency",
        "Transition_Efficiency",
        "Settled_Efficiency",
        "Controllable_Loss_Rate",
        "Share",
        "Goals_per_Possession",
        "Shot_Possession_Rate",
        "SOG_Possession_Rate",
        "Goal_Possession_Rate",
    }
    for col in explicit_ratio_cols.intersection(out.columns):
        out[col] = out[col] * 100
    return out


def compact_match_summary(kpis: pd.DataFrame) -> pd.DataFrame:
    """Return a coach-friendly one-row-per-team KPI summary."""
    cols = [
        "Team",
        "Possessions",
        "Passes_per_Possession",
        "Offensive_Efficiency",
        "Shot_Generation_Rate",
        "Shot_Accuracy",
        "Shooting_Efficiency",
        "Transition_Frequency",
        "Transition_Efficiency",
        "Settled_Efficiency",
        "Controllable_Loss_Rate",
    ]
    out = percent_columns(kpis[cols].copy())
    rename = {
        "Passes_per_Possession": "Pass/Poss",
        "Offensive_Efficiency": "Off Eff %",
        "Shot_Generation_Rate": "Shot Gen %",
        "Shot_Accuracy": "Shot Acc %",
        "Shooting_Efficiency": "Shoot Eff %",
        "Transition_Frequency": "Transition %",
        "Transition_Efficiency": "Trans Eff %",
        "Settled_Efficiency": "Settled Eff %",
        "Controllable_Loss_Rate": "Loss %",
    }
    return out.rename(columns=rename)


def compact_possession_type_summary(table: pd.DataFrame) -> pd.DataFrame:
    """Coach-friendly Transition vs Settled table."""
    work = table[table["Possession_Type"].isin(["Transition", "Settled_Offense"])].copy()
    cols = [
        "Team",
        "Possession_Type",
        "Possessions",
        "Share",
        "Passes_per_Possession",
        "Shot_Generation_Rate",
        "Shot_Accuracy",
        "Shooting_Efficiency",
        "Goals_per_Possession",
        "Controllable_Loss_Rate",
    ]
    work = percent_columns(work[cols])
    return work.rename(columns={
        "Possession_Type": "Type",
        "Share": "Share %",
        "Passes_per_Possession": "Pass/Poss",
        "Shot_Generation_Rate": "Shot Gen %",
        "Shot_Accuracy": "Shot Acc %",
        "Shooting_Efficiency": "Shoot Eff %",
        "Goals_per_Possession": "Goal/Poss %",
        "Controllable_Loss_Rate": "Loss %",
    })


# ---------------------------------------------------------------------------
# Command-line entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Sixes Lacrosse match analysis engine v0.3.1")
    parser.add_argument("xlsx", help="Coding Template V1.2 (.xlsx) with 4 worksheets")
    parser.add_argument("--focus-team", default=None, help="Team for review lists (default: official Home)")
    parser.add_argument("--output", default="analysis_output", help="Directory for CSV outputs")
    args = parser.parse_args()

    df, workbook_meta, score_qc, score_audit = load_match_workbook(args.xlsx)
    if args.focus_team is not None and args.focus_team not in {workbook_meta.team_home, workbook_meta.team_away}:
        parser.error(f"--focus-team must be {workbook_meta.team_home} or {workbook_meta.team_away}")
    meta = workbook_meta
    if args.focus_team is not None:
        from dataclasses import replace
        meta = replace(meta, focus_team=args.focus_team)
    infraction_log = load_infraction_log(args.xlsx)
    package = analyze_match(df, meta, score_qc=score_qc, score_audit=score_audit,
                            infractions=infraction_log)
    export_analysis(package, args.output)

    print("\n=== SCORE RECONSTRUCTION ===")
    if score_audit.empty:
        print("Score_Check: no intermediate checkpoints (allowed).")
    else:
        print(score_audit.to_string(index=False))
    print(f"First observed possession: {int(df.iloc[0]['Possession_ID'])}; "
          f"start score: {meta.team_home} {meta.initial_score_home}:"
          f"{meta.initial_score_away} {meta.team_away}")
    end_home = int(df.iloc[-1]["Score_Home_Before"]) + (
        int(df.iloc[-1]["Goal"]) if df.iloc[-1]["Team"] == meta.team_home else 0)
    end_away = int(df.iloc[-1]["Score_Away_Before"]) + (
        int(df.iloc[-1]["Goal"]) if df.iloc[-1]["Team"] == meta.team_away else 0)
    print(f"Calculated final score: {meta.team_home} {end_home}:"
          f"{end_away} {meta.team_away} "
          f"(official {meta.official_final_score_home}:{meta.official_final_score_away})")

    print("\n=== DATA QC ===")
    if package["qc"].empty:
        print("No QC issues flagged.")
    else:
        n_error = int((package["qc"]["Severity"] == "error").sum())
        n_warning = int((package["qc"]["Severity"] == "warning").sum())
        n_info = int((package["qc"]["Severity"] == "info").sum())
        print(f"QC summary: {n_error} error(s), {n_warning} warning(s), {n_info} info note(s)")
        print(package["qc"].to_string(index=False))

    print("\n=== MATCH SUMMARY ===")
    print(compact_match_summary(package["kpis"]).round(1).to_string(index=False))

    print("\n=== TRANSITION VS SETTLED ===")
    print(compact_possession_type_summary(package["by_possession_type"]).round(1).to_string(index=False))

    print("\n=== POSSESSION TIME (observed intervals only; seconds) ===")
    print(package["time_by_team"].round(1).to_string(index=False))
    print("\n=== TIME BY POSSESSION TYPE ===")
    print(package["time_by_possession_type"].round(1).to_string(index=False))
    print("\n=== INFRACTIONS BY TEAM / TYPE ===")
    if package["infraction_by_type"].empty:
        print("No infraction events recorded.")
    else:
        print(package["infraction_by_type"].to_string(index=False))
    print("\n=== LOW-FREQUENCY EVENTS ===")
    low = package["kpis"][[
        "Team",
        "Defensive_Fouls_Drawn",
        "Timeouts",
        "Defensive_Fouls_Drawn_per_100_Possessions",
        "Timeouts_per_100_Possessions",
    ]].copy()
    low = low.rename(columns={
        "Defensive_Fouls_Drawn": "Def Fouls Drawn",
        "Timeouts": "Timeouts",
        "Defensive_Fouls_Drawn_per_100_Possessions": "Def Fouls/100 Poss",
        "Timeouts_per_100_Possessions": "Timeouts/100 Poss",
    })
    print(low.round(1).to_string(index=False))

    print(f"\nExported analysis files to: {Path(args.output).resolve()}")


if __name__ == "__main__":
    main()
