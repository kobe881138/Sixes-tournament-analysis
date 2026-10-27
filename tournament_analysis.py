#!/usr/bin/env python3
"""Tournament Analysis Engine v0.1.2 for Coding Template V1.2.

Does NOT modify match Excel files. Computes match-by-match and pooled team metrics;
tracks unlinked infractions and documented delayed-possession-award events.
Uses the tested single-match analysis.py v0.3.2 calculation definitions.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
import pandas as pd
import analysis as single

# Recognize project-specific event labels without changing single-match code on disk.
single.KNOWN_INFRACTION_TYPES = single.KNOWN_INFRACTION_TYPES | {
    "Loose_Ball_Push", "Faceoff_Violation"
}

INFRACTION_ALIASES = {
    "Loose_Ball_Pushing": "Loose_Ball_Push",
    "Other/Unknown": "Unknown",
}
ALL_VALID_INFRACTIONS = (
    single.KNOWN_INFRACTION_TYPES
    | {"Loose_Ball_Push", "Faceoff_Violation", "Unknown"}
)
EXPORTS = [
    "Match_Summary", "Possession_Master", "Infraction_Master",
    "Team_Match_KPI", "Team_Tournament_KPI", "Team_Type_Profile",
    "Team_End_Reasons", "QC_Report", "Coding_Exceptions",
]


def as_int(value, fallback=0):
    try:
        return int(value) if pd.notna(value) else fallback
    except (TypeError, ValueError):
        return fallback


def read_exceptions(path: Path) -> dict:
    if not path.is_file():
        return {}
    with path.open(encoding="utf-8") as f:
        value = json.load(f)
    if not isinstance(value, dict):
        raise ValueError("coding_exceptions.json must contain a JSON object")
    return value


def load_events_preserving_na_label(path: Path) -> pd.DataFrame:
    """Unlike pandas defaults, literal Excel N/A denotes an intentional code here."""
    raw = pd.read_excel(path, sheet_name="Infraction_Log", dtype=object, keep_default_na=False)
    missing = set(single.INFRACTION_COLUMNS) - set(raw.columns)
    if missing:
        raise ValueError(f"Infraction_Log missing columns: {sorted(missing)}")
    raw = raw[single.INFRACTION_COLUMNS].replace("", pd.NA).dropna(how="all")
    for c in ("Infraction_ID", "Possession_ID"):
        raw[c] = pd.to_numeric(raw[c], errors="coerce")
    for c in ("Infraction_Team", "Infraction_Type", "Possession_Change"):
        raw[c] = raw[c].astype("string").str.strip()
    raw["Possession_Change"] = raw["Possession_Change"].str.capitalize().replace({"N/a": "N/A"})
    raw["Game_Clock_Display"] = raw["Game_Clock"].apply(single.format_game_clock)
    raw["Game_Clock_Seconds"] = raw["Game_Clock"].apply(single.clock_seconds)
    return raw.reset_index(drop=True)


def transform_events(original: pd.DataFrame, record: pd.DataFrame, meta,
                     exceptions: dict, filename: str):
    """Return (all-event master, linked events for single engine, notes).

    The source's Possession_Change is NEVER overwritten in Infraction_Master.
    Adjusted link/label/Yes-No are explicit analysis-only fields.
    """
    events = original.copy()
    events["Source_File"] = filename
    events["Match_ID"] = meta.match_id
    events["Original_Infraction_Type"] = events["Infraction_Type"]
    events["Original_Possession_Change"] = events["Possession_Change"]
    events["Analysis_Infraction_Type"] = events["Infraction_Type"].replace(INFRACTION_ALIASES)
    events["Analysis_Possession_Change"] = events["Possession_Change"]
    events["Coding_Exception"] = ""
    events["Event_Scope"] = "Linked_Possession"
    notes = []
    valid_ids = set(record["Possession_ID"].dropna().astype(int))
    for idx, e in events.iterrows():
        event_id = as_int(e["Infraction_ID"], -1)
        pid = e["Possession_ID"]
        inf_type = str(e["Analysis_Infraction_Type"])
        if event_id < 1 or pd.isna(e["Infraction_ID"]) or float(e["Infraction_ID"]) != event_id:
            notes.append(dict(Match_ID=meta.match_id, Severity="error",
                              Issue="Infraction_ID must be a positive integer",
                              Infraction_ID=event_id, Possession_ID=pid))
        if pd.isna(e["Infraction_Team"]) or str(e["Infraction_Team"]) not in {meta.team_home, meta.team_away}:
            notes.append(dict(Match_ID=meta.match_id, Severity="error",
                              Issue="Infraction_Team must be an official match team",
                              Infraction_ID=event_id, Possession_ID=pid))
        if inf_type not in ALL_VALID_INFRACTIONS:
            notes.append(dict(Match_ID=meta.match_id, Severity="warning",
                              Issue=f"Unrecognized infraction label {inf_type}",
                              Infraction_ID=event_id, Possession_ID=pid))
        if pd.isna(pid):
            events.at[idx, "Event_Scope"] = "No_Established_Possession"
            if str(e["Possession_Change"]).upper() != "N/A":
                notes.append(dict(Match_ID=meta.match_id, Severity="error",
                                  Issue="Unlinked event requires Possession_Change=N/A",
                                  Infraction_ID=event_id, Possession_ID=pid))
            else:
                notes.append(dict(Match_ID=meta.match_id, Severity="info",
                                  Issue="Unlinked event retained; excluded from possession-based turnovers",
                                  Infraction_ID=event_id, Possession_ID=pid))
        elif as_int(pid, -1) not in valid_ids:
            notes.append(dict(Match_ID=meta.match_id, Severity="error",
                              Issue="Infraction references nonexistent Possession_ID",
                              Infraction_ID=event_id, Possession_ID=pid))
        # A delayed award is not an immediate possession change. Original
        # Yes / No / N/A is retained for audit in the event master.
        annotation = exceptions.get(meta.match_id, {}).get(str(event_id))
        if annotation:
            if annotation.get("type") != "Next_Quarter_Possession_Award":
                notes.append(dict(Match_ID=meta.match_id, Severity="error",
                                  Issue=f"Unsupported exception type: {annotation.get('type')}",
                                  Infraction_ID=event_id, Possession_ID=pid))
            elif (pd.isna(pid)
                  or record.loc[record["Possession_ID"] == pid].empty):
                notes.append(dict(Match_ID=meta.match_id, Severity="error",
                                  Issue="Documented next-quarter award no longer matches source; review exception",
                                  Infraction_ID=event_id, Possession_ID=pid))
            else:
                next_quarter = annotation.get("next_quarter")
                awarded_team = annotation.get("awarded_team")
                source_row = record.loc[record["Possession_ID"] == pid].iloc[0]
                next_rows = record.loc[record["Quarter"] == next_quarter].sort_values("Possession_ID")
                if (source_row["Quarter"] != annotation.get("source_quarter")
                    or next_rows.empty or next_rows.iloc[0]["Team"] != awarded_team
                    or next_rows.iloc[0]["Start_Type"] == "Faceoff"):
                    notes.append(dict(Match_ID=meta.match_id, Severity="error",
                                      Issue="Next-quarter award exception contradicts source possession and restart; review",
                                      Infraction_ID=event_id, Possession_ID=pid))
                else:
                    # This event determines how the *next* quarter starts; it does
                    # not take the current possession away. Source code may be
                    # Yes (award of next restart), No, or N/A (no immediate transfer).
                    # Keep the submitted source value in Infraction_Master.
                    original_change = (str(e["Possession_Change"])
                                       if pd.notna(e["Possession_Change"]) else "")
                    if original_change not in {"Yes", "No", "N/A"}:
                        notes.append(dict(Match_ID=meta.match_id, Severity="warning",
                                          Issue="Documented next-quarter award has blank/invalid source Possession_Change; review original value",
                                          Infraction_ID=event_id, Possession_ID=pid))
                    events.at[idx, "Coding_Exception"] = annotation.get("type")
                    events.at[idx, "Event_Scope"] = "Next_Quarter_Award"
                    # Keep event-level meaning explicit; the old single-match QC
                    # accepts only Yes/No, so its input gets a technical No below.
                    events.at[idx, "Analysis_Possession_Change"] = "N/A"
                    notes.append(dict(Match_ID=meta.match_id, Severity="info",
                                      Issue="Documented next-quarter award validated; no immediate possession change/turnover counted",
                                      Infraction_ID=event_id, Possession_ID=pid))
    events["Infraction_Key"] = events["Infraction_ID"].apply(
        lambda x: f"{meta.match_id}:I{as_int(x, -1)}")
    events["Possession_Key"] = events["Possession_ID"].apply(
        lambda x: f"{meta.match_id}:P{as_int(x)}" if pd.notna(x) else "")
    linked = events.loc[
        events["Event_Scope"].isin(["Linked_Possession", "Next_Quarter_Award"])
        & events["Possession_ID"].notna(), single.INFRACTION_COLUMNS
        + ["Analysis_Infraction_Type", "Analysis_Possession_Change"]
    ].copy()
    linked["Infraction_Type"] = linked["Analysis_Infraction_Type"]
    # Adapter for the unchanged single-match QC, which requires Yes/No.
    # N/A here means no immediate turnover; the master retains N/A and scope.
    linked["Possession_Change"] = linked["Analysis_Possession_Change"].replace({"N/A": "No"})
    linked = linked[single.INFRACTION_COLUMNS]
    linked["Game_Clock_Seconds"] = linked["Game_Clock"].apply(single.clock_seconds)
    linked["Game_Clock_Display"] = linked["Game_Clock"].apply(single.format_game_clock)
    # For event-level reporting we do not mix the derived values into source fields.
    return events, linked, notes


def load_one(path: Path, documented: dict):
    scored, meta, score_qc, audit = single.load_match_workbook(path)
    raw = load_events_preserving_na_label(path)
    master_events, linked, notes = transform_events(raw, scored, meta, documented, path.name)
    result = single.analyze_match(scored, meta, score_qc=score_qc,
                                  score_audit=audit, infractions=linked)
    qc = result["qc"].copy()
    if notes:
        qc = pd.concat([qc, pd.DataFrame(notes)], ignore_index=True)
    qc["Match_ID"] = meta.match_id
    qc["Source_File"] = path.name
    work = single.add_time_metrics(scored)
    work, _ = single.attach_infraction_counts(work, linked)
    work["Match_ID"] = meta.match_id
    work["Source_File"] = path.name
    work["Possession_Key"] = work["Possession_ID"].apply(lambda x: f"{meta.match_id}:P{int(x)}")
    work["Opponent"] = work["Team"].apply(lambda t: meta.team_away if t == meta.team_home else meta.team_home)
    home_goals = as_int(work.loc[work["Team"] == meta.team_home, "Goal"].sum()) + meta.initial_score_home
    away_goals = as_int(work.loc[work["Team"] == meta.team_away, "Goal"].sum()) + meta.initial_score_away
    score_match = ((home_goals, away_goals) ==
                   (meta.official_final_score_home, meta.official_final_score_away))
    summary = {
        "Match_ID": meta.match_id, "Source_File": path.name,
        "Match_Date_Label": f"{meta.match_id[:2]}/{meta.match_id[2:4]}", "Home": meta.team_home, "Away": meta.team_away,
        "Recorded_Possessions": len(work),
        "Initial_Home": meta.initial_score_home, "Initial_Away": meta.initial_score_away,
        "Calculated_Home": home_goals, "Calculated_Away": away_goals,
        "Official_Home": meta.official_final_score_home, "Official_Away": meta.official_final_score_away,
        "Score_Matches": score_match,
        "QC_Errors": int((qc["Severity"] == "error").sum()),
        "QC_Warnings": int((qc["Severity"] == "warning").sum()),
    }
    if not score_match:
        qc = pd.concat([qc, pd.DataFrame([{
            "Match_ID": meta.match_id, "Source_File": path.name, "Severity": "error",
            "Issue": "Tournament score does not match official final score", "Possession_ID": None,
        }])], ignore_index=True)
        summary["QC_Errors"] += 1
    return summary, work, master_events, qc, meta


def summarize_team_match(work, meta):
    kpis = single.core_kpi_table(work, [meta.team_home, meta.team_away])
    timings = single.time_summary(work)
    kpis = kpis.merge(timings, on=["Team", "Possessions"], how="left", validate="1:1")
    kpis.insert(0, "Match_ID", meta.match_id)
    kpis.insert(2, "Opponent", kpis["Team"].apply(
        lambda x: meta.team_away if x == meta.team_home else meta.team_home))
    kpis["Shot_Possessions"] = [int((work.loc[work["Team"] == t, "Shot"] > 0).sum()) for t in kpis["Team"]]
    kpis["Transition_Possessions"] = [int((work.loc[work["Team"] == t, "Possession_Type"] == "Transition").sum()) for t in kpis["Team"]]
    kpis["Settled_Possessions"] = [int((work.loc[work["Team"] == t, "Possession_Type"] == "Settled_Offense").sum()) for t in kpis["Team"]]
    kpis["Other_Possessions"] = kpis["Possessions"] - kpis["Transition_Possessions"] - kpis["Settled_Possessions"]
    kpis["Transition_Share_All"] = kpis["Transition_Possessions"] / kpis["Possessions"]
    return kpis


def summarize_tournament(master):
    kpis = single.core_kpi_table(master, sorted(master["Team"].unique()))
    timings = single.time_summary(master)
    kpis = kpis.merge(timings, on=["Team", "Possessions"], how="left", validate="1:1")
    kpis["Matches_Coded"] = kpis["Team"].apply(
        lambda t: int(master.loc[master["Team"] == t, "Match_ID"].nunique()))
    kpis["Shot_Possessions"] = [int((master.loc[master["Team"] == t, "Shot"] > 0).sum()) for t in kpis["Team"]]
    kpis["Transition_Possessions"] = [int((master.loc[master["Team"] == t, "Possession_Type"] == "Transition").sum()) for t in kpis["Team"]]
    kpis["Settled_Possessions"] = [int((master.loc[master["Team"] == t, "Possession_Type"] == "Settled_Offense").sum()) for t in kpis["Team"]]
    kpis["Other_Possessions"] = kpis["Possessions"] - kpis["Transition_Possessions"] - kpis["Settled_Possessions"]
    kpis["Transition_Share_All"] = kpis["Transition_Possessions"] / kpis["Possessions"]
    return kpis


def analyze_tournament(directory: Path, documented: dict):
    files = sorted(p for p in directory.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise ValueError(f"No .xlsx found in {directory}")
    seen = set()
    matches, possessions, events, qc_tables, match_kpis = [], [], [], [], []
    exception_records = []
    for path in files:
        # A single incomplete/old template can halt the whole tournament import.
        # Include the *filename* so coaches can locate the offending Match_Info
        # without editing or guessing at their five validated match workbooks.
        try:
            match, work, infra, qc, meta = load_one(path, documented)
        except (ValueError, KeyError, TypeError, OSError) as exc:
            raise ValueError(
                f"Unable to read '{path.name}': {exc}. "
                "Check this file's Match_Info sheet and ensure Match_Records "
                "contains one completed workbook per match."
            ) from exc
        if meta.match_id in seen:
            raise ValueError(f"Duplicate Match_ID={meta.match_id}. Remove older or duplicate file from {directory}")
        seen.add(meta.match_id)
        matches.append(match)
        possessions.append(work)
        events.append(infra)
        qc_tables.append(qc)
        match_kpis.append(summarize_team_match(work, meta))
        for _, e in infra.loc[infra["Coding_Exception"] != ""].iterrows():
            exception_records.append({
                "Match_ID": meta.match_id, "Infraction_ID": int(e["Infraction_ID"]),
                "Source_Possession_Change": e["Original_Possession_Change"],
                "Analytical_Possession_Change": e["Analysis_Possession_Change"],
                "Scope": e["Coding_Exception"],
                "Reason": documented.get(meta.match_id, {}).get(str(int(e["Infraction_ID"])), {}).get("explanation", ""),
            })
    master = pd.concat(possessions, ignore_index=True)
    infra_master = pd.concat(events, ignore_index=True)
    qc = pd.concat(qc_tables, ignore_index=True)
    match_frame = pd.DataFrame(matches)
    team_match = pd.concat(match_kpis, ignore_index=True)
    if master["Possession_Key"].duplicated().any():
        raise ValueError("Duplicate Possession_Key in tournament master")
    if infra_master["Infraction_Key"].duplicated().any():
        raise ValueError("Duplicate Infraction_Key in tournament master")
    # All-match ratio denominators are pooled underlying possessions and shots.
    team_total = summarize_tournament(master)
    types = single.breakdown(master, "Possession_Type")
    reasons = single.end_reason_summary(master)
    return {
        "Match_Summary": match_frame, "Possession_Master": master,
        "Infraction_Master": infra_master, "Team_Match_KPI": team_match,
        "Team_Tournament_KPI": team_total, "Team_Type_Profile": types,
        "Team_End_Reasons": reasons, "QC_Report": qc,
        "Coding_Exceptions": pd.DataFrame(exception_records, columns=[
            "Match_ID", "Infraction_ID", "Source_Possession_Change",
            "Analytical_Possession_Change", "Scope", "Reason"]),
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input-dir", default="Match_Records", help="Folder containing V1.2 Excel files")
    ap.add_argument("--output", default="output", help="Output folder for analysis CSVs")
    ap.add_argument("--exceptions", default="coding_exceptions.json", help="Documented exception JSON")
    ap.add_argument("--focus-team", default="TPE", help="Team code for terminal summary")
    args = ap.parse_args()
    input_dir = Path(args.input_dir)
    documented = read_exceptions(Path(args.exceptions))
    data = analyze_tournament(input_dir, documented)
    output = Path(args.output)
    output.mkdir(parents=True, exist_ok=True)
    for key in EXPORTS:
        data[key].to_csv(output / f"{key}.csv", index=False, encoding="utf-8-sig")
    report = data["QC_Report"]
    errors = report.loc[report["Severity"] == "error"]
    print(f"Matches {len(data['Match_Summary'])}, Possessions {len(data['Possession_Master'])}, Infractions {len(data['Infraction_Master'])}")
    print(f"QC: {len(errors)} errors, {(report['Severity'] == 'warning').sum()} warnings, {(report['Severity'] == 'info').sum()} infos")
    print(data["Match_Summary"][["Match_ID", "Recorded_Possessions", "Score_Matches", "QC_Errors"]].to_string(index=False))
    selected = data["Team_Tournament_KPI"].loc[data["Team_Tournament_KPI"]["Team"] == args.focus_team]
    if not selected.empty:
        print(f"\n{args.focus_team} pooled metrics (fractions, not percentages):")
        print(selected[["Team", "Matches_Coded", "Possessions", "Goals", "Shots", "SOG",
                        "Offensive_Efficiency", "Shot_Generation_Rate", "Shot_Accuracy",
                        "Transition_Possessions", "Settled_Possessions"]].to_string(index=False))
    print(f"\nSaved CSVs: {output.resolve()}")
    if not errors.empty:
        print("NOTE: QC errors remain. Do not publish these results before review.", file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
