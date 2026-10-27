"""Pure data helpers for the Sixes Tournament Coach Dashboard.

All denominators follow analysis.py v0.3.2. Never average match percentages
when computing a tournament's overall efficiency. This file has no UI imports.
"""
from __future__ import annotations
import pandas as pd

POSITIONS = ("Transition", "Settled_Offense", "Other")


def fraction(num: float, den: float) -> float:
    return float(num) / float(den) if den else float('nan')


def slice_possessions(master: pd.DataFrame, team: str,
                      matches: list[str], quarters: list[str]) -> pd.DataFrame:
    return master.loc[
        (master['Team'] == team)
        & master['Match_ID'].isin(matches)
        & master['Quarter'].isin(quarters)
    ].copy()


def summary(df: pd.DataFrame) -> dict:
    """Aggregate counts first, then calculate rates. Durations use valid times."""
    n = len(df)
    goals = int(df['Goal'].fillna(0).sum())
    shots = int(df['Shot'].fillna(0).sum())
    sog = int(df['Shot_on_Goal'].fillna(0).sum())
    passes = int(df['Pass'].fillna(0).sum())
    shot_pos = int((df['Shot'].fillna(0) > 0).sum())
    trans = int((df['Possession_Type'] == 'Transition').sum())
    settled = int((df['Possession_Type'] == 'Settled_Offense').sum())
    duration = pd.to_numeric(df['Possession_Duration_s'], errors='coerce').dropna()
    first_shot = pd.to_numeric(df['Time_to_First_Shot_s'], errors='coerce').dropna()
    return {
        'Possessions': n, 'Goals': goals, 'Shots': shots, 'SOG': sog,
        'Passes': passes, 'Shot_Possessions': shot_pos,
        'Transition_Possessions': trans, 'Settled_Possessions': settled,
        'Other_Possessions': n-trans-settled,
        'Offensive_Efficiency': fraction(goals,n),
        'Shot_Generation_Rate': fraction(shot_pos,n),
        'Shot_Accuracy': fraction(sog,shots),
        'Shooting_Efficiency': fraction(goals,shots),
        'Passes_per_Possession': fraction(passes,n),
        'Transition_Share_All': fraction(trans,n),
        'Settled_Share_All': fraction(settled,n),
        'Mean_Possession_Duration_s': float(duration.mean()) if len(duration) else float('nan'),
        'Median_Possession_Duration_s': float(duration.median()) if len(duration) else float('nan'),
        'Duration_Valid_n': len(duration),
        'Mean_Time_to_First_Shot_s': float(first_shot.mean()) if len(first_shot) else float('nan'),
        'Median_Time_to_First_Shot_s': float(first_shot.median()) if len(first_shot) else float('nan'),
        'First_Shot_Valid_n': len(first_shot),
    }


def match_trends(df: pd.DataFrame, match_info: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for match_id, sub in df.groupby('Match_ID', sort=False):
        meta = match_info.loc[match_info['Match_ID'] == match_id]
        if meta.empty:
            raise ValueError(f'Missing Match_Info for {match_id}')
        meta = meta.iloc[0]
        team = str(sub['Team'].iloc[0])
        opponent = str(meta['Away']) if str(meta['Home']) == team else str(meta['Home'])
        match_label = f"{meta['Match_Date_Label']} vs {opponent}"
        rows.append({'Match_ID': match_id, 'Opponent': opponent,
                     'Match_Label': match_label, **summary(sub)})
    return pd.DataFrame(rows).sort_values('Match_ID').reset_index(drop=True) if rows else pd.DataFrame()


def type_profile(df: pd.DataFrame) -> pd.DataFrame:
    total = len(df)
    rows=[]
    for tp in POSITIONS:
        sub=df.loc[df['Possession_Type']==tp]
        r=summary(sub)
        rows.append({'Possession_Type': tp, **r, 'Share': fraction(len(sub),total)})
    return pd.DataFrame(rows)


def end_reasons(df: pd.DataFrame) -> pd.DataFrame:
    out=df['End_Reason'].fillna('Unknown').value_counts().rename_axis('End_Reason').reset_index(name='Count')
    out['Share']=out['Count'].apply(lambda x:fraction(x,len(df)))
    return out


def time_validity(df: pd.DataFrame) -> tuple[int,int]:
    return (int(pd.to_numeric(df['Possession_Duration_s'],errors='coerce').notna().sum()),
            int(pd.to_numeric(df['Time_to_First_Shot_s'],errors='coerce').notna().sum()))
