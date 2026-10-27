"""Sixes Lacrosse · Tournament Coach Dashboard v0.1.4.

Usage: streamlit run tournament_dashboard.py
Reads full match Excel files through the tested tournament_analysis.py engine.
Single Match Dashboard v0.3.2 is available through Streamlit's pages sidebar.
"""
from __future__ import annotations
from html import escape
from pathlib import Path
import tempfile

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from tournament_analysis import analyze_tournament, read_exceptions
from tournament_dashboard_data import (
    end_reasons, match_trends, slice_possessions, summary, type_profile,
)

ROOT = Path(__file__).resolve().parent
DATA_ROOT = ROOT / 'Match_Records'
EXCEPTIONS_PATH = ROOT / 'coding_exceptions.json'
QUARTERS = ['Q1','Q2','Q3','Q4']
TYPE_LABELS = {'Transition':'Transition','Settled_Offense':'Settled Offense','Other':'Other / Unclassified'}
COLORS = {'Transition':'#4279AE','Settled_Offense':'#D29B5C','Other':'#9BA3B0'}
TEAL = '#376A9A'
SECONDARY = '#CC9358'

METRICS = {
    'Offensive_Efficiency': ('Offensive Efficiency', 'Goals / Possessions', 'percent'),
    'Shot_Generation_Rate': ('Shot Generation', 'Shot possessions / Possessions', 'percent'),
    'Shot_Accuracy': ('Shot Accuracy', 'Shots on goal / All shots', 'percent'),
    'Shooting_Efficiency': ('Shooting Efficiency', 'Goals / All shots', 'percent'),
    'Passes_per_Possession': ('Passes per Possession', 'Total passes / Possessions', 'number'),
    'Mean_Possession_Duration_s': ('Possession Duration (mean)', 'Game-clock elapsed seconds, valid observations only', 'seconds'),
    'Mean_Time_to_First_Shot_s': ('Time to First Shot (mean)', 'Seconds to first shot, shot possessions with valid clock only', 'seconds'),
}

st.set_page_config(page_title='Sixes · Tournament Analysis', page_icon='🥍', layout='wide')
st.markdown('''<style>
.stTabs [data-baseweb="tab-list"] {gap:0.45rem; padding:0.25rem 0 0.6rem;}
.stTabs button[data-baseweb="tab"] {padding:0.65rem 1.0rem !important; border-radius:0.6rem 0.6rem 0.2rem 0.2rem; background:rgba(90,113,139,.08);}
.stTabs button[data-baseweb="tab"] p {font-size:1.03rem !important; font-weight:650 !important;}
.stTabs button[data-baseweb="tab"][aria-selected="true"] {background:rgba(70,127,179,.16); box-shadow:inset 0 -3px 0 #4279AE;}
.coach-note {padding:0.85rem 1rem; border-left:3px solid #4279AE; background:rgba(70,127,179,.07); border-radius:0.4rem; margin:0.5rem 0 1rem;}
</style>''', unsafe_allow_html=True)


@st.cache_data(show_spinner='Reading matches and validating data…')
def load_local(folder: str, file_signature: tuple, exceptions_signature: tuple):
    del file_signature, exceptions_signature  # cache invalidation keys
    return analyze_tournament(Path(folder), read_exceptions(EXCEPTIONS_PATH))


@st.cache_data(show_spinner='Reading uploaded matches and validating data…')
def load_upload(file_tuples: tuple, exceptions_signature: tuple):
    del exceptions_signature
    with tempfile.TemporaryDirectory(prefix='sixes_tournament_') as tmp:
        for name, contents in file_tuples:
            dest=Path(tmp) / Path(name).name
            dest.write_bytes(contents)
        return analyze_tournament(Path(tmp), read_exceptions(EXCEPTIONS_PATH))


def file_signature(folder: Path) -> tuple:
    return tuple((p.name,p.stat().st_size,p.stat().st_mtime_ns) for p in sorted(folder.glob('*.xlsx'))
                 if not p.name.startswith('~$'))


def fmt(value, kind='number', digits=1):
    if pd.isna(value):
        return '—'
    if kind == 'percent':
        return f'{value*100:.{digits}f}%'
    if kind == 'int':
        return str(int(value))
    return f'{float(value):.{digits}f}'


def conversion_chart(stages: list[tuple[str, int]], total: int, *, color: str = TEAL) -> go.Figure:
    """Counts first, percentages relative to first stage; never mix shot and possession units."""
    labels = [label for label, _ in stages][::-1]
    counts = [int(value) for _, value in stages][::-1]
    percent = [f"{value / total * 100:.1f}%" if total else "—" for value in counts]
    texts = [f"{value:,}  ({pct})" for value, pct in zip(counts, percent)]
    fig = go.Figure(go.Bar(
        y=labels, x=counts, orientation='h',
        marker=dict(color=[color, '#7597B4', '#B5C1CE'][:len(counts)]),
        text=texts, textposition='outside', cliponaxis=False,
        customdata=percent,
        hovertemplate='%{y}<br>Count: %{x}<br>Share of starting stage: %{customdata}<extra></extra>',
    ))
    fig.update_layout(
        height=250, margin=dict(l=8, r=8, t=12, b=20),
        xaxis=dict(range=[0, max(1, total) * 1.32], visible=False),
        yaxis=dict(title=None, automargin=True, tickfont=dict(size=13)),
        showlegend=False, template='plotly_white', bargap=0.34,
        font=dict(size=12), paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
    )
    return fig


def offense_mix_chart(profile: pd.DataFrame) -> go.Figure:
    """100% composition based on all selected team possessions, including Other."""
    fig = go.Figure()
    for _, r in profile.iterrows():
        tp = str(r['Possession_Type'])
        count = int(r['Possessions'])
        share = float(r['Share']) if pd.notna(r['Share']) else 0.
        fig.add_bar(
            y=['Possession type'], x=[share * 100], orientation='h',
            name=TYPE_LABELS.get(tp,tp), marker_color=COLORS.get(tp, '#9BA3B0'),
            text=[f"{count} · {share*100:.1f}%" if share >= 0.11 else ''],
            textposition='inside', insidetextanchor='middle',
            customdata=[[count,share*100]],
            hovertemplate='%{fullData.name}<br>%{customdata[0]} possessions<br>%{customdata[1]:.1f}%<extra></extra>',
        )
    fig.update_layout(
        barmode='stack', height=148, margin=dict(l=0,r=0,t=8,b=45),
        xaxis=dict(range=[0,100],showgrid=False,zeroline=False,ticksuffix='%',fixedrange=True),
        yaxis=dict(visible=False, fixedrange=True),
        legend=dict(orientation='h',y=-0.25,x=0, font=dict(size=12)),
        paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
        template='plotly_white',
    )
    return fig


def _outcome_metric(label: str, zh: str, rate, num: int, den: int, description: str):
    with st.container(border=True):
        st.caption(f'{label} · {zh}')
        st.metric(label, fmt(rate, 'percent'))
        st.caption(f'{num:,} / {den:,} · {description}')


def overview_story(data: dict, work: pd.DataFrame):
    """RQ-driven Overview: outcome conversion first, then offensive organization.

    All values come from the same active possession selection as Trends and Patterns.
    Both conversion charts have explicit units and never mix denominators.
    """
    n = int(data['Possessions'])
    shot_pos = int(data['Shot_Possessions'])
    goals = int(data['Goals'])
    shots = int(data['Shots'])
    sog = int(data['SOG'])
    matches = int(data['Matches'])

    st.subheader('Offensive Profile · 進攻表現總覽')
    st.caption(f'Selected sample · {matches} match(es)  |  {n:,} possessions  |  {goals:,} goals')
    st.markdown('### 01 · How do possessions become goals?')
    st.caption('RQ1 — 進攻成效：我們能否創造射門機會，並將球權轉換成得分？')

    pcol, scol = st.columns(2, gap='large')
    with pcol:
        st.markdown('**Possession conversion · 以球權為單位**')
        st.plotly_chart(conversion_chart([
            ('All possessions', n),
            ('With ≥1 shot', shot_pos),
            ('Scoring possessions', goals),
        ], n), use_container_width=True, config={'displaylogo':False})
        st.caption('A scoring possession counts once, even when multiple shots occurred.')
    with scol:
        st.markdown('**Shot conversion · 以射門次數為單位**')
        st.plotly_chart(conversion_chart([
            ('All shots', shots),
            ('Shots on goal', sog),
            ('Goals', goals),
        ], shots, color='#BD8955'), use_container_width=True,
            config={'displaylogo':False})
        st.caption('Shots and shots on goal are event counts, not possession counts.')

    col1,col2,col3 = st.columns(3, gap='small')
    with col1:
        _outcome_metric('Shot Generation','產生射門的球權比例',
                        data['Shot_Generation_Rate'], shot_pos, n,
                        'shot possessions / possessions')
    with col2:
        _outcome_metric('Shot Accuracy','射正率',
                        data['Shot_Accuracy'], sog, shots,
                        'shots on goal / shots')
    with col3:
        _outcome_metric('Offensive Efficiency','球權得分率',
                        data['Offensive_Efficiency'], goals, n,
                        'goals / possessions')

    st.markdown('### 02 · How is our offense organized?')
    st.caption('RQ2 — 進攻組織：主要使用哪些進攻型態？後續可在 Offensive Patterns 比較各型態效率。')
    profile = type_profile(work)
    st.plotly_chart(offense_mix_chart(profile), use_container_width=True,
                    config={'displaylogo':False})
    detail = '  ·  '.join(
        f"{TYPE_LABELS.get(str(r['Possession_Type']),str(r['Possession_Type']))}: "
        f"{int(r['Possessions'])}/{n} ({fmt(r['Share'],'percent')})"
        for _,r in profile.iterrows())
    st.caption(detail)
    with st.expander('Additional indicators · 進攻節奏與其他指標'):
        c1,c2,c3 = st.columns(3)
        c1.metric('Passes / Possession',fmt(data['Passes_per_Possession']))
        c2.metric('Mean Time to First Shot',
                  fmt(data['Mean_Time_to_First_Shot_s']) + ' s' if pd.notna(data['Mean_Time_to_First_Shot_s']) else '—')
        c3.metric('Mean Possession Duration',
                  fmt(data['Mean_Possession_Duration_s']) + ' s' if pd.notna(data['Mean_Possession_Duration_s']) else '—')
        st.caption(f"First-shot time: valid n={int(data['First_Shot_Valid_n'])}; "
                   f"possession duration: valid n={int(data['Duration_Valid_n'])}. "
                   'Invalid or unobservable times are excluded, not zero.')
        st.caption(f"Shooting Efficiency (Goals / Shots): {fmt(data['Shooting_Efficiency'],'percent')} "
                   f"({goals}/{shots}).")

    st.markdown('### 03 · What should coaches examine next?')
    st.caption('RQ3 — 跨場比較：比較不同對手、比賽節次及進攻型態，找出值得回看影片的球權。')
    st.markdown("""**Offensive Trends**：哪些場次的 Shot Generation、Shot Accuracy 和 Offensive Efficiency 不同？

**Offensive Patterns**：Transition 與 Settled Offense 各有多少次球權，效率及球權流失原因如何？""")
    with st.expander('Indicator definitions · 指標定義'):
        st.markdown("""
- **Shot Generation** = possessions with ≥1 shot / total possessions.
- **Shot Accuracy** = shots on goal / shots; this uses individual shot counts.
- **Offensive Efficiency** = goals / possessions; in this coding scheme, each possession has at most one goal.
- **Shooting Efficiency** = goals / total shots; this is different from Offensive Efficiency.
- **Possession Type** = Transition, Settled Offense, or Other, classified by the first offensive threat and the defensive organization.
- **Timing** uses only reliable observed timestamps; missing values do not count as zero.

Tournament rates are computed from pooled counts **after applying team, match and quarter filters**. These comparisons are descriptive; opponent and game contexts differ.
""")


def chart(fig, height=340):
    fig.update_layout(height=height, template='plotly_white', margin=dict(l=12,r=15,t=22,b=30),
                      font=dict(size=13), hovermode='closest',
                      paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig, use_container_width=True, config={'displaylogo':False})


def matches_table(trends: pd.DataFrame):
    if trends.empty:
        return
    r=[]
    for _,v in trends.iterrows():
        r.append({
            'Match':v['Match_Label'],'Poss':int(v['Possessions']),'Goals':int(v['Goals']),
            'Off. Eff.':fmt(v['Offensive_Efficiency'],'percent'),
            'Shot Gen.':fmt(v['Shot_Generation_Rate'],'percent'),
            'Shot Acc.':fmt(v['Shot_Accuracy'],'percent'),
            'Shoot Eff.':fmt(v['Shooting_Efficiency'],'percent'),
            'Pass/Poss':fmt(v['Passes_per_Possession'],digits=2),
            'Time to 1st Shot':fmt(v['Mean_Time_to_First_Shot_s'])+' s' if pd.notna(v['Mean_Time_to_First_Shot_s']) else '—',
            'Time valid n':int(v['First_Shot_Valid_n']),
        })
    st.dataframe(pd.DataFrame(r), use_container_width=True, hide_index=True)


def show_trend_plot(trends: pd.DataFrame, metric: str, overall: dict):
    title,definition,unit=METRICS[metric]
    vals=trends[metric] * (100 if unit=='percent' else 1)
    pooled=overall[metric] * (100 if unit=='percent' else 1)
    tip_n=[]
    for _,r in trends.iterrows():
        if metric=='Offensive_Efficiency': tip_n.append(f"{int(r['Goals'])} goals / {int(r['Possessions'])} poss")
        elif metric=='Shot_Generation_Rate': tip_n.append(f"{int(r['Shot_Possessions'])} shot poss / {int(r['Possessions'])} poss")
        elif metric=='Shot_Accuracy': tip_n.append(f"{int(r['SOG'])} SOG / {int(r['Shots'])} shots")
        elif metric=='Shooting_Efficiency': tip_n.append(f"{int(r['Goals'])} goals / {int(r['Shots'])} shots")
        elif metric=='Mean_Time_to_First_Shot_s': tip_n.append(f"valid n = {int(r['First_Shot_Valid_n'])}")
        elif metric=='Mean_Possession_Duration_s': tip_n.append(f"valid n = {int(r['Duration_Valid_n'])}")
        else: tip_n.append(f"{int(r['Passes'])} passes / {int(r['Possessions'])} poss")
    fig=go.Figure()
    fig.add_scatter(x=trends['Match_Label'],y=vals,mode='lines+markers+text',
                    name=title,line=dict(color=TEAL,width=3),marker=dict(size=11),
                    text=[fmt(x,unit) for x in trends[metric]],textposition='top center',
                    customdata=[[x] for x in tip_n],
                    hovertemplate='%{x}<br>%{y:.1f}'+('%' if unit=='percent' else (' s' if unit=='seconds' else ''))+'<br>%{customdata[0]}<extra></extra>')
    if pd.notna(pooled):
        fig.add_hline(y=pooled,line_dash='dot',line_color=SECONDARY,
                      annotation_text='Pooled: '+fmt(overall[metric],unit),annotation_position='bottom right')
    fig.update_layout(yaxis_title=('%' if unit=='percent' else ('Seconds' if unit=='seconds' else 'Passes')),
                      xaxis_title=None,showlegend=False)
    if unit=='percent':
        fig.update_yaxes(range=[0,max(65,min(105,float(vals.max())*1.22))] if len(vals.dropna()) else [0,100])
    chart(fig,405)
    st.caption(f'Definition: {definition}. Dotted reference line = pooled result across the selected possessions. Chronological match order does not establish causality.')


def plot_type_distribution(profiles):
    fig=go.Figure()
    for _,r in profiles.iterrows():
        tp=r['Possession_Type']
        fig.add_bar(x=[TYPE_LABELS[tp]],y=[r['Possessions']],name=TYPE_LABELS[tp],
                    marker_color=COLORS[tp],text=[f"{int(r['Possessions'])} ({fmt(r['Share'],'percent')})"],
                    textposition='outside',cliponaxis=False,
                    hovertemplate=f"{TYPE_LABELS[tp]}<br>Possessions: %{{y}}<br>Share: {fmt(r['Share'],'percent')}<extra></extra>")
    fig.update_layout(showlegend=False,yaxis_title='Possessions',xaxis_title=None)
    chart(fig,325)


def plot_type_efficiency(profiles):
    fig=go.Figure()
    for _,r in profiles.iterrows():
        tp=r['Possession_Type']; rate=r['Offensive_Efficiency']
        if pd.notna(rate):
            fig.add_bar(x=[TYPE_LABELS[tp]],y=[rate*100],name=TYPE_LABELS[tp],
                        marker_color=COLORS[tp],text=[fmt(rate,'percent')],textposition='outside',
                        customdata=[[int(r['Goals']),int(r['Possessions'])]],
                        hovertemplate='%{x}<br>Goals/Poss: %{customdata[0]} / %{customdata[1]}<br>Efficiency: %{y:.1f}%<extra></extra>')
    fig.update_layout(showlegend=False,yaxis_title='Goals / Possessions (%)',xaxis_title=None,yaxis_range=[0,105])
    chart(fig,325)


def plot_end_reasons(df):
    ends=end_reasons(df).sort_values('Count',ascending=True)
    if ends.empty:
        st.info('No end reasons available for these filters.')
        return
    fig=go.Figure(go.Bar(x=ends['Count'], y=ends['End_Reason'],orientation='h',
                         marker_color=TEAL,text=ends.apply(lambda r:f"{int(r['Count'])} · {fmt(r['Share'],'percent')}",axis=1),
                         textposition='outside',cliponaxis=False,
                         customdata=ends[['Share']].to_numpy(),
                         hovertemplate='%{y}<br>Count: %{x}<br>Share: %{customdata[0]:.1%}<extra></extra>'))
    fig.update_layout(xaxis_title='Possessions',yaxis_title=None,xaxis_range=[0,max(3,float(ends['Count'].max())*1.36)])
    chart(fig,max(300,42*len(ends)+90))


def type_table(profiles):
    rows=[]
    for _,r in profiles.iterrows():
        rows.append({'Type':TYPE_LABELS[r['Possession_Type']],
                     'Possessions':int(r['Possessions']),'Share':fmt(r['Share'],'percent'),
                     'Goals':int(r['Goals']),'Off. Eff.':fmt(r['Offensive_Efficiency'],'percent'),
                     'Shot Gen.':fmt(r['Shot_Generation_Rate'],'percent'),
                     'Pass/Poss':fmt(r['Passes_per_Possession'],digits=2),
                     'Time to 1st Shot':fmt(r['Mean_Time_to_First_Shot_s'])+' s' if pd.notna(r['Mean_Time_to_First_Shot_s']) else '—',
                     'Valid n':int(r['First_Shot_Valid_n'])})
    st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)


# Sidebar: sources and global team/match/quarter filtering.
with st.sidebar:
    st.markdown('### Data source')
    source=st.radio('Load match files', ['Local Match_Records','Upload Excel'],label_visibility='collapsed')
    exceptions_sig=(EXCEPTIONS_PATH.stat().st_size,EXCEPTIONS_PATH.stat().st_mtime_ns) if EXCEPTIONS_PATH.exists() else (0,0)
    package=None
    try:
        if source=='Local Match_Records':
            sig=file_signature(DATA_ROOT)
            st.caption(f'{len(sig)} Excel file(s) in Match_Records/')
            with st.expander('Show Excel files being imported'):
                for name, _, _ in sig:
                    st.code(name, language=None)
            if sig:
                package=load_local(str(DATA_ROOT),sig,exceptions_sig)
            else:
                st.info('Put your V1.2 Excel files inside Match_Records/, then refresh.')
        else:
            uploads=st.file_uploader('Upload V1.2 match workbooks', type=['xlsx'],accept_multiple_files=True)
            if uploads:
                if len({u.name for u in uploads}) != len(uploads):
                    st.error('Repeated filenames: use one file per match.')
                else:
                    payload=tuple(sorted(((u.name,u.getvalue()) for u in uploads),key=lambda x:x[0]))
                    package=load_upload(payload,exceptions_sig)
    except Exception as e:
        st.error(f'Unable to combine matches: {e}')
        st.info('Review the Excel filenames above. Move unfinished templates, old copies or other .xlsx files out of Match_Records; do not delete or change valid match records.')

if package is None:
    st.title('🥍 Sixes Tournament Analysis')
    st.info('Add match Excel files to Match_Records/ or upload them from the sidebar to get started.')
    st.stop()

qc=package['QC_Report']; errs=qc.loc[qc['Severity']=='error']; warns=qc.loc[qc['Severity']=='warning']
match_meta=package['Match_Summary'].sort_values('Match_ID')
master=package['Possession_Master']

st.title('🥍 Sixes Tournament Analysis')
st.caption('Coach Dashboard v0.1.4 · Possession-first performance analysis · Coding Template V1.2')

if not errs.empty or (match_meta['Score_Matches'] == False).any():
    st.error(f"Data QC has {len(errs)} error(s) or score mismatches. Fix the source files before publishing these results.")
    st.dataframe(qc.loc[qc['Severity']=='error'],hide_index=True,use_container_width=True)
    st.stop()

with st.sidebar:
    st.markdown('### Analysis filters')
    teams=sorted(master['Team'].dropna().unique().tolist())
    team=st.selectbox('Team', teams, index=teams.index('TPE') if 'TPE' in teams else 0)
    available=sorted(master.loc[master['Team']==team,'Match_ID'].unique().tolist())
    name_map={mid: f"{mid[:2]}/{mid[2:4]} vs {master.loc[(master['Match_ID']==mid)&(master['Team']==team),'Opponent'].iloc[0]}" for mid in available}
    selected=st.multiselect('Matches',available,default=available,format_func=lambda mid:name_map[mid])
    quarters=st.multiselect('Quarters',QUARTERS,default=QUARTERS)
    st.caption('Filters apply to all three tabs. Selection recalculates numerator and denominator from possessions.')
    st.divider()
    st.caption(f'Files: {len(match_meta)} · Possessions: {len(master)} · QC: 0 errors / {len(warns)} warnings')

if not selected or not quarters:
    st.warning('Select at least one match and one quarter.')
    st.stop()

work=slice_possessions(master, team, selected, quarters)
if work.empty:
    st.warning('No possessions match the selected team, matches, and quarters.')
    st.stop()

overall=summary(work); overall['Matches']=work['Match_ID'].nunique()
trends=match_trends(work,match_meta)
not_full_tpe=team !='TPE'
if not_full_tpe:
    st.info(f"Showing {team}'s games available in the loaded files ({overall['Matches']} match(es)); this is not necessarily {team}'s full tournament schedule.")

st.markdown(f'<div class="coach-note"><b>{escape(team)} · {overall["Matches"]} match(es) · {len(work)} possessions</b> — All displayed metrics are recalculated from your current match and quarter filters.</div>',unsafe_allow_html=True)

one,two,three=st.tabs(['Tournament Overview','Offensive Trends','Offensive Patterns'])
with one:
    overview_story(overall, work)
    st.divider()
    st.markdown('#### Match-by-match indicators')
    matches_table(trends)
    with st.expander(f'Data quality · 0 errors / {len(warns)} warnings',expanded=False):
        st.caption('Warnings indicate potential ambiguity or documented game situations—not necessarily mistakes. Review before publishing externally.')
        st.dataframe(match_meta[['Match_ID','Recorded_Possessions','Score_Matches','QC_Errors','QC_Warnings']],hide_index=True,use_container_width=True)
        if not warns.empty:
            relevant=warns.loc[warns['Match_ID'].isin(selected)]
            st.dataframe(relevant[['Match_ID','Severity','Issue','Possession_ID','Infraction_ID']],hide_index=True,use_container_width=True)

with two:
    st.subheader('Offensive Trends')
    st.caption('Chronological match comparisons. Opponent strength and match context can explain part of the differences; do not interpret the line as a causal training trend.')
    chosen=st.selectbox('Performance indicator',list(METRICS),format_func=lambda k:METRICS[k][0])
    show_trend_plot(trends,chosen,overall)
    st.markdown('#### All matches · KPI table')
    matches_table(trends)
    st.caption('Time measures are game-clock elapsed seconds, not shot-clock values. Unobservable and unverified timestamps are excluded, not counted as zero.')

with three:
    st.subheader('Offensive Patterns')
    st.caption('Transition vs Settled Offense is defined by the first offensive threat and defensive organization, not simply by a fixed number of seconds.')
    profiles=type_profile(work)
    left,right=st.columns(2)
    with left:
        st.markdown('#### Possession Type · Distribution')
        plot_type_distribution(profiles)
    with right:
        st.markdown('#### Possession Type · Offensive Efficiency')
        plot_type_efficiency(profiles)
    type_table(profiles)
    st.divider()
    st.markdown('#### Explore possessions by type')
    chosen_type=st.selectbox('Possession Type',['All','Transition','Settled_Offense','Other'],format_func=lambda t:'All Possessions' if t=='All' else TYPE_LABELS[t])
    sub=work if chosen_type=='All' else work.loc[work['Possession_Type']==chosen_type]
    if sub.empty:
        st.info('No possessions of this type in the current selection.')
    else:
        type_s=summary(sub)
        a,b,c,d=st.columns(4)
        a.metric('Possessions',type_s['Possessions'])
        b.metric('Offensive Efficiency',fmt(type_s['Offensive_Efficiency'],'percent'))
        c.metric('Shot Generation',fmt(type_s['Shot_Generation_Rate'],'percent'))
        d.metric('Time to First Shot',fmt(type_s['Mean_Time_to_First_Shot_s'])+' s' if pd.notna(type_s['Mean_Time_to_First_Shot_s']) else '—')
        st.caption(f"Valid timing observations: First Shot n={type_s['First_Shot_Valid_n']}; Duration n={type_s['Duration_Valid_n']}. Excludes missing clocks and no-shot possessions for Time to First Shot.")
        left,right=st.columns([2,1])
        with left:
            st.markdown('##### Possession End Reasons')
            plot_end_reasons(sub)
        with right:
            st.markdown('##### Possession Start Types')
            starts=sub['Start_Type'].fillna('Unknown').value_counts()
            st.dataframe(pd.DataFrame({'Start Type':starts.index,'Possessions':starts.values}),hide_index=True,use_container_width=True)
            st.markdown('##### Timing Summary')
            st.dataframe(pd.DataFrame([
                {'Measure':'Possession Duration','Valid n':type_s['Duration_Valid_n'], 'Mean (s)':fmt(type_s['Mean_Possession_Duration_s']),'Median (s)':fmt(type_s['Median_Possession_Duration_s'])},
                {'Measure':'Time to First Shot','Valid n':type_s['First_Shot_Valid_n'], 'Mean (s)':fmt(type_s['Mean_Time_to_First_Shot_s']),'Median (s)':fmt(type_s['Median_Time_to_First_Shot_s'])},
            ]),hide_index=True,use_container_width=True)
        export_cols=['Match_ID','Possession_ID','Quarter','Game_Clock_Display','Team','Opponent','Start_Type','Start_Zone','Possession_Type','Pass','Shot','Shot_on_Goal','Goal','End_Reason','Possession_Duration_s','Time_to_First_Shot_s']
        st.download_button('Download filtered possession list (CSV)',
                           sub[export_cols].to_csv(index=False).encode('utf-8-sig'),
                           file_name=f'{team}_selected_possessions.csv',mime='text/csv')

st.divider()
st.caption('Data: analyst-coded match videos · 2026 APAC Sixes · Results are descriptive, not causal. For full detail on one match, open “Single Match” in the navigation sidebar.')
