"""
BHSQA KPI Dashboard – Streamlit app
Upload the hospital KPI reporting workbooks (July–June fiscal-year template),
see achievement against national BHSQA targets, drill down to the exact
source cells, and download Excel / PDF / PowerPoint reports.
"""
import dataclasses
import datetime as dt
import glob
import hashlib
import os

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import exporters as X
import kpi_engine as E

st.set_page_config(page_title='BHSQA KPI Dashboard', page_icon='🏥', layout='wide', initial_sidebar_state='expanded')

STATUS_COLORS = X.STATUS_COLORS
STATUS_ICON = {'Met': '🟢', 'Not met': '🔴', 'No data': '⚪', 'Not rated': '🟠'}
NAVY = X.NAVY

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
.kcard {border-radius: 10px; padding: 14px 16px; background: #fff; border: 1px solid #E3E8EF;
        border-top: 5px solid var(--c); box-shadow: 0 1px 3px rgba(16,24,40,.06); height: 100%;}
.kcard .v {font-size: 2.1rem; font-weight: 700; color: var(--c); line-height: 1.1;}
.kcard .l {font-size: .85rem; color: #5B6575;}
.htile {border-radius: 10px; padding: 10px 12px; background: #fff; border: 1px solid #E3E8EF;
        border-left: 6px solid var(--c); margin-bottom: 6px;}
.htile .n {font-size: .78rem; color: #5B6575; height: 2.2em; overflow: hidden;}
.htile .v {font-size: 1.45rem; font-weight: 700; color: #1F2937;}
.htile .t {font-size: .75rem; color: var(--c); font-weight: 600;}
.hero {background: linear-gradient(90deg, #1F3A5F, #2F5D8C); color: white; padding: 18px 24px; border-radius: 12px; margin-bottom: 14px;}
.hero h1 {color: white; font-size: 1.7rem; margin: 0;}
.hero p {color: #DCE6F2; margin: 2px 0 0 0;}
.ntile {border-radius: 8px; padding: 6px 8px; background: #fff; border: 1px solid #E3E8EF; border-top: 5px solid var(--c);
         margin-bottom: 6px; height: 108px; overflow: hidden;}
.ntile .c {font-weight: 700; font-size: .8rem; color: #1F3A5F;}
.ntile .d {font-size: .7rem; color: #5B6575; line-height: 1.15; height: 2.35em; overflow: hidden;}
.ntile .v {font-size: 1.0rem; font-weight: 700; color: #1F2937; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;}
.ntile .t {font-size: .68rem; color: var(--c); font-weight: 600;}
.badge {display:inline-block; padding: 3px 10px; border-radius: 12px; color: white; font-weight: 600; font-size: .85rem;}
</style>
""", unsafe_allow_html=True)


# --------------------------------------------------------------------------- loading
@st.cache_resource(show_spinner='Reading workbook…')
def _load(name: str, data: bytes):
    return E.load_book(name, data)


@st.cache_resource(show_spinner='Extracting BHSQA KPIs…')
def _extract(keys: tuple, start: str, end: str):
    books = [_load(n, d) for n, d in st.session_state['_files'] if hashlib.md5(d).hexdigest() in keys]
    return E.extract(books, start, end)


def get_files():
    up = st.sidebar.file_uploader('Upload KPI reporting workbooks (.xlsx)', type=['xlsx'], accept_multiple_files=True,
                                  help='One workbook per fiscal year (July–June), e.g. the 2025-26 and 2026-27 QAS sheets.')
    files = [(f.name, f.getvalue()) for f in up] if up else []
    if not files:  # optional: files committed in ./data for presentation mode
        for p in sorted(glob.glob(os.path.join(os.path.dirname(__file__), 'data', '*.xlsx'))):
            files.append((os.path.basename(p), open(p, 'rb').read()))
        if files:
            st.sidebar.caption(f'Using {len(files)} workbook(s) from the app\'s data folder.')
    return files


files = get_files()
if not files:
    st.markdown('<div class="hero"><h1>🏥 BHSQA KPI Dashboard</h1>'
                '<p>Upload the hospital KPI reporting workbooks in the sidebar to begin.</p></div>', unsafe_allow_html=True)
    st.info('Upload one workbook per fiscal year (July–June). The app detects the year, extracts only KPIs aligned with the '
            'national BHSQA KPI User Guide, compares them with the national targets and keeps a link to every source cell.')
    st.stop()

st.session_state['_files'] = files
try:
    books = [_load(n, d) for n, d in files]
except Exception as e:  # noqa
    st.error(f'Could not read a workbook: {e}')
    st.stop()

all_months = E.available_months(books)
# default period: from Jan of the latest calendar year in the files to the last month with data
probe = _extract(tuple(hashlib.md5(d).hexdigest() for _, d in files), all_months[0], all_months[-1])
with_data = sorted(probe.monthly.dropna(subset=['result']).month.unique())
last_data = with_data[-1] if with_data else all_months[-1]
default_start = max(all_months[0], f"{last_data[:4]}-01")

# --------------------------------------------------------------------------- sidebar
sb = st.sidebar
sb.markdown('### Reporting period')
labels = {k: E.month_label(k) for k in all_months}
start, end = sb.select_slider('Months', options=all_months, value=(default_start, last_data), format_func=lambda k: labels[k],
                              label_visibility='collapsed')
res = _extract(tuple(hashlib.md5(d).hexdigest() for _, d in files), start, end)
hosp = sb.text_input('Hospital name', value=res.hospital or 'Hospital')
res = dataclasses.replace(res, hospital=hosp)

sb.markdown('### Filters')
units_all = list(dict.fromkeys(res.kpis.unit))
units = sb.multiselect('Units', units_all, default=units_all, placeholder='All units')
groups = sb.multiselect('BHSQA group', ['A. Clinical', 'B. Managerial'], default=['A. Clinical', 'B. Managerial'])
incl_survey = sb.toggle('Include survey-based KPIs (satisfaction, 5S, waiting time)', value=True)

grp = {c[0]: c[4] for c in E.CAT}
K = res.kpis.copy()
K['group'] = K.bhsqa.map(grp)
K = K[K.unit.isin(units or units_all) & K.group.isin(groups or ['A. Clinical', 'B. Managerial'])]
if not incl_survey:
    K = K[K.typ != 'survey']

sb.caption('Files: ' + ', '.join(f'{b.name} ({b.fy_label})' for b in books))

# --------------------------------------------------------------------------- helpers
def card(col, value, label, color):
    col.markdown(f'<div class="kcard" style="--c:{color}"><div class="v">{value}</div><div class="l">{label}</div></div>',
                 unsafe_allow_html=True)


def fmt(v, uom=''):
    return X.fmt(v, uom)


def kpi_table(df, key, height=420):
    """selectable KPI table; returns selected sid or None"""
    show = df.assign(Status=df.status.map(lambda s: f'{STATUS_ICON.get(s, "")} {s}'),
                     Achievement=df.apply(lambda k: fmt(k.achievement, k.uom), axis=1))
    show = show[['sid', 'unit', 'bhsqa', 'bhsqa_name', 'kpi', 'Achievement', 'target', 'Status', 'months_with_data', 'uom']]
    show.columns = ['ID', 'Unit', 'BHSQA no.', 'BHSQA KPI (national description)', 'Unit KPI (as in sheet)', 'Value',
                    'BHSQA target', 'Status', 'Months', 'Unit of measure']
    ev = st.dataframe(show, key=key, hide_index=True, use_container_width=True, height=height,
                      on_select='rerun', selection_mode='single-row',
                      column_config={'ID': st.column_config.TextColumn(width='small'),
                                     'BHSQA KPI (national description)': st.column_config.TextColumn(width='medium'),
                                     'Unit KPI (as in sheet)': st.column_config.TextColumn(width='medium'),
                                     'Months': st.column_config.NumberColumn(help='Months with both numerator and denominator')})
    rows = ev.selection.rows if ev and ev.selection else []
    return show.iloc[rows[0]]['ID'] if rows else None


def trend_fig(k, m):
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=m.label, y=m.result, mode='lines+markers+text', line=dict(color=NAVY, width=3),
                             text=[fmt(v) if pd.notna(v) else '' for v in m.result], textposition='top center',
                             name='Monthly result', connectgaps=False))
    if k.dir in ('<', '<=') and pd.notna(k.hi):
        fig.add_hline(y=k.hi, line_dash='dash', line_color='#C62828', annotation_text=f'Target {k.target}')
    elif k.dir in ('>', '>=') and pd.notna(k.lo):
        fig.add_hline(y=k.lo, line_dash='dash', line_color='#C62828', annotation_text=f'Target {k.target}')
    elif k.dir == 'range':
        fig.add_hrect(y0=k.lo, y1=k.hi, fillcolor='#2E7D32', opacity=0.12, line_width=0, annotation_text=f'Target {k.target}')
    if pd.notna(k.achievement):
        fig.add_hline(y=k.achievement, line_dash='dot', line_color='#666', annotation_text='Period achievement',
                      annotation_position='bottom right')
    fig.update_layout(height=330, margin=dict(l=10, r=10, t=30, b=10), yaxis_title=k.uom, showlegend=False,
                      plot_bgcolor='white')
    fig.update_yaxes(gridcolor='#EEE')
    return fig


def highlight(frame, cells, rows):
    def sty(df):
        css = pd.DataFrame('', index=df.index, columns=df.columns)
        for r in df.index:
            if r in rows:
                css.loc[r, :] = 'background-color:#FFF8E1'
        for (r, c) in cells:
            col = E.get_column_letter(c)
            if r in css.index and col in css.columns:
                css.loc[r, col] = 'background-color:#FFD54F; font-weight:700; border:2px solid #E65100'
        return css
    return frame.style.apply(sty, axis=None)


def detail(sid, ctx):
    """drill-down: KPI → monthly values → original source cells"""
    k = res.kpis.set_index('sid').loc[sid]
    m = res.monthly[res.monthly.sid == sid]
    c = STATUS_COLORS.get(k.status, '#888')
    st.markdown(f"#### {k.unit} › {k.kpi}")
    st.markdown(f'<span class="badge" style="background:{c}">{k.status}</span> &nbsp; **{k.bhsqa}** – {k.bhsqa_name} '
                f'&nbsp;·&nbsp; hospital indicator {k.jcode}', unsafe_allow_html=True)
    a, b, cc, d = st.columns(4)
    a.metric('Period achievement', fmt(k.achievement, k.uom))
    b.metric('BHSQA target', k.target)
    cc.metric('Gap to target', '—' if pd.isna(k.gap) else f'{k.gap:+.2f}')
    d.metric('Months with data', f'{k.months_with_data} / {len(res.months)}')
    if k.note:
        st.caption(f'ℹ️ {k.note}')
    t1, t2, t3 = st.tabs(['📈 Trend', '🔢 Monthly numbers', '📂 Source data (mother sheet)'])
    with t1:
        st.plotly_chart(trend_fig(k, m), use_container_width=True)
    with t2:
        mt = m[['label', 'num', 'den', 'result']].rename(columns={'label': 'Month', 'num': 'Numerator / value',
                                                                  'den': 'Denominator', 'result': f'Result ({k.uom})'})
        if k.typ in ('survey', 'count'):
            mt = mt.drop(columns=['Denominator'])
        st.dataframe(mt, hide_index=True, use_container_width=True)
        if k.typ in ('ratio', 'comp', 'bor'):
            st.caption(f'Period: total numerator {fmt(k.total_num)} ÷ total denominator {fmt(k.total_den)} '
                       f'× {k.mult:g} = **{fmt(k.achievement, k.uom)}** (months with both values only)')
    with t3:
        tr = E.source_trace(res, sid)
        if tr.empty:
            st.warning('This KPI row was not found in the uploaded workbooks for the selected period.')
            return
        st.markdown('Every number above comes from these cells in the uploaded workbooks:')
        st.dataframe(tr, hide_index=True, use_container_width=True, height=min(38 * (len(tr) + 1), 360))
        st.download_button('⬇️ Download source trace (CSV)', tr.to_csv(index=False).encode(), f'source_{sid}.csv',
                           'text/csv', key=f'csv_{ctx}_{sid}')
        st.markdown('**Original sheets** – extracted cells are outlined in orange, their rows shaded:')
        srcs = E.source_rows(res, sid)
        tabs = st.tabs([f'{f[:40]} › {s}' for f, s in srcs])
        for tb, ((f, s), info) in zip(tabs, srcs.items()):
            with tb:
                fr = E.sheet_frame(res.books[f], s, info['rows'])
                st.dataframe(highlight(fr, info['cells'], info['rows']), use_container_width=True,
                             height=min(36 * (len(fr) + 1), 420))
                if st.toggle('Show the whole sheet', key=f'full_{ctx}_{sid}_{f}_{s}'):
                    full = E.sheet_frame(res.books[f], s)
                    st.dataframe(highlight(full, info['cells'], info['rows']), use_container_width=True, height=600)


# --------------------------------------------------------------------------- header
s = X.summary(res, K)
st.markdown(f'<div class="hero"><h1>🏥 {res.hospital} – Key Performance Indicators</h1>'
            f'<p>Achievement against national BHSQA targets &nbsp;•&nbsp; {E.month_label(start)} – {E.month_label(end)}'
            f' &nbsp;•&nbsp; {len(set(K.unit))} units &nbsp;•&nbsp; {len(K)} unit KPIs</p></div>', unsafe_allow_html=True)

tabs = st.tabs(['📊 Executive overview', '🏥 Units', '🔎 KPI explorer', '🧭 BHSQA KPIs', '😊 Surveys',
                '⚠️ Data quality', '⬇️ Reports'])

# ---- 1 overview
with tabs[0]:
    cs = st.columns(5)
    card(cs[0], s['total'], 'BHSQA unit KPIs tracked', NAVY)
    card(cs[1], s['met'], 'Met target', STATUS_COLORS['Met'])
    card(cs[2], s['notmet'], 'Not met', STATUS_COLORS['Not met'])
    card(cs[3], s['nodata'] + s['notrated'], 'No data / not rated', STATUS_COLORS['No data'])
    card(cs[4], f"{s['pct']:.0f}%" if s['pct'] is not None else '—', 'of rated KPIs met', NAVY)
    st.write('')
    left, right = st.columns([1, 2])
    with left:
        vc = K.status.value_counts().reindex(X.STATUS_ORDER).fillna(0)
        fig = go.Figure(go.Pie(labels=vc.index, values=vc.values, hole=.62, sort=False,
                               marker=dict(colors=[STATUS_COLORS[x] for x in vc.index])))
        fig.update_layout(height=320, margin=dict(l=0, r=0, t=30, b=0), title='Status of unit KPIs',
                          annotations=[dict(text=f"<b>{s['pct']:.0f}%</b><br>met" if s['pct'] is not None else '', showarrow=False,
                                            font=dict(size=22, color=NAVY))], legend=dict(orientation='h', y=-0.1))
        st.plotly_chart(fig, use_container_width=True)
    with right:
        st.markdown('##### Headline hospital indicators')
        heads = []
        for u, code in E.HEADLINE:
            r = res.kpis[(res.kpis.unit == u) & (res.kpis.bhsqa == code)]
            if not r.empty:
                heads.append(r.iloc[0])
        hc = st.columns(4)
        for i, k in enumerate(heads):
            c = STATUS_COLORS.get(k.status, '#888')
            hc[i % 4].markdown(f'<div class="htile" style="--c:{c}"><div class="n">{k.bhsqa} · {k.kpi}</div>'
                               f'<div class="v">{fmt(k.achievement, k.uom)}</div>'
                               f'<div class="t">{k.status} · target {k.target}</div></div>', unsafe_allow_html=True)
    st.markdown('##### National BHSQA KPI scorecard – hospital value of each national KPI')
    st.caption('Colour = status against the BHSQA target. Full list with unit-wise values in the 🧭 BHSQA KPIs tab.')
    NV = E.bhsqa_values(K)
    tile_col = dict(STATUS_COLORS, **{'Not tracked': '#B0B7C3'})
    for i in range(0, len(NV), 8):
        cols = st.columns(8)
        for c, r in zip(cols, NV.iloc[i:i + 8].itertuples()):
            col = tile_col.get(r.Status, '#999')
            c.markdown(f'<div class="ntile" style="--c:{col}" title="{E.CAT_NAME[r.BHSQA]} – {r.ValueText} (target '
                       f'{dict((x[0], x[2]) for x in E.CAT)[r.BHSQA]})"><div class="c">{r.BHSQA}</div>'
                       f'<div class="d">{E.CAT_NAME[r.BHSQA]}</div><div class="v">{r.ValueText}</div>'
                       f'<div class="t">{r.Status}</div></div>', unsafe_allow_html=True)
    st.markdown('##### Largest gaps to target (excluding surveys) – click a row to trace it to the source')
    gaps = X.top_gaps(K, 10)
    sel = kpi_table(gaps, 'gaps', height=390) if not gaps.empty else None
    if sel:
        with st.container(border=True):
            detail(sel, 'ov')

# ---- 2 units
with tabs[1]:
    t = K.pivot_table(index='unit', columns='status', values='sid', aggfunc='count', fill_value=0)
    for x in X.STATUS_ORDER:
        if x not in t:
            t[x] = 0
    t['rated'] = t['Met'] + t['Not met']
    t['% met'] = (t['Met'] / t['rated'].replace(0, float('nan')) * 100).astype(float).round(0)
    t = t.sort_values(['% met', 'Met'], ascending=[True, True])
    long = t[X.STATUS_ORDER].reset_index().melt('unit', var_name='Status', value_name='KPIs')
    fig = px.bar(long, y='unit', x='KPIs', color='Status', orientation='h', color_discrete_map=STATUS_COLORS,
                 category_orders={'Status': X.STATUS_ORDER, 'unit': list(t.index)}, height=max(420, 22 * len(t)))
    fig.update_layout(margin=dict(l=10, r=10, t=30, b=10), legend=dict(orientation='h', y=1.03), yaxis_title='',
                      title='KPI status by unit (sorted by % met)', plot_bgcolor='white')
    c1, c2 = st.columns([3, 2])
    with c1:
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        sc = t[['Met', 'Not met', 'No data', 'Not rated', '% met']].sort_values('% met', ascending=False)
        st.dataframe(sc, use_container_width=True, height=max(420, 22 * len(t)),
                     column_config={'% met': st.column_config.ProgressColumn('% met', min_value=0, max_value=100, format='%d%%')})
    st.divider()
    u = st.selectbox('Open a unit', list(t.index[::-1]))
    uk = K[K.unit == u]
    st.markdown(f'##### {u} – click a KPI to see its trend and the source cells')
    sel = kpi_table(uk, f'unit_{u}', height=min(40 * (len(uk) + 1), 460))
    if sel:
        with st.container(border=True):
            detail(sel, 'unit')

# ---- 3 explorer
with tabs[2]:
    f1, f2, f3 = st.columns([1, 1, 2])
    stf = f1.multiselect('Status', X.STATUS_ORDER, default=X.STATUS_ORDER)
    codes = sorted(K.bhsqa.unique(), key=lambda c: (int(''.join(ch for ch in c.split('-')[1] if ch.isdigit())), c))
    cf = f2.multiselect('BHSQA KPI', codes, placeholder='All')
    q = f3.text_input('Search KPI or unit', placeholder='e.g. caesarean, VAP, ICU…')
    ex = K[K.status.isin(stf)]
    if cf:
        ex = ex[ex.bhsqa.isin(cf)]
    if q:
        ql = q.lower()
        ex = ex[ex.kpi.str.lower().str.contains(ql) | ex.unit.str.lower().str.contains(ql) | ex.bhsqa_name.str.lower().str.contains(ql)]
    st.caption(f'{len(ex)} KPIs • click any row to drill down to the original sheet')
    sel = kpi_table(ex, 'explorer', height=430)
    if sel:
        with st.container(border=True):
            detail(sel, 'ex')

# ---- 4 BHSQA KPIs: number, description, value
with tabs[3]:
    nv = E.bhsqa_values(K)
    cat = pd.DataFrame(E.CAT, columns=['BHSQA', 'KPI', 'Target', 'Frequency', 'Group'])
    al = cat.merge(nv, on='BHSQA').merge(res.alignment[['BHSQA', 'Coverage']], on='BHSQA', how='left')
    order = ['Met', 'Not met', 'No data', 'Not rated', 'Not tracked']
    ncol = dict(STATUS_COLORS, **{'Not tracked': '#B0B7C3'})
    cs = st.columns(5)
    for c, o in zip(cs, order):
        card(c, int((al.Status == o).sum()), f'national KPIs – {o.lower()}', ncol[o])
    st.caption(f'{len(al)} BHSQA KPI entries (67 national KPIs; 10A/10B and 53a/53b listed separately). '
               'Value = hospital figure: pooled numerator ÷ denominator across the units that report the KPI, '
               'the hospital-wide QAS survey for satisfaction/waiting time, or the unit value when only one unit reports it.')
    f1, f2 = st.columns([1, 3])
    gsel = f1.multiselect('Group', ['A. Clinical', 'B. Managerial'], default=['A. Clinical', 'B. Managerial'], key='bh_grp')
    ssel = f2.multiselect('Status', order, default=order, key='bh_st')
    view = al[al.Group.isin(gsel or ['A. Clinical', 'B. Managerial']) & al.Status.isin(ssel or order)]
    icon = dict(STATUS_ICON, **{'Not tracked': '⬜'})
    tbl = pd.DataFrame({
        'BHSQA no.': view.BHSQA, 'National KPI description': view.KPI, 'BHSQA target': view.Target,
        'Hospital value': view.ValueText, 'Status': view.Status.map(lambda x: f'{icon.get(x, "")} {x}'),
        'How the value is derived': view.Basis,
        'Value by unit': view.ByUnit.map(lambda x: E.by_unit_text(x)),
        'Coverage': view.Coverage, 'Frequency': view.Frequency})
    sel_al = st.dataframe(tbl, hide_index=True, use_container_width=True, height=560, on_select='rerun',
                          selection_mode='single-row', key='align',
                          column_config={'National KPI description': st.column_config.TextColumn(width='large'),
                                         'Value by unit': st.column_config.TextColumn(width='large')})
    st.download_button('⬇️ Download this table (CSV)', tbl.to_csv(index=False).encode('utf-8-sig'),
                       'BHSQA_KPI_values.csv', 'text/csv', key='bh_csv')
    rows = sel_al.selection.rows if sel_al and sel_al.selection else []
    if rows:
        code = view.iloc[rows[0]].BHSQA
        sub = K[K.bhsqa == code]
        st.markdown(f'##### {code} – {E.CAT_NAME.get(code, "")}  ·  hospital value **{view.iloc[rows[0]].ValueText}**')
        if sub.empty:
            st.warning('This national KPI is not captured in any unit sheet of the uploaded workbooks.')
        else:
            st.caption('Units reporting this KPI – click one to see its trend and source cells')
            sid = kpi_table(sub, f'al_{code}', height=min(40 * (len(sub) + 1), 360))
            if sid:
                with st.container(border=True):
                    detail(sid, 'al')

# ---- 5 surveys
with tabs[4]:
    sv = K[K.bhsqa.isin(['KPI-54', 'KPI-55', 'KPI-45']) & K.achievement.notna()]
    if sv.empty:
        st.info('No survey results in the selected period.')
    else:
        sv = sv.assign(Survey=sv.bhsqa.map(X.SURVEY_CODES))
        fig = px.bar(sv, x='unit', y='achievement', color='Survey', barmode='group', height=480,
                     color_discrete_sequence=['#1F77B4', '#5DA5DA', '#B0C4DE'],
                     labels={'achievement': '%', 'unit': ''}, hover_data={'target': True})
        fig.add_hline(y=90, line_dash='dash', line_color='#C62828', annotation_text='BHSQA target >90%')
        fig.update_layout(margin=dict(l=10, r=10, t=40, b=10), title='Patient satisfaction and 5S-CQI by unit',
                          plot_bgcolor='white', legend=dict(orientation='h', y=1.08))
        st.plotly_chart(fig, use_container_width=True)
        wt = K[(K.bhsqa == 'KPI-57') & K.achievement.notna()]
        if not wt.empty:
            fig2 = px.bar(wt.sort_values('achievement'), x='achievement', y='unit', orientation='h', height=max(300, 28 * len(wt)),
                          color='status', color_discrete_map=STATUS_COLORS, labels={'achievement': 'minutes', 'unit': ''},
                          title='OPD waiting time for consultation (minutes)')
            fig2.add_vline(x=45, line_dash='dash', line_color='#C62828', annotation_text='Target <45 min')
            fig2.update_layout(plot_bgcolor='white', margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig2, use_container_width=True)
        st.caption('Surveys are bi-annual: the January column holds the Jan–Jun round and July the Jul–Dec round.')

# ---- 6 data quality
with tabs[5]:
    iss = res.issues[res.issues.Unit.isin(K.unit.unique())]
    c1, c2 = st.columns([1, 2])
    with c1:
        card(st, len(iss), 'issues found while extracting', '#EF6C00')
        bu = iss.groupby('Unit').size().sort_values(ascending=True).reset_index(name='issues')
        if not bu.empty:
            fig = px.bar(bu, x='issues', y='Unit', orientation='h', height=max(300, 22 * len(bu)),
                         labels={'Unit': ''}, color_discrete_sequence=['#EF6C00'])
            fig.update_layout(showlegend=False, margin=dict(l=10, r=10, t=20, b=10), plot_bgcolor='white')
            st.plotly_chart(fig, use_container_width=True)
    with c2:
        nod = K[K.status == 'No data'][['unit', 'bhsqa', 'kpi']].rename(columns={'unit': 'Unit', 'bhsqa': 'BHSQA', 'kpi': 'KPI'})
        st.markdown(f'##### KPIs with no data in the period ({len(nod)})')
        st.dataframe(nod, hide_index=True, use_container_width=True, height=300)
        st.markdown('##### Issue log')
        st.dataframe(iss, hide_index=True, use_container_width=True, height=360)

# ---- 7 reports
with tabs[6]:
    st.markdown('##### Download reports for the selected period and filters')
    st.caption(f'{len(K)} unit KPIs • {E.month_label(start)} – {E.month_label(end)} • {res.hospital}')
    key = (start, end, hosp, tuple(sorted(K.sid)))
    if st.button('⚙️ Prepare Excel, PDF and PowerPoint', type='primary'):
        with st.spinner('Building reports…'):
            st.session_state['reports'] = (key, X.excel_report(res, K), X.pdf_report(res, K), X.pptx_report(res, K))
    rep = st.session_state.get('reports')
    if rep and rep[0] == key:
        stem = f"{res.hospital.replace(' ', '_')}_BHSQA_KPI_{start}_to_{end}"
        a, b, c = st.columns(3)
        a.download_button('📗 Excel workbook', rep[1], f'{stem}.xlsx',
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
        b.download_button('📕 PDF report', rep[2], f'{stem}.pdf', 'application/pdf', use_container_width=True)
        c.download_button('📙 PowerPoint deck', rep[3], f'{stem}.pptx',
                          'application/vnd.openxmlformats-officedocument.presentationml.presentation', use_container_width=True)
        st.info('Excel: click a KPI name to jump to its data rows, and "Source" for the exact cells. '
                '"Open source" links in the Source Map sheet open the original workbook at that cell when the original '
                'files are saved in the same folder as the report.')
    elif rep:
        st.caption('Filters changed – prepare the reports again.')

st.caption(f'Targets: Revised BHSQA KPI User Guide (QASD, Department of Health Services, MoH). '
           f'Generated {dt.datetime.now():%d %b %Y %H:%M}.')
