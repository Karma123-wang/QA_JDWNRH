"""Excel / PDF / PowerPoint exports for the BHSQA KPI dashboard."""
from __future__ import annotations

import datetime as dt
import io
import os

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL

from kpi_engine import Result, month_label, HEADLINE, days_in, bhsqa_values, by_unit_text, CAT

STATUS_COLORS = {'Met': '#2E7D32', 'Not met': '#C62828', 'No data': '#9E9E9E', 'Not rated': '#EF6C00'}
SCORE_COLORS = dict(STATUS_COLORS, **{'Not tracked': '#8A94A6'})
STATUS_ORDER = ['Met', 'Not met', 'Not rated', 'No data']
NAVY = '#1F3A5F'


def fmt(v, uom=''):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return '—'
    if abs(v) >= 1000:
        s = f'{v:,.0f}'
    elif abs(v) >= 10:
        s = f'{v:.1f}'
    else:
        s = f'{v:.2f}'
    return s + ('%' if uom == '%' else '')


def summary(res: Result, kpis: pd.DataFrame | None = None):
    k = res.kpis if kpis is None else kpis
    c = k.status.value_counts()
    met, nm = int(c.get('Met', 0)), int(c.get('Not met', 0))
    return dict(total=len(k), met=met, notmet=nm, nodata=int(c.get('No data', 0)), notrated=int(c.get('Not rated', 0)),
                pct=(met / (met + nm) * 100) if met + nm else None)


def period_text(res: Result):
    return f"{month_label(res.months[0])} – {month_label(res.months[-1])}" if res.months else ''


# --------------------------------------------------------------------------- charts (matplotlib)
def _png(fig):
    b = io.BytesIO()
    fig.savefig(b, format='png', dpi=170, bbox_inches='tight')
    plt.close(fig)
    b.seek(0)
    return b


def chart_status_by_unit(kpis: pd.DataFrame, w=11, h=6):
    t = kpis.pivot_table(index='unit', columns='status', values='sid', aggfunc='count', fill_value=0)
    for s in STATUS_ORDER:
        if s not in t:
            t[s] = 0
    rated = (t['Met'] + t['Not met']).replace(0, 1)
    t = t.assign(_p=t['Met'] / rated).sort_values(['_p', 'Met'], ascending=True)
    fig, ax = plt.subplots(figsize=(w, h))
    left = [0] * len(t)
    for s in STATUS_ORDER:
        ax.barh(t.index, t[s], left=left, color=STATUS_COLORS[s], label=s, height=0.7)
        left = [a + b for a, b in zip(left, t[s])]
    ax.set_xlabel('Number of BHSQA KPIs')
    ax.legend(ncol=4, loc='lower right', frameon=False, fontsize=8)
    ax.tick_params(axis='y', labelsize=7.5)
    for sp in ('top', 'right'):
        ax.spines[sp].set_visible(False)
    ax.set_title('BHSQA KPI status by unit', loc='left', fontsize=12, color=NAVY, fontweight='bold')
    return _png(fig)


def chart_donut(s):
    vals = [s['met'], s['notmet'], s['notrated'], s['nodata']]
    fig, ax = plt.subplots(figsize=(4, 4))
    ax.pie(vals, colors=[STATUS_COLORS[x] for x in STATUS_ORDER], startangle=90, counterclock=False,
           wedgeprops=dict(width=0.35, edgecolor='white'))
    ax.text(0, 0.08, f"{s['pct']:.0f}%" if s['pct'] is not None else '—', ha='center', fontsize=24, fontweight='bold', color=NAVY)
    ax.text(0, -0.2, 'of rated KPIs met', ha='center', fontsize=9, color='#555')
    ax.legend([f'{a} ({b})' for a, b in zip(STATUS_ORDER, vals)], loc='lower center', bbox_to_anchor=(0.5, -0.18),
              ncol=2, frameon=False, fontsize=8)
    return _png(fig)


def chart_coverage(al: pd.DataFrame):
    order = ['Reported', 'Partly reported', 'Tracked – no data', 'Not tracked']
    cols = ['#2E7D32', '#F9A825', '#9E9E9E', '#C62828']
    v = [int((al.Coverage == o).sum()) for o in order]
    fig, ax = plt.subplots(figsize=(7, 2.6))
    ax.barh(order[::-1], v[::-1], color=cols[::-1])
    for i, x in enumerate(v[::-1]):
        ax.text(x + 0.4, i, str(x), va='center', fontsize=10)
    ax.set_title(f'Coverage of the {len(al)} BHSQA KPI entries (67 KPIs incl. 10A/B, 53a/b)', loc='left', fontsize=11, color=NAVY, fontweight='bold')
    for sp in ('top', 'right'):
        ax.spines[sp].set_visible(False)
    return _png(fig)


def chart_headlines(res: Result):
    items = []
    for unit, code in HEADLINE:
        r = res.kpis[(res.kpis.unit == unit) & (res.kpis.bhsqa == code)]
        if not r.empty:
            items.append(r.iloc[0])
    n = len(items)
    cols = 4
    rows = (n + cols - 1) // cols
    fig, axes = plt.subplots(rows, cols, figsize=(13, 2.5 * rows))
    axes = axes.flatten() if n > 1 else [axes]
    for ax, k in zip(axes, items):
        m = res.monthly[res.monthly.sid == k.sid]
        x = list(range(len(m)))
        ax.plot(x, m.result, marker='o', color=NAVY, lw=1.8, ms=3)
        if k.dir in ('<', '<=') and k.hi is not None:
            ax.axhline(k.hi, color='#C62828', ls='--', lw=1)
        elif k.dir in ('>', '>=') and k.lo is not None:
            ax.axhline(k.lo, color='#C62828', ls='--', lw=1)
        elif k.dir == 'range':
            ax.axhspan(k.lo, k.hi, color='#2E7D32', alpha=0.12)
        ax.set_xticks(x)
        ax.set_xticklabels([l[:3] for l in m.label], fontsize=6)
        ax.tick_params(axis='y', labelsize=7)
        c = STATUS_COLORS.get(k.status, '#555')
        ax.set_title(f"{k.bhsqa} {k.kpi[:30]}\n{fmt(k.achievement, k.uom)}  •  target {k.target}", fontsize=8, color=c, loc='left')
        for sp in ('top', 'right'):
            ax.spines[sp].set_visible(False)
    for ax in axes[n:]:
        ax.axis('off')
    fig.tight_layout()
    return _png(fig)


SURVEY_CODES = {'KPI-54': 'OPD satisfaction', 'KPI-55': 'IPD satisfaction', 'KPI-45': '5S-CQI'}


def is_survey(k):
    return k.typ == 'survey'


def top_gaps(kpis: pd.DataFrame, n=6):
    w = kpis[(kpis.status == 'Not met') & (kpis.typ != 'survey')].copy()
    if w.empty:
        return w

    def rel(k):
        thr = k.hi if k.dir in ('<', '<=') else k.lo
        if k.dir == 'range':
            thr = k.hi if k.achievement > k.hi else k.lo
        return abs(k.gap) / max(abs(thr or 0), 5)
    w['rel'] = w.apply(rel, axis=1)
    return w.sort_values('rel', ascending=False).head(n)


def chart_surveys(kpis: pd.DataFrame, w=12, h=5):
    sv = kpis[kpis.bhsqa.isin(SURVEY_CODES) & kpis.achievement.notna()]
    if sv.empty:
        return None
    t = sv.pivot_table(index='unit', columns='bhsqa', values='achievement', aggfunc='mean')
    t = t.reindex(columns=[c for c in SURVEY_CODES if c in t.columns])
    t = t.loc[t.mean(axis=1).sort_values().index]
    fig, ax = plt.subplots(figsize=(w, h))
    n = len(t.columns); bw = 0.8 / n
    cols = ['#1F77B4', '#5DA5DA', '#B0C4DE']
    for i, c in enumerate(t.columns):
        ax.bar([x + i * bw for x in range(len(t))], t[c], width=bw, label=SURVEY_CODES[c], color=cols[i])
    ax.axhline(90, color='#C62828', ls='--', lw=1.2); ax.text(len(t) - 0.5, 91, 'Target >90%', color='#C62828', fontsize=8, ha='right')
    ax.set_xticks([x + bw * (n - 1) / 2 for x in range(len(t))]); ax.set_xticklabels(t.index, rotation=55, ha='right', fontsize=7.5)
    ax.set_ylim(0, 105); ax.set_ylabel('%'); ax.legend(frameon=False, ncol=3, fontsize=8, loc='lower right')
    ax.set_title('Survey-based KPIs by unit (latest rounds in period)', loc='left', fontsize=12, color=NAVY, fontweight='bold')
    for sp in ('top', 'right'):
        ax.spines[sp].set_visible(False)
    return _png(fig)


# --------------------------------------------------------------------------- EXCEL
def excel_report(res: Result, kpis: pd.DataFrame | None = None) -> bytes:
    kpis = res.kpis if kpis is None else kpis
    AR = 'Arial'

    def F(**k):
        return Font(name=AR, size=k.pop('size', 10), **k)
    HF = PatternFill('solid', fgColor='1F4E78'); UF = PatternFill('solid', fgColor='D9E1F2'); RF = PatternFill('solid', fgColor='F2F2F2')
    thin = Side(style='thin', color='BFBFBF'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
    BLUE = F(color='0000FF'); GREEN = F(color='008000'); LINK = F(color='0563C1', underline='single')

    def header(ws, row, names, height=32):
        for i, n in enumerate(names, 1):
            c = ws.cell(row, i, n); c.font = F(bold=True, color='FFFFFF'); c.fill = HF; c.border = BOX
            c.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        ws.row_dimensions[row].height = height

    wb = Workbook()
    rm = wb.active; rm.title = 'Read Me'
    ta = wb.create_sheet('Target vs Achievement'); ds = wb.create_sheet('Data')
    sc = wb.create_sheet('Unit Scorecard'); al = wb.create_sheet('BHSQA Alignment')
    sm = wb.create_sheet('Source Map'); iss = wb.create_sheet('Data Issues')
    months = res.months; nm = len(months)
    FC = 6; LC = FC + nm - 1; TC = LC + 1
    fm, lm, tl = CL(FC), CL(LC), CL(TC)

    # ---- Source Map first (row numbers needed for links)
    sm['A1'] = f"Source map – every value used, with the exact file / sheet / cell it came from"; sm['A1'].font = F(bold=True, size=12)
    sm['A2'] = ("'Open source' links work when the original reporting workbooks are saved in the same folder as this file.")
    sm['A2'].font = F(italic=True, color='595959')
    header(sm, 3, ['KPI ID', 'Unit', 'KPI', 'Month', 'Role', 'File', 'Sheet', 'Cell', 'Metric in sheet', 'Value in sheet', 'Value used', 'Open source'])
    first_sm = {}
    r = 4
    for _, k in kpis.iterrows():
        for key in months:
            for role, lab in (('num', 'Numerator'), ('den', 'Denominator')):
                c = res.cells.get((k.sid, key, role))
                if c is None:
                    continue
                for x in (c if isinstance(c, list) else [c]):
                    first_sm.setdefault(k.sid, r)
                    vals = [k.sid, k.unit, k.kpi, month_label(key), lab, x.file, x.sheet, x.cell, x.label,
                            '' if x.raw is None else str(x.raw), x.value]
                    for i, v in enumerate(vals, 1):
                        sm.cell(r, i, v).font = F(size=9)
                    lk = sm.cell(r, 12, 'Open source')
                    lk.hyperlink = f"{x.file}#'{x.sheet}'!{x.cell}"; lk.font = LINK
                    r += 1
    for i, w in enumerate([8, 22, 40, 10, 12, 44, 22, 7, 40, 14, 10, 12], 1):
        sm.column_dimensions[CL(i)].width = w
    sm.freeze_panes = 'A4'; sm.auto_filter.ref = f'A3:L{max(r - 1, 3)}'

    # ---- Data sheet
    ds['A1'] = f"{res.hospital} – BHSQA KPI data, {period_text(res)}"; ds['A1'].font = F(bold=True, size=12)
    ds['A2'] = 'Blue = values extracted from the source sheets (edit to correct). Black = formulas. Blank = not reported.'
    ds['A2'].font = F(italic=True, color='595959')
    header(ds, 3, ['Unit', 'KPI', 'BHSQA KPI', 'Hospital ind.', 'Line'] + [month_label(m) for m in months] + ['Period total / result', 'Unit of measure', 'Notes'])
    res_row = {}
    r = 4; prev = None
    mon = res.monthly.set_index(['sid', 'month'])
    for _, k in kpis.iterrows():
        if k.unit != prev:
            ds.cell(r, 1, k.unit).font = F(bold=True, size=11)
            for c in range(1, TC + 3):
                ds.cell(r, c).fill = UF
            r += 1; prev = k.unit

        def lab(row, line):
            for c, v in ((1, k.unit), (2, k.kpi), (3, k.bhsqa), (4, k.jcode), (5, line)):
                ds.cell(row, c, v).font = F(bold=(line == 'Result'), color='808080' if c == 1 else None)
            for c in range(1, TC + 3):
                ds.cell(row, c).border = BOX
            ds.cell(row, TC + 1, k.uom).font = F()

        def vals(row, key):
            for i, m in enumerate(months):
                v = mon.loc[(k.sid, m), key] if (k.sid, m) in mon.index else None
                v = None if v is None or pd.isna(v) else float(v)
                c = ds.cell(row, FC + i, v); c.font = BLUE; c.number_format = '#,##0.##'
        start = r
        if k.typ in ('ratio', 'comp', 'bor', 'nurse'):
            nr = r; lab(r, 'Number of nursing staff' if k.typ == 'nurse' else 'Numerator'); vals(r, 'num'); r += 1
            if k.typ == 'bor':
                br = r; lab(r, 'Bed strength')
                for i, m in enumerate(months):
                    d = mon.loc[(k.sid, m), 'den'] if (k.sid, m) in mon.index else None
                    c = ds.cell(r, FC + i, None if d is None or pd.isna(d) else round(d / days_in(m), 2)); c.font = BLUE
                r += 1
                dyr = r; lab(r, 'Days in month')
                for i, m in enumerate(months):
                    ds.cell(r, FC + i, days_in(m)).font = BLUE
                r += 1
                dr = r; lab(r, 'Denominator (beds × days)')
                for i in range(nm):
                    col = CL(FC + i)
                    ds.cell(r, FC + i, f'=IF(AND(ISNUMBER({col}{br}),ISNUMBER({col}{dyr})),{col}{br}*{col}{dyr},"")').number_format = '#,##0'
                r += 1
            else:
                dr = r; lab(r, 'Number of beds' if k.typ == 'nurse' else 'Denominator'); vals(r, 'den'); r += 1
            for a, b in ((nr, dr), (dr, nr)):
                ds.cell(a, TC, f'=SUMPRODUCT(ISNUMBER({fm}{a}:{lm}{a})*ISNUMBER({fm}{b}:{lm}{b}),{fm}{a}:{lm}{a})').number_format = '#,##0.##'
            rr = r; lab(r, 'Result')
            mult = k.mult
            for i in range(nm):
                col = CL(FC + i)
                if k.typ == 'nurse':
                    f = f'=IF(AND(ISNUMBER({col}{nr}),ISNUMBER({col}{dr})),IF({col}{nr}=0,"",{col}{dr}/{col}{nr}),"")'
                else:
                    f = f'=IF(AND(ISNUMBER({col}{nr}),ISNUMBER({col}{dr})),IF({col}{dr}=0,"",{col}{nr}/{col}{dr}*{mult:g}),"")'
                ds.cell(r, FC + i, f)
            ds.cell(r, TC, f'=IF({tl}{nr}=0,"",{tl}{dr}/{tl}{nr})' if k.typ == 'nurse'
                    else f'=IF({tl}{dr}=0,"",{tl}{nr}/{tl}{dr}*{mult:g})')
        else:
            vr = r; lab(r, 'Value (count)' if k.typ == 'count' else 'Recorded value'); vals(r, 'num'); r += 1
            rr = r; lab(r, 'Result')
            for i in range(nm):
                col = CL(FC + i)
                ds.cell(r, FC + i, f'=IF(ISNUMBER({col}{vr}),{col}{vr},"")')
            agg = 'SUM' if k.typ == 'count' else 'AVERAGE'
            ds.cell(r, TC, f'=IF(COUNT({fm}{vr}:{lm}{vr})=0,"",{agg}({fm}{vr}:{lm}{vr}))')
        for c in range(FC, TC + 1):
            x = ds.cell(rr, c); x.font = F(bold=True); x.fill = RF; x.number_format = '0.00'
        ds.cell(start, TC + 2, k.note or '').font = F(size=9, color='595959')
        if k.sid in first_sm:
            lk = ds.cell(rr, TC + 2, 'View sources →'); lk.hyperlink = f"#'Source Map'!A{first_sm[k.sid]}"; lk.font = LINK
        res_row[k.sid] = rr
        r += 1
    ds.freeze_panes = 'F4'
    for i, w in enumerate([20, 42, 10, 9, 26] + [9] * nm + [12, 18, 60], 1):
        ds.column_dimensions[CL(i)].width = w

    # ---- Target vs Achievement
    ta['A1'] = f"{res.hospital} – unit-wise BHSQA target vs achievement, {period_text(res)}"; ta['A1'].font = F(bold=True, size=12)
    ta['A2'] = ('Click a KPI name to jump to its data rows; click "Source" for the exact cells in the original reporting sheets. '
                'Achievement = period total numerator ÷ total denominator (months with both values) or mean of survey rounds.')
    ta['A2'].font = F(italic=True, color='595959')
    TFC = FC + 1
    hdr = ['Unit', 'Unit KPI', 'BHSQA no.', 'BHSQA KPI (national description)', 'Hospital ind.', 'Unit of measure'] + [month_label(m) for m in months] + \
          ['Period achievement', 'BHSQA target', 'Rule', 'Lower', 'Upper', 'Status', 'Gap to target', 'Source', 'Notes']
    header(ta, 3, hdr, 42)
    c_ach = TFC + nm; c_t = c_ach + 1; c_rule = c_t + 1; c_lo = c_rule + 1; c_hi = c_lo + 1; c_st = c_hi + 1; c_gap = c_st + 1
    tr = 4
    for _, k in kpis.iterrows():
        rr = res_row[k.sid]
        ta.cell(tr, 1, k.unit)
        kc = ta.cell(tr, 2, k.kpi); kc.hyperlink = f"#'Data'!B{rr}"; kc.font = LINK
        ta.cell(tr, 3, k.bhsqa); ta.cell(tr, 4, k.bhsqa_name); ta.cell(tr, 5, k.jcode); ta.cell(tr, 6, k.uom)
        for i in range(nm):
            c = ta.cell(tr, TFC + i, f"=Data!{CL(FC + i)}{rr}"); c.font = GREEN; c.number_format = '0.00'
        c = ta.cell(tr, c_ach, f"=Data!{tl}{rr}"); c.font = F(bold=True, color='008000'); c.number_format = '0.00'
        ta.cell(tr, c_t, k.target); ta.cell(tr, c_rule, k.dir)
        ta.cell(tr, c_lo, None if pd.isna(k.lo) else k.lo); ta.cell(tr, c_hi, None if pd.isna(k.hi) else k.hi)
        P, R_, Lo, Hi = (f'{CL(c_ach)}{tr}', f'{CL(c_rule)}{tr}', f'{CL(c_lo)}{tr}', f'{CL(c_hi)}{tr}')
        ta.cell(tr, c_st, (f'=IF({R_}="info","No BHSQA target",IF(NOT(ISNUMBER({P})),"No data",IF({R_}="norate","Not rated",'
                           f'IF({R_}="<",IF({P}<{Hi},"Met","Not met"),IF({R_}="<=",IF({P}<={Hi},"Met","Not met"),'
                           f'IF({R_}=">",IF({P}>{Lo},"Met","Not met"),IF({R_}=">=",IF({P}>={Lo},"Met","Not met"),'
                           f'IF(AND({P}>={Lo},{P}<={Hi}),"Met","Not met"))))))))')).font = F(bold=True)
        ta.cell(tr, c_gap, (f'=IF(OR({R_}="info",{R_}="norate",NOT(ISNUMBER({P}))),"",IF(OR({R_}="<",{R_}="<="),{P}-{Hi},'
                            f'IF(OR({R_}=">",{R_}=">="),{P}-{Lo},IF({P}<{Lo},{P}-{Lo},IF({P}>{Hi},{P}-{Hi},0)))))')).number_format = '+0.00;-0.00;0'
        if k.sid in first_sm:
            s = ta.cell(tr, c_gap + 1, 'Source'); s.hyperlink = f"#'Source Map'!A{first_sm[k.sid]}"; s.font = LINK
        ta.cell(tr, c_gap + 2, k.note or '').font = F(size=9)
        for c in range(1, c_gap + 3):
            ta.cell(tr, c).border = BOX
        tr += 1
    last = tr - 1
    stc = CL(c_st)
    rng = f'{stc}4:{stc}{max(last, 4)}'
    for val, fill, col in (('Met', 'C6EFCE', '006100'), ('Not met', 'FFC7CE', '9C0006'),
                           ('No data', 'EDEDED', '404040'), ('Not rated', 'FCE4D6', '833C0B')):
        ta.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=[f'"{val}"'],
                                                      fill=PatternFill('solid', fgColor=fill), font=Font(name=AR, color=col, bold=True)))
    ta.freeze_panes = 'G4'; ta.auto_filter.ref = f'A3:{CL(c_gap + 2)}{max(last, 4)}'
    for i, w in enumerate([22, 40, 10, 40, 8, 16] + [8] * nm + [12, 24, 6, 7, 7, 12, 10, 8, 50], 1):
        ta.column_dimensions[CL(i)].width = w

    # ---- Scorecard
    TV = "'Target vs Achievement'!"
    sc['A1'] = f'Unit scorecard – {period_text(res)}'; sc['A1'].font = F(bold=True, size=12)
    header(sc, 3, ['Unit', 'BHSQA KPIs', 'Met', 'Not met', 'No data', 'Not rated', '% met (of rated)'], 36)
    r = 4
    for u in dict.fromkeys(kpis.unit):
        sc.cell(r, 1, u).font = F(bold=True)
        sc.cell(r, 2, f'=COUNTIF({TV}$A$4:$A${last},A{r})')
        for c, s in ((3, 'Met'), (4, 'Not met'), (5, 'No data'), (6, 'Not rated')):
            sc.cell(r, c, f'=COUNTIFS({TV}$A$4:$A${last},A{r},{TV}${stc}$4:${stc}${last},"{s}")')
        sc.cell(r, 7, f'=IF(C{r}+D{r}=0,"",C{r}/(C{r}+D{r}))').number_format = '0%'
        for c in range(1, 8):
            sc.cell(r, c).border = BOX
        r += 1
    sc.cell(r, 1, 'Hospital total').font = F(bold=True)
    for c in range(2, 7):
        sc.cell(r, c, f'=SUM({CL(c)}4:{CL(c)}{r - 1})').font = F(bold=True)
    sc.cell(r, 7, f'=IF(C{r}+D{r}=0,"",C{r}/(C{r}+D{r}))').number_format = '0%'
    for c in range(1, 8):
        sc.cell(r, c).fill = UF; sc.cell(r, c).border = BOX
    for i, w in enumerate([28, 11, 8, 9, 9, 10, 14], 1):
        sc.column_dimensions[CL(i)].width = w

    # ---- Alignment
    al['A1'] = 'National BHSQA KPIs – alignment with hospital reporting sheets'; al['A1'].font = F(bold=True, size=12)
    al['A2'] = '* = benchmark set by professional judgement (BHSQA KPI User Guide).'; al['A2'].font = F(italic=True, color='595959')
    al['A1'] = f'National BHSQA KPIs – number, description and hospital value, {period_text(res)}'
    nv = bhsqa_values(kpis).merge(res.alignment[['BHSQA', 'KPI', 'Group', 'Target', 'Frequency', 'Coverage']], on='BHSQA')
    header(al, 3, ['BHSQA no.', 'National KPI description', 'Group', 'BHSQA target', 'Hospital value', 'Status',
                   'How the value is derived', 'Value by unit', 'Coverage', 'Frequency'])
    for i, r in enumerate(nv.itertuples(), 4):
        vals = [r.BHSQA, r.KPI, r.Group, r.Target, r.ValueText, r.Status, r.Basis, by_unit_text(r.ByUnit), r.Coverage, r.Frequency]
        for j, v in enumerate(vals, 1):
            c = al.cell(i, j, v); c.font = F(size=9 if j in (7, 8) else 10, bold=(j in (1, 5))); c.border = BOX
            c.alignment = Alignment(wrap_text=j in (2, 7, 8), vertical='top')
    lastA = 3 + len(nv)
    for val, fill in (('Met', 'C6EFCE'), ('Not met', 'FFC7CE'), ('No data', 'EDEDED'), ('Not rated', 'FCE4D6'), ('Not tracked', 'D9D9D9')):
        al.conditional_formatting.add(f'F4:F{lastA}', CellIsRule(operator='equal', formula=[f'"{val}"'], fill=PatternFill('solid', fgColor=fill)))
    al.cell(lastA + 2, 1, 'Hospital value = pooled numerator ÷ denominator across the units reporting the KPI (Data sheet totals); '
            'satisfaction and OPD waiting time use the hospital-wide QAS survey; surveys reported by several units are averaged.').font = F(italic=True, size=9)
    for i, w in enumerate([10, 46, 13, 24, 22, 11, 26, 70, 18, 12], 1):
        al.column_dimensions[CL(i)].width = w
    al.freeze_panes = 'C4'

    # ---- Issues
    iss['A1'] = 'Data-quality issues detected during extraction'; iss['A1'].font = F(bold=True, size=12)
    header(iss, 3, list(res.issues.columns))
    for i, row in enumerate(res.issues.itertuples(index=False), 4):
        for j, v in enumerate(row, 1):
            iss.cell(i, j, v).font = F(size=9)
    for i, w in enumerate([24, 40, 10, 60, 70], 1):
        iss.column_dimensions[CL(i)].width = w

    # ---- Read me
    lines = [(f'{res.hospital} – BHSQA KPI report, {period_text(res)}', F(bold=True, size=14)), ('', None),
             (f'Generated {dt.datetime.now():%d %b %Y %H:%M} from: ' + '; '.join(f"{b.name} ({b.fy_label})" for b in res.books.values()), F()),
             ('Targets: Revised BHSQA KPI User Guide, QASD, Department of Health Services, Ministry of Health.', F()), ('', None),
             ('Sheets', F(bold=True)),
             ('• Target vs Achievement – one row per unit KPI with monthly results, period achievement, target and status.', F()),
             ('• Data – numerators/denominators (blue, editable) and result formulas.', F()),
             ('• Unit Scorecard – met / not met / no data per unit.', F()),
             ('• BHSQA KPI Scorecard – every national KPI number with its description, target, hospital value, status and value by unit.', F()),
             ('• Source Map – the exact file › sheet › cell of every value. Keep the original workbooks in the same folder to use "Open source".', F()),
             ('• Data Issues – blanks, text entries and impossible values found.', F()), ('', None),
             ('Status: Met / Not met against BHSQA target; No data = no month with both values; Not rated = BHSQA target cannot be applied to the data as recorded.', F())]
    for i, (t, f) in enumerate(lines, 1):
        c = rm.cell(i, 1, t)
        if f:
            c.font = f
    rm.column_dimensions['A'].width = 150
    al.title = 'BHSQA KPI Scorecard'
    wb.move_sheet('Target vs Achievement', offset=0)
    b = io.BytesIO(); wb.save(b)
    return b.getvalue()


# --------------------------------------------------------------------------- PDF
def _font():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    base = os.path.join(matplotlib.get_data_path(), 'fonts', 'ttf')
    try:
        pdfmetrics.registerFont(TTFont('DejaVu', os.path.join(base, 'DejaVuSans.ttf')))
        pdfmetrics.registerFont(TTFont('DejaVu-Bold', os.path.join(base, 'DejaVuSans-Bold.ttf')))
        return 'DejaVu', 'DejaVu-Bold'
    except Exception:
        return 'Helvetica', 'Helvetica-Bold'


def pdf_report(res: Result, kpis: pd.DataFrame | None = None) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, KeepTogether)

    kpis = res.kpis if kpis is None else kpis
    FN, FB = _font()
    H1 = ParagraphStyle('h1', fontName=FB, fontSize=20, textColor=colors.HexColor(NAVY), spaceAfter=6, leading=24)
    H2 = ParagraphStyle('h2', fontName=FB, fontSize=13, textColor=colors.HexColor(NAVY), spaceBefore=6, spaceAfter=6)
    BODY = ParagraphStyle('b', fontName=FN, fontSize=9, leading=12)
    SMALL = ParagraphStyle('s', fontName=FN, fontSize=7.5, leading=9.5)
    SMB = ParagraphStyle('sb', parent=SMALL, fontName=FB)
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=14 * mm, rightMargin=14 * mm, topMargin=14 * mm, bottomMargin=14 * mm,
                            title=f'{res.hospital} BHSQA KPI report')
    W = landscape(A4)[0] - 28 * mm
    s = summary(res, kpis)
    el = []

    def footer(c, d):
        c.saveState(); c.setFont(FN, 7); c.setFillColor(colors.grey)
        c.drawString(14 * mm, 8 * mm, f'{res.hospital} • BHSQA KPI report • {period_text(res)}')
        c.drawRightString(landscape(A4)[0] - 14 * mm, 8 * mm, f'Page {d.page}')
        c.restoreState()

    def tbl(data, widths, header_bg=NAVY, status_col=None, zebra=True):
        t = Table(data, colWidths=widths, repeatRows=1)
        st = [('FONT', (0, 0), (-1, 0), FB, 8), ('FONT', (0, 1), (-1, -1), FN, 7.5),
              ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(header_bg)), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
              ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#BFBFBF')), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]
        if zebra:
            for i in range(1, len(data)):
                if i % 2 == 0:
                    st.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#F5F7FA')))
        if status_col is not None:
            for i in range(1, len(data)):
                v = data[i][status_col]
                v = v.text if hasattr(v, 'text') else v
                if v in SCORE_COLORS:
                    st.append(('TEXTCOLOR', (status_col, i), (status_col, i), colors.HexColor(SCORE_COLORS[v])))
                    st.append(('FONT', (status_col, i), (status_col, i), FB, 7.5))
        t.setStyle(TableStyle(st))
        return t

    # cover + executive summary
    el += [Paragraph(f'{res.hospital} – Key Performance Indicators', H1),
           Paragraph(f'Achievement against national BHSQA targets • {period_text(res)}', H2),
           Paragraph(f'Source: ' + '; '.join(f'{b.name} ({b.fy_label})' for b in res.books.values()) +
                     f'. Targets: Revised BHSQA KPI User Guide (QASD, MoH). Generated {dt.date.today():%d %b %Y}.', BODY),
           Spacer(1, 6)]
    cards = [[Paragraph(f'<font size=18><b>{v}</b></font><br/>{l}', ParagraphStyle('c', fontName=FN, fontSize=8, alignment=1, leading=20))
              for v, l in ((s['total'], 'BHSQA KPIs tracked'), (s['met'], 'Met'), (s['notmet'], 'Not met'),
                           (s['nodata'], 'No data'), (f"{s['pct']:.0f}%" if s['pct'] is not None else '—', 'of rated KPIs met'))]]
    ct = Table(cards, colWidths=[W / 5] * 5, rowHeights=[22 * mm])
    ct.setStyle(TableStyle([('BOX', (i, 0), (i, 0), 1, colors.HexColor(c)) for i, c in
                            enumerate([NAVY, STATUS_COLORS['Met'], STATUS_COLORS['Not met'], STATUS_COLORS['No data'], NAVY])] +
                           [('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    el += [ct, Spacer(1, 8)]
    el.append(Table([[Image(chart_donut(s), width=62 * mm, height=66 * mm),
                      Image(chart_coverage(res.alignment), width=W - 72 * mm, height=(W - 72 * mm) * 2.6 / 7)]],
                    colWidths=[68 * mm, W - 68 * mm]))
    el.append(PageBreak())
    nv = bhsqa_values(kpis).merge(res.alignment[['BHSQA', 'KPI', 'Target']], on='BHSQA')
    el.append(Paragraph('National BHSQA KPI scorecard – number, description and hospital value', H2))
    el.append(Paragraph('Hospital value = pooled numerator ÷ denominator across reporting units; satisfaction and OPD waiting '
                        'time from the hospital-wide QAS survey; values by unit in the last column.', SMALL))
    el.append(Spacer(1, 4))
    data = [['BHSQA no.', 'National KPI description', 'BHSQA target', 'Hospital value', 'Status', 'Value by unit']]
    for r in nv.itertuples():
        data.append([r.BHSQA, Paragraph(r.KPI, SMALL), Paragraph(r.Target, SMALL), Paragraph(f'<b>{r.ValueText}</b>', SMALL),
                     r.Status, Paragraph(by_unit_text(r.ByUnit, limit=8)[:420], SMALL)])
    el += [tbl(data, [18 * mm, 62 * mm, 36 * mm, 34 * mm, 18 * mm, W - 168 * mm], status_col=4), PageBreak()]
    el += [Paragraph('Headline hospital indicators', H2), Image(chart_headlines(res), width=W, height=W * 7.5 / 13 * 0.8)]
    el.append(PageBreak())
    el += [Paragraph('Status by unit', H2), Image(chart_status_by_unit(kpis, 11, 7), width=W * 0.85, height=W * 0.85 * 7 / 11)]
    sv = chart_surveys(kpis)
    if sv:
        el += [PageBreak(), Paragraph('Patient satisfaction and 5S-CQI surveys', H2), Image(sv, width=W, height=W * 5 / 12)]
    el.append(PageBreak())

    # not met
    nmk = kpis[kpis.status == 'Not met']
    el.append(Paragraph(f'KPIs not meeting the BHSQA target ({len(nmk)}, of which {int((nmk.typ == "survey").sum())} survey-based)', H2))
    data = [['Unit', 'BHSQA KPI (no. and description)', 'Unit KPI', 'Value', 'Target', 'Gap', 'Months']]
    for _, k in nmk.iterrows():
        data.append([Paragraph(k.unit, SMALL), Paragraph(f'<b>{k.bhsqa}</b> {k.bhsqa_name}', SMALL), Paragraph(k.kpi, SMALL),
                     fmt(k.achievement, k.uom), Paragraph(k.target, SMALL), fmt(k.gap), str(k.months_with_data)])
    el += [tbl(data, [40 * mm, 68 * mm, 62 * mm, 22 * mm, 40 * mm, 16 * mm, 14 * mm]), PageBreak()]

    # unit-wise detail
    el.append(Paragraph('Unit-wise target vs achievement', H2))
    for u in dict.fromkeys(kpis.unit):
        sub = kpis[kpis.unit == u]
        data = [['BHSQA', 'KPI', 'UoM'] + [l[:3] for l in map(month_label, res.months)] + ['Period', 'Target', 'Status']]
        mon = res.monthly.set_index(['sid', 'month'])
        for _, k in sub.iterrows():
            mv = [fmt(mon.loc[(k.sid, m), 'result']) if (k.sid, m) in mon.index else '—' for m in res.months]
            data.append([k.bhsqa, Paragraph(k.kpi, SMALL), Paragraph(k.uom, SMALL)] + mv +
                        [fmt(k.achievement, k.uom), Paragraph(k.target, SMALL), k.status])
        nmc = len(res.months)
        mw = min(13 * mm, (W - 175 * mm) / max(nmc, 1))
        widths = [16 * mm, 60 * mm, 26 * mm] + [mw] * nmc + [18 * mm, 33 * mm, 19 * mm]
        el.append(KeepTogether([Paragraph(u, ParagraphStyle('u', fontName=FB, fontSize=10, spaceBefore=6, spaceAfter=3)),
                                tbl(data, widths, status_col=len(data[0]) - 1)]))
    el.append(PageBreak())


    # issues
    el.append(Paragraph(f'Data-quality issues ({len(res.issues)})', H2))
    data = [['Unit', 'KPI', 'Month', 'Issue']]
    for row in res.issues.head(250).itertuples(index=False):
        data.append([Paragraph(row.Unit, SMALL), Paragraph(row.KPI, SMALL), row.Month, Paragraph(row.Issue, SMALL)])
    el.append(tbl(data, [45 * mm, 75 * mm, 20 * mm, 130 * mm]))
    doc.build(el, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


# --------------------------------------------------------------------------- PPTX
def pptx_report(res: Result, kpis: pd.DataFrame | None = None) -> bytes:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN

    kpis = res.kpis if kpis is None else kpis
    s = summary(res, kpis)
    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    navy = RGBColor(0x1F, 0x3A, 0x5F)

    def rgb(h):
        h = h.lstrip('#'); return RGBColor(int(h[:2], 16), int(h[2:4], 16), int(h[4:], 16))

    def title(sl, t, sub=None):
        bar = sl.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(0.9)); bar.fill.solid(); bar.fill.fore_color.rgb = navy
        bar.line.fill.background()
        tb = sl.shapes.add_textbox(Inches(0.4), Inches(0.12), Inches(12.5), Inches(0.7)).text_frame
        tb.text = t; p = tb.paragraphs[0]; p.font.size = Pt(26); p.font.bold = True; p.font.color.rgb = RGBColor(255, 255, 255)
        if sub:
            st = sl.shapes.add_textbox(Inches(0.4), Inches(0.95), Inches(12.5), Inches(0.4)).text_frame
            st.text = sub; st.paragraphs[0].font.size = Pt(13); st.paragraphs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        ft = sl.shapes.add_textbox(Inches(0.4), Inches(7.05), Inches(12.5), Inches(0.3)).text_frame
        ft.text = f'{res.hospital} • BHSQA KPIs • {period_text(res)}'; ft.paragraphs[0].font.size = Pt(9)
        ft.paragraphs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    def table(sl, rows, top, widths, status_col=None, font=10):
        n, m = len(rows), len(rows[0])
        shp = sl.shapes.add_table(n, m, Inches(0.4), top, Inches(sum(widths)), Inches(0.3 * n))
        t = shp.table
        for j, w in enumerate(widths):
            t.columns[j].width = Inches(w)
        for i, r in enumerate(rows):
            for j, v in enumerate(r):
                c = t.cell(i, j); c.text = str(v)
                p = c.text_frame.paragraphs[0]; p.font.size = Pt(font if i else font); p.font.bold = (i == 0)
                c.margin_top = c.margin_bottom = Inches(0.02)
                if i == 0:
                    c.fill.solid(); c.fill.fore_color.rgb = navy; p.font.color.rgb = RGBColor(255, 255, 255)
                if status_col is not None and j == status_col and v in SCORE_COLORS:
                    p.font.color.rgb = rgb(SCORE_COLORS[v]); p.font.bold = True
        return t

    # 1 title
    sl = prs.slides.add_slide(blank)
    bg = sl.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height); bg.fill.solid(); bg.fill.fore_color.rgb = navy; bg.line.fill.background()
    tb = sl.shapes.add_textbox(Inches(0.8), Inches(2.4), Inches(11.5), Inches(3)).text_frame
    tb.text = f'{res.hospital}'; tb.paragraphs[0].font.size = Pt(44); tb.paragraphs[0].font.bold = True
    tb.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    for txt, sz in ((f'Key Performance Indicators vs national BHSQA targets', 26), (period_text(res), 22),
                    (f'Prepared {dt.date.today():%d %B %Y}', 14)):
        p = tb.add_paragraph(); p.text = txt; p.font.size = Pt(sz); p.font.color.rgb = RGBColor(0xDD, 0xE6, 0xF0)

    # 2 executive summary
    sl = prs.slides.add_slide(blank); title(sl, 'Executive summary', 'Achievement of BHSQA KPIs reported by hospital units')
    cards = [(s['total'], 'BHSQA KPIs tracked', NAVY), (s['met'], 'Met target', STATUS_COLORS['Met']),
             (s['notmet'], 'Not met', STATUS_COLORS['Not met']), (s['nodata'], 'No data', STATUS_COLORS['No data']),
             (f"{s['pct']:.0f}%" if s['pct'] is not None else '—', 'of rated KPIs met', NAVY)]
    for i, (v, l, c) in enumerate(cards):
        box = sl.shapes.add_shape(1, Inches(0.4 + i * 2.5), Inches(1.5), Inches(2.3), Inches(1.4))
        box.fill.solid(); box.fill.fore_color.rgb = RGBColor(255, 255, 255); box.line.color.rgb = rgb(c); box.line.width = Pt(2)
        tf = box.text_frame; tf.text = str(v); tf.paragraphs[0].font.size = Pt(36); tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = rgb(c); tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        p = tf.add_paragraph(); p.text = l; p.font.size = Pt(12); p.font.color.rgb = RGBColor(0x44, 0x44, 0x44); p.alignment = PP_ALIGN.CENTER
    sl.shapes.add_picture(chart_coverage(res.alignment), Inches(0.4), Inches(3.3), width=Inches(7.4))
    worst = top_gaps(kpis, 6)
    tb = sl.shapes.add_textbox(Inches(8.1), Inches(3.3), Inches(4.9), Inches(3.5)).text_frame; tb.word_wrap = True
    tb.text = 'Largest gaps to target'; tb.paragraphs[0].font.bold = True; tb.paragraphs[0].font.size = Pt(14)
    tb.paragraphs[0].font.color.rgb = navy
    for _, k in worst.iterrows():
        p = tb.add_paragraph(); p.text = f'• {k.unit}: {k.kpi} – {fmt(k.achievement, k.uom)} (target {k.target})'; p.font.size = Pt(11)

    # national scorecard
    nv = bhsqa_values(kpis).merge(res.alignment[['BHSQA', 'KPI', 'Target']], on='BHSQA')
    rows = [[r.BHSQA, r.KPI[:62], r.Target, r.ValueText[:40], r.Status] for r in nv.itertuples()]
    per = 14
    for i in range(0, len(rows), per):
        sl = prs.slides.add_slide(blank)
        title(sl, f'National BHSQA KPI scorecard ({i // per + 1}/{(len(rows) + per - 1) // per})',
              'BHSQA KPI number, national description, target and hospital value')
        table(sl, [['BHSQA no.', 'National KPI description', 'BHSQA target', 'Hospital value', 'Status']] + rows[i:i + per],
              Inches(1.45), [1.2, 5.6, 2.6, 2.1, 1.0], status_col=4, font=10)

    # 3 headline indicators
    sl = prs.slides.add_slide(blank); title(sl, 'Headline hospital indicators', 'Monthly trend; red dashed line = BHSQA target limit, green band = target range')
    pic = chart_headlines(res)
    sl.shapes.add_picture(pic, Inches(0.9), Inches(1.4), height=Inches(5.55))

    # 4 status by unit
    sl = prs.slides.add_slide(blank); title(sl, 'Status by unit')
    sl.shapes.add_picture(chart_status_by_unit(kpis, 13, 6.3), Inches(0.4), Inches(1.05), width=Inches(12.5))

    # 5 not met
    nmk = kpis[(kpis.status == 'Not met') & (kpis.typ != 'survey')]
    rows = [[k.unit, k.bhsqa, k.kpi[:55], fmt(k.achievement, k.uom), k.target, str(k.months_with_data)] for _, k in nmk.iterrows()]
    for i in range(0, len(rows), 14):
        sl = prs.slides.add_slide(blank)
        title(sl, f'Clinical & managerial KPIs not meeting target ({len(rows)})' + (f' – {i // 14 + 1}' if len(rows) > 14 else ''),
              'Survey-based KPIs are shown on the next slide')
        table(sl, [['Unit', 'BHSQA', 'KPI', 'Achievement', 'Target', 'Months']] + rows[i:i + 14], Inches(1.45),
              [2.5, 1.0, 5.0, 1.4, 2.0, 0.8], font=10)
    sv = chart_surveys(kpis, 13, 5.6)
    if sv:
        sl = prs.slides.add_slide(blank)
        nsv = int(((kpis.status == 'Not met') & (kpis.typ == 'survey')).sum())
        title(sl, 'Patient satisfaction & 5S-CQI surveys', f'{nsv} survey-based KPIs below the >90% target')
        sl.shapes.add_picture(sv, Inches(0.4), Inches(1.45), width=Inches(12.5))

    # 6 alignment
    sl = prs.slides.add_slide(blank); title(sl, 'Alignment with national BHSQA KPIs')
    sl.shapes.add_picture(chart_coverage(res.alignment), Inches(0.4), Inches(1.2), width=Inches(7.5))
    nt = res.alignment[res.alignment.Coverage.isin(['Not tracked', 'Tracked – no data'])]
    tb = sl.shapes.add_textbox(Inches(8.1), Inches(1.2), Inches(4.9), Inches(5.6)).text_frame; tb.word_wrap = True
    tb.text = 'Not tracked / no data in period'; tb.paragraphs[0].font.bold = True; tb.paragraphs[0].font.size = Pt(14)
    for _, a in nt.iterrows():
        p = tb.add_paragraph(); p.text = f'{a.BHSQA}: {a.KPI} ({a.Coverage.lower()})'; p.font.size = Pt(10)

    # 7 data quality
    sl = prs.slides.add_slide(blank); title(sl, 'Data quality', f'{len(res.issues)} issues detected during extraction')
    top = res.issues.groupby('Issue').size().sort_values(ascending=False).head(12)
    tbl = [['Issue type', 'Count']] + [[i[:95], str(c)] for i, c in top.items()]
    table(sl, tbl, Inches(1.5), [10.5, 1.5], font=10)

    # appendix: unit tables
    for u in dict.fromkeys(kpis.unit):
        sub = kpis[kpis.unit == u]
        rows = [[k.bhsqa, f'{k.bhsqa_name[:60]}\n(unit KPI: {k.kpi[:50]})', fmt(k.achievement, k.uom), k.target, k.status,
                 str(k.months_with_data)] for _, k in sub.iterrows()]
        for i in range(0, len(rows), 8):
            sl = prs.slides.add_slide(blank); title(sl, f'{u}', 'Appendix – unit detail')
            table(sl, [['BHSQA no.', 'BHSQA KPI description', 'Value', 'Target', 'Status', 'Months']] + rows[i:i + 8], Inches(1.4),
                  [1.0, 5.5, 1.5, 2.6, 1.3, 0.9], status_col=4, font=11)
    b = io.BytesIO(); prs.save(b)
    return b.getvalue()
