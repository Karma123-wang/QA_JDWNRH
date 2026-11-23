"""
KPI extraction engine
---------------------
Reads hospital KPI reporting workbooks (one per fiscal year, July–June, in the
QAS/QASD template), extracts only the KPIs that align with the national
BHSQA KPI User Guide, and keeps an exact source reference (file, sheet, cell)
for every number so the dashboard can drill back to the "mother" data.
"""
from __future__ import annotations

import calendar
import datetime as dt
import io
import re
from dataclasses import dataclass, field

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- months
MON = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
       'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}


def month_label(key: str) -> str:
    y, m = key.split('-')
    return f"{calendar.month_abbr[int(m)]} {y}"


def days_in(key: str) -> int:
    y, m = map(int, key.split('-'))
    return calendar.monthrange(y, m)[1]


# --------------------------------------------------------------------------- books
@dataclass
class Book:
    name: str
    wb: openpyxl.Workbook
    fy_start: int
    hospital: str = ''
    _cols: dict = field(default_factory=dict)

    @property
    def fy_label(self):
        return f"FY {self.fy_start}-{str(self.fy_start + 1)[2:]}"

    def months(self):
        return [f"{self.fy_start}-{m:02d}" for m in range(7, 13)] + \
               [f"{self.fy_start + 1}-{m:02d}" for m in range(1, 7)]

    def sheet(self, aliases):
        al = [a.strip().lower() for a in aliases]
        for ws in self.wb.worksheets:
            if ws.title.strip().lower() in al:
                return ws
        return None

    def month_cols(self, ws):
        """calendar month key -> column index for a unit sheet"""
        if ws.title in self._cols:
            return self._cols[ws.title]
        out = {}
        for r in range(1, 7):
            vals = [str(ws.cell(r, c).value or '').strip().lower()[:3] for c in range(1, ws.max_column + 1)]
            if 'jul' in vals and 'jan' in vals:
                for i, v in enumerate(vals):
                    if v in MON:
                        m = MON[v]
                        key = f"{self.fy_start if m >= 7 else self.fy_start + 1}-{m:02d}"
                        out.setdefault(key, i + 1)
                break
        self._cols[ws.title] = out
        return out


def _detect_fy(wb, filename):
    for ws in wb.worksheets[:8]:
        for r in range(1, 4):
            nums = []
            for c in range(1, min(ws.max_column, 25) + 1):
                v = ws.cell(r, c).value
                try:
                    f = float(v)
                    if 2000 <= f <= 2100 and float(f).is_integer():
                        nums.append(int(f))
                except (TypeError, ValueError):
                    continue
            for a, b in zip(nums, nums[1:]):
                if b == a + 1:
                    return a
    m = re.search(r'(20\d\d)\D{0,6}(?:20)?(\d\d)', filename)
    if m:
        return int(m.group(1))
    raise ValueError(f"Could not detect the fiscal year of '{filename}'. "
                     "Put the years (e.g. 2026 2027) in row 1 of the unit sheets or in the file name.")


def _detect_hospital(wb):
    for ws in wb.worksheets[:6]:
        for r in range(1, 4):
            for c in range(1, 4):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip() and v.strip().isupper() and 3 <= len(v.strip()) <= 30:
                    return v.strip()
    return ''


def load_book(name: str, data: bytes) -> Book:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    return Book(name=name, wb=wb, fy_start=_detect_fy(wb, name), hospital=_detect_hospital(wb))


# --------------------------------------------------------------------------- units
UNITS = [  # display name, sheet-name aliases (any year)
    ('Medical Record', ['Medical record', 'Medical Record']),
    ('Gynae & Obstetrics IPD', ['Gyne', 'Gynae', 'Gynae IPD']),
    ('Medicine IPD', ['Medicine IPD']),
    ('Dermatology IPD', ['Dermatology IPD']),
    ('Psychiatry IPD', ['Psychiatry IPD']),
    ('Pediatric IPD', ['Pediatric IPD']),
    ('Surgery IPD', ['Surgery IPD']),
    ('Orthopedic IPD', ['Orthopedic IPD']),
    ('ENT IPD', ['ENT IPD']),
    ('Eye IPD', ['EYE IPD', 'Eye IPD']),
    ('Nutrition', ['Nutrition']),
    ('PICU', ['PICU']),
    ('AICU (MICU I & II)', ['AICU', 'AICU(MICU I & II)', 'AICU (MICU I & II)']),
    ('HDU (MICU III)', ['HDU', 'HDU (MICUIII)', 'HDU (MICU III)']),
    ('NICU', ['NICU']),
    ('CCU', ['CCU']),
    ('SICU', ['Surgical ICU', 'SICU']),
    ('Paediatric HDU', ['Ped. HDU', 'Ped HDU', 'Paediatric HDU']),
    ('Main OT', ['OT']),
    ('Eye OT', ['Eye OT']),
    ('Gynae OT', ['Gynae OT']),
    ('Anaesthesia', ['Anaesthesia']),
    ('Blood Bank', ['Blood Bank']),
    ('Lab & Pathology', ['Lab and Pathology']),
    ('Radiodiagnosis & Imaging', ['Radiodiagnosis and Imaging']),
    ('Pharmacy', ['Pharmacy']),
    ('CMS', ['CMS']),
    ('Nursing', ['Nursing']),
    ('IPC & AMSU', ['ICWM', 'IPC+AMSU', 'IPC']),
    ('AMSU', ['AMSU']),
    ('CHD', ['CHD']),
    ('BMES', ['BMED', 'BMES']),
    ('HRD', ['HRD']),
    ('QAS', ['QAD', 'QAS']),
    ('Planning', ['Planning']),
    ('Referral', ['Referral']),
    ('Procurement', ['Procurement']),
    ('HIMS', ['HIMD', 'HIMS']),
    ('Physiotherapy', ['Physiotherapy']),
    ('Dentistry', ['Dentistry']),
    ('Emergency', ['Emergency']),
    ('MCH Emergency', ['MCH Emergency Unit', 'MCH Emergency']),
]
ALIASES = dict(UNITS)


# --------------------------------------------------------------------------- locating rows
@dataclass
class Loc:
    unit: str
    kpi: str = ''           # KPI-name prefix (column B, carried down)
    off: int = 0             # row offset from the matched row
    metric: str | None = None  # metric prefix (column D)
    gen: str | None = None   # general-section keyword instead of kpi prefix


def _find(ws, loc: Loc):
    if ws is None:
        return None
    if loc.gen:
        in_gen = False
        for r in range(3, ws.max_row + 1):
            if 'general kpi' in str(ws.cell(r, 1).value or '').lower():
                in_gen = True
                continue
            b = str(ws.cell(r, 2).value or '').strip().lower()
            if in_gen and b and any(k in b for k in loc.gen.split('|')):
                return r + loc.off
        return None
    cur, k = None, loc.kpi.lower()
    for r in range(3, ws.max_row + 1):
        b = ws.cell(r, 2).value
        if b not in (None, ''):
            cur = str(b).strip().lower()
        d = str(ws.cell(r, 4).value or '').strip().lower()
        if cur and cur.startswith(k) and (loc.metric is None or d.startswith(loc.metric.lower())):
            return r + loc.off
    return None


def resolve(book: Book, locs):
    for loc in locs:
        ws = book.sheet(ALIASES.get(loc.unit, [loc.unit]))
        row = _find(ws, loc)
        if row:
            return ws, row
    return None, None


# --------------------------------------------------------------------------- value conversion
def conv(v):
    """returns (number or None, note or None)"""
    if v is None or v == '' or isinstance(v, bool):
        return None, None
    if isinstance(v, dt.time):
        return round(v.hour * 60 + v.minute + v.second / 60, 2), None
    if isinstance(v, dt.timedelta):
        return round(v.total_seconds() / 60, 2), None
    if isinstance(v, dt.datetime):
        return round((v - dt.datetime(1899, 12, 30)).total_seconds() / 60, 2), None
    if isinstance(v, (int, float)):
        return float(v), None
    s = str(v).strip()
    m = re.fullmatch(r'(\d+)[:;](\d+)[:;](\d+)', s)
    if m:
        return int(m[1]) * 60 + int(m[2]) + int(m[3]) / 60, f'"{s}" read as {m[1]}h {m[2]}m'
    m = re.fullmatch(r'([\d.]+)\s*hours?', s, re.I)
    if m:
        return float(m[1]) * 60, f'"{s}" read as {float(m[1]) * 60:g} min'
    m = re.fullmatch(r'([\d.]+)\s*%', s)
    if m:
        return float(m[1]), None
    m = re.fullmatch(r'([\d.]+)\s*days?', s, re.I)
    if m:
        return float(m[1]), None
    if s.upper() in ('-', '_', 'NA', 'NR', 'N/A'):
        return None, None
    return None, f'text "{s[:50]}" ignored'


# --------------------------------------------------------------------------- BHSQA catalogue
CAT = [
 ('KPI-1', 'Time for initial assessment of emergency patients', 'Red <15 min; Yellow 15–30 min; Green 30–60 min', 'Monthly', 'A. Clinical'),
 ('KPI-2', 'Time for initial assessment of an indoor patient', '≤45 minutes', 'Monthly', 'A. Clinical'),
 ('KPI-3', 'Care plan with desired outcome documented & countersigned (IPD)', '>90%*', 'Monthly', 'A. Clinical'),
 ('KPI-4', 'Screening for nutritional needs (IPD)', '>90%*', 'Monthly', 'A. Clinical'),
 ('KPI-5', 'Nursing care plan documented (IPD)', '>90%*', 'Monthly', 'A. Clinical'),
 ('KPI-6', 'Test result errors in clinical laboratory', '<0.6%', 'Monthly', 'A. Clinical'),
 ('KPI-7', 'Radiological examination errors', '<5%*', 'Monthly', 'A. Clinical'),
 ('KPI-8', 'Re-dos in clinical laboratory', '<5%*', 'Monthly', 'A. Clinical'),
 ('KPI-9', 'Re-dos in radiological services', '<5%*', 'Monthly', 'A. Clinical'),
 ('KPI-10A', 'Laboratory reports correlating with clinical diagnoses', '>90%', 'Monthly', 'A. Clinical'),
 ('KPI-10B', 'Radiological reports correlating with clinical diagnoses', '>90%', 'Monthly', 'A. Clinical'),
 ('KPI-11', 'Adherence to safety precautions – clinical laboratory', '>90%', 'Monthly', 'A. Clinical'),
 ('KPI-12', 'Adherence to safety precautions – radiology', '>90%', 'Monthly', 'A. Clinical'),
 ('KPI-13', 'Clinical laboratory test result TAT (% within TAT)', '>90%', 'Monthly', 'A. Clinical'),
 ('KPI-14', 'Radiological examination result TAT (% within TAT)', '>90%*', 'Monthly', 'A. Clinical'),
 ('KPI-15', 'Unavailability of laboratory reagents', '0 day', 'Biannually', 'A. Clinical'),
 ('KPI-16', 'Medication error rate (in-patient)', '0 per 1000 patient days', 'Monthly', 'A. Clinical'),
 ('KPI-17', 'Medication charts with error-prone abbreviations', '<8.4%', 'Monthly', 'A. Clinical'),
 ('KPI-18', 'In-patients who developed adverse drug reaction(s)', '<5%', 'Monthly', 'A. Clinical'),
 ('KPI-19', 'Medicine re-order levels in the medical store', '0%', 'Monthly', 'A. Clinical'),
 ('KPI-20', 'Compliance with DRA (BFDA) requirement', '>90%*', '—', 'A. Clinical'),
 ('KPI-21', 'Modification of anaesthesia plan', '<1%', 'Monthly', 'A. Clinical'),
 ('KPI-22', 'Unplanned ventilation following anaesthesia', '<1%', 'Monthly', 'A. Clinical'),
 ('KPI-23', 'Adverse anaesthesia events', '<1%', 'Monthly', 'A. Clinical'),
 ('KPI-24', 'Anaesthesia-related mortality rate', '<1%', 'Monthly', 'A. Clinical'),
 ('KPI-25', 'Re-scheduling of surgeries', '<10%', 'Monthly', 'A. Clinical'),
 ('KPI-26', 'Procedure to prevent wrong site/patient/surgery adhered to', '100%', 'Monthly', 'A. Clinical'),
 ('KPI-27', 'Appropriate prophylactic antibiotic within time frame', '100%', 'Monthly', 'A. Clinical'),
 ('KPI-28', 'Caesarean section rate', '10–15%', 'Monthly', 'A. Clinical'),
 ('KPI-29', 'Transfusion reactions', '0%', 'Monthly', 'A. Clinical'),
 ('KPI-30', 'Blood and blood components wasted', '<10.9%', 'Monthly', 'A. Clinical'),
 ('KPI-31', 'Blood component usage', '>90%', 'Monthly', 'A. Clinical'),
 ('KPI-32', 'TAT for issue of blood and blood components', 'Within 45 minutes', 'Monthly', 'A. Clinical'),
 ('KPI-33', 'Catheter-associated urinary tract infection (CAUTI)', '<7.5 per 1000 catheter-days', 'Monthly', 'A. Clinical'),
 ('KPI-34', 'Ventilator-associated pneumonia (VAP)', '<10 per 1000 ventilator days', 'Monthly', 'A. Clinical'),
 ('KPI-35', 'Central line-associated bloodstream infection (CLABSI)', '<5 per 1000 central-line days', 'Monthly', 'A. Clinical'),
 ('KPI-36', 'Surgical site infection (SSI)', '<5%', 'Monthly', 'A. Clinical'),
 ('KPI-37', 'Compliance with infection control & medical waste management', '>90%*', 'Biannually', 'A. Clinical'),
 ('KPI-38', 'Infant mortality rate (IMR)', '<23.85 per 1000 live births', '—', 'A. Clinical'),
 ('KPI-39', 'Maternal mortality ratio (MMR)', '<275 per 100,000 live births', 'Monthly', 'A. Clinical'),
 ('KPI-40', 'Return to ICU within 48 hours', '<11%', 'Monthly', 'A. Clinical'),
 ('KPI-41', 'Return to emergency within 72 h with similar complaints', '<5.8%', 'Monthly', 'A. Clinical'),
 ('KPI-42', 'Re-intubation rate in ICUs', '<12%', 'Monthly', 'A. Clinical'),
 ('KPI-43', 'Hand hygiene knowledge and practices', '>90%*', 'Biannually', 'A. Clinical'),
 ('KPI-44', 'Medicine prescription in capital letters and legible', '>90%*', '—', 'A. Clinical'),
 ('KPI-45', '5S-CQI implementation', '>90%*', 'Biannually', 'A. Clinical'),
 ('KPI-46', 'Incidences of patient fall in IPD', '≥10% reduction/year vs previous year', 'Monthly', 'B. Managerial'),
 ('KPI-47', 'Variations observed in mock drills', '–', '—', 'B. Managerial'),
 ('KPI-48', 'Incidence of bedsore after admission', '<0.54 per 1000 patient days', 'Monthly', 'B. Managerial'),
 ('KPI-49', 'Staff provided with complete PrEP', '100%*', '—', 'B. Managerial'),
 ('KPI-50', 'Average length of stay (ALOS)', '≤6 days', 'Monthly', 'B. Managerial'),
 ('KPI-51', 'Bed occupancy rate', '70–80%', 'Monthly', 'B. Managerial'),
 ('KPI-52', 'Critical equipment downtime', 'As per BMED requirement', 'Monthly', 'B. Managerial'),
 ('KPI-53(a)', 'Nurse-to-patient ratio – wards', '1:6', '—', 'B. Managerial'),
 ('KPI-53(b)', 'Nurse-to-patient ratio – ICUs', '1:1', '—', 'B. Managerial'),
 ('KPI-54', 'Outpatient satisfaction index', '>90%*', 'Biannually', 'B. Managerial'),
 ('KPI-55', 'Inpatient satisfaction index', '>90%*', 'Biannually', 'B. Managerial'),
 ('KPI-56', 'Patients discharged before 11 am', '>90%*', 'Monthly', 'B. Managerial'),
 ('KPI-57', 'OPD waiting time (consultation)', '<45 minutes', 'Biannually', 'B. Managerial'),
 ('KPI-58', 'Employee satisfaction index', '>90%*', '—', 'B. Managerial'),
 ('KPI-59', 'Employee attrition rate', '<10%', '—', 'B. Managerial'),
 ('KPI-60', 'Employee absenteeism rate', '0%', '—', 'B. Managerial'),
 ('KPI-61', 'Incidence of blood and body fluid exposure', '0', 'Monthly', 'B. Managerial'),
 ('KPI-62', 'Sentinel events reported, collected & analysed in time frame', '100%', '6-monthly', 'B. Managerial'),
 ('KPI-63', 'Incidence of needle-stick injury', '<2.18 per 1000 staff working days', 'Monthly', 'B. Managerial'),
 ('KPI-64', 'Medical records not having discharge summary', '0%', 'Monthly', 'B. Managerial'),
 ('KPI-65', 'Medical records not having ICD codification', '<10%', 'Monthly', 'B. Managerial'),
 ('KPI-66', 'Missing medical records', '<10%', 'Monthly', 'B. Managerial'),
 ('KPI-67', 'Follow-up of referrals', '>95%', 'Monthly', 'B. Managerial'),
]
CAT_NAME = {c[0]: c[1] for c in CAT}

# target rule per (sub)code: (text, dir, lo, hi). dir: < <= > >= range norate
T = {
 'KPI-1R': ('Red triage <15 min', '<', None, 15), 'KPI-1Y': ('Yellow triage ≤30 min', '<=', None, 30),
 'KPI-1G': ('Green triage ≤60 min', '<=', None, 60),
 'KPI-2': ('≤45 minutes', '<=', None, 45), 'KPI-3': ('>90%', '>', 90, None), 'KPI-4': ('>90%', '>', 90, None),
 'KPI-5': ('>90%', '>', 90, None), 'KPI-6': ('<0.6%', '<', None, 0.6), 'KPI-7': ('<5%', '<', None, 5),
 'KPI-8': ('<5%', '<', None, 5), 'KPI-9': ('<5%', '<', None, 5), 'KPI-10A': ('>90%', '>', 90, None),
 'KPI-10B': ('>90%', '>', 90, None), 'KPI-11': ('>90%', '>', 90, None), 'KPI-12': ('>90%', '>', 90, None),
 'KPI-15': ('0 days', '<=', None, 0), 'KPI-16': ('0 per 1000 patient days', '<=', None, 0),
 'KPI-17': ('<8.4%', '<', None, 8.4), 'KPI-18': ('<5%', '<', None, 5), 'KPI-19': ('0%', '<=', None, 0),
 'KPI-21': ('<1%', '<', None, 1), 'KPI-22': ('<1%', '<', None, 1), 'KPI-23': ('<1%', '<', None, 1),
 'KPI-24': ('<1%', '<', None, 1), 'KPI-25': ('<10%', '<', None, 10), 'KPI-26': ('100%', '>=', 100, None),
 'KPI-27': ('100%', '>=', 100, None), 'KPI-28': ('10–15%', 'range', 10, 15), 'KPI-29': ('0%', '<=', None, 0),
 'KPI-30': ('<10.9%', '<', None, 10.9), 'KPI-31': ('>90%', '>', 90, None), 'KPI-32': ('≤45 minutes', '<=', None, 45),
 'KPI-33': ('<7.5 per 1000 catheter-days', '<', None, 7.5), 'KPI-34': ('<10 per 1000 ventilator days', '<', None, 10),
 'KPI-35': ('<5 per 1000 central-line days', '<', None, 5), 'KPI-36': ('<5%', '<', None, 5),
 'KPI-37': ('>90%', '>', 90, None), 'KPI-39': ('<275 per 100,000 live births', '<', None, 275),
 'KPI-40': ('<11%', '<', None, 11), 'KPI-41': ('<5.8%', '<', None, 5.8), 'KPI-42': ('<12%', '<', None, 12),
 'KPI-43': ('>90%', '>', 90, None), 'KPI-45': ('>90%', '>', 90, None),
 'KPI-46': ('≥10% reduction/year vs previous year', 'norate', None, None),
 'KPI-47': ('No numeric target', 'norate', None, None), 'KPI-48': ('<0.54 per 1000 patient days', '<', None, 0.54),
 'KPI-49': ('100%', '>=', 100, None), 'KPI-50': ('≤6 days', '<=', None, 6), 'KPI-51': ('70–80%', 'range', 70, 80),
 'KPI-52': ('As per BMED requirement', 'norate', None, None), 'KPI-53(b)': ('1:1 (≤1 patient per nurse)', '<=', None, 1),
 'KPI-54': ('>90%', '>', 90, None), 'KPI-55': ('>90%', '>', 90, None), 'KPI-56': ('>90%', '>', 90, None),
 'KPI-57': ('<45 minutes', '<', None, 45), 'KPI-58': ('>90%', '>', 90, None), 'KPI-59': ('<10%', '<', None, 10),
 'KPI-60': ('0%', '<=', None, 0), 'KPI-61': ('0 exposures', '<=', None, 0), 'KPI-62': ('100%', '>=', 100, None),
 'KPI-63': ('<2.18 per 1000 staff working days', 'norate', None, None),
 'KPI-64': ('0%', '<=', None, 0), 'KPI-65': ('<10%', '<', None, 10), 'KPI-66': ('<10%', '<', None, 10),
 'KPI-67': ('>95%', '>', 95, None),
}


def bh(code):
    return code[:-1] if code in ('KPI-1R', 'KPI-1Y', 'KPI-1G') else code


# --------------------------------------------------------------------------- KPI specifications
@dataclass
class Spec:
    unit: str
    kpi: str
    code: str              # target code (KPI-1R etc.)
    jcode: str             # hospital's own indicator code
    typ: str               # ratio | bor | comp | nurse | count | survey
    num: list = field(default_factory=list)
    den: list = field(default_factory=list)
    extra: list = field(default_factory=list)   # [(label, [Loc])] for comp
    mult: float = 100
    uom: str = '%'
    note: str = ''
    frac: bool = False      # survey values stored as fractions (0.85) -> %
    sid: str = ''

    @property
    def bhsqa(self):
        return bh(self.code)


def _specs():
    S = []

    def R(unit, kpi, code, jcode, k, off=0, mult=100, uom='%', note='', doff=1, den=None, num=None, metric=None):
        S.append(Spec(unit, kpi, code, jcode, 'ratio', num or [Loc(unit, k, off, metric)],
                      den or [Loc(unit, k, off + doff, metric)], mult=mult, uom=uom, note=note))

    def SV(unit, kpi, code, jcode, locs, uom='%', note='', frac=True):
        S.append(Spec(unit, kpi, code, jcode, 'survey', locs, uom=uom, note=note, frac=frac))

    def CT(unit, kpi, code, jcode, locs, uom='count', note=''):
        S.append(Spec(unit, kpi, code, jcode, 'count', locs, uom=uom, note=note))

    BI = 'Bi-annual survey: value sits in Jan (Jan–Jun round) or Jul (Jul–Dec round) column'

    def general(unit, waiting=True, ipd=True, opd=True, blood=False):
        if opd:
            SV(unit, 'OPD satisfaction index (survey)', 'KPI-54', '51', [Loc(unit, gen='opd satis|opd sats')], note=BI)
        if ipd:
            SV(unit, 'IPD satisfaction index (survey)', 'KPI-55', '52', [Loc(unit, gen='ipd satis')], note=BI)
        SV(unit, '5S-CQI implementation (survey)', 'KPI-45', '–', [Loc(unit, gen='5s')], note=BI)
        if waiting:
            SV(unit, 'OPD waiting time – consultation (survey)', 'KPI-57', '53c',
               [Loc(unit, gen='waiting time for opd')], uom='minutes', frac=False, note=BI)
        if blood:
            SV(unit, 'Urgent blood transfusion TAT (survey)', 'KPI-32', '26', [Loc(unit, gen='urgent blood|stat blood')],
               uom='minutes', frac=False, note='Quarterly survey value (minutes)')

    MR = 'Medical Record'
    DD = [Loc(MR, 'Mortality rate', 1)]           # discharges + deaths
    IPDAYS = [Loc(MR, 'Average length of stay')]  # inpatient days
    SURG = [Loc(MR, 'Health Facility Surgery')]

    R(MR, 'Average length of stay (ALOS)', 'KPI-50', '47a', 'Average length of stay', mult=1, uom='days')
    S.append(Spec(MR, 'Bed occupancy rate (BOR)', 'KPI-51', '47b', 'bor', [Loc(MR, 'Bed occupancy rate')],
                  [Loc(MR, 'Health Facility Bed')], note='Denominator = bed strength × days in month'))
    R(MR, 'Records without discharge summary', 'KPI-64', '63', '% of medica records not', den=DD,
      note='Denominator = discharges + deaths (Mortality rate row)')
    R(MR, 'Records without ICD codification', 'KPI-65', '64', '% of medical records not having cod', den=DD,
      note='Denominator = discharges + deaths (Mortality rate row)')
    R(MR, 'Missing medical records', 'KPI-66', '66', '% of missing records')
    R(MR, 'Follow-up of referrals (in-country)', 'KPI-67', '73a', 'Follow up of referrals')

    def ipd(unit, surg=False, dc=True, dcoff=1):
        R(unit, 'Time for initial assessment of indoor patient', 'KPI-2', '1a', 'Time for initial assessment',
          mult=1, uom='minutes (avg)', note='Sum of minutes ÷ number of patients')
        R(unit, 'Care plan documented & countersigned', 'KPI-3', '2', 'Percentage of cases (in-patient)')
        if surg:
            R(unit, 'Re-scheduling of surgeries', 'KPI-25', '18', '% of rescheduling')
        if dc:
            R(unit, 'Patients discharged before 11 AM', 'KPI-56', '74', 'Patient discharged before', doff=dcoff)

    G = 'Gynae & Obstetrics IPD'
    ipd(G, surg=True, dc=False)
    R(G, 'Maternal mortality ratio (MMR)', 'KPI-39', '39', 'Maternal mortality', mult=100000, uom='per 100,000 live births')
    R(G, 'Caesarean section rate', 'KPI-28', '72', 'Caesarean rate')
    general(G)
    for u in ('Medicine IPD', 'Dermatology IPD', 'Pediatric IPD'):
        ipd(u); general(u)
    ipd('Psychiatry IPD', dcoff=2); general('Psychiatry IPD')
    for u in ('Surgery IPD', 'Orthopedic IPD', 'ENT IPD', 'Eye IPD'):
        ipd(u, surg=True); general(u, blood=(u == 'Surgery IPD'))
    N = 'Nutrition'
    for i, w in enumerate(['Medical ward', 'AICU', 'Paediatric ward']):
        R(N, f'Nutritional screening – {w}', 'KPI-4', '3', '% of cases wherein screening', off=2 * i)
    general(N, waiting=False, ipd=False, opd=False)
    for u in ('PICU', 'AICU (MICU I & II)', 'HDU (MICU III)', 'NICU', 'CCU', 'SICU', 'Paediatric HDU'):
        R(u, 'Return to ICU within 48 hours', 'KPI-40', '32', 'Return to ICU')
        R(u, 'Re-intubation rate (within 48 h)', 'KPI-42', '34', 'Reintubation')
        general(u, waiting=False, opd=False, blood=(u.startswith('AICU')))
    for u in ('Main OT', 'Eye OT', 'Gynae OT'):
        R(u, 'Safe-surgery procedure adhered (wrong site/patient/surgery)', 'KPI-26', '19', "% of cases where the hospital")
        general(u, waiting=False, ipd=False, opd=False)
    A = 'Anaesthesia'
    R(A, 'Modification of anaesthesia plan', 'KPI-21', '13', '% of modification')
    R(A, 'Unplanned ventilation following anaesthesia', 'KPI-22', '14', '% of unplanned ventilation')
    R(A, 'Adverse anaesthesia events', 'KPI-23', '15', '% of adverse anaesthesia')
    R(A, 'Anaesthesia-related mortality rate', 'KPI-24', '16', 'Anaesthesia related mortality')
    general(A, waiting=False, ipd=False, opd=False)
    B = 'Blood Bank'
    R(B, 'Transfusion reactions', 'KPI-29', '23', '% of transfusion reactions')
    R(B, 'Blood & blood components wasted', 'KPI-30', '24a', '% of blood and blood products wasted')
    S.append(Spec(B, 'Blood component usage', 'KPI-31', '25', 'comp', [],
                  [Loc(B, '% of blood component usage (WB)', 1)],
                  extra=[('PRC', [Loc(B, '% of blood component usage (PRC)')]),
                         ('PC', [Loc(B, '% of blood component usage (PC)')]),
                         ('FFP', [Loc(B, '% of blood component usage (FFP)')])],
                  note='BHSQA: component units (PRC+PC+FFP) ÷ all units transfused (incl. whole blood)'))
    general(B, waiting=False, ipd=False, opd=False, blood=True)
    LB = 'Lab & Pathology'
    R(LB, 'Test result (reporting) errors', 'KPI-6', '5b', 'No. of repor')
    R(LB, 'Re-dos in laboratory', 'KPI-8', '6b', 'Rate of re-dos')
    R(LB, 'Lab reports correlating with clinical diagnosis', 'KPI-10A', '7b', '% of reports correlating')
    R(LB, 'Adherence to safety precautions (lab)', 'KPI-11', '8b', '% of adherence to safety')
    CT(LB, 'Unavailability of laboratory reagents', 'KPI-15', '38b', [Loc(LB, 'Unavailability of consumables')], uom='days')
    general(LB, waiting=False, ipd=False, opd=False)
    RD = 'Radiodiagnosis & Imaging'
    R(RD, 'Radiological examination errors', 'KPI-7', '5a', 'Number of errors reported')
    R(RD, 'Re-dos in radiology', 'KPI-9', '6a', 'Rate of re-dos')
    R(RD, 'Radiology reports correlating with clinical diagnosis', 'KPI-10B', '7a', '% of reports correlating')
    R(RD, 'Adherence to safety precautions (radiology)', 'KPI-12', '8a', '% of adherence to safety')
    general(RD, waiting=False, ipd=False)
    PH = 'Pharmacy'
    R(PH, 'Medication error rate', 'KPI-16', '9', 'Incidence of medication error', mult=1000, uom='per 1000 patient days')
    R(PH, 'Adverse drug reactions (in-patients)', 'KPI-18', '10', '% of admissions with adverse')
    R(PH, 'Medication charts with error-prone abbreviations', 'KPI-17', '11', '% of medication charts')
    general(PH, waiting=False, ipd=False)
    R('CMS', 'Medicines below re-order level', 'KPI-19', '35', 'Drugs reorder levels')
    general('CMS', waiting=False, ipd=False, opd=False)
    NU = 'Nursing'
    R(NU, 'Nursing care plan documented', 'KPI-5', '4', '% of cases wherein Nursing')
    R(NU, 'Patient falls in IPD', 'KPI-46', '44', 'Incidence of falls', den=IPDAYS, mult=1000, uom='per 1000 patient days',
      note='Denominator = hospital inpatient days (Medical Record)')
    R(NU, 'Bedsore (pressure ulcer) after admission', 'KPI-48', '45', 'Incidence of hospital associated', den=IPDAYS,
      mult=1000, uom='per 1000 patient days', note='Denominator = hospital inpatient days (Medical Record)')
    S.append(Spec(NU, 'Nurse-to-patient ratio – ICU', 'KPI-53(b)', '50', 'nurse', [Loc(NU, 'Nurse- patient ratio')],
                  [Loc(NU, 'Nurse- patient ratio', 1)], uom='patients per nurse', note='Result = beds ÷ nursing staff'))
    general(NU, waiting=False, ipd=False, opd=False)
    IP = 'IPC & AMSU'
    R(IP, 'CAUTI', 'KPI-33', '27', 'Catheter associated', mult=1000, uom='per 1000 catheter-days')
    R(IP, 'VAP', 'KPI-34', '28', 'Pneumonia rate', mult=1000, uom='per 1000 ventilator days')
    R(IP, 'CLABSI', 'KPI-35', '29', 'Central line associated', mult=1000, uom='per 1000 central-line days')
    R(IP, 'Surgical site infection (SSI)', 'KPI-36', '30', 'Surgical site', den=SURG,
      note='Denominator = surgeries performed (Medical Record)')
    R(IP, 'Prophylactic antibiotic within time frame', 'KPI-27', '20', '',
      num=[Loc(IP, '% of cases who'), Loc('AMSU', '% of cases who')],
      den=[Loc(IP, '% of cases who', 1), Loc('AMSU', '% of cases who', 1)])
    R(IP, 'Staff with complete PrEP', 'KPI-49', '46', '',
      num=[Loc(IP, '% of staff provided'), Loc('CHD', '% of staff provided')],
      den=[Loc(IP, '% of staff provided', 1), Loc('CHD', '% of staff provided', 1)])
    R(IP, 'Blood & body-fluid exposure (IPD)', 'KPI-61', '61', 'Incidence of blood body fluid exposures (IPD)', den=IPDAYS,
      mult=1000, uom='per 1000 patient days', note='Target 0 exposures; BHSQA denominator is staff working days')
    R(IP, 'Blood & body-fluid exposure (OPD)', 'KPI-61', '61', 'Incidence of blood body fluid exposures (OPD)',
      mult=1000, uom='per 1000 OPD visits', note='Target 0 exposures')
    R(IP, 'Needle-stick injuries', 'KPI-63', '62', 'Incidences of needle', mult=1000, uom='per 1000 OPD visits',
      note='Sheet uses OPD visits; BHSQA uses staff working days → not rated')
    R(IP, 'Hand hygiene compliance', 'KPI-43', '69', 'Compliance to hand hygiene')
    R(IP, 'IC & MWM clinical audit score', 'KPI-37', '70', 'ICWM clinical audit')
    general(IP, waiting=False, ipd=False, opd=False)
    CT('BMES', 'Critical equipment downtime', 'KPI-52', '49', [Loc('BMES', 'Critical equipment downtime')], uom='hours')
    general('BMES', waiting=False, ipd=False, opd=False)
    H = 'HRD'
    R(H, 'Employee satisfaction index', 'KPI-58', '55', 'Employee satisfaction')
    R(H, 'Employee attrition rate (monthly)', 'KPI-59', '56', 'Employee attrition', note='BHSQA <10% is annual')
    R(H, 'Employee absenteeism (unauthorised)', 'KPI-60', '57', 'Staff absenteeism')
    general(H, waiting=False, ipd=False, opd=False)
    Q = 'QAS'
    SV(Q, 'Hospital OPD satisfaction index', 'KPI-54', '51', [Loc(Q, 'OPD satisfaction index')], note=BI)
    SV(Q, 'Hospital IPD satisfaction index', 'KPI-55', '52', [Loc(Q, 'IPD satisfaction index')], note=BI)
    SV(Q, 'Hospital OPD waiting time – consultation', 'KPI-57', '53c', [Loc(Q, 'Waiting time for OPD')], uom='minutes', frac=False)
    R(Q, 'Sentinel events analysed within time frame', 'KPI-62', '59', 'Number of sentinel')
    CT('Planning', 'Variations observed in mock drills', 'KPI-47', '43',
       [Loc('Planning', 'No. of variations observed'), Loc(Q, 'No. of variations observed')])
    R('Referral', 'Follow-up of referrals (ex-country)', 'KPI-67', '73b', 'Follow up of referrals')
    general('Referral', waiting=False, ipd=False, opd=False)
    general('Procurement', waiting=False, ipd=False, opd=False)
    general('HIMS', waiting=False, ipd=False, opd=False)
    general('CHD', waiting=False, ipd=False, opd=False)
    general('Planning', waiting=False, ipd=False, opd=False)
    general('Physiotherapy', ipd=False)
    D = 'Dentistry'
    R(D, 'Time for initial assessment of indoor patient', 'KPI-2', '1a', 'Time for initial', mult=1, uom='minutes (avg)')
    R(D, 'Re-scheduling of surgeries', 'KPI-25', '18', '% of rescheduling')
    general(D, ipd=False)
    for u in ('Emergency', 'MCH Emergency'):
        R(u, 'Initial assessment time – RED triage', 'KPI-1R', '1b', 'RED', mult=1, uom='minutes (avg)')
        R(u, 'Initial assessment time – YELLOW triage', 'KPI-1Y', '1b', 'YELLOW', mult=1, uom='minutes (avg)')
        R(u, 'Initial assessment time – GREEN triage', 'KPI-1G', '1b', 'GREEN', mult=1, uom='minutes (avg)')
        R(u, 'Return to emergency within 72 h', 'KPI-41', '33', 'Return to emergency')
        general(u, ipd=False, waiting=False, blood=True)
    for i, s in enumerate(S):
        s.sid = f"K{i + 1:03d}"
    return S


SPECS = _specs()

# headline KPIs for the executive page: (unit, code)
HEADLINE = [('Medical Record', 'KPI-50'), ('Medical Record', 'KPI-51'), ('Gynae & Obstetrics IPD', 'KPI-28'),
            ('Blood Bank', 'KPI-30'), ('IPC & AMSU', 'KPI-33'), ('IPC & AMSU', 'KPI-34'),
            ('IPC & AMSU', 'KPI-35'), ('IPC & AMSU', 'KPI-36'), ('IPC & AMSU', 'KPI-43'),
            ('QAS', 'KPI-55'), ('QAS', 'KPI-54'), ('Main OT', 'KPI-26')]


# --------------------------------------------------------------------------- extraction
@dataclass
class Cell:
    file: str
    sheet: str
    cell: str
    row: int
    col: int
    label: str
    raw: object
    value: float | None
    note: str | None = None

    def ref(self):
        return f"{self.file} › {self.sheet} › {self.cell}"


@dataclass
class Result:
    hospital: str
    months: list          # month keys in period
    kpis: pd.DataFrame    # one row per KPI (sid)
    monthly: pd.DataFrame  # sid × month
    cells: dict           # (sid, month, role) -> Cell or list[Cell]
    issues: pd.DataFrame
    alignment: pd.DataFrame
    books: dict           # file name -> Book


def _cell(book, ws, row, key):
    col = book.month_cols(ws).get(key)
    if not col:
        return None
    raw = ws.cell(row, col).value
    val, note = conv(raw)
    lab = str(ws.cell(row, 4).value or ws.cell(row, 2).value or '').strip()
    return Cell(book.name, ws.title.strip(), f"{get_column_letter(col)}{row}", row, col, lab, raw, val, note)


def _status(val, d, lo, hi):
    if d == 'info':
        return 'No BHSQA target'
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 'No data'
    if d == 'norate':
        return 'Not rated'
    ok = {'<': lambda: val < hi, '<=': lambda: val <= hi, '>': lambda: val > lo, '>=': lambda: val >= lo,
          'range': lambda: lo <= val <= hi}[d]()
    return 'Met' if ok else 'Not met'


def _gap(val, d, lo, hi):
    if val is None or d in ('info', 'norate') or pd.isna(val):
        return None
    if d in ('<', '<='):
        return val - hi
    if d in ('>', '>='):
        return val - lo
    return val - lo if val < lo else (val - hi if val > hi else 0.0)


def available_months(books):
    ks = set()
    for b in books:
        ks.update(b.months())
    return sorted(ks)


def extract(books: list[Book], start: str, end: str) -> Result:
    by_month = {}
    for b in sorted(books, key=lambda x: x.fy_start):
        for k in b.months():
            by_month.setdefault(k, b)
    months = [k for k in sorted(by_month) if start <= k <= end]
    cells, mrows, krows, issues = {}, [], [], []
    hospital = next((b.hospital for b in books if b.hospital), '')

    for s in SPECS:
        nums, dens = [], []
        for k in months:
            book = by_month[k]
            n = d = None
            if s.typ == 'comp':
                parts = []
                for lab, locs in s.extra:
                    ws, row = resolve(book, locs)
                    c = _cell(book, ws, row, k) if ws else None
                    if c:
                        parts.append(c)
                vals = [p.value for p in parts if p.value is not None]
                n = sum(vals) if vals else None
                cells[(s.sid, k, 'num')] = parts
            elif s.num:
                ws, row = resolve(book, s.num)
                c = _cell(book, ws, row, k) if ws else None
                if c:
                    if s.frac and c.value is not None and 0 < c.value <= 1:
                        c.value = round(c.value * 100, 2)
                    cells[(s.sid, k, 'num')] = c
                    n = c.value
                    if c.note:
                        issues.append((s.unit, s.kpi, month_label(k), c.ref(), c.note))
            if s.typ in ('ratio', 'comp', 'nurse', 'bor') and s.den:
                ws, row = resolve(book, s.den)
                c = _cell(book, ws, row, k) if ws else None
                if c:
                    cells[(s.sid, k, 'den')] = c
                    d = c.value
                    if c.note:
                        issues.append((s.unit, s.kpi, month_label(k), c.ref(), c.note))
            if s.typ == 'bor':
                cells[(s.sid, k, 'beds')] = cells.get((s.sid, k, 'den'))
                d = d * days_in(k) if d is not None else None
            nums.append(n); dens.append(d)

        if s.typ == 'bor':  # carry bed strength when blank
            beds = [cells.get((s.sid, k, 'beds')) for k in months]
            known = [(i, b.value) for i, b in enumerate(beds) if b and b.value is not None]
            for i, k in enumerate(months):
                if dens[i] is None and known:
                    j, v = min(known, key=lambda t: abs(t[0] - i))
                    dens[i] = v * days_in(k)
                    issues.append((s.unit, s.kpi, month_label(k), 'Medical Record › bed strength',
                                   f'Bed strength blank – {v:g} beds carried from {month_label(months[j])}'))

        res = []
        for i, k in enumerate(months):
            n, d = nums[i], dens[i]
            if s.typ in ('ratio', 'comp', 'bor'):
                r = n / d * s.mult if (n is not None and d not in (None, 0)) else None
                if n is not None and d is None:
                    issues.append((s.unit, s.kpi, month_label(k), '', 'Numerator recorded but denominator blank'))
                if d == 0 and n is not None:
                    issues.append((s.unit, s.kpi, month_label(k), '', 'Denominator recorded as 0'))
                if r is not None and s.uom == '%' and r > 100.5 and s.code not in ('KPI-51',):
                    issues.append((s.unit, s.kpi, month_label(k), '', f'Result {r:.1f}% exceeds 100% – numerator larger than denominator'))
            elif s.typ == 'nurse':
                r = d / n if (n not in (None, 0) and d is not None) else None
            else:
                r = n
            res.append(r)
            mrows.append(dict(sid=s.sid, month=k, label=month_label(k), num=n, den=d, result=r))

        # achievement
        if s.typ in ('ratio', 'comp', 'bor'):
            pairs = [(n, d) for n, d in zip(nums, dens) if n is not None and d is not None]
            sn, sd = sum(p[0] for p in pairs), sum(p[1] for p in pairs)
            ach = sn / sd * s.mult if pairs and sd else None
        elif s.typ == 'nurse':
            pairs = [(n, d) for n, d in zip(nums, dens) if n is not None and d is not None]
            sn, sd = sum(p[0] for p in pairs), sum(p[1] for p in pairs)
            ach = sd / sn if pairs and sn else None
        elif s.typ == 'count':
            v = [x for x in nums if x is not None]
            sn, sd = (sum(v) if v else None), None
            ach = sn
        else:
            v = [x for x in nums if x is not None]
            sn, sd = None, None
            ach = sum(v) / len(v) if v else None
        n_months = sum(1 for x in res if x is not None)

        tt, d_, lo, hi = T.get(s.code, ('No BHSQA target', 'info', None, None))
        note = s.note
        if s.code == 'KPI-2' and ach is not None and ach < 2:
            d_ = 'norate'
            note = 'Not rated: recorded "sum of time" gives an average under 2 minutes, which is not plausible.'
            issues.append((s.unit, s.kpi, '', '', 'Initial assessment "sum of time" too small to be a total – not rated'))
        krows.append(dict(sid=s.sid, unit=s.unit, kpi=s.kpi, bhsqa=s.bhsqa, bhsqa_name=CAT_NAME.get(s.bhsqa, ''),
                          jcode=s.jcode, typ=s.typ, uom=s.uom, total_num=sn, total_den=sd, achievement=ach,
                          months_with_data=n_months, target=tt, dir=d_, lo=lo, hi=hi,
                          status=_status(ach, d_, lo, hi), gap=_gap(ach, d_, lo, hi), note=note, mult=s.mult))

    kpis = pd.DataFrame(krows)
    monthly = pd.DataFrame(mrows)
    iss = pd.DataFrame(issues, columns=['Unit', 'KPI', 'Month', 'Source', 'Issue']).drop_duplicates()

    al = []
    for code, name, tgt, fr, grp in CAT:
        sub = kpis[kpis.bhsqa == code]
        units = ', '.join(dict.fromkeys(sub.unit))
        withd = int((sub.months_with_data > 0).sum())
        cov = ('Not tracked' if sub.empty else 'Tracked – no data' if withd == 0
               else 'Partly reported' if withd < len(sub) else 'Reported')
        al.append(dict(BHSQA=code, KPI=name, Group=grp, Target=tgt, Frequency=fr, Units=units or '—',
                       Rows=len(sub), WithData=withd, Met=int((sub.status == 'Met').sum()),
                       NotMet=int((sub.status == 'Not met').sum()), Coverage=cov))
    al = pd.DataFrame(al).merge(bhsqa_values(kpis), on='BHSQA', how='left')
    return Result(hospital, months, kpis, monthly, cells, iss, al, {b.name: b for b in books})


# --------------------------------------------------------------------------- national KPI values
def fmt_value(v, uom=''):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return '—'
    s = f'{v:,.0f}' if abs(v) >= 1000 else (f'{v:.1f}' if abs(v) >= 10 else f'{v:.2f}')
    if uom == '%':
        return s + '%'
    short = {'minutes (avg)': ' min', 'minutes': ' min', 'days': ' days', 'count': '', 'hours': ' h'}
    return s + short.get(uom, ' ' + uom if uom else '')


def _pool(rows):
    """combine unit rows of one BHSQA KPI into one hospital value"""
    if rows.empty or rows.uom.nunique() != 1 or rows.typ.nunique() != 1:
        return None
    typ = rows.typ.iloc[0]
    if typ in ('ratio', 'comp', 'bor'):
        den = rows.total_den.sum()
        return rows.total_num.sum() / den * rows.mult.iloc[0] if den else None
    if typ == 'nurse':
        num = rows.total_num.sum()
        return rows.total_den.sum() / num if num else None
    if typ == 'count':
        return rows.achievement.sum()
    return rows.achievement.mean()


NO_POOL = {'KPI-61', 'KPI-67'}  # different measures per unit – shown unit by unit


def bhsqa_values(kpis: pd.DataFrame) -> pd.DataFrame:
    """one row per national BHSQA KPI: hospital value, status and the value of every unit"""
    out = []
    for code, name, tgt, fr, grp in CAT:
        sub = kpis[kpis.bhsqa == code]
        have = sub[sub.achievement.notna()]
        rated = have[have.dir != 'norate']
        by_unit = [(r.unit + (' – ' + r.kpi.split('–')[-1].strip() if '–' in r.kpi and code in ('KPI-1', 'KPI-4') else ''),
                    fmt_value(r.achievement, r.uom), r.status) for r in sub.itertuples()]
        val, uom, basis, status = None, '', '', None
        if sub.empty:
            status, basis = 'Not tracked', 'Not in hospital sheets'
        elif code == 'KPI-1':
            parts, sts = [], []
            for sc, lab in (('KPI-1R', 'Red'), ('KPI-1Y', 'Yellow'), ('KPI-1G', 'Green')):
                r = kpis[(kpis.bhsqa == 'KPI-1') & kpis.kpi.str.contains(lab.upper()) & kpis.achievement.notna()]
                v = _pool(r)
                if v is not None:
                    parts.append(f'{lab} {v:.1f}')
                    t = T[sc]
                    sts.append(_status(v, t[1], t[2], t[3]))
            basis = f'Pooled across {len(set(have.unit))} unit(s)'
            uom = 'minutes (avg)'
            if parts:
                val = None
                status = 'Met' if all(x == 'Met' for x in sts) else 'Not met'
                out.append(dict(BHSQA=code, Value=None, ValueText=' · '.join(parts) + ' min', UoM=uom, Basis=basis,
                                Status=status, ByUnit=by_unit))
                continue
        else:
            q = have[have.unit == 'QAS']
            if code in ('KPI-54', 'KPI-55', 'KPI-57') and not q.empty:
                val, basis, uom = q.achievement.iloc[0], 'Hospital-wide survey (QAS)', q.uom.iloc[0]
            elif not rated.empty:
                val = None if code in NO_POOL else _pool(rated)
                uom = rated.uom.iloc[0] if rated.uom.nunique() == 1 else ''
                n = len(set(rated.unit))
                if val is not None:
                    basis = rated.unit.iloc[0] if n == 1 else (
                        f'Average of {n} units' if rated.typ.iloc[0] == 'survey' else f'Pooled across {n} units')
                else:
                    basis = 'Different units of measure – see unit values'
            elif not have.empty:
                uom, basis = have.uom.iloc[0], 'Not rated – see notes'
                val = _pool(have)
        if status is None:
            t = T.get(code)
            if val is not None and t:
                status = _status(val, t[1], t[2], t[3])
            elif not rated.empty:  # values that cannot be pooled
                st_ = set(rated.status)
                status = 'Not met' if 'Not met' in st_ else ('Met' if st_ == {'Met'} else 'No data')
            elif not have.empty:
                status = 'Not rated'
            else:
                status = 'No data'
        vt = fmt_value(val, uom)
        if val is None and not rated.empty:
            vt = ' · '.join(f'{x.kpi.split("(")[-1].rstrip(")") if "(" in x.kpi else x.unit}: {fmt_value(x.achievement, x.uom)}' for x in rated.itertuples())
            basis = 'Shown separately per unit'
        out.append(dict(BHSQA=code, Value=val, ValueText=vt, UoM=uom, Basis=basis,
                        Status=status, ByUnit=by_unit))
    return pd.DataFrame(out)


def by_unit_text(items, with_status=True, limit=None):
    items = items if limit is None else items[:limit]
    return '; '.join(f'{u}: {v}' + (f' ({s})' if with_status and s not in ('Met',) else '') for u, v, s in items)


# --------------------------------------------------------------------------- source helpers
def source_trace(res: Result, sid: str) -> pd.DataFrame:
    """one row per month and role with file / sheet / cell / raw value"""
    rows = []
    for k in res.months:
        for role, lab in (('num', 'Numerator'), ('den', 'Denominator')):
            c = res.cells.get((sid, k, role))
            if c is None:
                continue
            for x in (c if isinstance(c, list) else [c]):
                rows.append({'Month': month_label(k), 'Role': lab if not isinstance(c, list) else f'{lab} part',
                             'File': x.file, 'Sheet': x.sheet, 'Cell': x.cell, 'Metric in sheet': x.label,
                             'Value in sheet': '' if x.raw is None else str(x.raw), 'Value used': x.value})
    return pd.DataFrame(rows)


def source_rows(res: Result, sid: str):
    """{(file, sheet): {'rows': set, 'cells': set((row,col))}}"""
    out = {}
    for (s, k, role), c in res.cells.items():
        if s != sid or c is None:
            continue
        for x in (c if isinstance(c, list) else [c]):
            e = out.setdefault((x.file, x.sheet), {'rows': set(), 'cells': set()})
            e['rows'].add(x.row); e['cells'].add((x.row, x.col))
    return out


def sheet_frame(book: Book, sheet: str, rows=None, pad=2):
    """original sheet as a DataFrame (rows numbered as in Excel)"""
    ws = next(w for w in book.wb.worksheets if w.title.strip() == sheet)
    maxc = min(ws.max_column, 26)
    if rows:
        r0, r1 = max(1, min(rows) - pad), min(ws.max_row, max(rows) + pad)
        # always include header rows
        rr = sorted(set(range(1, 4)) | set(range(r0, r1 + 1)))
    else:
        rr = list(range(1, min(ws.max_row, 60) + 1))
    data = []
    for r in rr:
        row = []
        for c in range(1, maxc + 1):
            v = ws.cell(r, c).value
            if isinstance(v, (dt.time, dt.datetime, dt.timedelta)):
                v = str(v)
            row.append('' if v is None else v)
        data.append(row)
    df = pd.DataFrame(data, index=rr, columns=[get_column_letter(c) for c in range(1, maxc + 1)])
    df.index.name = 'Row'
    return df.astype(str).replace({'None': ''})
