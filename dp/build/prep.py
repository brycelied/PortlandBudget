"""Build DP-level dataset for the Fall TAO dashboard mockup.

Usage: python3 -I prep.py <dp_export.xlsx> <bureau_sa.csv> <out.json>
"""
import sys, json, re, hashlib
import pandas as pd

src, sa_csv, out = sys.argv[1:4]
AMT = '2026-27 FALL Requested Total'
NONPROG = {'Contingency', 'Reserved for Future Expenditure', 'Ending Fund Balance', 'Debt Service Reserves'}

fin = pd.read_excel(src, sheet_name='Financial Data Export')
per = pd.read_excel(src, sheet_name='Personnel Export')
sa = pd.read_csv(sa_csv).set_index('Bureau')

def clean_sa(s):
    return re.sub(r'^[A-Z]{2}\d\s*-\s*', '', s).strip()

def txt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()

fin['bcode'] = fin['Bureau - Name'].str[:2]
fin['isgf'] = fin['Fund - Name'].str.startswith('100000')
fte = per.groupby('Form ID')['FTE'].sum()
gfte = per[per['Fund - Name'].astype(str).str.startswith('100000')].groupby('Form ID')['FTE'].sum()

NO_SI = re.compile(r'^(none|n/?a|na|no\b|there (is|are|will be) no|this (request|package|dp|decision package) (has|will have) no|not applicable|-|no impact)', re.I)

def fund_label(v):
    v = txt(v)
    m = re.match(r'^(\d+)\s*-\s*(.*)$', v)
    return f"{m.group(2).strip()} ({m.group(1)})" if m else v

def h(x):  # deterministic 0..1 from DP id
    return int(hashlib.md5(str(x).encode()).hexdigest()[:8], 16) / 0xFFFFFFFF

dps = []
for fid, g in fin.groupby('Form Header ID'):
    first = g.iloc[0]
    exp = g[g['Account Name'] == 'Expense']
    rev = g[g['Account Name'] == 'Revenue']
    prog = exp[~exp['Major Object Name'].isin(NONPROG)]
    gexp, gprog = exp[exp.isgf], prog[prog.isgf]
    si = txt(first['Service Impacts'])
    name = re.sub(r'^[\d,]+\s*[–-]\s*', '', txt(first['DP - Name']))
    rtype = first['Request Type Name']
    nfte = float(fte.get(fid, 0))
    pnet = float(prog[AMT].sum()) + float(per.loc[per['Form ID'] == fid, 'Total'].sum())  # incl. position costs
    lname = name.lower()

    # ---- MOCK service-impact flag + type (placeholder until bureaus tag DPs) ----
    real_si = bool(si) and not NO_SI.match(si) and len(si) > 40
    eligible = rtype in {'Recognizing New External Revenue', 'Technical Adjustments', 'Policy Set-aside',
                         'Encumbrance Carryover (GF only)', 'Fund Balance True Up', 'Realignment Implementation'}
    has_si = eligible and real_si and (rtype != 'Realignment Implementation' or h(fid) < .25)
    if rtype in ('Recognizing New External Revenue', 'Technical Adjustments') and real_si and not has_si:
        has_si = h(fid) < .5
    if has_si:
        if rtype == 'Realignment Implementation':
            t = 'Realign'
        elif nfte > 0 and pnet >= 0:
            t = 'Create' if re.search(r'\bnew\b|pilot|launch|establish|creat|start', lname) or h(fid) < .5 else 'Enhance'
        elif abs(pnet) < 1000:
            t = 'Realign'
        elif pnet < 0 and float(exp[AMT].sum()) + float(per.loc[per['Form ID'] == fid, 'Total'].sum()) >= 0:
            t = 'Realign'
        elif pnet < 0:
            t = 'Eliminate' if re.search(r'eliminat|discontinu|sunset|close|end of', lname) or h(fid) < .25 else 'Reduce'
        elif re.search(r'\bnew\b|pilot|launch|establish|creat|start', lname) or (nfte > 0 and h(fid) < .6):
            t = 'Create'
        else:
            t = 'Enhance'
    else:
        objs = set(exp.loc[exp[AMT] != 0, 'Major Object Name'])
        if rtype == 'GF Balancing Entry (CBO Only)':
            t = 'Balancing'
        elif rtype in ('Non GF Contingency', 'Contingency', 'Reserve for Future Expenditure Tech Adjustment'):
            t = 'Contingency or reserve change'
        elif rtype == 'Fund Balance True Up':
            t = 'Contingency or reserve change' if objs and objs <= NONPROG else 'New revenue'
        elif rtype == 'Recognizing New External Revenue':
            t = 'New revenue'
        elif rtype == 'Realignment Implementation':
            t = 'Realignment'
        elif rtype in ('Policy Set-aside', 'Encumbrance Carryover (GF only)'):
            t = 'CSL add'
        else:  # Technical Adjustments
            rev_moves = rev[AMT].abs().sum()
            if abs(exp[AMT].sum()) < 1 and rev_moves > 0:
                t = 'Revenue replacement'
            elif pnet > 0 and h(fid) < .45:
                t = 'CSL add'
            else:
                t = 'Realignment'

    ln = g.groupby(['Account Name', 'Major Object Name', 'Fund - Name'])[AMT].sum().reset_index()
    lines = [[0 if r['Account Name'] == 'Revenue' else 1, r['Major Object Name'], fund_label(r['Fund - Name']), int(r[AMT])]
             for _, r in ln.iterrows() if r[AMT]]
    pg = per[per['Form ID'] == fid]
    pos = []
    if len(pg):
        pa = pg.groupby(['PCN', 'Job Class - Name', 'Fund - Name', 'Home Fund Center -  Name'], dropna=False).agg(
            fte=('FTE', 'sum'), sal=('Salary', 'sum'), ben=('Benefit', 'sum'), sta=('Statutory', 'sum'), tot=('Total', 'sum'),
            j=('Justification', lambda x: txt(next((v for v in x if txt(v)), '')))).reset_index()
        for _, r in pa.sort_values(['fte', 'tot'], ascending=False).iterrows():
            pos.append([txt(r['PCN']), re.sub(r'^\d+\s*-\s*', '', txt(r['Job Class - Name'])), fund_label(r['Fund - Name']),
                        re.sub(r'^\S+\s*-\s*', '', txt(r['Home Fund Center -  Name'])), round(float(r.fte), 2),
                        int(r.sal), int(r.ben + r.sta), int(r.tot), r.j])
    # The financial export leaves out position-generated personnel costs; they are only in the
    # personnel export. Add them to Personnel requirements by fund so each package balances.
    pos_by_fund = pg.groupby('Fund - Name')['Total'].sum() if len(pg) else pd.Series(dtype=float)
    lk = {(l[0], l[1], l[2]): l for l in lines}
    for f, v in pos_by_fund.items():
        if not v: continue
        key = (1, 'Personnel', fund_label(f))
        if key in lk: lk[key][3] += int(v)
        else:
            lk[key] = [1, 'Personnel', fund_label(f), int(v)]; lines.append(lk[key])
    lines = [l for l in lines if l[3]]
    pos_gf = int(sum(v for f, v in pos_by_fund.items() if str(f).startswith('100000')))
    gexp_v = int(gexp[AMT].sum()) + pos_gf; e_v = int(exp[AMT].sum()) + int(pos_by_fund.sum())
    pers_fin = int(exp.loc[exp['Major Object Name'] == 'Personnel', AMT].sum())
    dps.append({
        'id': int(fid), 'name': name, 'b': first['bcode'],
        'rt': rtype, 'si': has_si, 'tp': t,
        'e': e_v, 'ge': gexp_v,
        'fte': round(nfte, 2), 'gfte': round(float(gfte.get(fid, 0)), 2),
        'gf': bool(g.isgf.any()), 'ngf': bool((~g.isgf).any()),
        'ln': lines, 'pos': pos, 'pf': pers_fin,
        'd': txt(first['Package Desc']) or txt(first['Justification']),
        'sit': si, 'eq': txt(first['Equity Outcomes and Impact Tracking']),
    })

# Mock only: make sure every type appears at least once in the demo
red = sorted([d for d in dps if d['tp'] == 'Reduce'], key=lambda d: d['e'])
if red and not any(d['tp'] == 'Eliminate' for d in dps):
    red[0]['tp'] = 'Eliminate'

bureaus = {}
for code in fin['bcode'].unique():
    r = sa.loc[code]
    bureaus[code] = {'n': r['Bureau Name'], 'sa': clean_sa(r['NEW Service Area - Name'])}

totals = {
    'expense': int(fin.loc[fin['Account Name'] == 'Expense', AMT].sum()) + int(per['Total'].sum()),
    'pos_cost': int(per['Total'].sum()),
    'revenue': int(fin.loc[fin['Account Name'] == 'Revenue', AMT].sum()),
    'lines': int(len(fin)), 'dps': len(dps), 'fte': round(float(per.FTE.sum()), 2),
}
json.dump({'dps': dps, 'bureaus': bureaus, 'totals': totals}, open(out, 'w'), separators=(',', ':'), ensure_ascii=False)

d = pd.DataFrame(dps)
print(totals)
print(d.groupby(['si', 'tp']).agg(n=('id', 'size'), e=('e', 'sum')))
print('sum e', d.e.sum(), 'ge', d['ge'].sum(), 'fte', d.fte.sum(), 'positions', d.pos.map(len).sum(), 'dps w/ pos', (d.pos.map(len) > 0).sum())
