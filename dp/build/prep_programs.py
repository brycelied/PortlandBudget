"""Add program-offer data to the DP dataset for the program dashboard.

Usage: python3 -I prep_programs.py <dp_export.xlsx> <adhoc.xlsx> <bureau_sa.csv> <dp.json> <descriptions.json|-> <out.json>

descriptions.json: {"<bureau>|<program offer code>": {"d": description, "s": services}} from the OpenBook adopted budget book.

dp.json comes from prep.py (DP-level data and mock tags). This script adds:
  - every adopted program offer (bureau x sub-program) with adopted requirements and FTE
  - each DP's change by program offer (expense requirements + position costs), all funds and GF
"""
import sys, json, re
import pandas as pd

src, adhoc, sa_csv, dpjson, descs, out = sys.argv[1:7]
AMT = '2026-27 FALL Requested Total'
ADO = 'FY2026-27 Adopted Budget'

def txt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()

def clean_sa(s):
    return re.sub(r'^[A-Z]{2}\d\s*-\s*', '', s).strip()

base = json.load(open(dpjson))
desc = json.load(open(descs)) if descs != '-' else {}

a = pd.read_excel(adhoc, sheet_name='Detail Export', usecols=range(39))
a['gf'] = a['Fund Number - Name'].astype(str).str.startswith('100 ')
fin = pd.read_excel(src, sheet_name='Financial Data Export')
per = pd.read_excel(src, sheet_name='Personnel Export')

# ---------- adopted program offers ----------
exp = a[a['Account Name'] == 'Expense']
fte = a[a['Account Name'] == 'FTE']
keycols = ['Bureau', 'Sub-Program']
names = (a.drop_duplicates(keycols).set_index(keycols)[['Program', 'Program Name', 'Sub-Program Name']])
ado = exp.groupby(keycols)[ADO].sum()
ado_gf = exp[exp.gf].groupby(keycols)[ADO].sum()
fte_all = fte.groupby(keycols)[ADO].sum()
fte_gf = fte[fte.gf].groupby(keycols)[ADO].sum()
pname_any = a.drop_duplicates('Program').set_index('Program')['Program Name']
spname_any = a.drop_duplicates('Sub-Program').set_index('Sub-Program')['Sub-Program Name']
fa_map = a.drop_duplicates(['Bureau', 'Functional Area']).set_index(['Bureau', 'Functional Area'])['Sub-Program']

programs = {}
def ensure(b, code, fallback_name=''):
    k = f'{b}|{code}'
    if k in programs:
        return k
    if (b, code) in names.index:
        r = names.loc[(b, code)]
        pcode, pn, n = r['Program'], txt(r['Program Name']), txt(r['Sub-Program Name'])
    else:
        pcode = code[:4] if code not in ('NA', '') else 'NA'
        pn = txt(pname_any.get(pcode, '')) or 'Other'
        n = fallback_name or txt(spname_any.get(code, '')) or code
        if code in ('NA', ''):
            pcode, pn, n = 'NA', 'Not assigned', 'No program offer assigned'
    programs[k] = {
        'k': k, 'b': b, 'c': code, 'pc': pcode, 'pn': pn, 'n': n,
        'a': int(round(ado.get((b, code), 0))), 'ag': int(round(ado_gf.get((b, code), 0))),
        'f': round(float(fte_all.get((b, code), 0)), 2), 'fg': round(float(fte_gf.get((b, code), 0)), 2),
        'd': desc.get(k, {}).get('d', ''), 's': desc.get(k, {}).get('s', ''),
        'new': (b, code) not in names.index,
    }
    return k

for (b, code) in ado.index:
    ensure(b, code)

# ---------- DP changes by program ----------
fin['b'] = fin['Bureau - Name'].str[:2]
fin['po'] = fin['Program Offer - Name'].str.extract(r'^(\S+)')[0].fillna('NA')
fin['pon'] = fin['Program Offer - Name'].str.replace(r'^\S+\s*[–-]\s*', '', regex=True)
fin['gf'] = fin['Fund - Name'].str.startswith('100000')
dp_b = fin.drop_duplicates('Form Header ID').set_index('Form Header ID')['b']

chg = {}  # (dp, key) -> [all, gf, fte, gfte]
def add(dp, k, v, g, f=0.0, fg=0.0):
    r = chg.setdefault((dp, k), [0, 0, 0.0, 0.0])
    r[0] += v; r[1] += g; r[2] += f; r[3] += fg

ex = fin[fin['Account Name'] == 'Expense']
for (dp, b, po, pon, gf), v in ex.groupby(['Form Header ID', 'b', 'po', 'pon', 'gf'])[AMT].sum().items():
    if not v:
        continue
    k = ensure(b, po, txt(pon))
    add(dp, k, int(v), int(v) if gf else 0)

# position costs (missing from the financial export) by functional area -> program offer
per['db'] = per['Form ID'].map(dp_b)
per['fa'] = per['Allocated Functional Area'].astype(str).str.extract(r'^(\S+)')[0]
per['gf'] = per['Fund - Name'].astype(str).str.startswith('100000')
main_po = {}
for dp, g in ex.groupby('Form Header ID'):
    pers = g[g['Major Object Name'] == 'Personnel']
    src_g = pers if len(pers) else g
    main_po[dp] = src_g.groupby('po')[AMT].apply(lambda s: s.abs().sum()).idxmax()

pos_prog = {}
for i, r in per.iterrows():
    b = r.db
    sp = fa_map.get((b, r.fa))
    if sp is None and r.fa and r.fa != '0' and len(r.fa) >= 6:
        sp = r.fa[:6]
    if sp is None:
        sp = main_po.get(r['Form ID'], 'NA')
    k = ensure(b, sp)
    pos_prog[i] = k
    add(r['Form ID'], k, int(r.Total), int(r.Total) if r.gf else 0, float(r.FTE), float(r.FTE) if r.gf else 0.0)
per['pk'] = pd.Series(pos_prog)

# attach to DPs
by_dp = {}
for (dp, k), (v, g, f, fg) in chg.items():
    if v or g or f or fg:
        by_dp.setdefault(dp, []).append([k, v, g, round(f, 2), round(fg, 2)])

# positions with program, aligned to prep.py's grouping (PCN, job class, fund, fund center)
pos_key = per.groupby(['Form ID', 'PCN', 'Job Class - Name', 'Fund - Name'], dropna=False)['pk'].agg(lambda s: s.mode().iat[0])

for d in base['dps']:
    d['pl'] = sorted(by_dp.get(d['id'], []), key=lambda r: -abs(r[1]))
    for p in d['pos']:
        # p = [PCN, job, fund label, fund center, fte, sal, ben, tot, just]
        m = per[(per['Form ID'] == d['id']) & (per['PCN'].astype(str) == p[0])]
        p.append(programs[m['pk'].mode().iat[0]]['n'] if len(m) else '')

sa = pd.read_csv(sa_csv).set_index('Bureau')
bureaus = {}
for b in sorted({p['b'] for p in programs.values()}):
    r = sa.loc[b]
    sa_name = r['NEW Service Area - Name']
    bureaus[b] = {'n': r['Bureau Name'], 'sa': clean_sa(sa_name) if isinstance(sa_name, str) else 'Community & Economic Development'}

out_d = {
    'programs': list(programs.values()),
    'dps': base['dps'],
    'bureaus': bureaus,
    'totals': {**base['totals'],
               'adopted': int(round(exp[ADO].sum())), 'adopted_gf': int(round(exp[exp.gf][ADO].sum())),
               'adopted_fte': round(float(fte[ADO].sum()), 2)},
}
json.dump(out_d, open(out, 'w'), separators=(',', ':'), ensure_ascii=False)

# ---------- checks ----------
P = pd.DataFrame(out_d['programs'])
pl = pd.DataFrame([r for d in base['dps'] for r in d['pl']], columns=['k', 'v', 'g', 'f', 'fg'])
print('programs', len(P), 'adopted', P.a.sum(), 'vs', out_d['totals']['adopted'], 'GF', P.ag.sum(), 'vs', out_d['totals']['adopted_gf'])
print('fte', P.f.sum(), 'vs', out_d['totals']['adopted_fte'])
print('DP change by program', pl.v.sum(), 'vs DP total', sum(d['e'] for d in base['dps']), '| GF', pl.g.sum(), 'vs', sum(d['ge'] for d in base['dps']))
print('DP fte', round(pl.f.sum(), 2), 'gf fte', round(pl.fg.sum(), 2))
print('programs touched', pl.k.nunique(), 'new program offers', int(P.new.sum()), 'with descriptions', int((P.d != '').sum()))
multi = pl.groupby('k').size(); print('programs with 2+ DPs', int((multi > 1).sum()))
