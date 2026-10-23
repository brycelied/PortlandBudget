"""Merge Fall decision packages into the adopted-budget treemap data.

Usage: python3 -I prep_treemap.py <treemap index.html> <dp_export.xlsx> <adhoc.xlsx> <prog.json> <out.json>

Reads the treemap's embedded data (D) and adds:
  D.dr    DP rows in the treemap's row format plus the DP index:
          [bureau, program, fund, major object, capital flag, sub-program, amount, dp]
  D.dp    DP metadata: id, name, bureau, request type, sample tags, text, FTE
  D.dpos  positions per DP: [dp, sub-program, job class, PCN, FTE, total cost]
  D.desc  "bureau|sub-program" -> {d: description, s: services} from the adopted budget book
prog.json (from prep_programs.py) supplies DP metadata, sample tags, positions and descriptions.
"""
import sys, json, re
import pandas as pd

tm_html, src, adhoc, progjson, out = sys.argv[1:6]
AMT = '2026-27 FALL Requested Total'

s = open(tm_html).read()
i = s.index('const D = ') + len('const D = ')
D = json.loads(s[i:s.index(';\nconst GF')])
P = json.load(open(progjson))

def txt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()

def idx(lst, name):
    if name not in lst:
        lst.append(name)
    return lst.index(name)

a = pd.read_excel(adhoc, sheet_name='Detail Export', usecols=range(39))
fin = pd.read_excel(src, sheet_name='Financial Data Export')
per = pd.read_excel(src, sheet_name='Personnel Export')

# lookups from the adopted budget
bcode_name = a.drop_duplicates('Bureau').set_index('Bureau')['Bureau Name'].map(txt).to_dict()
names = a.drop_duplicates(['Bureau', 'Sub-Program']).set_index(['Bureau', 'Sub-Program'])[['Program Name', 'Sub-Program Name']]
sp_any = a.drop_duplicates('Sub-Program').set_index('Sub-Program')[['Program Name', 'Sub-Program Name']]
pg_any = a.drop_duplicates('Program').set_index('Program')['Program Name']
proj_type = a.dropna(subset=['Type']).drop_duplicates('Project Code - Name').set_index('Project Code - Name')['Type'].to_dict()
fa_map = a.drop_duplicates(['Bureau', 'Functional Area']).set_index(['Bureau', 'Functional Area'])['Sub-Program'].to_dict()
fund_by_num = {f.split(' ')[0]: k for k, f in enumerate(D['fd'])}
mo_by_name = {m[1]: k for k, m in enumerate(D['mo'])}
mo_by_name['Fund Transfers - Expense'] = mo_by_name['Fund Transfers']

D['ci'] = []
ci_by_code = {}
def ci_idx(v):
    t = txt(v)
    code = t.split(' ')[0]
    if code not in ci_by_code:
        ci_by_code[code] = len(D['ci'])
        D['ci'].append(re.sub(r'^(\S+)\s*-\s*', r'\1 ', t))
    return ci_by_code[code]

def fund_label(v):
    v = txt(v)
    m = re.match(r'^(\d+)\s*-\s*(.*)$', v)
    return f"{m.group(2).strip()} ({m.group(1)})" if m else v

def bureau_idx(code):
    return D['bu'].index(bcode_name[code])

def names_for(b, code, fallback=''):
    if (b, code) in names.index:
        r = names.loc[(b, code)]
        return txt(r['Program Name']), txt(r['Sub-Program Name'])
    if code in sp_any.index:
        r = sp_any.loc[code]
        return txt(r['Program Name']), txt(r['Sub-Program Name'])
    if code in ('', 'NA', None):
        return 'Not assigned', 'No program offer assigned'
    return txt(pg_any.get(code[:4], '')) or 'Other', fallback or code

def fund_idx(fund_num_name):
    num = str(fund_num_name).split(' ')[0]
    if num not in fund_by_num:
        D['fd'].append(re.sub(r'\s+', ' ', str(fund_num_name)).strip())
        fund_by_num[num] = len(D['fd']) - 1
    return fund_by_num[num]

def capital(project, cip_prog):
    t = proj_type.get(project)
    if t:
        return 1 if t == 'CIPJ' else 0
    return 1 if (not str(project).startswith('0 ') and txt(cip_prog) not in ('', 'NA')) else 0

# ---------- DP metadata (order defines the dp index) ----------
dps = sorted(P['dps'], key=lambda d: d['id'])
dp_index = {d['id']: k for k, d in enumerate(dps)}
D['dp'] = [{'id': d['id'], 'n': d['name'], 'b': bureau_idx(d['b']), 'rt': d['rt'], 'si': d['si'], 'tp': d['tp'],
            'd': d['d'], 'sit': d['sit'], 'eq': d['eq'], 'fte': d['fte'], 'e': d['e'],
            'ln': d['ln'], 'pos': d['pos']} for d in dps]

# ---------- DP rows ----------
fin['b'] = fin['Bureau - Name'].str[:2]
fin['po'] = fin['Program Offer - Name'].str.extract(r'^(\S+)')[0].fillna('NA')
fin['pon'] = fin['Program Offer - Name'].str.replace(r'^\S+\s*[–-]\s*', '', regex=True).map(txt)
ex = fin[fin['Account Name'] == 'Expense'].copy()
agg = {}
def add(key, v):
    agg[key] = agg.get(key, 0) + v

for _, r in ex.iterrows():
    v = int(r[AMT])
    if not v:
        continue
    bi = bureau_idx(r.b)
    pn, spn = names_for(r.b, r.po, r.pon)
    key = (bi, idx(D['pg'], pn), fund_idx(r['Fund Number - Name']), mo_by_name[r['Major Object Name']],
           capital(r['Project Code - Name'], r['CIP Program Name']), idx(D['sp'], spn), dp_index[r['Form Header ID']], ci_idx(r['CommItem - Name']))
    add(key, v)

# position costs (not in the financial export): Personnel, operating, program offer from functional area
dp_b = fin.drop_duplicates('Form Header ID').set_index('Form Header ID')['b']
main_po = {}
for dp, g in ex.groupby('Form Header ID'):
    pers = g[g['Major Object Name'] == 'Personnel']
    gg = pers if len(pers) else g
    main_po[dp] = gg.groupby('po')[AMT].apply(lambda x: x.abs().sum()).idxmax()
pos_rows = []
per['fa'] = per['Allocated Functional Area'].astype(str).str.extract(r'^(\S+)')[0]
for _, r in per.iterrows():
    b = dp_b[r['Form ID']]
    sp = fa_map.get((b, r.fa))
    if sp is None and r.fa and r.fa != '0' and len(r.fa) >= 6:
        sp = r.fa[:6]
    if sp is None:
        sp = main_po.get(r['Form ID'], 'NA')
    pn, spn = names_for(b, sp)
    bi, pi, si = bureau_idx(b), idx(D['pg'], pn), idx(D['sp'], spn)
    fnum = str(r['Fund - Name']).split(' ')[0]
    # personnel export uses 6-digit fund codes (100000); treemap funds use the fund number (100)
    match = [k for k, f in enumerate(D['fd']) if f.split(' ')[0] == fnum[:3] or f.split(' ')[0] == fnum]
    fi = match[0] if match else fund_idx(f"{fnum[:3]} - {re.sub(r'^\\d+\\s*-\\s*', '', str(r['Fund - Name']))}")
    if int(r.Total):
        add((bi, pi, fi, mo_by_name['Personnel'], 0, si, dp_index[r['Form ID']], ci_idx(r['CommItem - Name'])), int(r.Total))
    pos_rows.append((dp_index[r['Form ID']], si, txt(r['Job Class - Name']), txt(r['PCN']), float(r.FTE), int(r.Total)))

D['dr'] = [[*k[:6], v, k[6], k[7]] for k, v in agg.items() if v]

# each DP's full resources and requirements by fund and commitment item: [account, major object, fund, ci, amount]
lc = {}
for _, r in fin.iterrows():
    v = int(r[AMT])
    if not v:
        continue
    k = (dp_index[r['Form Header ID']], 0 if r['Account Name'] == 'Revenue' else 1, r['Major Object Name'], fund_label(r['Fund - Name']), ci_idx(r['CommItem - Name']))
    lc[k] = lc.get(k, 0) + v
for _, r in per.iterrows():
    if int(r.Total):
        k = (dp_index[r['Form ID']], 1, 'Personnel', fund_label(r['Fund - Name']), ci_idx(r['CommItem - Name']))
        lc[k] = lc.get(k, 0) + int(r.Total)
for k, v in lc.items():
    if v:
        D['dp'][k[0]].setdefault('lc', []).append([k[1], k[2], k[3], k[4], v])
for d in D['dp']:
    d.pop('ln', None)
    d.setdefault('lc', [])

# adopted budget at commitment item level, same coding as the treemap rows: [bu, pg, fd, mo, cap, sp, amount, ci]
ae = a[a['Account Name'] == 'Expense']
mo_by_code = {m[0]: k for k, m in enumerate(D['mo'])}
bu_by_name = {n: k for k, n in enumerate(D['bu'])}
bagg = {}
bcols = ['Bureau Name', 'Program Name', 'Fund Number - Name', 'Major Object - Name', 'Type', 'Sub-Program Name', 'CommItem - Name', 'FY2026-27 Adopted Budget']
for bn, pn, fn, mo, ty, spn, ci, v in ae[bcols].itertuples(index=False):
    if not v:
        continue
    key = (bu_by_name[txt(bn)], idx(D['pg'], txt(pn)), fund_idx(fn), mo_by_code[str(mo).split('-')[0].strip()],
           1 if ty == 'CIPJ' else 0, idx(D['sp'], txt(spn)), ci_idx(ci))
    bagg[key] = bagg.get(key, 0) + v
D['bci'] = [[*k[:6], int(round(v)), k[6]] for k, v in bagg.items() if round(v)]

# positions: one line per DP, sub-program and PCN
pp = pd.DataFrame(pos_rows, columns=['dp', 'sp', 'job', 'pcn', 'fte', 'tot'])
pp = pp.groupby(['dp', 'sp', 'pcn', 'job'], as_index=False).agg(fte=('fte', 'sum'), tot=('tot', 'sum'))
D['dpos'] = [[int(r.dp), int(r.sp), re.sub(r'^\d+\s*-\s*', '', r.job), r.pcn, round(float(r.fte), 2), int(r.tot)]
             for r in pp.sort_values(['dp', 'fte', 'tot'], ascending=[True, False, False]).itertuples() if r.fte or r.tot]

# ---------- descriptions keyed by treemap bureau and sub-program ----------
desc = {}
for p in P['programs']:
    if not (p.get('d') or p.get('s')):
        continue
    bi = bureau_idx(p['b'])
    _, spn = names_for(p['b'], p['c'], p['n'])
    if spn in D['sp']:
        desc[f"{bi}|{D['sp'].index(spn)}"] = {'d': p['d'], 's': p.get('s', '')}
D['desc'] = desc
D['totals'] = P['totals']

json.dump(D, open(out, 'w'), separators=(',', ':'), ensure_ascii=False)

# ---------- checks ----------
dr = pd.DataFrame(D['dr'], columns=['bu', 'pg', 'fd', 'mo', 'cap', 'sp', 'v', 'dp', 'ci'])
gf = [k for k, f in enumerate(D['fd']) if f.startswith('100 ')][0]
print('base rows', len(D['rows']), 'base total', sum(r[6] for r in D['rows']))
print('dp rows', len(dr), 'dp total', dr.v.sum(), 'GF', dr.loc[dr.fd == gf, 'v'].sum(), 'capital', dr.loc[dr.cap == 1, 'v'].sum())
print('dps', len(D['dp']), 'with rows', dr.dp.nunique(), 'positions', len(D['dpos']), 'descriptions', len(desc))
print('new program names', len(D['pg']), 'new sub-program names', len(D['sp']), 'funds', len(D['fd']))

# commitment item checks
bci = pd.DataFrame(D['bci'], columns=['bu','pg','fd','mo','cap','sp','v','ci'])
base = pd.DataFrame(D['rows'], columns=['bu','pg','fd','mo','cap','sp','v'])
k = ['bu','pg','fd','mo','cap','sp']
m = base.groupby(k).v.sum().to_frame('rows').join(bci.groupby(k).v.sum().to_frame('ci'), how='outer').fillna(0)
print('bci rows', len(bci), 'total', bci.v.sum(), '| keys off by >$2:', int(((m.rows - m.ci).abs() > 2).sum()), 'of', len(m))
lcs = [sum(x[4] for x in d['lc'] if x[0] == 1) for d in D['dp']]
print('ci labels', len(D['ci']), '| DP requirement lines total', sum(lcs), '| DP rows total', sum(r[6] for r in D['dr']))
print('size of bci json', len(json.dumps(D['bci'], separators=(',', ':'))))
