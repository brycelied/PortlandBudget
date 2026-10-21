import json, sys
s, data, out = sys.argv[1:4]
t = open(s).read(); d = json.load(open(data))
t = t.replace('__DATA__', json.dumps(d, separators=(',',':'), ensure_ascii=False).replace('</', '<\\/'))
t = t.replace('__LINES__', f"{d['totals']['lines']:,}").replace('__NDP__', str(d['totals']['dps'])).replace('__TOTEXP__', f"${d['totals']['expense']:,}").replace('__POSCOST__', f"${d['totals']['pos_cost']:,}")
open(out,'w').write(t); print(len(t))
