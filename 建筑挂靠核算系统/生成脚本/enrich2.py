# -*- coding: utf-8 -*-
"""在 ev_913.json 上补：票据类型 / 德誉嘉返现 / 迅驰过账应付 / 合伙项目 / 工资扣抵
跑法：python3 parse_gk2.py && python3 enrich2.py   →  ev2_913.json
"""
import openpyxl, json, os, re, collections

HERE = os.path.dirname(os.path.abspath(__file__))
REF  = os.path.join(HERE, '..', '参考')
REC  = os.path.join(REF, '原对账明细_9.13.xlsx')
LEDG = os.path.join(REF, '原发票总台账.xlsx')
OUT  = os.path.join(HERE, 'ev2_913.json')

D = json.load(open(os.path.join(HERE, 'ev_913.json')))
EV, PROJ_ORDER, RAW = D['events'], D['proj_order'], D['raw_rows']
NORM = {'德誉佳': '德誉嘉'}
n = lambda s: NORM.get(str(s).strip(), str(s).strip()) if s is not None else ''
def norm_proj(s):
    s = re.sub(r'\s+', '', str(s or ''))
    s = s.replace('（', '(').replace('）', ')').replace('，', ',')
    return re.sub(r'^项目名称[:：]', '', s)

# ---------- ① 票据类型 ----------
lw = openpyxl.load_workbook(LEDG, data_only=True)['总台账']
TYPE_MAP = {'设备票': '机械设备票', '劳务票': '劳务票', '土建票': '土建票', '安装票': '安装票'}
byamt = {}
for r in range(13, 58):
    t, u, a = lw.cell(r, 9).value, lw.cell(r, 11).value, lw.cell(r, 13).value
    if t and u and isinstance(a, (int, float)):
        byamt[(n(u), round(float(a), 2))] = TYPE_MAP.get(str(t).strip(), '其他')

MECH_UNITS = {'德誉嘉', '迅驰', '安锐'}
MECH_KEY = ('机械费', '设备', '租赁', '施工机具', '特种设备')
def guess_type(unit, proj):
    if not unit: return ''
    p = proj or ''
    if any(k in p for k in MECH_KEY): return '机械设备票'
    if unit in MECH_UNITS: return '机械设备票'
    return '劳务票'

matched = collections.Counter()
for e in EV:
    if e['kind'] == '销项开票':
        t = byamt.get((e['unit'], round(e['amt'], 2)))
        e['itype'] = t or guess_type(e['unit'], e['proj'])
        matched['台账匹配' if t else '按规则推断'] += 1
    else:
        e['itype'] = ''
sale_pool = [e for e in EV if e['kind'] == '销项开票' and e.get('cost_due')]
used = set()
for e in EV:
    if e['kind'] != '成本票': continue
    hit = None
    for k, s0 in enumerate(sale_pool):
        if k in used: continue
        if s0['proj'] == e['proj'] and abs(float(s0['cost_due']) - float(e['amt'])) < 1.0:
            hit = (k, s0); break
    if hit:
        used.add(hit[0]); e['itype'] = hit[1]['itype']; matched['成本票·按应开金额认亲'] += 1
    else:
        e['itype'] = guess_type(n(e['payee']), e['proj']); matched['成本票·按单位推断'] += 1
print('① 票据类型：', dict(matched))

# ---------- ② 工资扣抵 ----------
WAGE = []
for r in range(13, 58):
    aj = lw.cell(r, 36).value
    if isinstance(aj, (int, float)) and aj:
        pj = None
        for rr in range(r, 12, -1):
            v = lw.cell(rr, 8).value
            if v and str(v).strip(): pj = norm_proj(v); break
        WAGE.append(dict(proj=pj or '', amt=round(float(aj), 2), row=r))
print(f'② 工资扣抵：{len(WAGE)} 笔 合计 {sum(x["amt"] for x in WAGE):,.2f}')

# ---------- ③ 德誉嘉 返管理费 4% ----------
rw = openpyxl.load_workbook(REC, data_only=True)
hit = 0
for e in EV:
    if e.get('rebate'):
        hit += 1
    else:
        e['rebate'] = 0.0
print(f'③ 德誉嘉返现：{hit} 笔，合计 {sum(e["rebate"] for e in EV):,.2f}')

# ---------- ④ 迅驰 过账应付 / 应扣税费 ----------
ws = rw['迅驰']
pass_rows, ded_rows = [], []
for r in range(5, ws.max_row + 1):
    aa, ab = ws.cell(r, 27).value, ws.cell(r, 28).value
    proj, memo = ws.cell(r, 2).value, ws.cell(r, 3).value
    if isinstance(aa, (int, float)) and aa:
        pass_rows.append(dict(proj=norm_proj(proj), memo=str(memo).strip() if memo else '',
                              amt=round(float(aa), 2), row=r))
    if isinstance(ab, (int, float)) and ab:
        ded_rows.append(dict(proj=norm_proj(proj), memo=str(memo).strip() if memo else '',
                             amt=round(float(ab), 2), row=r))
print(f'④ 迅驰过账：应付 {len(pass_rows)} 笔 合计 {sum(x["amt"] for x in pass_rows):,.2f}；'
      f'应扣税费 {len(ded_rows)} 笔 合计 {sum(x["amt"] for x in ded_rows):,.2f}')

# ---------- ⑤ 合伙项目 ----------
BIZ = {'德誉嘉 ': 4, '迅驰': 4, '华城': 4, '康欣': 4, '金沁': 4, '安锐': 5, '湖南锦泰': 4, '杰华电气': 4}
partner = set()
for sn, bc in BIZ.items():
    w = rw[sn]
    cur = ''
    for r in range(5, w.max_row + 1):
        p = w.cell(r, 2).value
        if p: cur = norm_proj(p)
        v = w.cell(r, bc).value
        if v and '合伙' in str(v) and cur: partner.add(cur)
print(f'⑤ 合伙项目：{len(partner)} 个')

json.dump(dict(events=EV, pass_rows=pass_rows, ded_rows=ded_rows, partner=sorted(partner),
               wage=WAGE, proj_order=PROJ_ORDER, raw_rows=RAW),
          open(OUT, 'w'), ensure_ascii=False)
print('\n已写', OUT)
