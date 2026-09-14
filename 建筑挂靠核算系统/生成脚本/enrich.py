# -*- coding: utf-8 -*-
"""在原 events.json 上补三个新维度：票据类型 / 德誉嘉返现 / 迅驰过账应付 / 合伙项目"""
import openpyxl, json, collections, datetime as dt

BASE = '/home/user/temp/建筑挂靠核算系统/生成脚本/'
REC  = '/root/.claude/uploads/9101e7cb-4f82-53a7-bbe8-41b0f93b12b0/c2938ae7-____9.041.xlsx'
LEDG = '/tmp/claude-0/-home-user-temp/9101e7cb-4f82-53a7-bbe8-41b0f93b12b0/scratchpad/a059/ref_台账1.xlsx'
NORM = {'德誉佳': '德誉嘉'}
n = lambda s: NORM.get(str(s).strip(), str(s).strip()) if s is not None else ''

EV = json.load(open(BASE + 'events.json'))

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
        key = (e['unit'], round(e['amt'], 2))
        t = byamt.get(key)
        e['itype'] = t or guess_type(e['unit'], e['proj'])
        matched['台账匹配' if t else '按规则推断'] += 1
    else:
        e['itype'] = ''
# 成本票：先按「同一项目、金额等于某张销项票的应开成本票」认亲，认不上再按收票单位猜
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

# ---------- ⑤ 工资扣抵（原总台账「劳务成本·工资扣抵」列） ----------
WAGE = []
for r in range(13, 58):
    aj = lw.cell(r, 36).value          # AJ 工资扣抵
    if isinstance(aj, (int, float)) and aj:
        # 项目名向上继承（原表只在每组第一行填）
        pj = None
        for rr in range(r, 12, -1):
            v = lw.cell(rr, 8).value
            if v and str(v).strip():
                pj = str(v).strip(); break
        WAGE.append(dict(proj=pj or '', amt=round(float(aj), 2), row=r))
print(f'⑤ 工资扣抵：{len(WAGE)} 笔 合计 {sum(x["amt"] for x in WAGE):,.2f}（原表 115,288.45）')

# ---------- ② 德誉嘉 返管理费 4% ----------
rw = openpyxl.load_workbook(REC, data_only=True)
ws = rw['德誉嘉 ']
reb = {}          # (金额) -> 返现额
for r in range(5, ws.max_row + 1):
    g, s = ws.cell(r, 7).value, ws.cell(r, 19).value
    if isinstance(g, (int, float)) and g and isinstance(s, (int, float)) and s:
        reb[round(float(g), 2)] = round(float(s), 2)
hit = 0
for e in EV:
    e['rebate'] = 0.0
    if e['kind'] == '销项开票' and e['unit'] == '德誉嘉':
        v = reb.get(round(e['amt'], 2))
        if v:
            e['rebate'] = v; hit += 1
print(f'② 德誉嘉返现：{hit} 笔，合计 {sum(e["rebate"] for e in EV):,.2f}（原表 7,935.67）')

# ---------- ③ 迅驰 过账应付 ----------
ws = rw['迅驰']
pass_rows, ded_rows = [], []
for r in range(5, ws.max_row + 1):
    aa, ab = ws.cell(r, 27).value, ws.cell(r, 28).value
    proj, memo = ws.cell(r, 2).value, ws.cell(r, 3).value
    if isinstance(aa, (int, float)) and aa:
        pass_rows.append(dict(proj=str(proj).strip() if proj else '', memo=str(memo).strip() if memo else '',
                              amt=round(float(aa), 2)))
    if isinstance(ab, (int, float)) and ab:
        ded_rows.append(dict(proj=str(proj).strip() if proj else '', memo=str(memo).strip() if memo else '',
                             amt=round(float(ab), 2)))
print(f'③ 迅驰过账：应付 {len(pass_rows)} 笔 合计 {sum(x["amt"] for x in pass_rows):,.2f}；'
      f'应扣税费 {len(ded_rows)} 笔 合计 {sum(x["amt"] for x in ded_rows):,.2f}')

# ---------- ④ 合伙项目 ----------
BIZ = {'德誉嘉 ': 4, '迅驰': 4, '华城': 4, '康欣': 4, '金沁': 4, '安锐': 5, '湖南锦泰': 4, '大太线老旧线路': 4}
partner = set()
for sn, bc in BIZ.items():
    w = rw[sn]
    for r in range(5, w.max_row + 1):
        v = w.cell(r, bc).value
        p = w.cell(r, 2).value
        if v and '合伙' in str(v) and p:
            partner.add(str(p).strip())
print(f'④ 合伙项目：{len(partner)} 个')
for p in sorted(partner): print('   ', p[:60])

json.dump(dict(events=EV, pass_rows=pass_rows, ded_rows=ded_rows, partner=sorted(partner), wage=WAGE),
          open('/tmp/claude-0/-home-user-temp/9101e7cb-4f82-53a7-bbe8-41b0f93b12b0/scratchpad/a059/ev2.json', 'w'),
          ensure_ascii=False)
print('\n已写 ev2.json')
