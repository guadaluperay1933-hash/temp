# -*- coding: utf-8 -*-
"""不看公式，直接用 ev2_913.json + jour.json 在 Python 里重算一遍，再跟重算过的 xlsx 逐项比。
跑法：python3 verify2.py <重算过的xlsx>"""
import sys, os, json, collections, openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
XL = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', '建筑挂靠业务核算系统.xlsx')
D = json.load(open(os.path.join(HERE, 'ev2_913.json')))
EV = [e for e in D['events'] if e['kind'] != '预提税费']
EV.sort(key=lambda e: (e['date'], e['src']))
PASS_ROWS, DED_ROWS, WAGE_ROWS = D['pass_rows'], D['ded_rows'], D.get('wage', [])
PROJ_ORDER = D['proj_order']
R2 = lambda x: round(x + 0.0, 2)
SHEET2UNIT = {'德誉嘉': '德誉嘉', '迅驰': '迅驰', '华城': '华城', '金沁': '金沁',
              '湖南锦泰': '湖南锦泰', '康欣': '康欣', '安锐': '安锐', '杰华电气': '杰华'}
HOLD = ['德誉嘉', '迅驰', '华城', '金沁', '湖南锦泰', '康欣', '安锐', '杰华']
KMAP = {'挂靠单位代收': '挂靠代收'}
PCODE = {p['full']: f'A{i+1:03d}' for i, p in enumerate(PROJ_ORDER)}

# ---------- 在 Python 里把「业务流水」重建一遍 ----------
rows = []
for e in EV:
    kind = KMAP.get(e['kind'], e['kind'])
    amt = float(e['amt'])
    src = SHEET2UNIT.get(e['src'].split('!')[0].strip(), e['src'].split('!')[0].strip())
    if kind == '销项开票':
        if e.get('from_cost_col'):
            fee, ocol = '不回成本票', '已开成本票'
        elif abs(e.get('mrate', 0.0)) < 1e-9:
            fee = '不回成本票' if abs(float(e.get('cost_due') or 0)) < 0.005 else '不扣管理费'
            ocol = '销售开票金额'
        else:
            fee, ocol = '扣管理费', '销售开票金额'
        mfee = 0.0 if fee == '不扣管理费' else R2(amt * round(e.get('mrate', 0.0), 6))
        due = 0.0 if fee == '不回成本票' else R2(amt - mfee)
        tax = tuple(R2(float(e.get(k) or 0)) for k in ('tax_v', 'tax_s', 'tax_y', 'tax_i'))
    else:
        fee = ''
        ocol = '已开成本票' if e.get('from_cost_col') else '销售开票金额'
        mfee, due, tax = 0.0, 0.0, (0.0, 0.0, 0.0, 0.0)
    rows.append(dict(kind=kind, payer=e['payer'], payee=e['payee'], amt=amt, mfee=mfee, due=due,
                     tax=tax, taxsum=R2(sum(tax)), cnt=e.get('count_in', '是'), src=src, fee=fee,
                     ocol=ocol, proj=PCODE.get(e['proj'], ''),
                     ar=R2(amt * (e['rebate'] / amt)) if (kind == '销项开票' and e.get('rebate') and amt) else 0.0))
for x in PASS_ROWS:
    rows.append(dict(kind='其他应付发生', payer='迅驰', payee='迅驰', amt=x['amt'], mfee=0, due=0,
                     tax=(0,)*4, taxsum=0, cnt='是', src='迅驰', fee='', ocol='销售开票金额',
                     proj=PCODE.get(x['proj'], ''), ar=0))
for x in WAGE_ROWS:
    rows.append(dict(kind='工资扣抵', payer='泓普', payee='康欣', amt=x['amt'], mfee=0, due=0,
                     tax=(0,)*4, taxsum=0, cnt='是', src='总台账', fee='', ocol='销售开票金额',
                     proj=PCODE.get(x['proj'], ''), ar=0))
for x in DED_ROWS:
    rows.append(dict(kind='其他应付扣税', payer='泓普', payee='迅驰', amt=x['amt'], mfee=0, due=0,
                     tax=(0,)*4, taxsum=0, cnt='是', src='迅驰', fee='', ocol='销售开票金额',
                     proj='', ar=0))
V = [r for r in rows if r['cnt'] == '是']

def S(pred, key='amt', src=None):
    return R2(sum(r[key] for r in (rows if src else V)
                 if pred(r) and (src is None or r['src'] == src)))

wb = openpyxl.load_workbook(XL, data_only=True)
res = []
def chk(label, want, got):
    got = float(got) if isinstance(got, (int, float)) else 0.0
    res.append((label, R2(want), R2(got), abs(R2(want) - R2(got)) < 0.05))

# ---------- ① 单位汇总 ----------
ws = wb['单位汇总']
COLS = {'开票额': 3, '应扣管理费': 4, '应到成本票': 5, '扣费的': 6, '不扣费的': 7, '已收成本票': 8,
        '还差成本票': 9, '应提税费': 10, '已交税': 11, '欠税未交': 12,
        '业主已付': 13, '业主未付': 14, '已转我方': 15, '代收未转': 16, '管理费已结算': 17, '扣质保金': 18}
rowof = {}
for r in range(7, 47):
    u = ws.cell(row=r, column=1).value
    if u: rowof[u] = r
for u in HOLD:
    r = rowof[u]
    sale = lambda p: (lambda x: x['kind'] == '销项开票' and x['payer'] == u and p(x))
    chk(f'单位汇总·{u}·开票额', S(lambda x: x['kind'] == '销项开票' and x['payer'] == u),
        ws.cell(row=r, column=COLS['开票额']).value)
    chk(f'单位汇总·{u}·应扣管理费',
        S(lambda x: x['kind'] == '销项开票' and x['payer'] == u, 'mfee'),
        ws.cell(row=r, column=COLS['应扣管理费']).value)
    chk(f'单位汇总·{u}·应到成本票',
        S(lambda x: x['kind'] == '销项开票' and x['payer'] == u, 'due'),
        ws.cell(row=r, column=COLS['应到成本票']).value)
    chk(f'单位汇总·{u}·扣费的',
        S(lambda x: x['kind'] == '销项开票' and x['payer'] == u and x['fee'] == '扣管理费', 'due'),
        ws.cell(row=r, column=COLS['扣费的']).value)
    chk(f'单位汇总·{u}·不扣费的',
        S(lambda x: x['kind'] == '销项开票' and x['payer'] == u and x['fee'] == '不扣管理费', 'due'),
        ws.cell(row=r, column=COLS['不扣费的']).value)
    chk(f'单位汇总·{u}·已收成本票',
        S(lambda x: x['kind'] in ('成本票', '销项开票') and x['payee'] == u),
        ws.cell(row=r, column=COLS['已收成本票']).value)
    chk(f'单位汇总·{u}·应提税费', R2(sum(x['taxsum'] for x in V if x['src'] == u)),
        ws.cell(row=r, column=COLS['应提税费']).value)
    chk(f'单位汇总·{u}·已交税',
        R2(sum(x['amt'] for x in V if x['kind'] == '已交税' and x['src'] == u)),
        ws.cell(row=r, column=COLS['已交税']).value)
    chk(f'单位汇总·{u}·业主已付', S(lambda x: x['kind'] == '挂靠代收' and x['payee'] == u),
        ws.cell(row=r, column=COLS['业主已付']).value)
    chk(f'单位汇总·{u}·已转我方', S(lambda x: x['kind'] == '我方收款' and x['payer'] == u),
        ws.cell(row=r, column=COLS['已转我方']).value)
    chk(f'单位汇总·{u}·管理费已结算', S(lambda x: x['kind'] == '管理费结算' and x['payee'] == u),
        ws.cell(row=r, column=COLS['管理费已结算']).value)

# ---------- ② 8 张单位竖版明细的期间合计 ----------
for u in HOLD:
    ws = wb[f'{u}明细']
    mine = [x for x in rows if x['src'] == u]
    chk(f'{u}明细·销售开票金额',
        R2(sum(x['amt'] for x in mine if x['kind'] == '销项开票' and x['ocol'] != '已开成本票')),
        ws['H6'].value)
    chk(f'{u}明细·应扣管理费', R2(sum(x['mfee'] for x in mine)), ws['I6'].value)
    chk(f'{u}明细·应到成本票', R2(sum(x['due'] for x in mine)), ws['J6'].value)
    chk(f'{u}明细·已到成本票',
        R2(sum(x['amt'] for x in mine if x['ocol'] == '已开成本票')), ws['K6'].value)
    chk(f'{u}明细·应扣税费', R2(sum(x['taxsum'] for x in mine)), ws['Q6'].value)
    chk(f'{u}明细·已交税', R2(sum(x['amt'] for x in mine if x['kind'] == '已交税')), ws['R6'].value)
    chk(f'{u}明细·业主付给挂靠', R2(sum(x['amt'] for x in mine if x['kind'] == '挂靠代收')), ws['S6'].value)
    chk(f'{u}明细·挂靠转我方', R2(sum(x['amt'] for x in mine if x['kind'] == '我方收款')), ws['T6'].value)

# ---------- ③ 项目汇总（逐项目） ----------
ws = wb['项目汇总']
for i, p in enumerate(PROJ_ORDER):
    code = f'A{i+1:03d}'
    r = 7 + i
    assert ws.cell(row=r, column=1).value == code, (r, code, ws.cell(row=r, column=1).value)
    chk(f'项目汇总·{code}·开票额', S(lambda x: x['kind'] == '销项开票' and x['proj'] == code),
        ws.cell(row=r, column=3).value)
    chk(f'项目汇总·{code}·应到成本票',
        S(lambda x: x['kind'] == '销项开票' and x['proj'] == code, 'due'),
        ws.cell(row=r, column=5).value)

# ---------- ④ 税费台账 ----------
ws = wb['税费台账']
for r in range(7, 20):
    u = ws.cell(row=r, column=1).value
    if u not in HOLD: continue
    for j, k in enumerate(('tax_v', 'tax_s', 'tax_y', 'tax_i')):
        chk(f'税费台账·{u}·{k}', R2(sum(x['tax'][j] for x in V if x['src'] == u)),
            ws.cell(row=r, column=3 + j).value)

# ---------- ⑤ 往来台账：德誉嘉返现 ----------
ws = wb['往来台账']
chk('往来台账·返现发生额合计', R2(sum(x['ar'] for x in V if x['ar'] > 0)), None
    if False else ws.cell(row=6, column=3).value)

ok = sum(1 for x in res if x[3])
print(f'共 {len(res)} 项，一致 {ok} 项，不一致 {len(res)-ok} 项')
for lab, w, g, good in res:
    if not good: print(f'  ✗ {lab}: 重算 {w:,.2f}  表里 {g:,.2f}  差 {g-w:,.2f}')
if ok == len(res): print('✓ 全部一致')
