# -*- coding: utf-8 -*-
"""不看公式，直接用 ev2.json + jour.json 的原始数据在 Python 里重算一遍，
   再跟重算过的 xlsx 逐项比。跑法：python3 verify.py <重算过的xlsx>"""
import sys, os, json, openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
XL = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', '建筑挂靠业务核算系统.xlsx')
D2 = json.load(open(os.path.join(HERE, 'ev2.json')))
EV = [e for e in D2['events'] if e['kind'] != '预提税费']
EV.sort(key=lambda e: (e['date'], e['src']))
PASS_ROWS, DED_ROWS, WAGE_ROWS = D2['pass_rows'], D2['ded_rows'], D2.get('wage', [])
PARTNER = set(D2['partner'])
JOUR = json.load(open(os.path.join(HERE, 'jour.json')))
R2 = lambda x: round(x + 0.0, 2)

# 单位档案参数（与 gk_build.py 的 UNITS 一致）
UP = {
 '德誉嘉': dict(up=.13, our=.13, diff='否', mode='—',  add=0,   stamp=0,       inc=0),
 '迅驰':   dict(up=.09, our=.03, diff='是', mode='税差法',   add=.12, stamp=.000588, inc=0),
 '华城':   dict(up=.03, our=.01, diff='是', mode='全额销项法', add=.12, stamp=.0003,   inc=0),
 '康欣':   dict(up=.03, our=.03, diff='是', mode='全额销项法', add=.06, stamp=0,       inc=.002),
 '金沁':   dict(up=.09, our=.03, diff='是', mode='税差法',   add=.12, stamp=.0006,   inc=0),
 '安锐':   dict(up=.13, our=.13, diff='是', mode='税差法',   add=.06, stamp=0,       inc=0),
 '湖南锦泰': dict(up=.09, our=.09, diff='是', mode='全额销项法', add=.12, stamp=0,       inc=0),
 '杰华':   dict(up=.13, our=.13, diff='是', mode='全额销项法', add=.12, stamp=0,       inc=0),
}
HOLD = list(UP)
KMAP = {'挂靠单位代收': '挂靠代收'}
MFEE_RATE = {}
for _e in EV:
    if _e['kind'] == '销项开票' and _e['amt']:
        _k = (_e['payer'], round(_e['amt'], 2))
        MFEE_RATE[_k] = max(MFEE_RATE.get(_k, 0.0), round((_e.get('mfee') or 0) / _e['amt'], 6))
ETYPE_FIX = {'账户费年费': '账户年费', '耗材费用': '耗材费'}
rows = []
for e in EV:
    if e.get('count_in', '是') != '是': continue
    kind = KMAP.get(e['kind'], e['kind'])
    amt = float(e['amt'])
    mfee = R2(amt * MFEE_RATE.get((e['payer'], round(amt, 2)), 0.0)) if kind == '销项开票' else 0.0
    due = R2(amt - mfee) if kind == '销项开票' else 0.0
    reb = R2(e.get('rebate', 0) or 0)
    p = UP.get(e['payer'], {})
    v = a = st = ic = 0.0
    if kind == '销项开票' and p:
        if p['diff'] == '是' and p['up']:
            v = R2(amt / (1 + p['up']) * p['up'] - (due / (1 + p['our']) * p['our'] if p['our'] else 0)) \
                if p['mode'] == '税差法' else R2(amt / (1 + p['up']) * p['up'])
        a = R2(v * p['add']); st = R2(amt * p['stamp']); ic = R2(amt / (1 + p['up']) * p['inc'])
    rows.append(dict(kind=kind, payer=e['payer'], payee=e['payee'], proj=e['proj'], amt=amt,
                     mfee=mfee, due=due, reb=reb, v=v, a=a, st=st, ic=ic, itype=e.get('itype', '')))
for x in PASS_ROWS: rows.append(dict(kind='其他应付发生', payer='迅驰', payee='迅驰', proj=x['proj'],
                                     amt=x['amt'], mfee=0, due=0, reb=0, v=0, a=0, st=0, ic=0, itype=''))
for x in WAGE_ROWS: rows.append(dict(kind='工资扣抵', payer='泓普', payee='康欣', proj=x['proj'],
                                     amt=x['amt'], mfee=0, due=0, reb=0, v=0, a=0, st=0, ic=0, itype='劳务票'))
for x in DED_ROWS: rows.append(dict(kind='其他应付扣税', payer='泓普', payee='迅驰', proj='',
                                    amt=x['amt'], mfee=0, due=0, reb=0, v=0, a=0, st=0, ic=0, itype=''))

CLS = lambda t: '劳务类' if t in ('劳务票', '土建票', '安装票') else ('机械类' if t == '机械设备票'
       else ('材料类' if t == '材料票' else ('其他' if t else '')))
S = lambda f: R2(sum(x['amt'] for x in rows if f(x)))
SF = lambda fld, f: R2(sum(x[fld] for x in rows if f(x)))
hold = lambda u: u in HOLD
# 末层：该行开票方没有在同一项目上从别人手里收到销项票
sale = [x for x in rows if x['kind'] == '销项开票']
recv = {(x['proj'], x['payee']) for x in sale}
LAST = lambda x: x['kind'] == '销项开票' and (x['proj'], x['payer']) not in recv

OWNER = {'民能', '铜梁供电'}
PCODE_J = {}          # 日记账「项目」→ 项目编号 的映射，目前为空（原表填的不是项目编号）
NOT_COST = {'税费', '管理费', '借款', '还借款', '转备用金', '其他应付支付', '其他应收收回'}
TOPS = lambda x: x['kind'] == '销项开票' and x['payee'] in OWNER and x['proj']
PN = lambda x: x['proj'] in PARTNER          # 合伙项目
def profit(sel):
    inc = R2(sum(x['amt'] for x in rows if TOPS(x) and sel(x)))
    fee = R2(sum(x['mfee'] for x in rows if x['kind'] == '销项开票' and x['proj'] and sel(x)))
    tax = R2(sum(x['v'] + x['a'] + x['st'] + x['ic'] for x in rows
                 if x['kind'] == '销项开票' and x['proj'] and sel(x)))
    reb = R2(sum(x['reb'] for x in rows if x['kind'] == '销项开票' and x['proj'] and sel(x)))
    gross = R2(inc - fee - tax + reb)
    # 导入时日记账的项目编号列留空（原表那一列是工地/部门标注），所以按项目算的实际支出应当是 0
    spend = R2(sum(j['exp'] for j in JOUR
                   if PCODE_J.get(j.get('proj', '')) and j['etype'] not in NOT_COST))
    duep = R2(sum(x['amt'] for x in rows if x['kind'] == '其他应付发生' and x['proj'] and sel(x))
              - sum(x['amt'] for x in rows if x['kind'] == '其他应付扣税' and x['proj'] and sel(x)))
    return dict(inc=inc, fee=fee, tax=tax, reb=reb, gross=gross, spend=spend, duep=duep,
                mine=R2(gross - spend - duep))
PA, PS, PP = profit(lambda x: True), profit(lambda x: not PN(x)), profit(PN)

CASES = [
 ('项目利润 合计·业主端开票额', '项目利润', 'G6', PA['inc']),
 ('项目利润 合计·各层管理费',   '项目利润', 'H6', PA['fee']),
 ('项目利润 合计·应提税费',     '项目利润', 'I6', PA['tax']),
 ('项目利润 合计·返现',         '项目利润', 'J6', PA['reb']),
 ('项目利润 合计·票面毛利',     '项目利润', 'K6', PA['gross']),
 ('项目利润 合计·项目实际支出', '项目利润', 'M6', PA['spend']),
 ('项目利润 合计·应转合伙方',   '项目利润', 'O6', PA['duep']),
 ('项目利润 合计·归属我方',     '项目利润', 'P6', PA['mine']),
 ('项目利润 自营小计·开票额',   '项目利润', 'G7', PS['inc']),
 ('项目利润 自营小计·归属我方', '项目利润', 'P7', PS['mine']),
 ('项目利润 合伙小计·开票额',   '项目利润', 'G8', PP['inc']),
 ('项目利润 合伙小计·归属我方', '项目利润', 'P8', PP['mine']),
 ('单位汇总 合计·开票额',      '单位汇总', 'C6', S(lambda x: x['kind']=='销项开票' and hold(x['payer']))),
 ('单位汇总 合计·应扣管理费',  '单位汇总', 'D6', SF('mfee', lambda x: x['kind']=='销项开票' and hold(x['payer']))),
 ('单位汇总 合计·应到成本票',  '单位汇总', 'E6', SF('due',  lambda x: x['kind']=='销项开票' and hold(x['payer']))),
 ('单位汇总 合计·已收成本票',  '单位汇总', 'F6',
  R2(sum(x['amt'] for x in rows if x['kind'] in ('成本票','销项开票') and hold(x['payee'])))),
 ('单位汇总 合计·应提税费',    '单位汇总', 'H6',
  R2(sum(x['v']+x['a']+x['st']+x['ic'] for x in rows if x['kind']=='销项开票' and hold(x['payer'])))),
 ('单位汇总 合计·已交税',      '单位汇总', 'I6', S(lambda x: x['kind']=='已交税' and hold(x['payer']))),
 ('单位汇总 合计·业主已付',    '单位汇总', 'K6', S(lambda x: x['kind']=='挂靠代收' and hold(x['payee']))),
 ('单位汇总 合计·已转我方',    '单位汇总', 'M6', S(lambda x: x['kind']=='我方收款' and hold(x['payer']))),
 ('单位汇总 合计·管理费已结算','单位汇总', 'O6', S(lambda x: x['kind']=='管理费结算' and hold(x['payee']))),
 ('单位汇总 合计·扣质保金',    '单位汇总', 'P6', S(lambda x: x['kind']=='扣质保金' and hold(x['payer']))),
 ('税费台账 合计·应提增值税',  '税费台账', 'C6', SF('v',  lambda x: x['kind']=='销项开票' and hold(x['payer']))),
 ('税费台账 合计·应提附加税',  '税费台账', 'D6', SF('a',  lambda x: x['kind']=='销项开票' and hold(x['payer']))),
 ('税费台账 合计·应提印花税',  '税费台账', 'E6', SF('st', lambda x: x['kind']=='销项开票' and hold(x['payer']))),
 ('税费台账 合计·应提所得税',  '税费台账', 'F6', SF('ic', lambda x: x['kind']=='销项开票' and hold(x['payer']))),
 ('项目汇总 合计·开票额',      '项目汇总', 'C6', S(lambda x: x['kind']=='销项开票' and x['proj'])),
 ('链条核算 合计·我方应开',    '链条核算', 'M6', SF('due', lambda x: LAST(x) and x['proj'])),
 ('链条核算 合计·我方已开',    '链条核算', 'N6', S(lambda x: x['kind']=='成本票' and x['proj'])),
 ('链条核算 合计·工资扣抵',    '链条核算', 'O6', S(lambda x: x['kind']=='工资扣抵' and x['proj'])),
 ('发票缺口 劳务·应开',        '发票缺口', 'D6', SF('due', lambda x: LAST(x) and x['proj'] and CLS(x['itype'])=='劳务类')),
 ('发票缺口 劳务·已开',        '发票缺口', 'E6', S(lambda x: x['kind']=='成本票' and x['proj'] and CLS(x['itype'])=='劳务类')),
 ('发票缺口 劳务·工资扣抵',    '发票缺口', 'F6', S(lambda x: x['kind']=='工资扣抵' and x['proj'])),
 ('发票缺口 机械·应开',        '发票缺口', 'H6', SF('due', lambda x: LAST(x) and x['proj'] and CLS(x['itype'])=='机械类')),
 ('发票缺口 机械·已开',        '发票缺口', 'I6', S(lambda x: x['kind']=='成本票' and x['proj'] and CLS(x['itype'])=='机械类')),
 ('往来台账 其他应收·发生',    '往来台账', 'C6', SF('reb', lambda x: x['kind']=='销项开票' and hold(x['payer']))),
 ('往来台账 其他应付·发生',    '往来台账', 'H6', S(lambda x: x['kind']=='其他应付发生' and hold(x['payee']))),
 ('往来台账 其他应付·代扣税费','往来台账', 'I6', S(lambda x: x['kind']=='其他应付扣税' and hold(x['payee']))),
 ('代收台账 合计·业主付给单位','代收台账', 'G6', S(lambda x: x['kind']=='挂靠代收')),
 ('代收台账 合计·转给我方',    '代收台账', 'H6', S(lambda x: x['kind']=='我方收款')),
 ('费用统计 合计·收入',        '费用统计', 'E6', R2(sum(j['inc'] for j in JOUR))),
 ('费用统计 合计·支出',        '费用统计', 'F6', R2(sum(j['exp'] for j in JOUR))),
]
wb = openpyxl.load_workbook(XL, data_only=True)
ok = bad = 0
for name, sh, cell, want in CASES:
    got = wb[sh][cell].value
    got = R2(got) if isinstance(got, (int, float)) else got
    if isinstance(got, (int, float)) and abs(got - want) < 0.05:
        ok += 1; print(f'  ✓ {name:28s} {got:>16,.2f}')
    else:
        bad += 1; print(f'  ✗ {name:28s} 表内={got}  独立重算={want:,.2f}')
# 单位子表 == 单位汇总 对应行
print()
usum = wb['单位汇总']
u2row = {usum.cell(r, 1).value: r for r in range(7, 47) if usum.cell(r, 1).value}
for u in HOLD:
    sub = wb.get(f'{u}明细') if hasattr(wb, 'get') else (wb[f'{u}明细'] if f'{u}明细' in wb.sheetnames else None)
    if sub is None: continue
    a, b = sub['C6'].value or 0, usum.cell(u2row[u], 3).value or 0
    if abs(a - b) < 0.05: ok += 1; print(f'  ✓ {u}明细 开票额合计 = 单位汇总 {u} 行  {a:,.2f}')
    else: bad += 1; print(f'  ✗ {u}明细 {a} ≠ 单位汇总 {b}')
# 自营 + 合伙 必须等于全部
w = openpyxl.load_workbook(XL, data_only=True) if False else wb
for col in ('G', 'K', 'P'):
    a3, b3, c3 = (w['项目利润'][f'{col}6'].value or 0, w['项目利润'][f'{col}7'].value or 0,
                  w['项目利润'][f'{col}8'].value or 0)
    if abs(a3 - b3 - c3) < 0.05:
        ok += 1; print(f'  ✓ 项目利润 {col} 列：自营 + 合伙 = 全部  {a3:,.2f}')
    else:
        bad += 1; print(f'  ✗ 项目利润 {col} 列：全部 {a3} ≠ 自营 {b3} + 合伙 {c3}')
print(f'\n{ok}/{ok+bad} 项一致' + ('' if bad == 0 else f'，{bad} 项不符'))
sys.exit(1 if bad else 0)
