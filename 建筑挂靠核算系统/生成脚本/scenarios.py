# -*- coding: utf-8 -*-
"""换参数跑一轮：年度 / 起止日期 / 选择单位 改了以后，各查询表算出来的数
   必须跟 Python 按同样条件重算的数一致。跑法：python3 scenarios.py <xlsx>"""
import sys, os, json, shutil, subprocess, datetime as dt, openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', '建筑挂靠业务核算系统.xlsx')
TMP = os.environ.get('GK_TMP', '/tmp/claude-0/-home-user-temp/9101e7cb-4f82-53a7-bbe8-41b0f93b12b0/scratchpad/a059')
RECALC = '/mnt/skills/public/xlsx/scripts/recalc.py'
D2 = json.load(open(os.path.join(HERE, 'ev2.json')))
EV = [e for e in D2['events'] if e['kind'] != '预提税费']
PASS_ROWS, DED_ROWS, WAGE_ROWS = D2['pass_rows'], D2['ded_rows'], D2.get('wage', [])
JOUR = json.load(open(os.path.join(HERE, 'jour.json')))
R2 = lambda x: round(x + 0.0, 2)
MFEE_RATE = {}
for _e in EV:
    if _e['kind'] == '销项开票' and _e['amt']:
        _k = (_e['payer'], round(_e['amt'], 2))
        MFEE_RATE[_k] = max(MFEE_RATE.get(_k, 0.0), round((_e.get('mfee') or 0) / _e['amt'], 6))
ETYPE_FIX = {'账户费年费': '账户年费', '耗材费用': '耗材费'}
KMAP = {'挂靠单位代收': '挂靠代收'}
HOLD = ['德誉嘉', '迅驰', '华城', '康欣', '金沁', '安锐', '湖南锦泰', '杰华']
CLS = lambda t: '劳务类' if t in ('劳务票', '土建票', '安装票') else ('机械类' if t == '机械设备票' else
      ('材料类' if t == '材料票' else ('其他' if t else '')))

FLOW = []
for e in EV:
    if e.get('count_in', '是') != '是': continue
    amt = float(e['amt']); kind = KMAP.get(e['kind'], e['kind'])
    mfee = R2(amt * MFEE_RATE.get((e['payer'], round(amt, 2)), 0.0)) if kind == '销项开票' else 0.0
    FLOW.append(dict(d=e['date'], kind=kind, payer=e['payer'], payee=e['payee'], proj=e['proj'],
                     amt=amt, mfee=mfee, due=R2(amt - mfee) if kind == '销项开票' else 0.0,
                     reb=R2(e.get('rebate', 0) or 0), itype=e.get('itype', '')))
for x in PASS_ROWS: FLOW.append(dict(d='2026-03-27', kind='其他应付发生', payer='迅驰', payee='迅驰',
                                     proj=x['proj'], amt=x['amt'], mfee=0, due=0, reb=0, itype=''))
for x in WAGE_ROWS: FLOW.append(dict(d='2026-06-30', kind='工资扣抵', payer='泓普', payee='康欣',
                                     proj=x['proj'], amt=x['amt'], mfee=0, due=0, reb=0, itype='劳务票'))
for x in DED_ROWS: FLOW.append(dict(d='2026-03-27', kind='其他应付扣税', payer='泓普', payee='迅驰',
                                    proj='', amt=x['amt'], mfee=0, due=0, reb=0, itype=''))
sale = [x for x in FLOW if x['kind'] == '销项开票']
recv = {(x['proj'], x['payee']) for x in sale}
LAST = lambda x: x['kind'] == '销项开票' and (x['proj'], x['payer']) not in recv
def rng(a, b):
    return [x for x in FLOW if (a is None or x['d'] >= a) and (b is None or x['d'] <= b)]
def jrng(a, b):
    return [j for j in JOUR if (a is None or j['date'] >= a) and (b is None or j['date'] <= b)]

SCEN = [
 ('①全部期间',          None, None, None, None),
 ('②年度=2026',         '单位汇总', 'B3', 2026, ('2026-01-01', '2026-12-31')),
 ('③年度=2025',         '单位汇总', 'B3', 2025, ('2025-01-01', '2025-12-31')),
 ('④2026-06 单月',      '单位汇总', ('D3', 'F3'), (dt.date(2026,6,1), dt.date(2026,6,30)), ('2026-06-01', '2026-06-30')),
 ('⑤2026上半年',        '单位汇总', ('D3', 'F3'), (dt.date(2026,1,1), dt.date(2026,6,30)), ('2026-01-01', '2026-06-30')),
]
SHEETS_YEAR = ['单位汇总', '项目汇总', '链条核算', '发票缺口', '往来台账', '税费台账', '代收台账', '费用统计',
               '单位项目明细']
fails = 0
for name, sh, cell, val, (lo, hi) in [(s[0], s[1], s[2], s[3], s[4] or (None, None)) for s in SCEN]:
    path = os.path.join(TMP, f'sc_{abs(hash(name))%10000}.xlsx')
    shutil.copy(SRC, path)
    wb = openpyxl.load_workbook(path)
    if sh:
        for s2 in SHEETS_YEAR:
            w = wb[s2]
            off = 2 if s2 == '单位项目明细' else 0     # 单位项目明细 多一个「选择单位」在前面
            if isinstance(cell, tuple):
                for c0, v0 in zip(cell, val):
                    col = chr(ord(c0[0]) + off); w[f'{col}3'] = v0
            else:
                col = chr(ord(cell[0]) + off); w[f'{col}3'] = val
        for u in HOLD:                                  # 单位子表：年度在 D3、起止在 F3/H3
            w = wb[f'{u}明细']
            if isinstance(cell, tuple):
                for c0, v0 in zip(cell, val): w[f'{chr(ord(c0[0])+2)}3'] = v0
            else: w[f'{chr(ord(cell[0])+2)}3'] = val
    wb.save(path); wb.close()
    out = subprocess.run(['python3', RECALC, path], capture_output=True, text=True).stdout
    st = json.loads(out[out.index('{'):out.rindex('}')+1])
    R = openpyxl.load_workbook(path, data_only=True)
    F = rng(lo, hi); J = jrng(lo, hi)
    hold = lambda u: u in HOLD
    S = lambda f, fld='amt': R2(sum(x[fld] for x in F if f(x)))
    want = {
      ('单位汇总','C6'): S(lambda x: x['kind']=='销项开票' and hold(x['payer'])),
      ('单位汇总','D6'): S(lambda x: x['kind']=='销项开票' and hold(x['payer']), 'mfee'),
      ('单位汇总','K6'): S(lambda x: x['kind']=='挂靠代收' and hold(x['payee'])),
      ('项目汇总','C6'): S(lambda x: x['kind']=='销项开票' and x['proj']),
      ('链条核算','M6'): S(lambda x: LAST(x) and x['proj'], 'due'),
      ('链条核算','N6'): S(lambda x: x['kind']=='成本票' and x['proj']),
      ('发票缺口','D6'): S(lambda x: LAST(x) and x['proj'] and CLS(x['itype'])=='劳务类', 'due'),
      ('发票缺口','H6'): S(lambda x: LAST(x) and x['proj'] and CLS(x['itype'])=='机械类', 'due'),
      ('发票缺口','F6'): S(lambda x: x['kind']=='工资扣抵' and x['proj']),
      ('往来台账','C6'): S(lambda x: x['kind']=='销项开票' and hold(x['payer']), 'reb'),
      ('往来台账','H6'): S(lambda x: x['kind']=='其他应付发生' and hold(x['payee'])),
      ('往来台账','I6'): S(lambda x: x['kind']=='其他应付扣税' and hold(x['payee'])),
      ('代收台账','G6'): S(lambda x: x['kind']=='挂靠代收'),
      ('代收台账','H6'): S(lambda x: x['kind']=='我方收款'),
      ('费用统计','E6'): R2(sum(j['inc'] for j in J)),
      ('费用统计','F6'): R2(sum(j['exp'] for j in J)),
      ('康欣明细','C6'): S(lambda x: x['kind']=='销项开票' and x['payer']=='康欣' and x['proj']),
      ('金沁明细','C6'): S(lambda x: x['kind']=='销项开票' and x['payer']=='金沁' and x['proj']),
    }
    bad = []
    for (s2, c2), w2 in want.items():
        g2 = R2(R[s2][c2].value) if isinstance(R[s2][c2].value, (int, float)) else R[s2][c2].value
        if not isinstance(g2, (int, float)) or abs(g2 - w2) >= 0.05: bad.append(f'{s2}!{c2} 表内={g2} 应为={w2:,.2f}')
    flag = '✓' if (st['total_errors'] == 0 and not bad) else '✗'
    if flag == '✗': fails += 1
    print(f'{flag} {name:16s} 重算 {st["status"]} 错误 {st["total_errors"]:3d}  校验 {len(want)-len(bad)}/{len(want)}')
    for b in bad: print('      ', b)
    R.close()
print(('\n全部场景通过' if not fails else f'\n{fails} 个场景不通过'))
sys.exit(1 if fails else 0)
