# -*- coding: utf-8 -*-
"""核对《外汇资金台账.xlsx》。用法: python3 verify_fx.py <重算过的xlsx> [交付件xlsx]

重算件里银行子表是 #NAME?（LibreOffice 没有 FILTER 函数），那是重算工具的毛病，
交付件里公式是大写的、没动过。这里核对所有能核对的：
汇率取数、折算、逐账户余额、四张报表、核对表勾稽项。
"""
import sys, datetime, collections
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

got = openpyxl.load_workbook(sys.argv[1], data_only=True)
out = openpyxl.load_workbook(sys.argv[2]) if len(sys.argv) > 2 else None
pa, ra, fl, ba = got['参数设置'], got['月度汇率'], got['流水录入'], got['账户余额表']
mr, kr, dr, qr, ck = got['月度汇报表'], got['收支分类报表'], got['资金日报表'], got['区间查询'], got['核对表']
A0, R0, B0, N_ACC, N_ROW, N_CUR, N_CAT = 11, 6, 6, 30, 2000, 8, 30
ok = bad = 0


def chk(n, g, w, tol=0.02):
    global ok, bad
    good = abs(g - w) <= tol if isinstance(g, (int, float)) and isinstance(w, (int, float)) else g == w
    if good:
        ok += 1
    else:
        bad += 1
        print(f'  ✗ {n}: 实际 {g!r}  应为 {w!r}')


def N(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


# ── 把原始资料读成 Python 结构 ────────────────────────────────────────
BASE = pa['B4'].value
YEAR = pa['B5'].value
RPT = pa['B6'].value
EOM = pa['B7'].value
accs = []
for r in range(A0, A0 + N_ACC):
    if not pa.cell(r, 3).value:
        continue
    accs.append(dict(bank=pa.cell(r, 2).value, name=pa.cell(r, 3).value, cur=pa.cell(r, 5).value,
                     open=N(pa.cell(r, 7).value), orate=pa.cell(r, 8).value, row=r))
curs = [ra.cell(5, 2 + j).value for j in range(N_CUR)]
months = [ra.cell(6 + m, 1).value for m in range(12)]
book = {c: {} for c in curs}
eomr = {c: {} for c in curs}
for m in range(12):
    for j, c in enumerate(curs):
        v = ra.cell(6 + m, 2 + j).value
        if isinstance(v, (int, float)):
            book[c][m] = v
        v2 = ra.cell(6 + m, 11 + j).value
        if isinstance(v2, (int, float)):
            eomr[c][m] = v2
cats = {}
for r in range(11, 11 + N_CAT):
    if pa.cell(r, 15).value:
        cats[pa.cell(r, 15).value] = pa.cell(r, 16).value


def rate(d, cur, eom=False):
    """复刻表里的取数规则：往前找最近一个填过的月份。"""
    if cur == BASE:
        return 1.0
    if d is None:
        return None
    src = eomr if eom else book
    best = None
    for m in range(12):
        if months[m] <= d and m in src.get(cur, {}):
            best = src[cur][m]
    if best is None and eom:
        return rate(d, cur, eom=False)
    return best


flows = []
for r in range(R0, R0 + N_ROW):
    if not fl.cell(r, 2).value:
        continue
    d = fl.cell(r, 1).value
    a = fl.cell(r, 2).value
    acc = next((x for x in accs if x['name'] == a), None)
    cur = acc['cur'] if acc else None
    h = fl.cell(r, 9).value
    rt = h if isinstance(h, (int, float)) else rate(d, cur)
    inc, exp = N(fl.cell(r, 7).value), N(fl.cell(r, 8).value)
    net = round(inc - exp, 2)
    flows.append(dict(row=r, d=d, acc=a, cur=cur, cat=fl.cell(r, 5).value,
                      attr=cats.get(fl.cell(r, 5).value), inc=inc, exp=exp, net=net,
                      rate=rt, cny=(None if rt is None else round(net * rt, 2)),
                      prj=fl.cell(r, 16).value, party=fl.cell(r, 17).value))
print(f'账户 {len(accs)} 个，流水 {len(flows)} 笔，币种 {len([c for c in curs if c])} 个')

print('══ 1. 汇率取数与折算 ══')
e = 0
for f in flows:
    g = fl.cell(f['row'], 10).value
    if (g is None and f['rate'] is not None) or (g is not None and f['rate'] is None) \
            or (g is not None and f['rate'] is not None and abs(g - f['rate']) > 1e-9):
        e += 1
        if e <= 3:
            print(f'   第{f["row"]}行 表里 {g} / 应为 {f["rate"]}')
chk('逐笔应用汇率', e, 0)
e = sum(1 for f in flows
        if abs(N(fl.cell(f['row'], 12).value) - N(f['cny'])) > 0.02)
chk('逐笔折人民币 = 原币净额 × 应用汇率', e, 0)
chk('实际成交汇率优先于当月记账汇率',
    fl.cell(12, 10).value, 7.16)
chk('没有汇率的币种，折人民币留空（不是 0）',
    fl.cell(20, 12).value in (None, ''), True)

print('══ 2. 逐账户即时余额 ══')
run, runc = collections.defaultdict(float), collections.defaultdict(float)
for a in accs:
    run[a['name']] = a['open']
    orate = a['orate'] if isinstance(a['orate'], (int, float)) else rate(EOM, a['cur'], eom=True)
    runc[a['name']] = round(a['open'] * orate, 2) if orate else 0
e1 = e2 = 0
for f in flows:
    run[f['acc']] = round(run[f['acc']] + f['net'], 2)
    runc[f['acc']] = round(runc[f['acc']] + N(f['cny']), 2)
    if abs(N(fl.cell(f['row'], 13).value) - run[f['acc']]) > 0.02:
        e1 += 1
    if abs(N(fl.cell(f['row'], 14).value) - runc[f['acc']]) > 0.02:
        e2 += 1
chk('账户余额(原币) 逐行', e1, 0)
chk('账面人民币余额 逐行', e2, 0)

print('══ 3. 账户余额表 ══')
for i, a in enumerate(accs):
    r = B0 + i
    inc = round(sum(f['inc'] for f in flows if f['acc'] == a['name']), 2)
    exp = round(sum(f['exp'] for f in flows if f['acc'] == a['name']), 2)
    end = round(a['open'] + inc - exp, 2)
    chk(f'{a["name"]} 期末原币', N(ba.cell(r, 8).value), end)
    chk(f'{a["name"]} 账面人民币成本', N(ba.cell(r, 11).value), runc[a['name']])
    er = rate(EOM, a['cur'], eom=True)
    if er:
        chk(f'{a["name"]} 期末重估', N(ba.cell(r, 10).value), round(end * er, 2))
        chk(f'{a["name"]} 汇兑损益', N(ba.cell(r, 12).value), round(end * er - runc[a['name']], 2))
TR = B0 + N_ACC
chk('没账户的那几十行是空的，不会算出 0',
    [ba.cell(r, 3).value for r in range(B0 + len(accs), TR)],
    [None] * (TR - B0 - len(accs)))
tot = round(sum(N(ba.cell(B0 + i, 10).value) for i in range(len(accs))), 2)
chk('期末人民币合计', N(ba.cell(TR, 10).value), tot)

print('══ 4. 月度汇报表 ══')
for m in range(1, 13):
    r = 5 + m
    lo = datetime.datetime(YEAR, m, 1)
    hi = datetime.datetime(YEAR + (m == 12), m % 12 + 1, 1)
    inm = [f for f in flows if f['d'] and lo <= f['d'] < hi]
    w_in = round(sum(f['cny'] for f in inm if f['attr'] != '内部转账' and N(f['cny']) > 0), 2)
    w_out = round(-sum(f['cny'] for f in inm if f['attr'] != '内部转账' and N(f['cny']) < 0), 2)
    chk(f'{m}月 收入', N(mr.cell(r, 2).value), w_in)
    chk(f'{m}月 支出', N(mr.cell(r, 3).value), w_out)
    w_end = round(sum(runc[a['name']] for a in accs) * 0, 2)   # 占位，下面单独算
    cum = round(sum(N(a['open']) * 0 for a in accs), 2)
opening_cny = round(sum(N(pa.cell(a['row'], 9).value) for a in accs), 2)
for m in range(1, 13):
    r = 5 + m
    hi = datetime.datetime(YEAR + (m == 12), m % 12 + 1, 1)
    w = round(opening_cny + sum(N(f['cny']) for f in flows if f['d'] and f['d'] < hi), 2)
    chk(f'{m}月 月末资金总额', N(mr.cell(r, 7).value), w)

print('══ 5. 收支分类报表 ══')
K0 = 12
names = list(cats)
e = 0
for i, c in enumerate(names):
    r = K0 + i
    for m in range(1, 13):
        lo = datetime.datetime(YEAR, m, 1)
        hi = datetime.datetime(YEAR + (m == 12), m % 12 + 1, 1)
        sign = -1 if cats[c] == '支出' else 1
        w = round(sign * sum(N(f['cny']) for f in flows
                             if f['cat'] == c and f['d'] and lo <= f['d'] < hi), 2)
        if abs(N(kr.cell(r, 2 + m).value) - w) > 0.02:
            e += 1
chk('类别 × 12 个月 逐格', e, 0)
w_inc = round(sum(N(f['cny']) for f in flows if f['attr'] == '收入'), 2)
w_exp = round(-sum(N(f['cny']) for f in flows if f['attr'] == '支出'), 2)
chk('收入合计', N(kr['O5'].value), w_inc)
chk('支出合计', N(kr['O6'].value), w_exp)
chk('经营净额', N(kr['O8'].value), round(w_inc - w_exp, 2))
chk('内部转账净额', N(kr['O7'].value),
    round(sum(N(f['cny']) for f in flows if f['attr'] == '内部转账'), 2))

print('══ 6. 资金日报表 ══')
d0, d1 = dr['B3'].value, dr['D3'].value
days = (d1 - d0).days + 1
e = 0
for k in range(min(days, 31)):
    r = 7 + k
    day = d0 + datetime.timedelta(days=k)
    w_in = round(sum(N(f['cny']) for f in flows
                     if f['d'] == day and f['attr'] != '内部转账' and N(f['cny']) > 0), 2)
    w_out = round(-sum(N(f['cny']) for f in flows
                       if f['d'] == day and f['attr'] != '内部转账' and N(f['cny']) < 0), 2)
    if abs(N(dr.cell(r, 2).value) - w_in) > 0.02 or abs(N(dr.cell(r, 3).value) - w_out) > 0.02:
        e += 1
chk('按日收付 逐天', e, 0)
D2 = 40
day = dr.cell(D2 + 1, 2).value
for i, a in enumerate(accs):
    r = D2 + 3 + i
    prev = round(a['open'] + sum(f['net'] for f in flows if f['acc'] == a['name'] and f['d'] < day), 2)
    inc = round(sum(f['inc'] for f in flows if f['acc'] == a['name'] and f['d'] == day), 2)
    exp = round(sum(f['exp'] for f in flows if f['acc'] == a['name'] and f['d'] == day), 2)
    chk(f'{a["name"]} 昨日余额', N(dr.cell(r, 4).value), prev)
    chk(f'{a["name"]} 本日余额', N(dr.cell(r, 7).value), round(prev + inc - exp, 2))

print('══ 7. 区间查询 ══')
q0, q1 = qr['B3'].value, qr['D3'].value
sel = [f for f in flows if f['d'] and q0 <= f['d'] <= q1]
chk('收入总额', N(qr['B5'].value),
    round(sum(N(f['cny']) for f in sel if f['attr'] != '内部转账' and N(f['cny']) > 0), 2))
chk('支出总额', N(qr['D5'].value),
    round(-sum(N(f['cny']) for f in sel if f['attr'] != '内部转账' and N(f['cny']) < 0), 2))
prj_row = 43          # 区间查询：类别块 9~38、小计 39，项目块标题 41、表头 42、数据 43 起
e = 0
prjs = [pa.cell(11 + i, 18).value for i in range(40)]
for i, p in enumerate(prjs):
    if not p:
        continue
    r = prj_row + i
    w = round(sum(N(f['cny']) for f in sel if f['prj'] == p and N(f['cny']) > 0), 2)
    if abs(N(qr.cell(r, 2).value) - w) > 0.02:
        e += 1
chk('按项目 收入 逐行', e, 0)

print('══ 8. 核对表 ══')
for r in range(5, 5 + 24):
    if ck.cell(r, 2).value != '勾稽':
        continue
    v = ck.cell(r, 4).value
    chk('勾稽 · ' + str(ck.cell(r, 3).value)[:28], round(N(v), 2), 0)
print('   待办项：', {str(ck.cell(r, 3).value)[:20]: ck.cell(r, 4).value
                   for r in range(5, 5 + 24)
                   if ck.cell(r, 2).value == '待办' and N(ck.cell(r, 4).value) != 0})

if out is not None:
    print('══ 9. 交付件结构 ══')
    subs = [s for s in out.sheetnames if s in [a['name'] for a in accs]]
    chk('银行子表张数 = 账户数', len(subs), len(accs))
    for s in subs:
        v = out[s]['A6'].value
        chk(f'{s} 子表是 FILTER 且套了 IF',
            isinstance(v, ArrayFormula) and 'IF(流水录入!$A$6:$S$2005="","",' in v.text, True)
        chk(f'{s} B2 = 本表账户名', out[s]['B2'].value, s)
    chk('FILTER 没被改成小写',
        sum(1 for s in subs if '_xlfn._xlws.FILTER' in out[s]['A6'].value.text), len(subs))

print(f'\n══ 结果：{ok} 项通过，{bad} 项不通过 ══')
sys.exit(1 if bad else 0)
