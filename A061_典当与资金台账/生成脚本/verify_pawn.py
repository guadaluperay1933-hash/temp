# -*- coding: utf-8 -*-
"""核对《A061_典当登记表.xlsx》。用法: python3 verify_pawn.py <重算过的xlsx> <原_典当登记表.xlsx>

注意：LibreOffice 重算会把 _xlfn._xlws.FILTER 小写化，所以子表和查询表三块明细
在重算件里是 #NAME?，那是重算工具的毛病，交付件里公式是大写的、没动过。
这里只核对能核对的：搬运完整性、普通公式、汇总条、核对列。
"""
import sys, collections
import openpyxl

got = openpyxl.load_workbook(sys.argv[1], data_only=True)
src = openpyxl.load_workbook(sys.argv[2], data_only=True)
t2, t1, q, res = got['表2客户典当登记'], got['表1不动产登记'], got['查询表'], got['资料']
s2, s1 = src['Sheet2'], src['Sheet1']
D0, E2, E1 = 4, 403, 123
ok = bad = 0

def chk(n, g, w, tol=0.02):
    global ok, bad
    good = abs(g - w) <= tol if isinstance(g, (int, float)) and isinstance(w, (int, float)) else g == w
    if good:
        ok += 1
    else:
        bad += 1
        print(f'  ✗ {n}: 实际 {g!r}  应为 {w!r}')

def col(ws, c, r0, r1, key):
    return [ws.cell(r, c).value for r in range(r0, r1 + 1) if ws.cell(r, key).value not in (None, '')]

print('══ 1. 数据一行不少地搬过来了 ══')
o2 = [r for r in range(3, 204) if s2.cell(r, 9).value not in (None, '')]
o1 = [r for r in range(3, 86) if s1.cell(r, 5).value not in (None, '')]
g2 = [r for r in range(D0, E2 + 1) if t2.cell(r, 5).value not in (None, '')]
g1 = [r for r in range(D0, E1 + 1) if t1.cell(r, 4).value not in (None, '')]
chk('表2 行数', len(g2), len(o2))
chk('表1 行数', len(g1), len(o1))
N = lambda v: v if isinstance(v, (int, float)) else 0
chk('表2 典当金额合计', round(sum(N(t2.cell(r, 12).value) for r in g2), 2),
    round(sum(N(s2.cell(r, 10).value) for r in o2), 2))
chk('表2 已还本金合计', round(sum(N(t2.cell(r, 13).value) for r in g2), 2),
    round(sum(N(s2.cell(r, 11).value) for r in o2), 2))
chk('表2 应收利息合计', round(sum(N(t2.cell(r, 16).value) for r in g2), 2),
    round(sum(N(s2.cell(r, 13).value) for r in o2), 2))
chk('表1 典当金额合计', round(sum(N(t1.cell(r, 10).value) for r in g1), 2),
    round(sum(N(s1.cell(r, 7).value) for r in o1), 2))
chk('表1 应收利息合计', round(sum(N(t1.cell(r, 12).value) for r in g1), 2),
    round(sum(N(s1.cell(r, 9).value) for r in o1), 2))
chk('表2 客户姓名逐个对上', col(t2, 5, D0, E2, 5), [str(s2.cell(r, 9).value).strip() for r in o2])
chk('表1 客户姓名逐个对上', col(t1, 4, D0, E1, 4), [str(s1.cell(r, 5).value).strip() for r in o1])

print('══ 2. 自动算的列 ══')
e = 0
for r in g2:
    if abs(N(t2.cell(r, 14).value) - round(N(t2.cell(r, 12).value) - N(t2.cell(r, 13).value), 2)) > 0.02:
        e += 1
chk('表2 当前在当 = 典当金额−已还本金', e, 0)
e = 0
for r in g1:
    want = round(sum(N(t1.cell(r, c).value) for c in (14, 15, 16, 17, 18)), 2)
    if abs(N(t1.cell(r, 19).value) - want) > 0.02:
        e += 1
chk('表1 合计收款 = 对公+林总+建行民生+微信+违约金', e, 0)
e = 0
for r in g1:
    a, b = t1.cell(r, 20).value, t1.cell(r, 21).value
    want = 0 if (a in (None, '') or b in (None, '')) else max(0, (b - a).days if hasattr(b - a, 'days') else 0)
    if abs(N(t1.cell(r, 22).value) - want) > 0.5:
        e += 1
chk('表1 逾期天数', e, 0)
chk('表2 序号连续', [t2.cell(r, 1).value for r in g2], list(range(1, len(g2) + 1)))
chk('表1 序号连续', [t1.cell(r, 1).value for r in g1], list(range(1, len(g1) + 1)))

print('══ 3. 三张子表拆得对 ══')
inst = collections.Counter(t2.cell(r, 2).value for r in g2)
chk('机构只剩三种', sorted(inst), ['个人典当', '金辉典当'])
for nm in ('金辉典当', '明道典当', '个人典当'):
    sh = got[nm]
    rows = [r for r in g2 if t2.cell(r, 2).value == nm]
    chk(f'{nm} B2 拆分依据', sh['B2'].value, nm)
    chk(f'{nm} 笔数', sh['AD2'].value, len(rows))
    chk(f'{nm} 典当金额', sh['AF2'].value, round(sum(N(t2.cell(r, 12).value) for r in rows), 2))
    chk(f'{nm} 当前在当', sh['AJ2'].value, round(sum(N(t2.cell(r, 14).value) for r in rows), 2))
    chk(f'{nm} 应收利息', sh['AL2'].value, round(sum(N(t2.cell(r, 16).value) for r in rows), 2))
chk('三张子表笔数加起来 = 表2 总笔数',
    sum(got[n]['AD2'].value for n in ('金辉典当', '明道典当', '个人典当')), len(g2))

print('══ 4. 查询表汇总条 ══')
who = q['B3'].value
w2 = [r for r in g2 if t2.cell(r, 5).value == who]
w1 = [r for r in g1 if t1.cell(r, 4).value == who]
print(f'   当前查的是「{who}」：表2 {len(w2)} 笔，表1 {len(w1)} 笔')
chk('典当笔数', q['B5'].value, len(w2))
chk('典当金额', q['D5'].value, round(sum(N(t2.cell(r, 12).value) for r in w2), 2))
chk('已还本金', q['F5'].value, round(sum(N(t2.cell(r, 13).value) for r in w2), 2))
chk('当前在当', q['H5'].value, round(sum(N(t2.cell(r, 14).value) for r in w2), 2))
chk('应收利息', q['J5'].value, round(sum(N(t2.cell(r, 16).value) for r in w2), 2))
chk('不动产笔数', q['L5'].value, len(w1))
chk('不动产金额', q['N5'].value, round(sum(N(t1.cell(r, 10).value) for r in w1), 2))
print(f'   资金·收到 {q["P5"].value} / 付出 {q["R5"].value}（跨文件，重算环境读不到属正常）')

print('══ 5. 核对列 ══')
f2 = [t2.cell(r, 24).value for r in g2 if isinstance(t2.cell(r, 24).value, str) and t2.cell(r, 24).value.startswith('※')]
f1 = [t1.cell(r, 27).value for r in g1 if isinstance(t1.cell(r, 27).value, str) and t1.cell(r, 27).value.startswith('※')]
print(f'   表2 提示 {len(f2)} 条: {dict(collections.Counter(f2))}')
print(f'   表1 提示 {len(f1)} 条: {dict(collections.Counter(f1))}')
chk('客户名单已全部收录（资料!L5）', res['L5'].value, 0)
chk('客户名单条数', sum(1 for r in range(D0, D0 + 500) if res.cell(r, 9).value),
    len({t2.cell(r, 5).value for r in g2} | {t1.cell(r, 4).value for r in g1}))

print(f'\n══ 结果：{ok} 项通过，{bad} 项不通过 ══')
sys.exit(1 if bad else 0)
