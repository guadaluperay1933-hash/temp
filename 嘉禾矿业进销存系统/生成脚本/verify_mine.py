# -*- coding: utf-8 -*-
"""核对《嘉禾矿业进销存系统.xlsx》。用法: python3 verify_mine.py <重算过的xlsx>"""
import sys, os, json, collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
D = json.load(open(os.path.join(HERE, 'mine_data.json'), encoding='utf-8'))
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
it, fe, ib, ob, ba, st, pc, mr, ar, ck = (wb['物料档案'], wb['费用台账'], wb['入库单'], wb['出库单'],
                                          wb['到货批次'], wb['库存台账'], wb['月度盘点'],
                                          wb['领导月报'], wb['应收对账'], wb['核对表'])
D0, S0 = 4, 5
N_IT, N_IN, N_OUT, N_BAT = 1200, 2500, 4000, 200
ok = bad = 0
RATE = 370.0

def chk(n, g, w, tol=1.0):
    global ok, bad
    good = abs(g - w) <= tol if isinstance(g, (int, float)) and isinstance(w, (int, float)) else g == w
    if good:
        ok += 1
    else:
        bad += 1
        print(f'  ✗ {n}: 实际 {g!r}  应为 {w!r}')

N = lambda v: v if isinstance(v, (int, float)) else 0

print('══ 1. 搬运完整性 ══')
chk('物料档案条数', sum(1 for r in range(D0, D0 + N_IT) if it[f'A{r}'].value), len(D['items']))
chk('入库单行数', sum(1 for r in range(D0, D0 + N_IN) if ib[f'G{r}'].value), len(D['inbound']))
chk('出库单行数', sum(1 for r in range(D0, D0 + N_OUT) if ob[f'D{r}'].value), len(D['outbound']))
chk('费用台账行数', sum(1 for r in range(D0, D0 + 500) if fe[f'C{r}'].value), len(D['fees']))
chk('到货批次条数', sum(1 for r in range(S0, S0 + N_BAT) if ba[f'A{r}'].value), len(D['batches']))
src_qty = sum(x['qty'] for x in D['inbound'])
chk('入库数量合计', round(sum(N(ib[f'K{r}'].value) for r in range(D0, D0 + N_IN)), 3), round(src_qty, 3), 0.01)
src_amt = sum(x['amount'] for x in D['inbound'])
chk('入库原币金额合计', round(sum(N(ib[f'O{r}'].value) for r in range(D0, D0 + N_IN)), 2), round(src_amt, 2), 0.05)
chk('费用原币合计', round(sum(N(fe[f'G{r}'].value) for r in range(D0, D0 + 500)), 2),
    round(sum(x['amount'] for x in D['fees']), 2), 0.01)

print('══ 2. 币种折算 ══')
chk('入库货值·先令 = 原币 × 汇率',
    round(sum(N(ib[f'P{r}'].value) for r in range(D0, D0 + N_IN))),
    round(sum(round(x['amount'] * RATE) for x in D['inbound'])), 5)
chk('费用·先令 = 原币 × 汇率',
    round(sum(N(fe[f'I{r}'].value) for r in range(D0, D0 + 500))),
    round(sum(round(x['amount'] * RATE) for x in D['fees'])), 5)

print('══ 3. 海运费分摊 ══')
byb = collections.defaultdict(lambda: dict(val=0, fee=0, noprice=0, alloc=0, rows=0))
for r in range(D0, D0 + N_IN):
    k = ib[f'E{r}'].value
    if not k or not ib[f'G{r}'].value:
        continue
    byb[k]['val'] += N(ib[f'P{r}'].value)
    byb[k]['alloc'] += N(ib[f'U{r}'].value)
    byb[k]['rows'] += 1
    if ib[f'Y{r}'].value == '待补价':
        byb[k]['noprice'] += 1
for r in range(D0, D0 + 500):
    k = fe[f'C{r}'].value
    if k:
        byb[k]['fee'] += N(fe[f'I{r}'].value)
for k, v in byb.items():
    if v['fee'] == 0:
        chk(f'{k} 没费用就不该有分摊', round(v['alloc']), 0)
    elif v['noprice'] > 0:
        chk(f'{k} 有 {v["noprice"]} 行待补价 → 整批挂账不摊', round(v['alloc']), 0)
    else:
        chk(f'{k} 分摊合计 = 费用合计（一分不差）', round(v['alloc']), round(v['fee']), 0.5)
chk('全表 已摊 + 挂账 = 可摊费用总额',
    round(N(ba['J4'].value) + N(ba['K4'].value)), round(N(ba['H4'].value)), 1)

print('══ 4. 库存勾稽（期初+入库+盘盈+退库−自用−外销−盘亏−退货 = 期末）══')
for tag, cols, tot in (('金额', ('H', 'J', 'M', 'O', 'Q', 'S', 'U', 'W'), 'Y'),
                       ('数量', ('G', 'I', 'L', 'N', 'P', 'R', 'T', 'V'), 'X')):
    want = sum(N(st[f'{c}4'].value) * (1 if i < 4 else -1) for i, c in enumerate(cols))
    chk(f'{tag}勾稽', round(N(st[f'{tot}4'].value), 2), round(want, 2), 0.02)
e = 0
for r in range(S0, S0 + N_IT):
    if not st[f'A{r}'].value:
        continue
    w = (N(st[f'G{r}'].value) + N(st[f'I{r}'].value) + N(st[f'L{r}'].value) + N(st[f'N{r}'].value)
         - N(st[f'P{r}'].value) - N(st[f'R{r}'].value) - N(st[f'T{r}'].value) - N(st[f'V{r}'].value))
    if abs(N(st[f'X{r}'].value) - w) > 0.002:
        e += 1
chk('库存台账逐行数量勾稽', e, 0)
chk('期末金额 = 入库成本合计（首期没有出库）',
    round(N(st['Y4'].value)), round(sum(N(ib[f'V{r}'].value) for r in range(D0, D0 + N_IN))
                                    - sum(N(ob[f'S{r}'].value) for r in range(D0, D0 + N_OUT))), 2)

print('══ 5. 月加权单位成本 ══')
e = 0
for r in range(S0, S0 + N_IT):
    if not st[f'A{r}'].value:
        continue
    den = N(st[f'G{r}'].value) + N(st[f'I{r}'].value)
    w = round((N(st[f'H{r}'].value) + N(st[f'J{r}'].value)) / den, 2) if den else 0
    if abs(N(st[f'K{r}'].value) - w) > 0.02:
        e += 1
chk('月加权单位成本逐行', e, 0)

print('══ 6. 售价规则 ══')
e_a = e_b = n_a = n_b = 0
for r in range(D0, D0 + N_IT):
    if not it[f'A{r}'].value:
        continue
    src, p, sug = it[f'H{r}'].value, N(it[f'I{r}'].value), N(it[f'J{r}'].value)
    if src == 'A国内':
        n_a += 1
        if abs(sug - round(p * 2 * RATE)) > 1:
            e_a += 1
    elif src == 'B当地':
        n_b += 1
        if abs(sug - round(p * 1.03)) > 1:
            e_b += 1
chk(f'国内货 售价=进价×2×370（{n_a} 个）', e_a, 0)
chk(f'当地货 售价=进价×1.03（{n_b} 个）', e_b, 0)

print('══ 7. 月报与核对表 ══')
chk('月报 勾稽差额 = 0', round(N(mr['B46'].value)), 0)
chk('月报 采购总成本 = 库存台账入库金额', round(N(mr['B12'].value)), round(N(st['J4'].value)), 2)
chk('月报 折人民币 = 先令 / 汇率', round(N(mr['B13'].value), 2), round(N(mr['B12'].value) / RATE, 2), 0.02)
chk('月报 两矿自用合计', round(N(mr['B27'].value)), round(N(mr['B25'].value) + N(mr['B26'].value)))
# 首期就有 3 条勾稽红，全部来自原表本来就缺的数据，不是公式错：
#   9  外销没售价 1 笔 —— 原表 439 行那只卖给三矿的空柜从来没写过价
#   14 入库数量要大于 0 的 4 行 —— 土工膜 3 个柜 + 发电机组，原表数量栏是空的
#   16 上面这些的合计
red = [ck[f'A{r}'].value for r in range(6, 23) if ck[f'E{r}'].value == '※异常']
chk('核对表 勾稽红项 = 3（全部是原表缺数据，公式没错）', len(red), 3)
print('   勾稽红项：' + ' | '.join(x[:30] for x in red))
print(f'   勾稽 {ck["B3"].value}/{ck["C3"].value} 通过；待办项：'
      + ', '.join(f'{ck[f"A{r}"].value.split(" ")[0]}={ck[f"C{r}"].value}'
                  for r in range(6, 23) if ck[f'E{r}'].value == '△待补'))

print('══ 8. 红字提示 ══')
fi = collections.Counter(ib[f'AC{r}'].value for r in range(D0, D0 + N_IN)
                         if isinstance(ib[f'AC{r}'].value, str) and ib[f'AC{r}'].value[:1] in '※△')
fo = collections.Counter(ob[f'W{r}'].value for r in range(D0, D0 + N_OUT)
                         if isinstance(ob[f'W{r}'].value, str) and ob[f'W{r}'].value[:1] in '※△')
for k, v in fi.most_common():
    print(f'   入库 {v:4d} × {k}')
for k, v in fo.most_common():
    print(f'   出库 {v:4d} × {k}')
chk('入库真错（※）只剩原表本来就没填数量的那几行',
    sum(v for k, v in fi.items() if k.startswith('※')), 4)

print(f'\n══ 结果：{ok} 项通过，{bad} 项不通过 ══')
sys.exit(1 if bad else 0)
