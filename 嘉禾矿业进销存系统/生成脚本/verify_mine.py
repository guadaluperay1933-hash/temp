# -*- coding: utf-8 -*-
"""核对《嘉禾矿业进销存系统.xlsx》。

用法: python3 verify_mine.py <重算过的xlsx> [交付件xlsx]
第二个参数给了的话，还会检查【到货批次】确实是公式生成的（不是手打的静态值）。
"""
import sys, os, json, collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
D = json.load(open(os.path.join(HERE, 'mine_data.json'), encoding='utf-8'))
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
wb = wb
it, fe, ib, ob, ba, st, pc, mr, ar, ck = (wb['物料档案'], wb['费用台账'], wb['入库单'], wb['出库单'],
                                          wb['到货批次'], wb['库存台账'], wb['月度盘点'],
                                          wb['领导月报'], wb['应收对账'], wb['核对表'])
D0, S0 = 4, 5
N_IT, N_IN, N_OUT, N_BAT = 1800, 4000, 5000, 200
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
ALL_IN = D['inbound'] + D.get('new_inbound', []) + D.get('buy', [])
ALL_OUT = D.get('outbound', []) + D.get('use', []) + D.get('lend', [])
chk('入库单行数（货柜+新柜+本地采购）', sum(1 for r in range(D0, D0 + N_IN) if ib[f'G{r}'].value), len(ALL_IN))
chk('出库单行数（原表+使用+借调）', sum(1 for r in range(D0, D0 + N_OUT) if ob[f'D{r}'].value), len(ALL_OUT))
chk('费用台账行数（国内+新柜+达市）', sum(1 for r in range(D0, D0 + 800) if fe[f'C{r}'].value),
    len(D['fees']) + len(D.get('new_fees', [])) + len(D.get('local_fees', [])))
chk('物料档案条数', sum(1 for r in range(D0, D0 + N_IT) if it[f'A{r}'].value), len(D['items']))
src_qty = sum(x['qty'] for x in ALL_IN)
chk('入库数量合计', round(sum(N(ib[f'K{r}'].value) for r in range(D0, D0 + N_IN)), 3), round(src_qty, 3), 0.01)
src_amt = sum(x.get('amount', 0) for x in ALL_IN)
chk('入库原币金额合计', round(sum(N(ib[f'O{r}'].value) for r in range(D0, D0 + N_IN)), 2), round(src_amt, 2), 0.05)
chk('费用原币合计', round(sum(N(fe[f'G{r}'].value) for r in range(D0, D0 + 800)), 2),
    round(sum(x['amount'] for x in D['fees']) + sum(x['cny'] for x in D.get('new_fees', []))
          + sum(x['tzs'] for x in D.get('local_fees', [])), 2), 1)
chk('达市到港费用折先令合计', round(sum(N(fe[f'I{r}'].value) for r in range(D0, D0 + 800)
                                       if fe[f'F{r}'].value == 'TZS')),
    round(sum(x['tzs'] for x in D.get('local_fees', []))), 2)

print('══ 2. 币种折算 ══')
chk('入库货值·先令 = 原币 × 汇率',
    round(sum(N(ib[f'P{r}'].value) for r in range(D0, D0 + N_IN))),
    round(sum(round(x.get('amount', 0) * RATE) for x in ALL_IN)), 5)

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
for r in range(D0, D0 + 800):
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
# 库存台账只收「账期之内 + 类型对 + 状态正常」的行，核对时要用同一把尺子，
# 否则 312 行没日期的、以及类型不是采购入库的会算进来
import datetime as _dt
P0, P1 = _dt.datetime(2026, 1, 1), _dt.datetime(2026, 8, 31)
def _in(d):
    return isinstance(d, _dt.datetime) and P0 <= d <= P1
chk('库存台账 本月入库金额 = 入库单同口径合计', round(N(st['J4'].value)),
    round(sum(N(ib[f'V{r}'].value) for r in range(D0, D0 + N_IN)
              if ib[f'X{r}'].value == '采购入库' and ib[f'Z{r}'].value == '正常'
              and _in(ib[f'C{r}'].value))), 3)
chk('库存台账 自用消耗 = 出库单同口径合计', round(N(st['Q4'].value)),
    round(sum(N(ob[f'U{r}'].value) for r in range(D0, D0 + N_OUT)
              if ob[f'K{r}'].value == '自用消耗' and ob[f'W{r}'].value == '正常'
              and _in(ob[f'C{r}'].value))), 3)
chk('借调不进出库成本',
    round(sum(N(ob[f'U{r}'].value) for r in range(D0, D0 + N_OUT)
              if ob[f'K{r}'].value == '矿区借调')), 0)

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
MR0 = 26
chk('月报 各矿区合计', round(N(mr[f'C{MR0+7}'].value)),
    round(sum(N(mr[f'C{MR0+i}'].value) for i in range(7))))


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
chk('借调笔数', sum(1 for r in range(D0, D0 + N_OUT) if ob[f'K{r}'].value == '矿区借调'),
    len(D.get('lend', [])))
chk('重复嫌疑物料数', sum(1 for r in range(D0, D0 + 200)
                          if wb['参数表'].cell(r, 23).value), len(D.get('dup_codes', [])))

print('══ 9. 到货批次是公式长出来的，不是手打的 ══')
S0, N_BAT = 5, 200
IN_END, F_END = D0 + N_IN - 1, D0 + 800 - 1
# 入库单 + 费用台账里到底有几个不重复的批次
bset, border = set(), []
for r in range(D0, IN_END + 1):
    v = ib.cell(r, 5).value
    if v and v not in bset:
        bset.add(v); border.append(v)
fee_only = []
for r in range(D0, F_END + 1):
    v = fe.cell(r, 3).value
    if v and v not in bset:
        bset.add(v); fee_only.append(v)
got_b = [ba.cell(r, 1).value for r in range(S0, S0 + N_BAT) if ba.cell(r, 1).value]
chk('批次行数 = 入库单+费用台账里的不重复批次数', len(got_b), len(border) + len(fee_only))
chk('批次顺序 = 在入库单里第一次出现的顺序', got_b, border + fee_only)
chk('入库单的「批次首现」序号连号',
    [ib.cell(r, 30).value for r in range(D0, IN_END + 1) if ib.cell(r, 30).value],
    list(range(1, len(border) + 1)))
e1 = e2 = e3 = 0
first_row = {}
for r in range(D0, IN_END + 1):
    v = ib.cell(r, 5).value
    if v and v not in first_row:
        first_row[v] = r
for i, b in enumerate(got_b):
    r = S0 + i
    fr = first_row.get(b)
    if fr:
        if ba.cell(r, 3).value != ib.cell(fr, 3).value:
            e1 += 1
        if ba.cell(r, 4).value != ib.cell(fr, 4).value:
            e2 += 1
    want = '' if str(b).startswith('本地') else str(b).rsplit('-', 1)[0]
    if (ba.cell(r, 2).value or '') != want:
        e3 += 1
chk('到港日期 = 该批次第一笔入库的日期', e1, 0)
chk('货源自动带出来', e2, 0)
chk('柜号从批次号里截出来（本地批次留空）', e3, 0)
chk('只有费用没有货的批次也被捞出来了',
    all(ba.cell(S0 + got_b.index(b), 5).value in (0, None) for b in fee_only), True)

if len(sys.argv) > 2:
    out = openpyxl.load_workbook(sys.argv[2])
    ob2 = out['到货批次']
    for c, nm in ((1, '批次号'), (2, '集装箱号'), (3, '到港日期'), (4, '货源')):
        vals = [ob2.cell(r, c).value for r in range(S0, S0 + N_BAT)]
        chk(f'{nm}列整列都是公式', all(isinstance(v, str) and v.startswith('=') for v in vals), True)
    chk('入库单「批次首现」列是公式',
        str(out['入库单'].cell(D0, 30).value or '').startswith('='), True)
    chk('费用台账「批次首现」列是公式',
        str(out['费用台账'].cell(D0, 13).value or '').startswith('='), True)

print(f'\n══ 结果：{ok} 项通过，{bad} 项不通过 ══')
sys.exit(1 if bad else 0)
