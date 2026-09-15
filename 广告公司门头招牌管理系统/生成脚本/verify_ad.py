# -*- coding: utf-8 -*-
"""对重算后的工作簿逐项核对。用法: python3 verify_ad.py <重算过的xlsx>"""
import sys, os, json, collections, datetime
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = json.load(open(os.path.join(HERE, 'ad_data.json'), encoding='utf-8'))
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
ok = bad = 0
def chk(name, got, want, tol=0.02):
    global ok, bad
    if isinstance(want, (int, float)) and isinstance(got, (int, float)):
        good = abs(got - want) <= tol
    else:
        good = got == want
    if good:
        ok += 1
    else:
        bad += 1
        print(f'  ✗ {name}: 实际 {got!r}  应为 {want!r}')
    return good

def cell(sheet, addr):
    v = wb[sheet][addr].value
    return 0 if v is None else v

D0, S0 = 4, 5
print('══ 1. 资金：账户余额 vs 日记账原始数据 ══')
inc = sum(x['income'] for x in DATA['journal'])
out = sum(x['outgo'] for x in DATA['journal'])
chk('微信1 累计收入', cell('账户余额', 'C7'), round(inc, 2))
chk('微信1 累计支出', cell('账户余额', 'D7'), round(out, 2))
chk('微信1 当前余额', cell('账户余额', 'E7'), round(inc - out, 2))
chk('合计行 = 各账户之和', cell('账户余额', 'E14'), round(inc - out, 2))
chk('微信1 笔数', cell('账户余额', 'H7'), len(DATA['journal']))
chk('原表最后一天的结余对得上', cell('账户余额', 'E7'), 4795.83)
# 日记账最后一行的「本账户余额」应等于账户余额表
jw = wb['资金日记账']
last = D0 + len(DATA['journal']) - 1
chk('日记账最后一行余额 = 账户余额表', jw[f'L{last}'].value, cell('账户余额', 'E7'))
chk('日记账最后一行全账户合计', jw[f'M{last}'].value, cell('账户余额', 'E14'))

print('══ 2. 项目成本：项目核算 vs 三张来源表 ══')
cost = collections.defaultdict(lambda: collections.defaultdict(float))
M2G = {'主材': '成本-主材', '外发加工': '成本-主材', '辅材': '成本-辅材',
       '五金耗材': '成本-辅材', '其他材料': '成本-其他'}
W2G = {'设计费': '成本-人工', '制作工时': '成本-人工', '安装工时': '成本-人工',
       '加班费': '成本-人工', '外叫小工': '成本-人工', '其他人工': '成本-人工',
       '吊车/租车': '成本-车运', '运费/搬运': '成本-车运'}
for x in DATA['purchases']:
    cost[x['project']][M2G[x['mcat']]] += x['amount']
for x in DATA['labour']:
    cost[x['project']][W2G[x['wcat']]] += x['amount']
CAT2G = {}
for a, b in [('主材采购', '成本-主材'), ('外发加工', '成本-主材'), ('辅材采购', '成本-辅材'),
             ('网购材料耗材', '成本-辅材'), ('五金耗材', '成本-辅材'), ('设计费', '成本-人工'),
             ('制作工资', '成本-人工'), ('安装工资', '成本-人工'), ('加班费', '成本-人工'),
             ('外叫小工', '成本-人工'), ('运费/快递费', '成本-车运'),
             ('车费/油费/过路费', '成本-车运'), ('吊车/租车费', '成本-车运'),
             ('项目餐费补助', '成本-其他'), ('其他直接费用', '成本-其他')]:
    CAT2G[a] = b
for x in DATA['journal']:
    g = CAT2G.get(x['cat'])
    if g and x['project']:
        cost[x['project']][g] += x['outgo'] - x['income']
kw = wb['项目核算']
name2row = {}
for i in range(400):
    r = S0 + i
    nm = kw[f'B{r}'].value
    if nm:
        name2row[nm] = r
chk('项目核算行数 = 导入项目数', len(name2row), len(DATA['projects']))
GCOL = {'成本-主材': 'L', '成本-辅材': 'M', '成本-人工': 'N', '成本-车运': 'O', '成本-其他': 'P'}
miss = 0
for p in DATA['projects']:
    r = name2row.get(p['name'])
    if not r:
        print('  ✗ 项目核算里找不到', p['name']); bad += 1; continue
    for g, col in GCOL.items():
        want = round(cost[p['name']].get(g, 0.0), 2)
        got = kw[f'{col}{r}'].value or 0
        if abs(got - want) > 0.02:
            print(f'  ✗ {p["name"]} {g}: 实际 {got}  应为 {want}'); bad += 1; miss += 1
        else:
            ok += 1
print(f'   （{len(DATA["projects"])} 个项目 × 5 个成本口径逐格核对，对不上 {miss} 格）')

print('══ 3. 17 张原项目表的成本一分不差 ══')
orig = collections.defaultdict(float)
for x in DATA['purchases'] + DATA['labour']:
    orig[x['project']] += x['amount']
for p in DATA['projects']:
    if not p['src']:
        continue
    r = name2row[p['name']]
    jr = sum(round(x['outgo'] - x['income'], 2) for x in DATA['journal']
             if x['project'] == p['name'] and CAT2G.get(x['cat']))
    chk(f'{p["src"]} 成本合计', kw[f'Q{r}'].value or 0, round(orig[p['name']] + jr, 2))

print('══ 4. 合同额 / 已收款 / 未收款 ══')
recv_j = collections.defaultdict(float)
for x in DATA['journal']:
    if x['cat'] in ('项目收款', '预收定金', '退还客户款') and x['project']:
        recv_j[x['project']] += x['income'] - x['outgo']
order_amt = collections.defaultdict(float)
for x in DATA['orders']:
    order_amt[x['project']] += round((x['qty'] or 1) * x['price'], 2)
tot_c = tot_r = 0.0
for p in DATA['projects']:
    r = name2row[p['name']]
    wc = round(p['contract'] + order_amt.get(p['name'], 0.0), 2)
    wr = round(p['opening_recv'] + recv_j.get(p['name'], 0.0), 2)
    tot_c += wc; tot_r += wr
    chk(f'{p["name"]} 合同总额', kw[f'I{r}'].value or 0, wc)
    chk(f'{p["name"]} 已收款', kw[f'J{r}'].value or 0, wr)
chk('项目核算 合计·合同额', cell('项目核算', 'I4'), round(tot_c, 2))
chk('项目核算 合计·已收款', cell('项目核算', 'J4'), round(tot_r, 2))
chk('项目核算 合计·未收款 = 合同-已收', cell('项目核算', 'K4'),
    round(cell('项目核算', 'I4') - cell('项目核算', 'J4'), 2))
chk('项目核算 合计·毛利 = 合同-成本', cell('项目核算', 'R4'),
    round(cell('项目核算', 'I4') - cell('项目核算', 'Q4'), 2))
chk('成本合计 = 五项之和', cell('项目核算', 'Q4'),
    round(sum(cell('项目核算', f'{c}4') for c in 'LMNOP'), 2))

print('══ 5. 客户应收 与 项目台账 一致 ══')
cu = wb['客户应收']
by_cust_c = collections.defaultdict(float); by_cust_r = collections.defaultdict(float)
for p in DATA['projects']:
    r = name2row[p['name']]
    by_cust_c[p['customer']] += kw[f'I{r}'].value or 0
    by_cust_r[p['customer']] += kw[f'J{r}'].value or 0
n = 0
for i in range(250):
    r = S0 + i
    c = cu[f'A{r}'].value
    if not c:
        continue
    n += 1
    chk(f'客户 {c} 合同额', cu[f'C{r}'].value or 0, round(by_cust_c[c], 2))
    chk(f'客户 {c} 已收款', cu[f'D{r}'].value or 0, round(by_cust_r[c], 2))
chk('客户应收 合计 = 项目核算 合计', cell('客户应收', 'C4'), cell('项目核算', 'I4'))
chk('客户应收 已收合计 = 项目核算', cell('客户应收', 'D4'), cell('项目核算', 'J4'))
print(f'   （{n} 个客户逐个核对）')

print('══ 6. 供应商对账 ══')
sp = wb['供应商对账']
buy = collections.defaultdict(float); lab = collections.defaultdict(float)
pay = collections.defaultdict(float); direct = collections.defaultdict(float)
for x in DATA['purchases']:
    buy[x['supplier']] += x['amount']
for x in DATA['labour']:
    lab[x['person']] += x['amount']
LOAN = {'借出款', '收回借出款', '借入款', '归还借入款', '股东投入',
        '股东取款', '员工借支', '收回员工借支'}
loan = collections.defaultdict(float)
for x in DATA['journal']:
    if x['party']:
        if x['cat'] in LOAN:
            loan[x['party']] += x['outgo'] - x['income']
        else:
            pay[x['party']] += x['outgo'] - x['income']
        if CAT2G.get(x['cat']):
            direct[x['party']] += x['outgo']
n = 0
for i in range(150):
    r = S0 + i
    s = sp[f'A{r}'].value
    if not s:
        continue
    n += 1
    chk(f'供应商 {s} 采购发生', sp[f'B{r}'].value or 0, round(buy.get(s, 0.0), 2))
    chk(f'供应商 {s} 付款净额', sp[f'E{r}'].value or 0, round(pay.get(s, 0.0), 2))
    chk(f'供应商 {s} 借贷净额', sp[f'G{r}'].value or 0, round(loan.get(s, 0.0), 2))
    chk(f'供应商 {s} 应付余额',
        sp[f'F{r}'].value or 0,
        round((sp[f'B{r}'].value or 0) + (sp[f'C{r}'].value or 0)
              + (sp[f'D{r}'].value or 0) - (sp[f'E{r}'].value or 0), 2))
print(f'   （{n} 家供应商逐个核对）')

print('══ 7. 利润表 ══')
li = wb['利润表']
done = [p for p in DATA['projects'] if p['done_date']]
rev = collections.defaultdict(float); cst = collections.defaultdict(float)
for p in done:
    r = name2row[p['name']]
    ym = int(p['done_date'][:4]) * 100 + int(p['done_date'][5:7])
    rev[ym] += kw[f'I{r}'].value or 0
    cst[ym] += kw[f'Q{r}'].value or 0
for m in range(1, 13):
    col = chr(ord('A') + m)
    chk(f'{m}月 主营业务收入', li[f'{col}6'].value or 0, round(rev.get(202600 + m, 0.0), 2))
chk('全年 主营业务收入', li['N6'].value or 0, round(sum(rev.values()), 2))
chk('营业收入合计 = 1+2', li['N8'].value or 0,
    round((li['N6'].value or 0) + (li['N7'].value or 0), 2))
chk('营业成本合计 = 3..8', li['N16'].value or 0,
    round(sum(li[f'N{r}'].value or 0 for r in range(10, 16)), 2))
chk('毛利 = 收入 - 成本', li['N17'].value or 0,
    round((li['N8'].value or 0) - (li['N16'].value or 0), 2))
chk('营业利润 = 毛利 - 期间费用', li['N24'].value or 0,
    round((li['N17'].value or 0) - (li['N23'].value or 0), 2))
chk('净利润 = 营业利润 + 营业外', li['N27'].value or 0,
    round((li['N24'].value or 0) + (li['N25'].value or 0) - (li['N26'].value or 0), 2))
chk('完工项目成本合计进了利润表', round(sum(li[f'N{r}'].value or 0 for r in range(10, 15)), 2),
    round(sum(cst.values()), 2))
chk('在制项目个数', li['B31'].value, len(DATA['projects']) - len(done))
chk('在制+完工 = 全部合同额',
    round((li['B32'].value or 0) + (li['N6'].value or 0), 2), round(tot_c, 2))
chk('应收账款 = 项目核算合计', li['B35'].value or 0, cell('项目核算', 'K4'))
chk('资金余额 = 账户余额合计', li['B37'].value or 0, cell('账户余额', 'E14'))

print('══ 8. 录入体检 ══')
us = wb['使用说明']
n_wait = sum(1 for x in DATA['journal'] if x['cat'] == '待分类')
chk('待分类笔数', us['C24'].value, n_wait)
chk('待分类金额', us['C25'].value, round(sum(x['income']+x['outgo'] for x in DATA['journal'] if x['cat']=='待分类'),2))
chk('成本类支出没填项目笔数', us['C26'].value,
    sum(1 for i in range(len(DATA['journal'])) if jw[f'P{D0+i}'].value == '※成本类支出没填项目'))
chk('未填完工日期的项目数', us['C28'].value, len(DATA['projects']) - len(done))
chk('日记账其它红字提示笔数', us['C30'].value,
    sum(1 for i in range(len(DATA['journal']))
        if isinstance(jw[f'P{D0+i}'].value, str) and jw[f'P{D0+i}'].value.startswith('※')
        and jw[f'P{D0+i}'].value != '※成本类支出没填项目'))
for lab_, addr in (('采购单', 'C31'), ('生产安装', 'C32'), ('项目台账', 'C33'), ('订单明细', 'C34')):
    v = us[addr].value
    print(f'   {lab_} 红字提示 {v} 条')
print()
print(f'══ 结果：{ok} 项通过，{bad} 项不通过 ══')
sys.exit(1 if bad else 0)
