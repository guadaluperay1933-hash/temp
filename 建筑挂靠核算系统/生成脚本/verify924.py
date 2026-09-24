# -*- coding: utf-8 -*-
"""9.24 版独立核算：不看公式，只读【业务流水】【资金日记账】里的录入值，在 Python 里把
   两组回款（欠业主未付款 / 应收挂靠方）、单位汇总、项目汇总、8 张单位明细、日记账余额、费用统计
   全部重算一遍，再跟重算过的 xlsx 里公式算出来的数逐项比。
跑法：python3 verify924.py <重算过的xlsx>"""
import sys, os, collections, datetime as dt, openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
XL = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', '建筑挂靠业务核算系统.xlsx')
wbv = openpyxl.load_workbook(XL, data_only=True)
wbf = openpyxl.load_workbook(XL, data_only=False)
R2 = lambda x: round(x + 0.0, 2)
N = lambda v: float(v) if isinstance(v, (int, float)) else 0.0
BAD = []
NCHK = [0]
def chk(name, a, b, tol=0.05):
    NCHK[0] += 1
    if a is None or b is None or abs(N(a) - N(b)) > tol:
        BAD.append((name, a, b))

# ---------- 单位档案 ----------
wu = wbv['单位档案']
UTYPE = {}
for r in range(6, 46):
    if wu[f'A{r}'].value: UTYPE[wu[f'A{r}'].value] = wu[f'C{r}'].value
HOLD = [u for u, t in UTYPE.items() if t == '挂靠单位']

# ---------- 业务流水：只读录入值 ----------
wf, wff = wbv['业务流水'], wbf['业务流水']
FL = []
for r in range(5, 2005):
    if wf[f'B{r}'].value is None: continue
    kind = wf[f'D{r}'].value
    ocol_raw = wff[f'AM{r}'].value
    if isinstance(ocol_raw, str) and ocol_raw.startswith('='):
        ocol = '已开成本票' if kind == '成本票' else '销售开票金额'
    else:
        ocol = ocol_raw
    src = wf[f'AI{r}'].value
    payer, payee = wf[f'E{r}'].value, wf[f'F{r}'].value
    eh, fh = UTYPE.get(payer) == '挂靠单位', UTYPE.get(payee) == '挂靠单位'
    if src: ubel = src
    elif kind in ('成本票', '挂靠代收', '业主扣质保金', '管理费结算', '工资扣抵'):
        ubel = payee if fh else (payer if eh else '')
    else:
        ubel = payer if eh else (payee if fh else '')
    amt, rate, flag = N(wf[f'J{r}'].value), N(wf[f'K{r}'].value), wf[f'M{r}'].value
    mfee = 0.0 if kind != '销项开票' or flag == '不扣管理费' else R2(amt * rate)
    due = 0.0 if kind != '销项开票' or flag == '不回成本票' else R2(amt - mfee)
    d = wf[f'B{r}'].value
    FL.append(dict(r=r, date=d, proj=wf[f'C{r}'].value, kind=kind, payer=payer, payee=payee, amt=amt,
                   flag=flag, ocol=ocol, src=src, ubel=ubel, mfee=mfee, due=due,
                   tax=sum(N(wf[f'{c}{r}'].value) for c in 'PQRS'), cnt=wf[f'AK{r}'].value))
    # 公式列自己也核一下：管理费、应到成本票、归属单位
    chk(f'业务流水第{r}行 管理费', wf[f'X{r}'].value, mfee, 0.005)
    chk(f'业务流水第{r}行 应到成本票', wf[f'Y{r}'].value, due, 0.005)
    if wf[f'AN{r}'].value != ubel: BAD.append((f'业务流水第{r}行 归属单位', wf[f'AN{r}'].value, ubel))
print(f'业务流水 {len(FL)} 笔')

def tot(rows):
    t = collections.Counter()
    for x in rows:
        k, a = x['kind'], x['amt']
        if k == '销项开票' and x['ocol'] != '已开成本票':
            t['sale'] += a; t['a_ar'] += a
            if x['flag'] != '不回成本票': t['b_ar'] += a
        if k == '销项开票':
            t['mfee'] += x['mfee']; t['due'] += x['due']
        if x['ocol'] == '已开成本票': t['done'] += a
        if k == '代垫应收': t['b_ar'] += a
        if k == '挂靠代收': t['a_got'] += a
        if k == '业主扣质保金': t['a_bond'] += a
        if k == '管理费结算': t['b_fee'] += a
        if k == '扣质保金': t['b_bond'] += a
        if k == '我方收款': t['b_got'] += a
        if k == '已交税': t['paid'] += a
        t['tax'] += x['tax']
    t['a_left'] = t['sale'] - t['a_got']
    t['b_left'] = t['b_ar'] - t['b_fee'] - t['b_bond'] - t['b_got']
    t['transit'] = t['a_got'] - t['b_got'] - t['b_fee'] - t['b_bond']
    t['gap'] = t['due'] - t['done']
    return t

# ---------- 单位汇总 ----------
su = wbv['单位汇总']
HCOL = {str(su.cell(5, c).value).replace('\n', '').replace('(截至截止日)', ''): c for c in range(1, 30) if su.cell(5, c).value}
MAPU = {'开票额(该单位开出)': 'sale', '应扣管理费': 'mfee', '应到成本票': 'due', '已收成本票': 'done',
        '业主已付给挂靠单位': 'a_got', '欠业主未付款余额': 'a_left', '应收挂靠方金额': 'b_ar',
        '管理费已结算': 'b_fee', '扣质保金': 'b_bond', '挂靠单位已转我方': 'b_got',
        '应收挂靠方余额': 'b_left', '挂靠单位代收未转': 'transit', '还差成本票': 'gap'}
for k in MAPU: assert k in HCOL, ('单位汇总缺列', k)
UT = {}
for r in range(7, 47):
    u = su[f'A{r}'].value
    if not u: continue
    t = tot([x for x in FL if x['ubel'] == u]); UT[u] = t
    for lab, key in MAPU.items():
        chk(f'单位汇总 {u} {lab}', su.cell(r, HCOL[lab]).value, t[key])
tall = tot([x for x in FL if x['ubel']])
for lab, key in MAPU.items():
    chk(f'单位汇总 合计 {lab}', su.cell(6, HCOL[lab]).value, tall[key])

# ---------- 8 张单位明细：第 7 行期间合计 + 末行滚动余额 ----------
for u in HOLD:
    ws = wbv[f'{u}明细']
    hc = {str(ws.cell(6, c).value).split('\n')[0]: c for c in range(1, 40) if ws.cell(6, c).value}
    t = UT[u]
    for lab, key in [('销售开票金额', 'sale'), ('应扣管理费', 'mfee'), ('应到成本票', 'due'), ('已到成本票', 'done'),
                     ('应收工程款', 'a_ar'), ('业主扣', 'a_bond'), ('开票已回款', 'a_got'), ('开票未到账', 'a_left'),
                     ('应收金额', 'b_ar'), ('扣管理费', 'b_fee'), ('质保金', 'b_bond'), ('已到账', 'b_got'),
                     ('未到账余额', 'b_left')]:
        chk(f'{u}明细 合计 {lab}', ws.cell(7, hc[lab]).value, t[key])
    n = sum(1 for x in FL if x['ubel'] == u)
    last = 8 + n - 1
    chk(f'{u}明细 笔数', sum(1 for r in range(8, 408) if ws[f'B{r}'].value not in (None, '')), n, 0)
    chk(f'{u}明细 末行欠业主未付款余额', ws.cell(last, hc['开票未到账']).value, t['a_left'])
    chk(f'{u}明细 末行应收挂靠方余额', ws.cell(last, hc['未到账余额']).value, t['b_left'])
    if ws.cell(last + 1, hc['未到账余额']).value not in (None, ''):
        BAD.append((f'{u}明细 第{last+1}行应为空', ws.cell(last + 1, hc['未到账余额']).value, ''))

# ---------- 项目汇总：逐项目 ----------
sp = wbv['项目汇总']
HP = {str(sp.cell(5, c).value).replace('\n', '').replace('(截至截止日)', ''): c for c in range(1, 30) if sp.cell(5, c).value}
NEG = []
for r in range(7, 307):
    code = sp[f'A{r}'].value
    if not code: continue
    t = tot([x for x in FL if x['proj'] == code and x['ubel']])
    for lab, key in MAPU.items():
        chk(f'项目汇总 {code} {lab}', sp.cell(r, HP[lab]).value, t[key])
    if t['a_left'] < -0.05 or t['b_left'] < -0.05:
        NEG.append((code, sp[f'B{r}'].value, R2(t['a_left']), R2(t['b_left'])))
for lab, key in MAPU.items():
    chk(f'项目汇总 合计 {lab}', sp.cell(6, HP[lab]).value, tot([x for x in FL if x['ubel'] and x['proj']])[key])

# ---------- 单位项目明细（当前选的那家） ----------
sx = wbv['单位项目明细']
pick = sx['B3'].value
HX = {str(sx.cell(5, c).value).replace('\n', '').replace('(截至截止日)', ''): c for c in range(1, 30) if sx.cell(5, c).value}
for lab, key in MAPU.items():
    chk(f'单位项目明细[{pick}] 合计 {lab}', sx.cell(6, HX[lab]).value,
        tot([x for x in FL if x['ubel'] == pick and x['proj']])[key])

# ---------- 资金日记账 ----------
wj, wjf = wbv['资金日记账'], wbf['资金日记账']
ACC = [wj[f'O{r}'].value for r in range(6, 16) if wj[f'O{r}'].value]
OPEN = {wj[f'O{r}'].value: N(wj[f'P{r}'].value) for r in range(6, 16) if wj[f'O{r}'].value}
SUBT = {'本月合计', '本年累计', '过次页', '承前页', '上年结转'}
bal = collections.Counter(OPEN)
JR = []
for r in range(6, 2006):
    b, j, k, l, i = (wj[f'{c}{r}'].value for c in 'BJKLI')
    ok = isinstance(b, (dt.date, dt.datetime)) and j not in (None, '') and (N(k) or N(l)) \
        and str(i or '').replace(' ', '').replace('　', '') not in SUBT
    chk(f'日记账第{r}行 计入', wj[f'X{r}'].value, 1 if ok else 0, 0)
    if not ok: continue
    bal[j] += N(k) - N(l)
    chk(f'日记账第{r}行 本科目余额', wj[f'U{r}'].value, bal[j])
    JR.append(dict(r=r, acct=j, dr=N(k), cr=N(l), etype=wj[f'F{r}'].value, proj=wj[f'E{r}'].value, date=b))
for r in range(6, 16):
    a = wj[f'O{r}'].value
    if not a: continue
    chk(f'日记账 {a} 借方合计', wj[f'Q{r}'].value, sum(x['dr'] for x in JR if x['acct'] == a))
    chk(f'日记账 {a} 贷方合计', wj[f'R{r}'].value, sum(x['cr'] for x in JR if x['acct'] == a))
    chk(f'日记账 {a} 期末余额', wj[f'S{r}'].value, bal[a])
chk('日记账 余额合计', wj['S16'].value, sum(bal[a] for a in ACC))
print(f'资金日记账 {len(JR)} 笔，科目 {len(ACC)} 个')

# ---------- 费用统计 ----------
we = wbv['费用统计']
hdr = {we.cell(5, c).value: c for c in range(2, 12) if we.cell(5, c).value}
listed = set()
for r in range(7, 47):
    t = we[f'A{r}'].value
    if not t: continue
    listed.add(t)
    for a, c in hdr.items():
        chk(f'费用统计 {t}×{a}', we.cell(r, c).value,
            sum(x['cr'] - x['dr'] for x in JR if x['etype'] == t and x['acct'] == a))
    chk(f'费用统计 {t} 收入', we[f'L{r}'].value, sum(x['dr'] for x in JR if x['etype'] == t))
    chk(f'费用统计 {t} 支出', we[f'M{r}'].value, sum(x['cr'] for x in JR if x['etype'] == t))
chk('费用统计 没列到的类型 支出', we['M47'].value, sum(x['cr'] for x in JR if x['etype'] not in listed))
chk('费用统计 合计 支出', we['M6'].value, sum(x['cr'] for x in JR))
chk('费用统计 合计 收入', we['L6'].value, sum(x['dr'] for x in JR))

# ---------- 首页 ----------
wh = wbv['首页']
cards = {}
for row in wh.iter_rows(min_row=4, max_row=30):
    for c in row:
        if isinstance(c.value, str) and c.value and wh.cell(c.row + 1, c.column).value is not None:
            cards[c.value] = wh.cell(c.row + 1, c.column).value
chk('首页 资金余额', cards.get('资金余额（日记账各科目合计）'), sum(bal[a] for a in ACC))
chk('首页 欠业主未付款（8家相加）', cards.get('欠业主未付款（8家相加）'), tall['a_left'])
chk('首页 应收挂靠方（8家相加）', cards.get('应收挂靠方（8家相加）'), tall['b_left'])
chk('首页 还差成本票未开', cards.get('还差成本票未开'), sum(max(UT[u]['gap'], 0) for u in UT if UT[u]['gap'] > 0.5))

# ---------- 报告 ----------
print(f'\n共核 {NCHK[0]} 项，不一致 {len(BAD)} 项')
for b in BAD[:60]: print('  ✗', b)
print(f'\n按项目看余额为负的项目 {len(NEG)} 个（提示，不算错）：')
for x in NEG: print('  ', x)
sys.exit(1 if BAD else 0)
