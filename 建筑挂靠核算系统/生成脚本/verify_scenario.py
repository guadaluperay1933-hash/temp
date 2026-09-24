# -*- coding: utf-8 -*-
"""9.24 复核发现的几种录法 + 期间筛选，在副本上实测：
   ① 康欣开给金沁的成本票（计入汇总=否）要算到金沁头上
   ② 上一层挂靠单位付给下一层（德誉嘉→安锐 挂靠代收）要算到安锐头上
   ③ 泓普代交税、两边都不是挂靠单位 → 校验要报
   ④ 日记账粘进「本月合计」行 → 不计入、首页不报错
   ⑤ 明细 / 单位汇总填了期间：余额列＝截至截止日期的余额，发生额只算期间内
跑法：python3 verify_scenario.py <交付的xlsx> <临时目录>"""
import sys, os, shutil, subprocess, datetime as dt, openpyxl

SRC, TMP = sys.argv[1], sys.argv[2]
os.makedirs(TMP, exist_ok=True)
X = os.path.join(TMP, 'scenario.xlsx')
shutil.copy(SRC, X)
HERE = os.path.dirname(os.path.abspath(__file__))
RECALC = os.path.join(HERE, '..', '..', '工具', 'recalc_safe.py')
N = lambda v: float(v) if isinstance(v, (int, float)) else 0.0

base = openpyxl.load_workbook(SRC, data_only=True)
def last_row(ws, col='B', r0=5):
    r = r0
    while ws[f'{col}{r}'].value not in (None, ''): r += 1
    return r
wb = openpyxl.load_workbook(X)
wf = wb['业务流水']
r = last_row(base['业务流水'])
R1, R2, R3 = r, r + 1, r + 2
for rr, vals in ((R1, [dt.date(2026, 9, 24), 'A045', '成本票', '康欣', '金沁', '3%专票', '劳务票', '测试：康欣开给金沁的成本票', 1000]),
                 (R2, [dt.date(2026, 9, 24), 'A007', '挂靠代收', '德誉嘉', '安锐', None, None, '测试：德誉嘉付给安锐', 200]),
                 (R3, [dt.date(2026, 9, 24), 'A001', '已交税', '泓普', '税局', None, None, '测试：泓普代交税（两边都不是挂靠单位）', 50])):
    for c, v in zip('BCDEFGHIJ', vals):
        if v is not None: wf[f'{c}{rr}'] = v
wf[f'AK{R1}'] = '否'
wj = wb['资金日记账']
JR = last_row(base['资金日记账'], 'B', 6)
for c, v in zip('BIJKL', [dt.date(2026, 9, 30), '本月合计', '现金', 8000, 1433]):
    wj[f'{c}{JR}'] = v
# 期间：康欣明细截止 2026-07-31；金沁明细起 2026-07-01；单位汇总 2026 年
wb['康欣明细']['H3'] = dt.date(2026, 7, 31)
wb['金沁明细']['F3'] = dt.date(2026, 7, 1)
wb['单位汇总']['B3'] = 2026
wb.save(X)
subprocess.run([sys.executable, RECALC, X, '900'], check=True, capture_output=True)
v = openpyxl.load_workbook(X, data_only=True)
BAD = []
def chk(name, a, b, tol=0.05):
    ok = (a == b) if isinstance(b, str) else abs(N(a) - N(b)) <= tol
    print(('  ✓ ' if ok else '  ✗ ') + f'{name}: {a!r}  应为 {b!r}')
    if not ok: BAD.append(name)

f = v['业务流水']
print('① ② ③ 新录行的归属单位与校验')
chk('康欣→金沁 成本票 归属', f[f'AN{R1}'].value, '金沁')
chk('德誉嘉→安锐 挂靠代收 归属', f[f'AN{R2}'].value, '安锐')
chk('泓普代交税 校验提示', f[f'O{R3}'].value, '两边都不是挂靠单位，算不到哪家头上')
b0, b1 = base['金沁明细'], v['金沁明细']
b0k, b1k = base['康欣明细'], v['康欣明细']
# 金沁明细这次填了起始日期，看全期间的已到成本票用单位汇总（年度 2026 不影响）
su0, su1 = base['单位汇总'], v['单位汇总']
def urow(ws, u):
    return next(rr for rr in range(7, 47) if ws[f'A{rr}'].value == u)
# 单位汇总这次填了年度 2026：2025 年的已收成本票会掉出去，预期值要扣掉
fb = base['业务流水']
def done2025(u):
    return sum(N(fb[f'J{rr}'].value) for rr in range(5, 2005)
               if fb[f'AN{rr}'].value == u and fb[f'AM{rr}'].value == '已开成本票'
               and isinstance(fb[f'B{rr}'].value, dt.datetime) and fb[f'B{rr}'].value.year == 2025)
chk('单位汇总 金沁 已收成本票 +1000（2026 年）', N(su1[f'H{urow(su1, "金沁")}'].value) - N(su0[f'H{urow(su0, "金沁")}'].value),
    1000 - done2025('金沁'))
chk('单位汇总 康欣 已收成本票 不变（2026 年）', N(su1[f'H{urow(su1, "康欣")}'].value) - N(su0[f'H{urow(su0, "康欣")}'].value),
    0 - done2025('康欣'))
chk('单位汇总 安锐 业主已付 +200', N(su1[f'M{urow(su1, "安锐")}'].value) - N(su0[f'M{urow(su0, "安锐")}'].value), 200)
chk('单位汇总 德誉嘉 业主已付 不变', N(su1[f'M{urow(su1, "德誉嘉")}'].value) - N(su0[f'M{urow(su0, "德誉嘉")}'].value), 0)
print('④ 日记账本月合计行')
chk('日记账 计入', v['资金日记账'][f'X{JR}'].value, 0)
chk('日记账 校验', v['资金日记账'][f'Y{JR}'].value, '√ 合计/结转行（不计入）')
d21 = [c.value for row in v['首页'].iter_rows(min_row=18, max_row=30) for c in row
       if isinstance(c.value, str) and c.value.startswith(('✓', '✗')) and '笔，科目' in c.value]
chk('首页 日记账校验 仍通过', (d21 or [''])[0][:1], '✓')
chk('现金期末余额不变', v['资金日记账']['S6'].value, base['资金日记账']['S6'].value)
print('⑤ 期间筛选：余额＝截至截止日期')
def bal_upto(ws_base, end, plus, minus):
    """在全期间的基线明细上，按日期把截止日期之前的行加起来（基线没有筛选）"""
    hc = {str(ws_base.cell(6, c).value).split('\n')[0]: c for c in range(1, 40) if ws_base.cell(6, c).value}
    t = 0.0
    for rr in range(8, 408):
        d = ws_base[f'B{rr}'].value
        if not isinstance(d, (dt.date, dt.datetime)) or d.date() > end if isinstance(d, dt.datetime) else False: continue
        if isinstance(d, dt.datetime): d = d.date()
        if not isinstance(d, dt.date) or d > end: continue
        t += sum(N(ws_base.cell(rr, hc[k]).value) for k in plus) - sum(N(ws_base.cell(rr, hc[k]).value) for k in minus)
    return round(t, 2)
hc = {str(b1k.cell(6, c).value).split('\n')[0]: c for c in range(1, 40) if b1k.cell(6, c).value}
end = dt.date(2026, 7, 31)
chk('康欣明细(截至7.31) 欠业主未付款余额', b1k.cell(7, hc['开票未到账']).value, bal_upto(b0k, end, ['应收工程款'], ['开票已回款']))
chk('康欣明细(截至7.31) 应收挂靠方余额', b1k.cell(7, hc['未到账余额']).value,
    bal_upto(b0k, end, ['应收金额'], ['扣管理费', '质保金', '已到账']))
hq = {str(b1.cell(6, c).value).split('\n')[0]: c for c in range(1, 40) if b1.cell(6, c).value}
chk('金沁明细(7.1 起) 余额列＝全期余额', b1.cell(7, hq['未到账余额']).value, b0.cell(7, hq['未到账余额']).value)
chk('金沁明细(7.1 起) 应收工程款只算 7.1 以后', b1.cell(7, hq['应收工程款']).value,
    round(sum(N(b0.cell(rr, hq['应收工程款']).value) for rr in range(8, 408)
              if isinstance(b0[f'B{rr}'].value, dt.datetime) and b0[f'B{rr}'].value.date() >= dt.date(2026, 7, 1)), 2))
for u in ('金沁', '康欣', '杰华'):
    chk(f'单位汇总(2026) {u} 应收挂靠方余额≥0 且等于全期', su1[f'S{urow(su1, u)}'].value, su0[f'S{urow(su0, u)}'].value)
print(f'\n不一致 {len(BAD)} 项')
sys.exit(1 if BAD else 0)
