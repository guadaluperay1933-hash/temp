# -*- coding: utf-8 -*-
"""新增挂靠单位实测：照【操作流程】第八部分的步骤，在副本上加一家「海川」
   ① 单位档案加一行（类型＝挂靠单位）  ② 项目档案加一个海川的项目
   ③ 复制【杰华明细】改名「海川明细」、B3 选海川
   ④ 业务流水录 5 笔：开票、成本票、业主付给海川、海川转我方、管理费结算
看：海川明细逐笔和第 7 行；单位汇总 / 往来台账 / 税费台账 / 单位项目明细多出海川；原来 8 家一格没变；
    J3 的三种提示（B3 空着、B3 不是挂靠单位、复制了表没改 B3）
跑法：python3 verify_newunit.py <交付的xlsx> <临时目录>"""
import sys, os, shutil, subprocess, datetime as dt, openpyxl

SRC, TMP = sys.argv[1], sys.argv[2]
os.makedirs(TMP, exist_ok=True)
X = os.path.join(TMP, 'newunit.xlsx')
shutil.copy(SRC, X)
HERE = os.path.dirname(os.path.abspath(__file__))
RECALC = os.path.join(HERE, '..', '..', '工具', 'recalc_safe.py')
N = lambda v: float(v) if isinstance(v, (int, float)) else 0.0
SUBS = ['迅驰明细', '康欣明细', '德誉嘉明细', '华城明细', '金沁明细', '湖南锦泰明细', '安锐明细', '杰华明细']
BAD = []


def chk(name, a, b, tol=0.005):
    ok = (a == b) if isinstance(b, str) or b is None else abs(N(a) - N(b)) <= tol
    print(('  ✓ ' if ok else '  ✗ ') + f'{name}: {a!r}  应为 {b!r}')
    if not ok:
        BAD.append(name)


def first_empty(ws, col, r0, r1):
    return next(r for r in range(r0, r1 + 1) if ws[f'{col}{r}'].value in (None, ''))


base = openpyxl.load_workbook(SRC, data_only=True)
wb = openpyxl.load_workbook(X)
# ① 单位档案
wu = wb['单位档案']
ur = first_empty(wu, 'A', 6, 45)
for c, v in zip('ABCDEFGLM', ['海川', '海川建设（测试）', '挂靠单位', 0.03, 0.03, '否', '全额销项法', '否', '正常']):
    wu[f'{c}{ur}'] = v
# ② 项目档案
wp = wb['项目档案']
pr = first_empty(wp, 'A', 6, 305)
CODE = 'A900'
for c, v in zip('ABCDEFGIN', [CODE, '海川测试项目（全称）', '海川测试项目', '自营项目', '框架合同', '民能', '海川', '泓普', '在建']):
    wp[f'{c}{pr}'] = v
# ③ 复制杰华明细 → 海川明细，B3 选海川；另外三张故意录错，看 J3 提示
cp = wb.copy_worksheet(wb['杰华明细']); cp.title = '海川明细'; cp['B3'] = '海川'
e1 = wb.copy_worksheet(wb['杰华明细']); e1.title = '测试没改B3明细'          # B3 还是杰华
e2 = wb.copy_worksheet(wb['杰华明细']); e2.title = '测试民能明细'; e2['B3'] = '民能'
e3 = wb.copy_worksheet(wb['杰华明细']); e3.title = '测试空B3明细'; e3['B3'] = None
# ④ 业务流水
wf = wb['业务流水']
fr = first_empty(wf, 'B', 5, 2004)
ROWS = [
    (dt.date(2026, 9, 25), CODE, '销项开票', '海川', '民能', '3%专票', '劳务票', '2026.9.25海川开票到民能：测试', 10000, 0.02, 0, '扣管理费'),
    (dt.date(2026, 9, 26), CODE, '成本票', '泓普', '海川', '3%专票', '劳务票', '2026.9.26泓普开成本票到海川：测试', 9000),
    (dt.date(2026, 9, 27), CODE, '挂靠代收', '民能', '海川', None, None, '2026.9.27民能付款到海川：测试', 10000),
    (dt.date(2026, 9, 28), CODE, '我方收款', '海川', '泓普', None, None, '2026.9.28海川转泓普：测试', 9800),
    (dt.date(2026, 9, 28), CODE, '管理费结算', '泓普', '海川', None, None, '2026.9.28结管理费：测试', 200),
]
for i, vals in enumerate(ROWS):
    for c, v in zip('BCDEFGHIJKLM', vals):
        if v is not None:
            wf[f'{c}{fr + i}'] = v
wb['单位项目明细']['B3'] = '海川'
wb.save(X)
subprocess.run([sys.executable, RECALC, X, '900'], check=True, capture_output=True)
v = openpyxl.load_workbook(X, data_only=True)

print('① 业务流水：5 笔都 √，都归到海川')
f = v['业务流水']
for i in range(len(ROWS)):
    chk(f'业务流水第 {fr + i} 行 校验', f[f'O{fr + i}'].value, '√')
    chk(f'业务流水第 {fr + i} 行 归属单位', f[f'AN{fr + i}'].value, '海川')

print('② 海川明细（复制杰华明细、B3 选海川）')
h = v['海川明细']
chk('标题跟着 B3', h['A1'].value, '海川 · 对账明细（给领导看的逐笔明细）')
chk('J3 没有提示', h['J3'].value, '全部期间')
chk('D7 笔数', h['D7'].value, '5 笔（本表共 5 笔）')
exp_memo = ['2026.9.25海川开票到民能：测试', '【成本票】2026.9.26泓普开成本票到海川：测试',
            '【挂靠代收】2026.9.27民能付款到海川：测试', '【我方收款】2026.9.28海川转泓普：测试',
            '【管理费结算】2026.9.28结管理费：测试']
for i, m in enumerate(exp_memo):
    chk(f'第 {8 + i} 行摘要', h[f'D{8 + i}'].value, m)
chk('第 13 行空', h['D13'].value or '', '')
for col, lab, exp in [('H', '销售开票金额', 10000), ('I', '应扣管理费', 200), ('J', '应到成本票', 9800),
                      ('K', '已到成本票', 9000), ('M', '剩余开票金额', 800),
                      ('U', '应收工程款', 10000), ('W', '开票已回款', 10000), ('X', '欠业主未付款余额', 0),
                      ('Y', '应收金额', 10000), ('Z', '扣管理费', 200), ('AB', '已到账', 9800), ('AC', '应收挂靠方余额', 0)]:
    chk(f'第 7 行 {lab}', h[f'{col}7'].value, exp)
chk('最后一笔滚动余额 剩余开票金额', h['M12'].value, 800)

print('③ 汇总表自动多出海川，并且没被折叠')
for sh, cols in [('单位汇总', {'C': 10000, 'D': 200, 'E': 9800, 'H': 9000, 'I': 800, 'M': 10000, 'N': 0,
                              'O': 10000, 'P': 200, 'R': 9800, 'S': 0, 'U': '有'}),
                 ('往来台账', {'C': 0, 'H': 0}), ('税费台账', {'G': 0, 'H': 0})]:
    ws = v[sh]
    r = next((rr for rr in range(7, 47) if ws[f'A{rr}'].value == '海川'), None)
    chk(f'{sh} 有海川一行', r is not None and 'yes', 'yes')
    if r:
        chk(f'{sh} 海川那一行没折叠', str(ws.row_dimensions[r].hidden), 'False')
        for c, exp in cols.items():
            chk(f'{sh} 海川 {ws[f"{c}5"].value}', ws[f'{c}{r}'].value, exp)
x = v['单位项目明细']
xr = next((rr for rr in range(7, 320) if x[f'A{rr}'].value == CODE), None)
chk('单位项目明细 选海川能看到新项目', xr is not None and 'yes', 'yes')
if xr:
    chk('单位项目明细 海川 开票额', x[f'C{xr}'].value, 10000)

print('④ 原来 8 家一格没变')
for sn in SUBS:
    a, b = base[sn], v[sn]
    diff = [c for c in range(1, 38) if a.cell(7, c).value != b.cell(7, c).value
            and not (isinstance(a.cell(7, c).value, (int, float)) and isinstance(b.cell(7, c).value, (int, float))
                     and abs(a.cell(7, c).value - b.cell(7, c).value) < 0.005)]
    chk(f'{sn} 第 7 行合计不变', len(diff), 0)
    chk(f'{sn} J3 没有提示', b['J3'].value, a['J3'].value)
su0, su1 = base['单位汇总'], v['单位汇总']
for rr in range(7, 15):
    chk(f'单位汇总 {su0[f"A{rr}"].value} 一行不变',
        sum(1 for c in range(3, 21) if abs(N(su0.cell(rr, c).value) - N(su1.cell(rr, c).value)) > 0.005), 0)
chk('首页 业务流水校验仍通过', (v['首页']['D20'].value or '')[:1], '✓')

print('⑤ J3 提示')
chk('复制了表没改 B3', v['测试没改B3明细']['J3'].value, '⚠ 表名跟本表单位「杰华」对不上：复制过来的表要在 B3 选新单位')
chk('B3 选了业主', v['测试民能明细']['J3'].value, '⚠ 「民能」在【单位档案】里类型不是「挂靠单位」，这张表取不到数')
chk('B3 空着', v['测试空B3明细']['J3'].value, '⚠ 请在 B3 选本表单位')
chk('B3 空着 → 明细是空的', v['测试空B3明细']['D8'].value or '', '')
print(f'\n不一致 {len(BAD)} 项')
sys.exit(1 if BAD else 0)
