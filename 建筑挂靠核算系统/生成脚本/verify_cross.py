# -*- coding: utf-8 -*-
"""需求 3：【单位汇总】【项目汇总】【单位项目明细】和 8 张单位竖版明细，
   同一个数在四个地方必须一模一样。这里逐项对。
   跑法：python3 verify_cross.py <重算过的xlsx>
"""
import sys, collections, openpyxl

path = sys.argv[1]
wb = openpyxl.load_workbook(path, data_only=True)
UNITS = ['德誉嘉','迅驰','华城','金沁','湖南锦泰','康欣','安锐','杰华']
ok = bad = 0
def chk(name, a, b, tol=0.05):
    global ok, bad
    a = float(a or 0); b = float(b or 0)
    if abs(a-b) <= tol: ok += 1
    else:
        bad += 1
        print('  ✗ %-44s %14.2f  vs %14.2f   差 %12.2f' % (name, a, b, a-b))

def find_col(ws, hdr_row, name):
    for c in range(1, ws.max_column+1):
        v = ws.cell(hdr_row, c).value
        if v and str(v).replace('\n','') == name.replace('\n',''):
            return c
    return None

# ---------- 单位汇总 ----------
su = wb['单位汇总']
hr_u = next(r for r in range(1, 20) if su.cell(r,1).value == '单位简称')
U = {}
for r in range(hr_u+1, hr_u+30):
    u = su.cell(r,1).value
    if u in UNITS: U[u] = r
cols_u = {n: find_col(su, hr_u, n) for n in
          ('开票额(该单位开出)','应扣管理费','应到成本票','其中·按净额(开票额−管理费)',
           '其中·按全额(不扣管理费)','已收成本票','还差成本票')}
print('【单位汇总】表头行 %d，找到单位 %d 家，列：%s' % (hr_u, len(U), {k:v for k,v in cols_u.items() if v}))

# ---------- 项目汇总：按单位加总 ----------
sp = wb['项目汇总']
hr_p = next(r for r in range(1, 20) if sp.cell(r,1).value == '项目编号')
c_short = find_col(sp, hr_p, '项目简称')
cols_p = {n: find_col(sp, hr_p, n) for n in cols_u}
Pagg = collections.defaultdict(lambda: collections.defaultdict(float))
nproj = 0
for r in range(hr_p+1, sp.max_row+1):
    code = sp.cell(r,1).value
    if not code or str(code) in ('合  计','合计'): continue
    short = str(sp.cell(r, c_short).value or '')
    u = short[short.rfind('(')+1:short.rfind(')')] if '(' in short and short.endswith(')') else None
    if u not in UNITS: continue
    nproj += 1
    for n, c in cols_p.items():
        if c: Pagg[u][n] += float(sp.cell(r,c).value or 0)
print('【项目汇总】表头行 %d，能认出归属单位的项目 %d 个' % (hr_p, nproj))

# ---------- 8 张单位竖版明细：期间合计行 ----------
DET = {}
for u in UNITS:
    ws = wb[f'{u}明细']
    DET[u] = dict(开票额=ws['H7'].value, 管理费=ws['I7'].value, 应到=ws['J7'].value, 已到=ws['K7'].value)

print('\n=== ① 单位汇总 vs 8 张单位明细的期间合计 ===')
for u in UNITS:
    r = U[u]
    chk(f'{u}·开票额',     su.cell(r, cols_u['开票额(该单位开出)']).value, DET[u]['开票额'])
    chk(f'{u}·应扣管理费', su.cell(r, cols_u['应扣管理费']).value,        DET[u]['管理费'])
    chk(f'{u}·应到成本票', su.cell(r, cols_u['应到成本票']).value,        DET[u]['应到'])
    chk(f'{u}·已收成本票', su.cell(r, cols_u['已收成本票']).value,        DET[u]['已到'])

print('\n=== ② 项目汇总 合计行 vs 单位汇总 合计行（两张表看同一批流水，合计必须相等）===')
tr_u = next(r for r in range(1, 20) if str(su.cell(r,1).value or '').replace(' ','') == '合计')
tr_p = next(r for r in range(1, 20) if str(sp.cell(r,1).value or '').replace(' ','') == '合计')
for n in ('开票额(该单位开出)','应扣管理费','应到成本票','其中·按净额(开票额−管理费)',
          '其中·按全额(不扣管理费)','已收成本票'):
    if cols_u[n] and cols_p[n]:
        chk(f'合计·{n}', su.cell(tr_u, cols_u[n]).value, sp.cell(tr_p, cols_p[n]).value)

print('\n=== ②b 单位项目明细 合计行 vs 单位汇总 里选中那家单位的那一行 ===')
sd = wb['单位项目明细']
hr_d = next(r for r in range(1, 20) if sd.cell(r,1).value == '项目编号')
tr_d = next(r for r in range(1, 20) if str(sd.cell(r,1).value or '').replace(' ','') == '合计')
pick = sd['B3'].value
cols_d = {n: find_col(sd, hr_d, n) for n in cols_u}
print(f'   【单位项目明细】当前选的是「{pick}」')
if pick in U:
    for n in ('开票额(该单位开出)','应扣管理费','应到成本票','已收成本票'):
        if cols_u[n] and cols_d[n]:
            chk(f'{pick}·{n}', sd.cell(tr_d, cols_d[n]).value, su.cell(U[pick], cols_u[n]).value)

print('\n=== ③ 应到成本票 = 按净额 + 按全额 ===')
for u in UNITS:
    r = U[u]
    tot = float(su.cell(r, cols_u['应到成本票']).value or 0)
    a = float(su.cell(r, cols_u['其中·按净额(开票额−管理费)']).value or 0)
    b = float(su.cell(r, cols_u['其中·按全额(不扣管理费)']).value or 0)
    chk(f'{u}·净额+全额=应到', a+b, tot)

print('\n=== ④ 还差成本票 = 应到 − 已收 ===')
for u in UNITS:
    r = U[u]
    d = float(su.cell(r, cols_u['应到成本票']).value or 0) - float(su.cell(r, cols_u['已收成本票']).value or 0)
    chk(f'{u}·还差', su.cell(r, cols_u['还差成本票']).value, d)

print('\n' + '='*56)
print('共核 %d 项，一致 %d 项，不一致 %d 项' % (ok+bad, ok, bad))
sys.exit(1 if bad else 0)
