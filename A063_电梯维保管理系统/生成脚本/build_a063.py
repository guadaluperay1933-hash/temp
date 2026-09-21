# -*- coding: utf-8 -*-
"""生成《A063 电梯维保管理系统.xlsx》

原来两张表（电梯资料 + 合同应收款）的问题：合同、开票、收款全挤在一行，
一个客户分几次收、有的开票有的开收据，就记不下来了；合同没回来也没地方挂账。

这一版拆成「合同 → 应收 → 收款」三张流水，用「客户｜项目」一个键串起来：
  · 合同没回来照样建行（合同状态选「未签 / 已签未回」），应收照挂，提醒里单列
  · 合同款和合同外款分开登记（费用类别自动判款别），收款可以合并收、也可以指定冲哪一张
  · 收款按票据类型分成 发票 / 收据 / 现金无票 三路，应收拆成 已开票应收 / 未开票应收
  · 应收和收款都不分年度、一直往下记，换年度只改《基础资料》的管理年度，
    上年欠款自动结转成本年期初 —— 天然支持连续多年
  · 维保状态分 正常维保 / 安装免保 / 技术免保 / 质保期内 / 暂停 / 已解约，免保到期自动提醒

跑法：python3 build_a063.py [输出.xlsx]
"""
import sys, os, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from a063_style import *
from a063_data import *
from a063_test import *

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'A063_电梯维保管理系统.xlsx')

wb = openpyxl.Workbook()
wb.remove(wb.active)
SH = {}


def sheet(name, hidden=False):
    ws = wb.create_sheet(name)
    if hidden:
        ws.sheet_state = 'hidden'
    SH[name] = ws
    return ws


def name(nm, ref):
    if nm in wb.defined_names:
        del wb.defined_names[nm]
    wb.defined_names.add(DefinedName(nm, attr_text=ref))


DVN = 0


def dv(ws, sqref, formula, msg=None, block=False):
    global DVN
    DVN += 1
    d = DataValidation(type='list', formula1=formula, allow_blank=True,
                       showErrorMessage=block, showInputMessage=bool(msg))
    if msg:
        d.promptTitle, d.prompt = '填写提示', msg
    if block:
        d.errorTitle, d.error = '只能从下拉里选', '这一列要跟【基础资料】的清单一致，请从下拉里选。'
    ws.add_data_validation(d)
    d.add(sqref)


def dynlist(sheet_name, col, r0, r1):
    """按整列非空条数自动伸缩的下拉区"""
    rng = f"'{sheet_name}'!${col}${r0}:${col}${r1}"
    return f'=OFFSET({sheet_name}!${col}${r0},0,0,MAX(1,COUNTA({rng})),1)'


# 各表行区间
L0, L1 = D0, D0 + N_LIFT - 1        # 电梯资料
K0, K1 = D0, D0 + N_CON - 1         # 维保合同 / 应收对账
A0, A1 = D0, D0 + N_AR - 1          # 应收登记
R0, R1 = D0, D0 + N_RC - 1          # 收款登记
BH = 6                              # 基础资料的表头行（上面 3~5 行放参数）
B0 = BH + 1                         # 基础资料清单首行
BE = B0 + N_LIST - 1                # 基础资料清单末行
T0, T1 = D0, D0 + N_CON - 1         # 分期收款计划 · 按合同汇总
X0 = T1 + 4                         # 分期收款计划 · 分期明细首行
X1 = X0 + N_CON * MAX_TERM - 1

# ══════════════════════════════════════════════════════════════
# ① 基础资料
# ══════════════════════════════════════════════════════════════
ws = sheet('基础资料')
title(ws, '基 础 资 料 · 清单和参数都在这里维护', 'S',
      '★ 淡黄＝要你填，浅灰＝公式自动算。这一页改了，后面所有表的下拉和提醒跟着变。\n'
      '★ 「费用类别」右边那列「款别」决定这笔钱算合同款还是合同外款；'
      '「票据类型」右边那列「收款归类」决定收款算发票、收据还是现金无票；'
      '「收费周期」右边那列「每年期数」决定分期付款一年分几期收 —— 这三列别乱改。',
      color=C_BASE)
PARAMS = [
    (3, 'A', 'C', 'D', 'E', 'F', 'K', '管理年度', YEAR,
     '报表按这个年度算「本年」，上年欠款自动结转成本年期初。过年只改这一格，流水一行都不用搬。'),
    (3, 'L', 'N', 'O', 'P', 'Q', 'S', '合同到期提前提醒(天)', 60, '合同止期在这个天数内就进提醒。'),
    (4, 'A', 'C', 'D', 'E', 'F', 'K', '年检 / 校验提前提醒(天)', 30,
     '年检、限速器校验、载重试验到期前多少天开始提醒。'),
    (4, 'L', 'N', 'O', 'P', 'Q', 'S', '欠款账龄预警(天)', 90, '最早一笔没收清的应收拖过这个天数就报警。'),
    (5, 'A', 'C', 'D', 'E', 'F', 'S', '分期应收款日 ＝ 期初后(天)', 0,
     '★ 分期付款用：合同约定「每期期初付」就填 0，「期初后 15 天内付清」就填 15，'
     '「期末付」就填每期天数。《分期收款计划》按这个算逾期。'),
]
for rr, l0, l1, v0, v1, n0, n1, lt, vv, nt in PARAMS:
    put(ws, f'{l0}{rr}', lt, font=F_TOT, fill=FILL_SEC, align=CR)
    for cc in range(ord(l0), ord(l1) + 1):
        put(ws, f'{chr(cc)}{rr}', None, font=F_TOT, fill=FILL_SEC)
    ws.merge_cells(f'{l0}{rr}:{l1}{rr}')
    put(ws, f'{l0}{rr}', lt, font=F_TOT, fill=FILL_SEC, align=CR)
    put(ws, f'{v0}{rr}', vv, font=F_BIG, fill=FILL_IN, fmt='0')
    for cc in range(ord(v0), ord(v1) + 1):
        put(ws, f'{chr(cc)}{rr}', None, font=F_BIG, fill=FILL_IN)
    ws.merge_cells(f'{v0}{rr}:{v1}{rr}')
    put(ws, f'{v0}{rr}', vv, font=F_BIG, fill=FILL_IN, fmt='0')
    put(ws, f'{n0}{rr}', nt, font=F_NOTE, align=CL_, border=None)
    ws.merge_cells(f'{n0}{rr}:{n1}{rr}')
    ws.row_dimensions[rr].height = 24

LISTS = [('A', '片区 / 站点', AREAS), ('C', '梯型', LIFT_TYPES), ('E', '维保性质', MAINT_KIND),
         ('G', '维保状态', MAINT_STATE), ('I', '合同状态', CON_STATE)]
for col, hdr_, vals in LISTS:
    put(ws, f'{col}{BH}', hdr_, font=F_HDR, fill=PatternFill('solid', fgColor=C_BASE))
    for i in range(N_LIST):
        put(ws, f'{col}{B0+i}', vals[i] if i < len(vals) else None, font=F_IN, fill=FILL_IN, align=CL_)
PAIRS = [('K', '收费周期', 'L', '每年期数 ★', PAY_CYCLE),
         ('M', '费用类别', 'N', '款别 ★', FEE_ITEMS),
         ('P', '票据类型', 'Q', '收款归类 ★', BILL_TYPES)]
for c1, h1, c2, h2, data in PAIRS:
    put(ws, f'{c1}{BH}', h1, font=F_HDR, fill=PatternFill('solid', fgColor=C_BASE))
    put(ws, f'{c2}{BH}', h2, font=F_HDR, fill=PatternFill('solid', fgColor=C_BASE))
    for i in range(N_LIST):
        a, b = data[i] if i < len(data) else (None, None)
        put(ws, f'{c1}{B0+i}', a, font=F_IN, fill=FILL_IN, align=CL_)
        put(ws, f'{c2}{B0+i}', b, font=F_IN, fill=FILL_IN, fmt='0' if c2 == 'L' else None)
put(ws, f'S{BH}', '收款方式', font=F_HDR, fill=PatternFill('solid', fgColor=C_BASE))
for i in range(N_LIST):
    put(ws, f'S{B0+i}', PAY_WAYS[i] if i < len(PAY_WAYS) else None, font=F_IN, fill=FILL_IN, align=CL_)
ws.row_dimensions[BH].height = 30

widths(ws, {'A': 16, 'B': 2, 'C': 14, 'D': 2, 'E': 12, 'F': 20, 'G': 14, 'H': 8, 'I': 14,
            'J': 20, 'K': 14, 'L': 11, 'M': 20, 'N': 13, 'O': 8, 'P': 18, 'Q': 13, 'R': 2, 'S': 14})
dv(ws, f'N{B0}:N{BE}', '"' + ','.join(['合同款', '合同外款']) + '"', block=True)
dv(ws, f'Q{B0}:Q{BE}', '"' + ','.join(['发票', '收据', '现金无票']) + '"', block=True)
ws.freeze_panes = f'A{B0}'
page(ws)

name('管理年度', '基础资料!$D$3')
name('合同提醒天', '基础资料!$O$3')
name('年检提醒天', '基础资料!$D$4')
name('账龄预警天', '基础资料!$O$4')
name('账期天数', '基础资料!$D$5')
for nm, col in [('片区表', 'A'), ('梯型表', 'C'), ('维保性质表', 'E'), ('维保状态表', 'G'),
                ('合同状态表', 'I'), ('收费周期表', 'K'), ('每年期数表', 'L'), ('费用类别表', 'M'),
                ('费用款别', 'N'), ('票据类型表', 'P'), ('票据归类表', 'Q'), ('收款方式表', 'S')]:
    name(nm, f'基础资料!${col}${B0}:${col}${BE}')
print('  ✓ 基础资料')

# ══════════════════════════════════════════════════════════════
# ② 电梯资料
# ══════════════════════════════════════════════════════════════
ws = sheet('电梯资料')
HD_L = ['序号', '片区 ★', '使用单位 ★', '项目名称 ★', '客户项目(自动键)', '型号', '设备注册代码',
        '救援标识牌编号', '生产厂家', '梯型 ★', '层', '站', '门', '载重', '速度', '台量 ★',
        '设备编号', '使用编号', '制造日期', '下次年检日期', '限速器校验(下次)', '载重试验125%(下次)',
        '维保人员', '维保性质 ★', '维保状态 ★', '免保到期日', '年检提醒', '校验提醒', '状态提醒',
        '注意事项 / 备注', '核对']
title(ws, '电 梯 资 料 · 一台梯子一行', 'AE',
      '★ 淡黄＝手工填，浅灰＝公式自动算，别往灰格里打字。\n'
      '★ E 列「客户项目」＝使用单位＋项目名称，是全表通用的那把钥匙，后面《维保合同》《应收登记》《收款登记》都按它对上号；'
      '同一个单位有几个项目，就按项目分开写。\n'
      '★ 「维保状态」选 安装免保 / 技术免保 / 质保期内 的，一定要填「免保到期日」，到期前后会在《到期与提醒》里跳出来提示转收费。')
headers(ws, HDR, HD_L)
put(ws, f'A{HDR-1}', '合计 →', font=F_TOT, fill=FILL_TOT, align=CR)
put(ws, f'P{HDR-1}', f'=SUM($P${L0}:$P${L1})', font=F_TOT, fill=FILL_TOT, fmt=NUM)
put(ws, f'C{HDR-1}', f'=COUNTIF($C${L0}:$C${L1},"?*")&" 个单位行"', font=F_TOT, fill=FILL_TOT)
put(ws, f'Y{HDR-1}', f'=COUNTIF($Y${L0}:$Y${L1},"正常维保")&" 台正常维保"', font=F_TOT, fill=FILL_TOT)
put(ws, f'Z{HDR-1}', f'=(COUNTIF($Y${L0}:$Y${L1},"安装免保")+COUNTIF($Y${L0}:$Y${L1},"技术免保")'
                     f'+COUNTIF($Y${L0}:$Y${L1},"质保期内"))&" 台免保"', font=F_TOT, fill=FILL_TOT)
ws.row_dimensions[HDR - 1].height = 20

for r in range(L0, L1 + 1):
    put(ws, f'A{r}', f'=IF($C{r}="","",ROW()-{HDR})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    for c in ['B', 'C', 'D', 'F', 'G', 'H', 'I', 'J', 'Q', 'R', 'W', 'X', 'Y', 'AD']:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL_ if c in ('C', 'D', 'AD') else C)
    for c in ['K', 'L', 'M']:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, fmt='0')
    for c in ['N', 'O']:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN)
    put(ws, f'P{r}', None, font=F_IN, fill=FILL_IN, fmt=NUM)
    for c in ['S', 'T', 'U', 'V', 'Z']:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, fmt=DATEF)
    put(ws, f'E{r}', f'=IF($C{r}="","",$C{r}&" ｜ "&$D{r})', font=F_AUTO, fill=FILL_AUTO, align=CL_)
    put(ws, f'AA{r}', f'=IF($T{r}="","",IF($T{r}<TODAY(),"已过期 "&TEXT(TODAY()-$T{r},"0")&" 天",'
                      f'IF($T{r}-TODAY()<=年检提醒天,"还剩 "&TEXT($T{r}-TODAY(),"0")&" 天","")))',
        font=F_WARN, fill=FILL_AUTO)
    put(ws, f'AB{r}', f'=IF(COUNT($U{r}:$V{r})=0,"",IF(MIN($U{r}:$V{r})<TODAY(),'
                      f'"已过期 "&TEXT(TODAY()-MIN($U{r}:$V{r}),"0")&" 天",'
                      f'IF(MIN($U{r}:$V{r})-TODAY()<=年检提醒天,"还剩 "&TEXT(MIN($U{r}:$V{r})-TODAY(),"0")&" 天","")))',
        font=F_WARN, fill=FILL_AUTO)
    put(ws, f'AC{r}', f'=IF($C{r}="","",IF(OR($Y{r}="安装免保",$Y{r}="技术免保",$Y{r}="质保期内"),'
                      f'IF($Z{r}="","▲免保没填到期日",'
                      f'IF($Z{r}<TODAY(),"▲免保已到期 该转收费了",'
                      f'IF($Z{r}-TODAY()<=合同提醒天,"免保还剩 "&TEXT($Z{r}-TODAY(),"0")&" 天","免保中"))),'
                      f'IF($Y{r}="暂停维保","暂停中 不计费",IF($Y{r}="已解约","已解约",'
                      f'IF(COUNTIF(合同键,$E{r})=0,"▲还没建合同行","")))))', font=F_WARN, fill=FILL_AUTO)
    put(ws, f'AE{r}', f'=IF($C{r}="","",IF($D{r}="","★项目名称没填",IF(N($P{r})<=0,"★台量要填",'
                      f'IF(AND(OR($Y{r}="安装免保",$Y{r}="技术免保",$Y{r}="质保期内"),$Z{r}=""),"★免保到期日没填",'
                      f'IF(COUNTIF(合同键,$E{r})=0,"★还没建合同行","OK")))))', font=F_AUTO, fill=FILL_AUTO)

for i, L in enumerate(ALL_LIFTS):
    r, x = L0 + i, LIFT_EXTRA[i]
    unit, proj, model, regno, resq, maker, kind, ce, zhan, men, load, spd, tai, devno, useno = L
    area, man, mk, mstate, freeend, made, chk1, chk2, chk3, memo = x
    for col, v in zip('BCDFGHIJKLMNOPQR',
                      [area, unit, proj, model, regno, resq, maker, kind, ce, zhan, men, load, spd, tai, devno, useno]):
        ws[f'{col}{r}'] = v
    for col, v in zip(['S', 'T', 'U', 'V', 'Z'], [made, chk1, chk2, chk3, freeend]):
        ws[f'{col}{r}'] = v
    ws[f'W{r}'], ws[f'X{r}'], ws[f'Y{r}'] = man, mk, mstate
    ws[f'AD{r}'] = memo

widths(ws, {'A': 5, 'B': 13, 'C': 30, 'D': 18, 'E': 34, 'F': 16, 'G': 22, 'H': 14, 'I': 18, 'J': 12,
            'K': 5, 'L': 5, 'M': 5, 'N': 9, 'O': 9, 'P': 6, 'Q': 16, 'R': 11, 'S': 12, 'T': 12,
            'U': 13, 'V': 15, 'W': 10, 'X': 10, 'Y': 11, 'Z': 12, 'AA': 14, 'AB': 14, 'AC': 20,
            'AD': 24, 'AE': 14})
dv(ws, f'B{L0}:B{L1}', dynlist('基础资料', 'A', B0, BE), block=True)
dv(ws, f'J{L0}:J{L1}', dynlist('基础资料', 'C', B0, BE), block=True)
dv(ws, f'X{L0}:X{L1}', dynlist('基础资料', 'E', B0, BE), block=True)
dv(ws, f'Y{L0}:Y{L1}', dynlist('基础资料', 'G', B0, BE),
   msg='安装免保 / 技术免保 / 质保期内 这三种，记得把右边「免保到期日」填上。', block=True)
ws.freeze_panes = 'F5'
ws.auto_filter.ref = f'A{HDR}:AE{L1}'
page(ws, titles=f'{HDR}:{HDR}')

name('电梯键', f'电梯资料!$E${L0}:$E${L1}')
name('电梯片区', f'电梯资料!$B${L0}:$B${L1}')
name('电梯单位', f'电梯资料!$C${L0}:$C${L1}')
name('电梯台量', f'电梯资料!$P${L0}:$P${L1}')
name('电梯状态', f'电梯资料!$Y${L0}:$Y${L1}')
print('  ✓ 电梯资料')

# ══════════════════════════════════════════════════════════════
# ③ 维保合同
# ══════════════════════════════════════════════════════════════
ws = sheet('维保合同')
HD_K = ['序号', '片区(自动)', '使用单位 ★', '项目名称 ★', '客户项目(自动键)', '合同编号', '合同状态 ★',
        '合同起 ★', '合同止 ★', '合同台量 ★', '在册台量(自动)', '合同年费(含税) ★', '收费周期 ★',
        '维保性质 ★', '免保台量(自动)', '收费台量(自动)', '到期提醒', '合同回签提醒', '台量核对',
        '本年应收', '本年已收', '欠款余额(累计)', '备注', '核对']
title(ws, '维 保 合 同 · 一个项目一行，合同没回来照样先建行', 'X',
      '★ 一个「客户项目」只能建一行，重复了 X 列会报「★键重复」。\n'
      '★ 合同没及时回来：F 列合同编号先空着，G 列选「已签未回 / 未签 / 续签中」，照样可以到《应收登记》挂应收；'
      'R 列会一直提醒「▲合同未回」，《到期与提醒》里也单列一块，方便催合同。\n'
      '★ T/U/V 三列是自动汇总，明细看《应收登记》《收款登记》，对账看《应收对账》。')
headers(ws, HDR, HD_K)
put(ws, f'A{HDR-1}', '合计 →', font=F_TOT, fill=FILL_TOT, align=CR)
for c, f in [('J', 'SUM'), ('K', 'SUM'), ('L', 'SUM'), ('O', 'SUM'), ('P', 'SUM'),
             ('T', 'SUM'), ('U', 'SUM'), ('V', 'SUM')]:
    put(ws, f'{c}{HDR-1}', f'={f}(${c}${K0}:${c}${K1})', font=F_TOT, fill=FILL_TOT,
        fmt=NUM if c in ('J', 'K', 'O', 'P') else MONEY)
put(ws, f'C{HDR-1}', f'=COUNTIF($C${K0}:$C${K1},"?*")&" 个项目"', font=F_TOT, fill=FILL_TOT)
put(ws, f'R{HDR-1}', f'=COUNTIF($R${K0}:$R${K1},"▲*")&" 份合同未回"', font=F_WARN, fill=FILL_TOT)
ws.row_dimensions[HDR - 1].height = 20

for r in range(K0, K1 + 1):
    put(ws, f'A{r}', f'=IF($C{r}="","",ROW()-{HDR})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    for c in ['C', 'D', 'F', 'G', 'M', 'N', 'W']:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL_ if c in ('C', 'D', 'W') else C)
    for c in ['H', 'I']:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, fmt=DATEF)
    put(ws, f'J{r}', None, font=F_IN, fill=FILL_IN, fmt=NUM)
    put(ws, f'L{r}', None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'B{r}', f'=IF($C{r}="","",IFERROR(INDEX(电梯片区,MATCH($E{r},电梯键,0)),""))',
        font=F_AUTO, fill=FILL_AUTO)
    put(ws, f'E{r}', f'=IF($C{r}="","",$C{r}&" ｜ "&$D{r})', font=F_AUTO, fill=FILL_AUTO, align=CL_)
    put(ws, f'K{r}', f'=IF($C{r}="","",SUMIF(电梯键,$E{r},电梯台量))', font=F_AUTO, fill=FILL_AUTO, fmt=NUM)
    put(ws, f'O{r}', f'=IF($C{r}="","",SUMIFS(电梯台量,电梯键,$E{r},电梯状态,"安装免保")'
                     f'+SUMIFS(电梯台量,电梯键,$E{r},电梯状态,"技术免保")'
                     f'+SUMIFS(电梯台量,电梯键,$E{r},电梯状态,"质保期内"))', font=F_AUTO, fill=FILL_AUTO, fmt=NUM)
    put(ws, f'P{r}', f'=IF($C{r}="","",$K{r}-$O{r})', font=F_AUTO, fill=FILL_AUTO, fmt=NUM)
    put(ws, f'Q{r}', f'=IF($C{r}="","",IF($G{r}="已终止","",IF($I{r}="","▲没填合同止期",'
                     f'IF($I{r}<TODAY(),"已到期 "&TEXT(TODAY()-$I{r},"0")&" 天",'
                     f'IF($I{r}-TODAY()<=合同提醒天,"将到期 "&TEXT($I{r}-TODAY(),"0")&" 天","")))))',
        font=F_WARN, fill=FILL_AUTO)
    put(ws, f'R{r}', f'=IF($C{r}="","",IF($G{r}="","▲合同状态没选",'
                     f'IF(OR($G{r}="已签未回",$G{r}="未签",$G{r}="续签中"),"▲合同未回（"&$G{r}&"）","")))',
        font=F_WARN, fill=FILL_AUTO)
    put(ws, f'S{r}', f'=IF($C{r}="","",IF(N($J{r})=N($K{r}),"一致",'
                     f'"★合同"&TEXT(N($J{r}),"0")&"台 / 在册"&TEXT(N($K{r}),"0")&"台"))',
        font=F_AUTO, fill=FILL_AUTO)
    put(ws, f'T{r}', f'=IF($C{r}="","",ROUND(SUMIFS(应收金额,应收键,$E{r},应收年度,管理年度),2))',
        font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'U{r}', f'=IF($C{r}="","",ROUND(SUMIFS(收款金额,收款键,$E{r},收款年度,管理年度),2))',
        font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'V{r}', f'=IF($C{r}="","",ROUND(SUMIF(应收键,$E{r},应收金额)-SUMIF(收款键,$E{r},收款金额),2))',
        font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'X{r}', f'=IF($C{r}="","",IF($D{r}="","★项目名称没填",'
                     f'IF(COUNTIF($E${K0}:$E${K1},$E{r})>1,"★键重复",'
                     f'IF(AND($H{r}<>"",$I{r}<>"",$I{r}<$H{r}),"★合同起止反了",'
                     f'IF(N($K{r})=0,"★电梯资料里没有这个项目","OK")))))', font=F_AUTO, fill=FILL_AUTO)

for i, K in enumerate(CONTRACTS):
    r = K0 + i
    unit, proj, area, cno, cst, d0, d1, tai, fee, cyc, kind, memo = K
    ws[f'C{r}'], ws[f'D{r}'], ws[f'F{r}'], ws[f'G{r}'] = unit, proj, cno, cst
    ws[f'H{r}'], ws[f'I{r}'], ws[f'J{r}'], ws[f'L{r}'] = d0, d1, tai, fee
    ws[f'M{r}'], ws[f'N{r}'], ws[f'W{r}'] = cyc, kind, memo

# 分期付款的几个底数（隐藏辅助列，《分期收款计划》直接读这几列）
for r in range(K0, K1 + 1):
    put(ws, f'Y{r}', f'=IF($C{r}="",0,IFERROR(INDEX(每年期数表,MATCH($M{r},收费周期表,0)),0))',
        font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'Z{r}', f'=IF(OR($C{r}="",$H{r}="",$I{r}=""),0,MAX(1,ROUND(($I{r}+1-$H{r})/30.4375,0)))',
        font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'AA{r}', f'=IF($C{r}="",0,ROUND(N($L{r})*$Z{r}/12,2))', font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'AB{r}', f'=IF(N($Y{r})<=0,0,IF($Y{r}=1,1,MIN({MAX_TERM},MAX(1,ROUND($Z{r}*$Y{r}/12,0)))))',
        font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'AC{r}', f'=IF(N($AB{r})=0,0,ROUND($AA{r}/$AB{r},2))', font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
for _c in ['Y', 'Z', 'AA', 'AB', 'AC']:
    ws.column_dimensions[_c].hidden = True
for _c, _h in [('Y', '每年期数'), ('Z', '合同月数'), ('AA', '合同总额'),
               ('AB', '分期期数'), ('AC', '每期金额')]:
    put(ws, f'{_c}{HDR}', _h, font=F_HDR, fill=PatternFill('solid', fgColor=C_MAIN))

widths(ws, {'A': 5, 'B': 13, 'C': 30, 'D': 18, 'E': 34, 'F': 14, 'G': 11, 'H': 12, 'I': 12,
            'J': 9, 'K': 10, 'L': 12, 'M': 11, 'N': 10, 'O': 11, 'P': 11, 'Q': 15, 'R': 20,
            'S': 18, 'T': 13, 'U': 13, 'V': 14, 'W': 26, 'X': 22, 'Y': 10, 'Z': 10,
            'AA': 13, 'AB': 10, 'AC': 12})
dv(ws, f'G{K0}:G{K1}', dynlist('基础资料', 'I', B0, BE),
   msg='合同还没回来就选「已签未回 / 未签 / 续签中」，照样能挂应收。', block=True)
dv(ws, f'M{K0}:M{K1}', dynlist('基础资料', 'K', B0, BE), block=True)
dv(ws, f'N{K0}:N{K1}', dynlist('基础资料', 'E', B0, BE), block=True)
ws.freeze_panes = 'F5'
ws.auto_filter.ref = f'A{HDR}:X{K1}'
page(ws, titles=f'{HDR}:{HDR}')

name('合同键', f'维保合同!$E${K0}:$E${K1}')
name('合同单位', f'维保合同!$C${K0}:$C${K1}')
name('合同项目', f'维保合同!$D${K0}:$D${K1}')
name('合同片区', f'维保合同!$B${K0}:$B${K1}')
name('合同台量', f'维保合同!$K${K0}:$K${K1}')
name('合同状态', f'维保合同!$G${K0}:$G${K1}')
print('  ✓ 维保合同')


def dynlist_f(sheet_name, col, r0, r1):
    """公式列的下拉区：COUNTA 会把「返回空串的公式格」也算成有内容，所以用 COUNTIF(...,"?*")"""
    rng = f'{sheet_name}!${col}${r0}:${col}${r1}'
    return f'=OFFSET({sheet_name}!${col}${r0},0,0,MAX(1,COUNTIF({rng},"?*")),1)'


KEYS = [c[0] + ' ｜ ' + c[1] for c in CONTRACTS]

# ══════════════════════════════════════════════════════════════
# ④ 应收登记
# ══════════════════════════════════════════════════════════════
ws = sheet('应收登记')
HD_A = ['序号', '应收单号(自动)', '登记日期 ★', '客户项目 ★', '使用单位(自动)', '项目名称(自动)',
        '费用类别 ★', '款别(自动)', '计费期间起', '计费期间止', '应收金额(含税) ★', '开票日期',
        '发票号', '开票金额', '开票状态', '已收(自动核销)', '未收余额', '年度(自动)', '备注', '核对',
        '开票年度', '指定冲销额', '净应收', '同项目累计净应收', '行号', '归期日']
title(ws, '应 收 登 记 · 该收人家多少钱，一笔一行', 'T',
      '★ 合同款、合同外款都记在这一张表里：G 列选「费用类别」，H 列会自动判成「合同款」还是「合同外款」，'
      '不用分两张表记，最后在《应收对账》里自动拆开给你看。\n'
      '★ 合同没回来一样能挂：D 列只要在《维保合同》里建过行就能选到（哪怕合同状态是「未签」）。\n'
      '★ 开票和收款是两码事：这里填的是「开没开票、开了多少」；钱几时到、开的是发票还是收据，记到《收款登记》。\n'
      '★ P/Q 两列是自动核销出来的：收款里指定了单号的先冲那一张，没指定的按「先挂先冲」往下顺，所以每张单都看得出还差多少。',
      color=C_BIZ)
headers(ws, HDR, HD_A, fill_color=C_BIZ)
put(ws, f'A{HDR-1}', '合计 →', font=F_TOT, fill=FILL_TOT, align=CR)
for c in ['K', 'N', 'P', 'Q']:
    put(ws, f'{c}{HDR-1}', f'=ROUND(SUM(${c}${A0}:${c}${A1}),2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'C{HDR-1}', f'=COUNT($C${A0}:$C${A1})&" 笔"', font=F_TOT, fill=FILL_TOT)
put(ws, f'T{HDR-1}', f'=IF(COUNTIF($T${A0}:$T${A1},"★*")=0,"全部 OK",'
                     f'COUNTIF($T${A0}:$T${A1},"★*")&" 笔要改")', font=F_WARN, fill=FILL_TOT)
ws.row_dimensions[HDR - 1].height = 20

for r in range(A0, A1 + 1):
    put(ws, f'A{r}', f'=IF($C{r}="","",ROW()-{HDR})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'B{r}', f'=IF($C{r}="","","YS"&TEXT($C{r},"yyyymm")&"-"&TEXT(ROW()-{HDR},"0000"))',
        font=F_AUTO, fill=FILL_AUTO)
    put(ws, f'C{r}', None, font=F_IN, fill=FILL_IN, fmt=DATEF)
    put(ws, f'D{r}', None, font=F_IN, fill=FILL_IN, align=CL_)
    put(ws, f'E{r}', f'=IF($D{r}="","",IFERROR(INDEX(合同单位,MATCH($D{r},合同键,0)),'
                     f'IFERROR(TRIM(LEFT($D{r},FIND("｜",$D{r})-1)),$D{r})))', font=F_AUTO, fill=FILL_AUTO, align=CL_)
    put(ws, f'F{r}', f'=IF($D{r}="","",IFERROR(INDEX(合同项目,MATCH($D{r},合同键,0)),'
                     f'IFERROR(TRIM(MID($D{r},FIND("｜",$D{r})+1,99)),"")))', font=F_AUTO, fill=FILL_AUTO, align=CL_)
    put(ws, f'G{r}', None, font=F_IN, fill=FILL_IN, align=CL_)
    put(ws, f'H{r}', f'=IF($G{r}="","",IFERROR(INDEX(费用款别,MATCH($G{r},费用类别表,0)),"合同外款"))',
        font=F_AUTO, fill=FILL_AUTO)
    for c in ['I', 'J', 'L']:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, fmt=DATEF)
    put(ws, f'K{r}', None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'M{r}', None, font=F_IN, fill=FILL_IN)
    put(ws, f'N{r}', None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'O{r}', f'=IF($C{r}="","",IF(N($N{r})<=0,"未开票",'
                     f'IF(ROUND(N($N{r})-N($K{r}),2)>=0,"已开票",'
                     f'IFERROR("部分开票 "&TEXT(N($N{r})/N($K{r}),"0%"),"部分开票"))))', font=F_AUTO, fill=FILL_AUTO)
    put(ws, f'P{r}', f'=IF($C{r}="","",ROUND($V{r}+MIN($W{r},MAX(0,'
                     f'(SUMIF(收款键,$D{r},收款金额)-SUMIFS(收款金额,收款键,$D{r},收款关联,"<>"))'
                     f'-($X{r}-$W{r}))),2))', font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'Q{r}', f'=IF($C{r}="","",ROUND(N($K{r})-N($P{r}),2))', font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'R{r}', f'=IF($C{r}="","",YEAR($C{r}))', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'S{r}', None, font=F_IN, fill=FILL_IN, align=CL_)
    put(ws, f'T{r}', f'=IF($C{r}="","",IF($D{r}="","★没选客户项目",'
                     f'IF(COUNTIF(合同键,$D{r})=0,"★《维保合同》里没这个项目",'
                     f'IF(N($K{r})<=0,"★应收金额要填",'
                     f'IF(ROUND(N($N{r})-N($K{r}),2)>0.001,"★开票金额大于应收",'
                     f'IF(AND($L{r}<>"",N($N{r})<=0),"★填了开票日期没填开票金额",'
                     f'IF(AND($I{r}<>"",$J{r}<>"",$J{r}<$I{r}),"★计费期间起止反了","OK")))))))',
        font=F_AUTO, fill=FILL_AUTO)
    put(ws, f'U{r}', f'=IF($L{r}="","",YEAR($L{r}))', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'V{r}', f'=IF($C{r}="","",ROUND(SUMIF(收款关联,$B{r},收款金额),2))',
        font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'W{r}', f'=IF($C{r}="","",MAX(0,ROUND(N($K{r})-N($V{r}),2)))', font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'X{r}', f'=IF($C{r}="","",ROUND(SUMIFS(应收净额,应收键,$D{r},应收行号,"<="&ROW()),2))',
        font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'Y{r}', f'=IF($C{r}="","",ROW())', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'Z{r}', f'=IF($C{r}="","",IF($I{r}<>"",$I{r},$C{r}))', font=F_AUTO, fill=FILL_AUTO, fmt=DATEQ)

for i, A in enumerate(AR_ROWS):
    r = A0 + i
    d, ci, item, p0, p1, amt, idate, ino, iamt, memo = A
    ws[f'C{r}'], ws[f'D{r}'], ws[f'G{r}'] = d, KEYS[ci - 1], item
    ws[f'I{r}'], ws[f'J{r}'], ws[f'K{r}'] = p0, p1, amt
    ws[f'L{r}'], ws[f'M{r}'] = idate, (ino or None)
    ws[f'N{r}'] = iamt if iamt else None
    ws[f'S{r}'] = memo or None

widths(ws, {'A': 5, 'B': 16, 'C': 12, 'D': 34, 'E': 28, 'F': 17, 'G': 15, 'H': 10, 'I': 12, 'J': 12,
            'K': 13, 'L': 12, 'M': 12, 'N': 12, 'O': 12, 'P': 13, 'Q': 13, 'R': 8, 'S': 26, 'T': 24,
            'U': 10, 'V': 12, 'W': 12, 'X': 16, 'Y': 8, 'Z': 12})
for c in 'UVWXYZ':
    ws.column_dimensions[c].hidden = True
dv(ws, f'D{A0}:D{A1}', dynlist_f('维保合同', 'E', K0, K1),
   msg='从《维保合同》建过的「客户项目」里选。没有就先去《维保合同》补一行，合同状态选「未签」也行。')
dv(ws, f'G{A0}:G{A1}', dynlist('基础资料', 'M', B0, BE),
   msg='维保费＝合同款；维修/配件/报检/年检等＝合同外款。右边 H 列自动判。', block=True)
ws.freeze_panes = 'E5'
ws.auto_filter.ref = f'A{HDR}:T{A1}'
page(ws, titles=f'{HDR}:{HDR}')

name('应收单号', f'应收登记!$B${A0}:$B${A1}')
name('应收日期', f'应收登记!$C${A0}:$C${A1}')
name('应收键', f'应收登记!$D${A0}:$D${A1}')
name('应收类别', f'应收登记!$G${A0}:$G${A1}')
name('应收款别', f'应收登记!$H${A0}:$H${A1}')
name('应收金额', f'应收登记!$K${A0}:$K${A1}')
name('应收开票额', f'应收登记!$N${A0}:$N${A1}')
name('应收未收', f'应收登记!$Q${A0}:$Q${A1}')
name('应收年度', f'应收登记!$R${A0}:$R${A1}')
name('应收开票年度', f'应收登记!$U${A0}:$U${A1}')
name('应收净额', f'应收登记!$W${A0}:$W${A1}')
name('应收行号', f'应收登记!$Y${A0}:$Y${A1}')
name('应收归期日', f'应收登记!$Z${A0}:$Z${A1}')
name('应收已收', f'应收登记!$P${A0}:$P${A1}')
print('  ✓ 应收登记')

# ══════════════════════════════════════════════════════════════
# ⑤ 收款登记
# ══════════════════════════════════════════════════════════════
ws = sheet('收款登记')
HD_R = ['序号', '收款单号(自动)', '收款日期 ★', '客户项目 ★', '使用单位(自动)', '项目名称(自动)',
        '收款金额 ★', '票据类型 ★', '票据归类(自动)', '收款方式 ★', '款别 ★', '关联应收单号(选填)',
        '年度(自动)', '备注', '核对']
title(ws, '收 款 登 记 · 钱几时到、开的什么票，一笔一行', 'O',
      '★ 收款乱就乱在「有的开发票、有的开收据、有的干脆现金没票」：H 列选票据类型，'
      'I 列自动归成 发票 / 收据 / 现金无票 三路，《应收对账》里分三列给你看。\n'
      '★ 合同款和合同外款可以合并收：K 列选「合并(不细分)」，L 列不用填，系统按「先挂先冲」自动去冲最早那几笔应收。\n'
      '★ 想指定冲哪一张应收单（比如客户特意说这笔是付去年的维修费），就在 L 列把那张应收单号选上，冲销优先按你指定的来。\n'
      '★ 这张表不分年度、一直往下记，过年不用新建表。',
      color=C_BIZ)
headers(ws, HDR, HD_R, fill_color=C_BIZ)
put(ws, f'A{HDR-1}', '合计 →', font=F_TOT, fill=FILL_TOT, align=CR)
put(ws, f'G{HDR-1}', f'=ROUND(SUM($G${R0}:$G${R1}),2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'C{HDR-1}', f'=COUNT($C${R0}:$C${R1})&" 笔"', font=F_TOT, fill=FILL_TOT)
put(ws, f'I{HDR-1}', f'=ROUND(SUMIF($I${R0}:$I${R1},"发票",$G${R0}:$G${R1}),2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'J{HDR-1}', f'=ROUND(SUMIF($I${R0}:$I${R1},"收据",$G${R0}:$G${R1}),2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'K{HDR-1}', f'=ROUND(SUMIF($I${R0}:$I${R1},"现金无票",$G${R0}:$G${R1}),2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'O{HDR-1}', f'=IF(COUNTIF($O${R0}:$O${R1},"★*")=0,"全部 OK",'
                     f'COUNTIF($O${R0}:$O${R1},"★*")&" 笔要改")', font=F_WARN, fill=FILL_TOT)
ws.row_dimensions[HDR - 1].height = 20

for r in range(R0, R1 + 1):
    put(ws, f'A{r}', f'=IF($C{r}="","",ROW()-{HDR})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'B{r}', f'=IF($C{r}="","","SK"&TEXT($C{r},"yyyymm")&"-"&TEXT(ROW()-{HDR},"0000"))',
        font=F_AUTO, fill=FILL_AUTO)
    put(ws, f'C{r}', None, font=F_IN, fill=FILL_IN, fmt=DATEF)
    put(ws, f'D{r}', None, font=F_IN, fill=FILL_IN, align=CL_)
    put(ws, f'E{r}', f'=IF($D{r}="","",IFERROR(INDEX(合同单位,MATCH($D{r},合同键,0)),'
                     f'IFERROR(TRIM(LEFT($D{r},FIND("｜",$D{r})-1)),$D{r})))', font=F_AUTO, fill=FILL_AUTO, align=CL_)
    put(ws, f'F{r}', f'=IF($D{r}="","",IFERROR(INDEX(合同项目,MATCH($D{r},合同键,0)),'
                     f'IFERROR(TRIM(MID($D{r},FIND("｜",$D{r})+1,99)),"")))', font=F_AUTO, fill=FILL_AUTO, align=CL_)
    put(ws, f'G{r}', None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    for c in ['H', 'J', 'K']:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN)
    put(ws, f'I{r}', f'=IF($H{r}="","",IFERROR(INDEX(票据归类表,MATCH($H{r},票据类型表,0)),"现金无票"))',
        font=F_AUTO, fill=FILL_AUTO)
    put(ws, f'L{r}', None, font=F_IN, fill=FILL_IN)
    put(ws, f'M{r}', f'=IF($C{r}="","",YEAR($C{r}))', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'N{r}', None, font=F_IN, fill=FILL_IN, align=CL_)
    put(ws, f'O{r}', f'=IF($C{r}="","",IF($D{r}="","★没选客户项目",'
                     f'IF(COUNTIF(合同键,$D{r})=0,"★《维保合同》里没这个项目",'
                     f'IF(N($G{r})<=0,"★收款金额要填",'
                     f'IF($H{r}="","★票据类型要选",'
                     f'IF(AND($L{r}<>"",COUNTIF(应收单号,$L{r})=0),"★关联的应收单号不存在",'
                     f'IF(AND($L{r}<>"",IFERROR(INDEX(应收键,MATCH($L{r},应收单号,0)),"")<>$D{r}),'
                     f'"★关联单号不是这个客户的",'
                     f'IF(AND($L{r}<>"",ROUND(SUMIF(收款关联,$L{r},收款金额)'
                     f'-IFERROR(INDEX(应收金额,MATCH($L{r},应收单号,0)),0),2)>0.001),'
                     f'"★指定冲销超过那张单的应收","OK"))))))))', font=F_AUTO, fill=FILL_AUTO)

for i, RC in enumerate(RC_ROWS):
    r = R0 + i
    d, ci, amt, bt, way, kb, link, memo = RC
    ws[f'C{r}'], ws[f'D{r}'], ws[f'G{r}'] = d, KEYS[ci - 1], amt
    ws[f'H{r}'], ws[f'J{r}'], ws[f'K{r}'] = bt, way, kb
    if link:
        ws[f'L{r}'] = 'YS%s-%04d' % (AR_ROWS[link - 1][0].strftime('%Y%m'), link)
    ws[f'N{r}'] = memo or None

widths(ws, {'A': 5, 'B': 16, 'C': 12, 'D': 34, 'E': 28, 'F': 17, 'G': 13, 'H': 17, 'I': 11,
            'J': 11, 'K': 13, 'L': 16, 'M': 8, 'N': 26, 'O': 24})
dv(ws, f'D{R0}:D{R1}', dynlist_f('维保合同', 'E', K0, K1),
   msg='从《维保合同》建过的「客户项目」里选。')
dv(ws, f'H{R0}:H{R1}', dynlist('基础资料', 'P', B0, BE),
   msg='开的什么票就选什么：专票/普票/电子票 → 发票；收据 → 收据；什么都没开 → 无票。', block=True)
dv(ws, f'J{R0}:J{R1}', dynlist('基础资料', 'S', B0, BE), block=True)
dv(ws, f'K{R0}:K{R1}', '"' + ','.join(KUAN_BIE) + '"',
   msg='一次同时收了合同款和合同外款，就选「合并(不细分)」。', block=True)
dv(ws, f'L{R0}:L{R1}', dynlist_f('应收登记', 'B', A0, A1),
   msg='选填。要指定这笔钱冲哪一张应收单才填；不填就按「先挂先冲」自动核销。')
ws.freeze_panes = 'E5'
ws.auto_filter.ref = f'A{HDR}:O{R1}'
page(ws, titles=f'{HDR}:{HDR}')

name('收款单号', f'收款登记!$B${R0}:$B${R1}')
name('收款日期', f'收款登记!$C${R0}:$C${R1}')
name('收款键', f'收款登记!$D${R0}:$D${R1}')
name('收款金额', f'收款登记!$G${R0}:$G${R1}')
name('收款票据', f'收款登记!$H${R0}:$H${R1}')
name('收款归类', f'收款登记!$I${R0}:$I${R1}')
name('收款方式', f'收款登记!$J${R0}:$J${R1}')
name('收款款别', f'收款登记!$K${R0}:$K${R1}')
name('收款关联', f'收款登记!$L${R0}:$L${R1}')
name('收款年度', f'收款登记!$M${R0}:$M${R1}')
print('  ✓ 收款登记')

# ══════════════════════════════════════════════════════════════
# ⑥ 应收对账（一个项目一行，跟《维保合同》行对行）
# ══════════════════════════════════════════════════════════════
ws = sheet('应收对账')
HD_D = ['序号', '使用单位', '项目名称', '合同编号', '合同状态', '合同到期', '在册台量',
        '上年结转欠款', '本年应收合计', '其中：合同款', '其中：合同外款', '本年已开票', '本年未开票',
        '本年已收合计', '发票收款', '收据收款', '现金/无票收款', '已开票应收', '未开票应收',
        '欠款余额', '账龄(天)', '提醒', '备注', '核对', '客户项目键', '以后年度净额', '最早未清日期']
title(ws, '应 收 对 账 · 每个项目到底还欠多少，看这一张', 'X',
      '★ 这一张全是公式，不用填，只有 W 列备注可以写。行跟《维保合同》一一对应。\n'
      '★ 跨年就靠 H 列：上年及以前「应收 − 收款」的差额自动变成本年期初，所以 26 年的欠款会自己接到 27 年去；'
      '过年只要把《基础资料》D3 的「管理年度」改成 2027，整张表就是 27 年的口径了，流水一行都不用搬。\n'
      '★ 欠款拆两半：R「已开票应收」＝票开了钱没到；S「未开票应收」＝票还没开的那部分。R＋S＝T 欠款余额。\n'
      '★ N 列本年已收再按票据拆成 O 发票 / P 收据 / Q 现金无票 三路，跟实际收款方式对得上。',
      color=C_RPT)
headers(ws, HDR, HD_D, fill_color=C_RPT)
put(ws, f'A{HDR-1}', '合计 →', font=F_TOT, fill=FILL_TOT, align=CR)
put(ws, f'G{HDR-1}', f'=SUM($G${K0}:$G${K1})', font=F_TOT, fill=FILL_TOT, fmt=NUM)
for c in 'HIJKLMNOPQRST':
    put(ws, f'{c}{HDR-1}', f'=ROUND(SUM(${c}${K0}:${c}${K1}),2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'V{HDR-1}', f'=COUNTIF($V${K0}:$V${K1},"*▲*")&" 个项目要盯"', font=F_WARN, fill=FILL_TOT)
put(ws, f'X{HDR-1}', f'=IF(COUNTIF($X${K0}:$X${K1},"★*")=0,"全部平",'
                     f'COUNTIF($X${K0}:$X${K1},"★*")&" 行不平")', font=F_WARN, fill=FILL_TOT)
ws.row_dimensions[HDR - 1].height = 20

for r in range(K0, K1 + 1):
    put(ws, f'Y{r}', f'=IF(维保合同!$C{r}="","",维保合同!$E{r})', font=F_AUTO, fill=FILL_AUTO)
    put(ws, f'A{r}', f'=IF($Y{r}="","",ROW()-{HDR})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'B{r}', f'=IF($Y{r}="","",维保合同!$C{r})', font=F_IN, fill=FILL_AUTO, align=CL_)
    put(ws, f'C{r}', f'=IF($Y{r}="","",维保合同!$D{r})', font=F_IN, fill=FILL_AUTO, align=CL_)
    put(ws, f'D{r}', f'=IF($Y{r}="","",维保合同!$F{r})', font=F_AUTO, fill=FILL_AUTO)
    put(ws, f'E{r}', f'=IF($Y{r}="","",维保合同!$G{r})', font=F_AUTO, fill=FILL_AUTO)
    put(ws, f'F{r}', f'=IF($Y{r}="","",维保合同!$I{r})', font=F_AUTO, fill=FILL_AUTO, fmt=DATEQ)
    put(ws, f'G{r}', f'=IF($Y{r}="","",维保合同!$K{r})', font=F_AUTO, fill=FILL_AUTO, fmt=NUM)
    put(ws, f'H{r}', f'=IF($Y{r}="","",ROUND(SUMIFS(应收金额,应收键,$Y{r},应收年度,"<"&管理年度)'
                     f'-SUMIFS(收款金额,收款键,$Y{r},收款年度,"<"&管理年度),2))',
        font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($Y{r}="","",ROUND(SUMIFS(应收金额,应收键,$Y{r},应收年度,管理年度),2))',
        font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($Y{r}="","",ROUND(SUMIFS(应收金额,应收键,$Y{r},应收年度,管理年度,'
                     f'应收款别,"合同款"),2))', font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'K{r}', f'=IF($Y{r}="","",ROUND($I{r}-$J{r},2))', font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'L{r}', f'=IF($Y{r}="","",ROUND(SUMIFS(应收开票额,应收键,$Y{r},应收开票年度,管理年度),2))',
        font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'M{r}', f'=IF($Y{r}="","",ROUND($I{r}-$L{r},2))', font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'N{r}', f'=IF($Y{r}="","",ROUND(SUMIFS(收款金额,收款键,$Y{r},收款年度,管理年度),2))',
        font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    for c, g in [('O', '发票'), ('P', '收据'), ('Q', '现金无票')]:
        put(ws, f'{c}{r}', f'=IF($Y{r}="","",ROUND(SUMIFS(收款金额,收款键,$Y{r},收款年度,管理年度,'
                           f'收款归类,"{g}"),2))', font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'R{r}', f'=IF($Y{r}="","",ROUND(MAX(0,MIN($T{r},SUMIF(应收键,$Y{r},应收开票额)'
                     f'-SUMIF(收款键,$Y{r},收款金额))),2))', font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'S{r}', f'=IF($Y{r}="","",ROUND(MAX(0,$T{r})-$R{r},2))', font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'T{r}', f'=IF($Y{r}="","",ROUND(SUMIF(应收键,$Y{r},应收金额)-SUMIF(收款键,$Y{r},收款金额),2))',
        font=F_TOT, fill=FILL_KPI, fmt=MONEY)
    # 「最早一笔还没收清的应收是哪天」：AGGREGATE 有的表格软件不认，用 SUMPRODUCT+MIN+IF 这个到处都能跑的写法
    _mn = f'SUMPRODUCT(MIN(IF((应收键=$Y{r})*(应收未收>0.01),应收日期,999999)))'
    put(ws, f'AA{r}', f'=IF($Y{r}="","",IF({_mn}>=999999,"",{_mn}))',
        font=F_AUTO, fill=FILL_AUTO, fmt=DATEQ)
    put(ws, f'U{r}', f'=IF($AA{r}="","",IF($T{r}<=0.01,"",MAX(0,TODAY()-$AA{r})))',
        font=F_AUTO, fill=FILL_AUTO, fmt=DAYS)
    put(ws, f'V{r}', f'=IF($Y{r}="","",IF($T{r}>0.01,'
                     f'IF(AND(ISNUMBER($U{r}),$U{r}>=账龄预警天),"▲欠款拖了 "&TEXT($U{r},"0")&" 天","欠款未清"),'
                     f'IF($T{r}<-0.01,"预收/多收","已结清"))'
                     f'&IF(OR($E{r}="已签未回",$E{r}="未签",$E{r}="续签中")," ▲合同未回","")'
                     f'&IF(AND(ISNUMBER($F{r}),$F{r}<TODAY())," ▲合同已到期",""))',
        font=F_WARN, fill=FILL_AUTO, align=CL_)
    put(ws, f'W{r}', None, font=F_IN, fill=FILL_IN, align=CL_)
    put(ws, f'X{r}', f'=IF($Y{r}="","",'
                     f'IF(ROUND($H{r}+$I{r}-$N{r}+$Z{r}-$T{r},2)<>0,"★期初＋本年－已收≠余额",'
                     f'IF(ROUND($R{r}+$S{r}-MAX(0,$T{r}),2)<>0,"★已开票＋未开票≠欠款",'
                     f'IF(ROUND($J{r}+$K{r}-$I{r},2)<>0,"★合同款＋合同外款≠本年应收",'
                     f'IF(ROUND($O{r}+$P{r}+$Q{r}-$N{r},2)<>0,"★发票＋收据＋现金≠本年已收","OK")))))',
        font=F_AUTO, fill=FILL_AUTO)
    put(ws, f'Z{r}', f'=IF($Y{r}="","",ROUND(SUMIFS(应收金额,应收键,$Y{r},应收年度,">"&管理年度)'
                     f'-SUMIFS(收款金额,收款键,$Y{r},收款年度,">"&管理年度),2))',
        font=F_AUTO, fill=FILL_AUTO, fmt=MONEY)

widths(ws, {'A': 5, 'B': 30, 'C': 18, 'D': 13, 'E': 11, 'F': 12, 'G': 8, 'H': 13, 'I': 13, 'J': 13,
            'K': 13, 'L': 13, 'M': 13, 'N': 13, 'O': 12, 'P': 12, 'Q': 14, 'R': 13, 'S': 13,
            'T': 14, 'U': 10, 'V': 26, 'W': 22, 'X': 24, 'Y': 34, 'Z': 13, 'AA': 14})
for c in ['Y', 'Z', 'AA']:
    ws.column_dimensions[c].hidden = True
ws.freeze_panes = 'D5'
ws.auto_filter.ref = f'A{HDR}:X{K1}'
page(ws, titles=f'{HDR}:{HDR}')

name('对账键', f'应收对账!$Y${K0}:$Y${K1}')
name('对账欠款', f'应收对账!$T${K0}:$T${K1}')
print('  ✓ 应收对账')

# ══════════════════════════════════════════════════════════════
# ⑦ 分期收款计划（大合同按季 / 半月 / 半年付款，每期该回多少、回了没有）
# ══════════════════════════════════════════════════════════════
ws = sheet('分期收款计划')
title(ws, '分 期 收 款 计 划 · 按季 / 半月 / 半年付款的合同，每期该回多少、回了没有', 'V',
      '★ 全自动，不用填。排期是按《维保合同》的「合同起止 ＋ 收费周期 ＋ 合同年费」算出来的：\n'
      '　　合同总额 ＝ 年费 × 合同月数 ÷ 12；每期金额 ＝ 合同总额 ÷ 期数（尾差进最后一期）。\n'
      '★ 一年几期看《基础资料》的「每年期数」：一次性 1 期、半年一次 2 期、季度一次 4 期、月度 12 期、半月一次 24 期、按次不排期。\n'
      '★ 「应收款日」＝ 每期期初 ＋《基础资料》D5 的账期天数。过了这个日子还没收够，就算逾期。\n'
      '★ 「累计已回款」用的是核销口径（该期挂的合同款应收实际收到多少），不是简单按收款日期切，'
      '所以客户一笔钱把几期一起付了也算得准。\n'
      '★ 这里的「回款缺口」只算合同款里已经到期的那几期，跟《应收对账》的「欠款余额」不是一回事 ——'
      '后者还包含合同外的维修费配件费，以及还没到期但已经挂账开票的钱。两个数各看各的，不用对平。',
      color=C_DASH)

# ── 块① 按合同看回款进度 ──
HD_T = ['序号', '使用单位', '项目名称', '合同编号', '合同状态', '收费周期', '合同起', '合同止',
        '合同年费', '合同总额', '分期期数', '每期金额', '已到期期数', '累计计划回款',
        '累计已挂应收', '累计已回款', '回款缺口', '回款进度', '下一期应收款日', '下一期金额',
        '回款状态', '核对']
headers(ws, HDR, HD_T, fill_color=C_DASH, height=30)
for i in range(N_CON):
    r, cr = T0 + i, K0 + i
    ds, de = X0 + i * MAX_TERM, X0 + i * MAX_TERM + MAX_TERM - 1
    put(ws, f'A{r}', f'=IF(维保合同!$C{cr}="","",ROW()-{HDR})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    for c, sc in [('B', 'C'), ('C', 'D'), ('E', 'G'), ('F', 'M')]:
        put(ws, f'{c}{r}', f'=IF($A{r}="","",维保合同!${sc}{cr})', font=F_IN, fill=FILL_AUTO,
            align=CL_ if c in ('B', 'C') else C)
    put(ws, f'D{r}', f'=IF($A{r}="","",维保合同!$F{cr}&"")', font=F_IN, fill=FILL_AUTO)
    put(ws, f'G{r}', f'=IF($A{r}="","",维保合同!$H{cr})', font=F_IN, fill=FILL_AUTO, fmt=DATEQ)
    put(ws, f'H{r}', f'=IF($A{r}="","",维保合同!$I{cr})', font=F_IN, fill=FILL_AUTO, fmt=DATEQ)
    put(ws, f'I{r}', f'=IF($A{r}="","",维保合同!$L{cr})', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",维保合同!$AA{cr})', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'K{r}', f'=IF($A{r}="","",维保合同!$AB{cr})', font=F_IN, fill=FILL_AUTO, fmt=NUM)
    put(ws, f'L{r}', f'=IF($A{r}="","",维保合同!$AC{cr})', font=F_TOT, fill=FILL_KPI, fmt=MONEY)
    put(ws, f'M{r}', f'=IF($A{r}="","",COUNTIF($J${ds}:$J${de},1))', font=F_IN, fill=FILL_AUTO, fmt=NUM)
    put(ws, f'N{r}', f'=IF($A{r}="","",ROUND(SUMIF($J${ds}:$J${de},1,$G${ds}:$G${de}),2))',
        font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'O{r}', f'=IF($A{r}="","",ROUND(SUMIF($J${ds}:$J${de},1,$H${ds}:$H${de}),2))',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'P{r}', f'=IF($A{r}="","",ROUND(SUMIF($J${ds}:$J${de},1,$I${ds}:$I${de}),2))',
        font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'Q{r}', f'=IF($A{r}="","",ROUND($N{r}-$P{r},2))', font=F_WARN, fill=FILL_KPI, fmt=MONEY)
    put(ws, f'R{r}', f'=IF(OR($A{r}="",N($N{r})=0),"",$P{r}/$N{r})', font=F_IN, fill=FILL_AUTO, fmt=PCT)
    put(ws, f'S{r}', f'=IF($A{r}="","",IFERROR(INDEX($F${ds}:$F${de},MATCH(0,$J${ds}:$J${de},0)),""))',
        font=F_IN, fill=FILL_AUTO, fmt=DATEQ)
    put(ws, f'T{r}', f'=IF($A{r}="","",IFERROR(INDEX($G${ds}:$G${de},MATCH(0,$J${ds}:$J${de},0)),""))',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'U{r}', f'=IF($A{r}="","",IF(N($K{r})=0,"按次结算，不排期",'
                     f'IF($Q{r}<=0.01,"✔ 按期收清",'
                     f'"▲欠 "&TEXT($Q{r},"#,##0.00")&"（约 "&TEXT(ROUND($Q{r}/MAX($L{r},0.01),1),"0.0")&" 期）")))',
        font=F_WARN, fill=FILL_WARN, align=CL_)
    put(ws, f'V{r}', f'=IF($A{r}="","",IF(N($K{r})=0,"不排期",'
                     f'IF(ROUND(SUM($G${ds}:$G${de})-$J{r},2)<>0,"★各期合计≠合同总额","OK")))',
        font=F_AUTO, fill=FILL_AUTO)
put(ws, f'A{HDR-1}', '① 按合同看回款进度（一个合同一行，跟《维保合同》行对行）　合计 →',
    font=F_SEC, fill=FILL_TOT, align=CL_)
for _c in 'BCDEFGH':
    put(ws, f'{_c}{HDR-1}', None, font=F_SEC, fill=FILL_TOT)
ws.merge_cells(f'A{HDR-1}:H{HDR-1}')
put(ws, f'A{HDR-1}', '① 按合同看回款进度（一个合同一行，跟《维保合同》行对行）　合计 →',
    font=F_SEC, fill=FILL_TOT, align=CL_)
for c in 'IJNOPQ':
    put(ws, f'{c}{HDR-1}', f'=ROUND(SUM(${c}${T0}:${c}${T1}),2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'R{HDR-1}', f'=IFERROR($P${HDR-1}/$N${HDR-1},"")', font=F_TOT, fill=FILL_TOT, fmt=PCT)
put(ws, f'U{HDR-1}', f'=COUNTIF($U${T0}:$U${T1},"▲*")&" 个合同没按期收齐"', font=F_WARN, fill=FILL_TOT)
put(ws, f'V{HDR-1}', f'=IF(COUNTIF($V${T0}:$V${T1},"★*")=0,"全部 OK",'
                     f'COUNTIF($V${T0}:$V${T1},"★*")&" 行要查")', font=F_WARN, fill=FILL_TOT)
ws.row_dimensions[HDR - 1].height = 20

# ── 块② 分期明细排期 ──
SEC2, HD2 = X0 - 2, X0 - 1
put(ws, f'A{SEC2}', f' ② 分期明细排期（一个合同最多排 {MAX_TERM} 期；用表头的筛选按「本期状态」挑出逾期的期）',
    font=F_SEC, fill=FILL_SEC, align=CL_)
ws.merge_cells(f'A{SEC2}:O{SEC2}')
ws.row_dimensions[SEC2].height = 22
headers(ws, HD2, ['合同序号', '客户项目', '期次', '本期区间起', '本期区间止', '应收款日',
                  '本期计划金额', '本期已挂应收', '本期已回款', '已到期', '累计计划',
                  '累计已挂', '累计已回款', '本期状态', '逾期天数'], fill_color=C_DASH, height=30)
for i in range(N_CON):
    cr = K0 + i
    ds = X0 + i * MAX_TERM
    ED = f'EDATE(维保合同!$H{cr},'
    for j in range(MAX_TERM):
        dr, n = ds + j, j + 1
        hf = f'INT(({n}-1)/2)'
        put(ws, f'A{dr}', f'=IF($B{dr}="","",{i+1})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
        put(ws, f'B{dr}', f'=IF(OR(维保合同!$C{cr}="",N(维保合同!$AB{cr})<{n}),"",维保合同!$E{cr})',
            font=F_AUTO, fill=FILL_AUTO, align=CL_)
        put(ws, f'C{dr}', f'=IF($B{dr}="","",{n})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
        put(ws, f'D{dr}', f'=IF($B{dr}="","",IF(维保合同!$Y{cr}={MAX_TERM},'
                          f'{ED}{hf})+MOD({n}-1,2)*15,{ED}({n}-1)*12/维保合同!$Y{cr})))',
            font=F_IN, fill=FILL_AUTO, fmt=DATEQ)
        put(ws, f'E{dr}', f'=IF($B{dr}="","",MIN(维保合同!$I{cr},IF(维保合同!$Y{cr}={MAX_TERM},'
                          f'IF(MOD({n}-1,2)=0,{ED}{hf})+14,{ED}{hf}+1)-1),'
                          f'{ED}{n}*12/维保合同!$Y{cr})-1)))', font=F_IN, fill=FILL_AUTO, fmt=DATEQ)
        put(ws, f'F{dr}', f'=IF($B{dr}="","",$D{dr}+N(账期天数))', font=F_TOT, fill=FILL_AUTO, fmt=DATEQ)
        put(ws, f'G{dr}', f'=IF($B{dr}="","",IF({n}=维保合同!$AB{cr},'
                          f'ROUND(维保合同!$AA{cr}-维保合同!$AC{cr}*(维保合同!$AB{cr}-1),2),'
                          f'维保合同!$AC{cr}))', font=F_TOT, fill=FILL_KPI, fmt=MONEY)
        put(ws, f'H{dr}', f'=IF($B{dr}="","",ROUND(SUMIFS(应收金额,应收键,$B{dr},应收款别,"合同款",'
                          f'应收归期日,">="&$D{dr},应收归期日,"<="&$E{dr}),2))',
            font=F_IN, fill=FILL_AUTO, fmt=MONEY)
        put(ws, f'I{dr}', f'=IF($B{dr}="","",ROUND(SUMIFS(应收已收,应收键,$B{dr},应收款别,"合同款",'
                          f'应收归期日,">="&$D{dr},应收归期日,"<="&$E{dr}),2))',
            font=F_IN, fill=FILL_AUTO, fmt=MONEY)
        put(ws, f'J{dr}', f'=IF($B{dr}="","",IF($F{dr}<=TODAY(),1,0))', font=F_AUTO, fill=FILL_AUTO, fmt='0')
        put(ws, f'K{dr}', f'=IF($B{dr}="","",ROUND(SUM($G${ds}:$G{dr}),2))', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
        put(ws, f'L{dr}', f'=IF($B{dr}="","",ROUND(SUM($H${ds}:$H{dr}),2))', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
        put(ws, f'M{dr}', f'=IF($B{dr}="","",ROUND(SUM($I${ds}:$I{dr}),2))', font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
        put(ws, f'N{dr}', f'=IF($B{dr}="","",IF($J{dr}=0,"未到期",'
                          f'IF($M{dr}>=$K{dr}-0.01,"✔ 已收清",'
                          f'IF($L{dr}<$K{dr}-0.01,"▲还没挂应收 "&TEXT($K{dr}-$L{dr},"#,##0.00"),'
                          f'"▲欠 "&TEXT($K{dr}-$M{dr},"#,##0.00")))))',
            font=F_WARN, fill=FILL_WARN, align=CL_)
        put(ws, f'O{dr}', f'=IF($B{dr}="","",IF(OR($J{dr}=0,$M{dr}>=$K{dr}-0.01),"",TODAY()-$F{dr}))',
            font=F_WARN, fill=FILL_AUTO, fmt=DAYS)

# 两块共用同一批列，列宽要同时照顾「块①的合同信息」和「块②的期次信息」
widths(ws, {'A': 8, 'B': 34, 'C': 18, 'D': 13, 'E': 12, 'F': 13, 'G': 13, 'H': 13, 'I': 13,
            'J': 13, 'K': 13, 'L': 13, 'M': 13, 'N': 24, 'O': 13, 'P': 13, 'Q': 13, 'R': 11,
            'S': 14, 'T': 13, 'U': 26, 'V': 20})
ws.freeze_panes = 'C5'
ws.auto_filter.ref = f'A{HD2}:O{X1}'
page(ws, titles=f'{HDR}:{HDR}')
PLAN_TOT, PLAN_D0, PLAN_D1 = HDR - 1, X0, X1
print('  ✓ 分期收款计划')

# ══════════════════════════════════════════════════════════════
# ⑦ _自动清单（隐藏）—— 提醒表的「压缩」辅助
# ══════════════════════════════════════════════════════════════
ws = sheet('_自动清单', hidden=True)
put(ws, 'A1', '这是自动生成的辅助表，专门给《到期与提醒》把散在各表里的提醒行压成连续清单用的，不用管它，也别删。',
    font=F_NOTE, align=CL_, border=None)
for col, lab in [('A', '合同提醒标记'), ('B', '累计'), ('D', '年检校验标记'), ('E', '累计'),
                 ('G', '免保标记'), ('H', '累计'), ('J', '欠款标记'), ('K', '累计')]:
    put(ws, f'{col}{HDR}', lab, font=F_HDR2, fill=FILL_SEC)
for r in range(K0, K1 + 1):
    put(ws, f'A{r}', f'=IF(维保合同!$C{r}="",0,IF(OR(维保合同!$Q{r}<>"",维保合同!$R{r}<>""),1,0))', fmt='0')
    put(ws, f'B{r}', f'=SUM($A${K0}:$A{r})', fmt='0')
    put(ws, f'J{r}', f'=IF(应收对账!$Y{r}="",0,IF(应收对账!$T{r}>0.01,1,0))', fmt='0')
    put(ws, f'K{r}', f'=SUM($J${K0}:$J{r})', fmt='0')
for r in range(L0, L1 + 1):
    put(ws, f'D{r}', f'=IF(电梯资料!$C{r}="",0,IF(OR(电梯资料!$AA{r}<>"",电梯资料!$AB{r}<>""),1,0))', fmt='0')
    put(ws, f'E{r}', f'=SUM($D${L0}:$D{r})', fmt='0')
    put(ws, f'G{r}', f'=IF(电梯资料!$C{r}="",0,IF(OR(电梯资料!$Y{r}="安装免保",电梯资料!$Y{r}="技术免保",'
                     f'电梯资料!$Y{r}="质保期内"),1,0))', fmt='0')
    put(ws, f'H{r}', f'=SUM($G${L0}:$G{r})', fmt='0')
widths(ws, {c: 14 for c in 'ABDEGHJK'})
print('  ✓ _自动清单')

# ══════════════════════════════════════════════════════════════
# ⑧ 到期与提醒
# ══════════════════════════════════════════════════════════════
CUM_C = f'_自动清单!$B${K0}:$B${K1}'      # 合同
CUM_J = f'_自动清单!$E${L0}:$E${L1}'      # 年检/校验
CUM_M = f'_自动清单!$H${L0}:$H${L1}'      # 免保
CUM_Q = f'_自动清单!$K${K0}:$K${K1}'      # 欠款
TOT_C, TOT_J, TOT_M, TOT_Q = f'_自动清单!$B${K1}', f'_自动清单!$E${L1}', f'_自动清单!$H${L1}', f'_自动清单!$K${K1}'


def pick(src, col, k, cum, r0, r1):
    """文本列：空单元格被 INDEX 取出来会变成 0，补个 &"" 让它还是空的"""
    return f'=IFERROR(INDEX({src}!${col}${r0}:${col}${r1},MATCH({k},{cum},0))&"","")'


def pickn(src, col, k, cum, r0, r1):
    """日期 / 金额列：保持数值，空的靠单元格格式藏掉"""
    return f'=IFERROR(INDEX({src}!${col}${r0}:${col}${r1},MATCH({k},{cum},0)),"")'


N_B1, N_B2, N_B3, N_B4 = 60, 80, 50, 60
ws = sheet('到期与提醒')
title(ws, '到 期 与 提 醒 · 该催的合同、该约的年检、该转收费的免保、该要的钱', 'L',
      '★ 整张表都是自动生成的，不用填。四块分别是：合同到期/没回签、电梯年检与校验到期、免保快到期该转收费、欠款账龄预警。\n'
      '★ 「剩余天数」是负的（红字）就是已经过期了。提前多少天开始提醒，在《基础资料》第 3~5 行的参数里改。\n'
      '★ 每块最多显示的条数写在小标题上，超了会提示；真要全量看，回各自的明细表用筛选。',
      color=C_WARN)

r_ = 4
put(ws, f'A{r_}', f'=" ① 合同到期 / 合同没回签 —— 共 "&{TOT_C}&" 条（这里最多列 {N_B1} 条）"',
    font=F_SEC, fill=FILL_SEC, align=CL_)
ws.merge_cells(f'A{r_}:L{r_}')
H1 = r_ + 1
headers(ws, H1, ['序号', '片区', '使用单位', '项目名称', '合同编号', '合同状态', '合同起', '合同止',
                 '剩余天数', '提醒', '合同年费', '备注'], fill_color=C_WARN, height=26)
B1 = H1 + 1
for i in range(N_B1):
    r, k = B1 + i, i + 1
    put(ws, f'A{r}', f'=IF({k}>{TOT_C},"",{k})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    for col, sc in [('B', 'B'), ('C', 'C'), ('D', 'D'), ('E', 'F'), ('F', 'G')]:
        put(ws, f'{col}{r}', pick('维保合同', sc, k, CUM_C, K0, K1), font=F_IN, fill=FILL_AUTO,
            align=CL_ if col in ('C', 'D') else C)
    put(ws, f'G{r}', pickn('维保合同', 'H', k, CUM_C, K0, K1), font=F_IN, fill=FILL_AUTO, fmt=DATEQ)
    put(ws, f'H{r}', pickn('维保合同', 'I', k, CUM_C, K0, K1), font=F_IN, fill=FILL_AUTO, fmt=DATEQ)
    put(ws, f'I{r}', f'=IF(N($H{r})=0,"",$H{r}-TODAY())', font=F_WARN, fill=FILL_AUTO, fmt=DAYS)
    put(ws, f'J{r}', f'=IF($A{r}="","",TRIM(IFERROR(INDEX(维保合同!$Q${K0}:$Q${K1},MATCH({k},{CUM_C},0))&"","")'
                     f'&" "&IFERROR(INDEX(维保合同!$R${K0}:$R${K1},MATCH({k},{CUM_C},0))&"","")))',
        font=F_WARN, fill=FILL_WARN, align=CL_)
    put(ws, f'K{r}', pickn('维保合同', 'L', k, CUM_C, K0, K1), font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'L{r}', pick('维保合同', 'W', k, CUM_C, K0, K1), font=F_AUTO, fill=FILL_AUTO, align=CL_)

r_ = B1 + N_B1 + 1
put(ws, f'A{r_}', f'=" ② 电梯年检 / 限速器校验 / 载重试验 到期 —— 共 "&{TOT_J}&" 台（这里最多列 {N_B2} 台）"',
    font=F_SEC, fill=FILL_SEC, align=CL_)
ws.merge_cells(f'A{r_}:L{r_}')
H2 = r_ + 1
headers(ws, H2, ['序号', '使用单位', '项目名称', '设备编号', '使用编号', '维保人员', '下次年检',
                 '限速器校验', '载重试验125%', '最早到期', '剩余天数', '提醒'], fill_color=C_WARN, height=26)
B2 = H2 + 1
for i in range(N_B2):
    r, k = B2 + i, i + 1
    put(ws, f'A{r}', f'=IF({k}>{TOT_J},"",{k})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    for col, sc in [('B', 'C'), ('C', 'D'), ('D', 'Q'), ('E', 'R'), ('F', 'W')]:
        put(ws, f'{col}{r}', pick('电梯资料', sc, k, CUM_J, L0, L1), font=F_IN, fill=FILL_AUTO,
            align=CL_ if col in ('B', 'C') else C)
    for col, sc in [('G', 'T'), ('H', 'U'), ('I', 'V')]:
        put(ws, f'{col}{r}', pickn('电梯资料', sc, k, CUM_J, L0, L1), font=F_IN, fill=FILL_AUTO, fmt=DATEQ)
    put(ws, f'J{r}', f'=IF(COUNT($G{r}:$I{r})=0,"",MIN($G{r}:$I{r}))', font=F_TOT, fill=FILL_AUTO, fmt=DATEQ)
    put(ws, f'K{r}', f'=IF(N($J{r})=0,"",$J{r}-TODAY())', font=F_WARN, fill=FILL_AUTO, fmt=DAYS)
    _a = f'IFERROR(INDEX(电梯资料!$AA${L0}:$AA${L1},MATCH({k},{CUM_J},0))&"","")'
    _b = f'IFERROR(INDEX(电梯资料!$AB${L0}:$AB${L1},MATCH({k},{CUM_J},0))&"","")'
    put(ws, f'L{r}', f'=IF($A{r}="","",TRIM(IF({_a}<>"","年检 "&{_a},"")'
                     f'&IF(AND({_a}<>"",{_b}<>"")," ｜ ","")'
                     f'&IF({_b}<>"","校验 "&{_b},"")))', font=F_WARN, fill=FILL_WARN, align=CL_)

r_ = B2 + N_B2 + 1
put(ws, f'A{r_}', f'=" ③ 免保（安装免保 / 技术免保 / 质保期内）到期该转收费 —— 共 "&{TOT_M}&" 台（这里最多列 {N_B3} 台）"',
    font=F_SEC, fill=FILL_SEC, align=CL_)
ws.merge_cells(f'A{r_}:L{r_}')
H3 = r_ + 1
headers(ws, H3, ['序号', '使用单位', '项目名称', '设备编号', '使用编号', '维保状态', '免保到期日',
                 '剩余天数', '提醒', '维保人员', '注意事项', ''], fill_color=C_WARN, height=26)
B3 = H3 + 1
for i in range(N_B3):
    r, k = B3 + i, i + 1
    put(ws, f'A{r}', f'=IF({k}>{TOT_M},"",{k})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    for col, sc in [('B', 'C'), ('C', 'D'), ('D', 'Q'), ('E', 'R'), ('F', 'Y')]:
        put(ws, f'{col}{r}', pick('电梯资料', sc, k, CUM_M, L0, L1), font=F_IN, fill=FILL_AUTO,
            align=CL_ if col in ('B', 'C') else C)
    put(ws, f'G{r}', pickn('电梯资料', 'Z', k, CUM_M, L0, L1), font=F_IN, fill=FILL_AUTO, fmt=DATEQ)
    put(ws, f'H{r}', f'=IF(N($G{r})=0,"",$G{r}-TODAY())', font=F_WARN, fill=FILL_AUTO, fmt=DAYS)
    put(ws, f'I{r}', pick('电梯资料', 'AC', k, CUM_M, L0, L1), font=F_WARN, fill=FILL_WARN, align=CL_)
    put(ws, f'J{r}', pick('电梯资料', 'W', k, CUM_M, L0, L1), font=F_IN, fill=FILL_AUTO)
    put(ws, f'K{r}', pick('电梯资料', 'AD', k, CUM_M, L0, L1), font=F_AUTO, fill=FILL_AUTO, align=CL_)
    put(ws, f'L{r}', None, font=F_AUTO, fill=FILL_AUTO)

r_ = B3 + N_B3 + 1
put(ws, f'A{r_}', f'=" ④ 欠款账龄预警 —— 还有欠款的项目共 "&{TOT_Q}&" 个（这里最多列 {N_B4} 个）"',
    font=F_SEC, fill=FILL_SEC, align=CL_)
ws.merge_cells(f'A{r_}:L{r_}')
H4 = r_ + 1
headers(ws, H4, ['序号', '使用单位', '项目名称', '合同状态', '合同到期', '欠款余额', '已开票应收',
                 '未开票应收', '账龄(天)', '提醒', '备注', ''], fill_color=C_WARN, height=26)
B4 = H4 + 1
for i in range(N_B4):
    r, k = B4 + i, i + 1
    put(ws, f'A{r}', f'=IF({k}>{TOT_Q},"",{k})', font=F_AUTO, fill=FILL_AUTO, fmt='0')
    for col, sc in [('B', 'B'), ('C', 'C'), ('D', 'E')]:
        put(ws, f'{col}{r}', pick('应收对账', sc, k, CUM_Q, K0, K1), font=F_IN, fill=FILL_AUTO,
            align=CL_ if col in ('B', 'C') else C)
    put(ws, f'E{r}', pickn('应收对账', 'F', k, CUM_Q, K0, K1), font=F_IN, fill=FILL_AUTO, fmt=DATEQ)
    for col, sc in [('F', 'T'), ('G', 'R'), ('H', 'S')]:
        put(ws, f'{col}{r}', pickn('应收对账', sc, k, CUM_Q, K0, K1),
            font=F_TOT if col == 'F' else F_IN, fill=FILL_KPI if col == 'F' else FILL_AUTO, fmt=MONEY)
    put(ws, f'I{r}', pickn('应收对账', 'U', k, CUM_Q, K0, K1), font=F_WARN, fill=FILL_AUTO, fmt=DAYS)
    put(ws, f'J{r}', pick('应收对账', 'V', k, CUM_Q, K0, K1), font=F_WARN, fill=FILL_WARN, align=CL_)
    put(ws, f'K{r}', pick('应收对账', 'W', k, CUM_Q, K0, K1), font=F_AUTO, fill=FILL_AUTO, align=CL_)
    put(ws, f'L{r}', None, font=F_AUTO, fill=FILL_AUTO)
put(ws, f'E{B4+N_B4}', '本块合计 →', font=F_TOT, fill=FILL_TOT, align=CR)
for col in ['F', 'G', 'H']:
    put(ws, f'{col}{B4+N_B4}', f'=ROUND(SUM(${col}${B4}:${col}${B4+N_B4-1}),2)',
        font=F_TOT, fill=FILL_TOT, fmt=MONEY)

widths(ws, {'A': 5, 'B': 30, 'C': 18, 'D': 16, 'E': 14, 'F': 13, 'G': 13, 'H': 13, 'I': 13,
            'J': 28, 'K': 22, 'L': 18})
ws.freeze_panes = 'A6'
page(ws)
print('  ✓ 到期与提醒')

name('应收开票日', f'应收登记!$L${A0}:$L${A1}')

# ══════════════════════════════════════════════════════════════
# ⑨ 月度汇总
# ══════════════════════════════════════════════════════════════
ws = sheet('月度汇总')
title(ws, '月 度 汇 总 · 本年度的应收、开票、收款，按月 / 类别 / 片区 / 票据 / 方式 五个口径', 'K',
      '★ 「本年度」＝《基础资料》D3 的管理年度。改那一格，整张表跟着换年。\n'
      '★ 全是公式，不用填。①按月 ②按费用类别 ③按片区 ④按票据类型 ⑤按收款方式，'
      '②③④⑤ 的行是跟着《基础资料》的清单走的，那边加一行这里就多一行。\n'
      '★ ① 最后一列「累计欠款余额」＝上年结转 ＋ 当年累计应收 － 当年累计收款，跟《应收对账》的欠款合计对得上。',
      color=C_RPT)
MSEC = lambda rr, txt: (put(ws, f'A{rr}', txt, font=F_SEC, fill=FILL_SEC, align=CL_),
                        ws.merge_cells(f'A{rr}:K{rr}'))

r_ = 4
MSEC(r_, '① 按月：应收 / 开票 / 收款')
HM = r_ + 1
headers(ws, HM, ['月份', '应收合计', '其中：合同款', '其中：合同外款', '开票金额', '收款合计',
                 '发票收款', '收据收款', '现金/无票收款', '当月应收－收款', '累计欠款余额'],
        fill_color=C_RPT, height=28)
M0 = HM + 1
for m in range(1, 13):
    r = M0 + m - 1
    d0 = f'DATE(管理年度,{m},1)'
    d1 = f'DATE(管理年度,{m+1},1)'
    put(ws, f'A{r}', f'=管理年度&"年{m:02d}月"', font=F_TOT, fill=FILL_AUTO)
    put(ws, f'B{r}', f'=ROUND(SUMIFS(应收金额,应收日期,">="&{d0},应收日期,"<"&{d1}),2)',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'C{r}', f'=ROUND(SUMIFS(应收金额,应收日期,">="&{d0},应收日期,"<"&{d1},应收款别,"合同款"),2)',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'D{r}', f'=ROUND($B{r}-$C{r},2)', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'E{r}', f'=ROUND(SUMIFS(应收开票额,应收开票日,">="&{d0},应收开票日,"<"&{d1}),2)',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'F{r}', f'=ROUND(SUMIFS(收款金额,收款日期,">="&{d0},收款日期,"<"&{d1}),2)',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    for col, g in [('G', '发票'), ('H', '收据'), ('I', '现金无票')]:
        put(ws, f'{col}{r}', f'=ROUND(SUMIFS(收款金额,收款日期,">="&{d0},收款日期,"<"&{d1},'
                             f'收款归类,"{g}"),2)', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'J{r}', f'=ROUND($B{r}-$F{r},2)', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'K{r}', f'=ROUND(应收对账!$H${HDR-1}+SUM($B${M0}:$B{r})-SUM($F${M0}:$F{r}),2)',
        font=F_TOT, fill=FILL_KPI, fmt=MONEY)
MT_ = M0 + 12
put(ws, f'A{MT_}', '全年合计', font=F_TOT, fill=FILL_TOT)
for col in 'BCDEFGHIJ':
    put(ws, f'{col}{MT_}', f'=ROUND(SUM(${col}${M0}:${col}${M0+11}),2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'K{MT_}', f'=$K{M0+11}', font=F_TOT, fill=FILL_TOT, fmt=MONEY)

r_ = MT_ + 2
MSEC(r_, '② 按费用类别：合同款 / 合同外款 各多少，开了多少票')
HF = r_ + 1
headers(ws, HF, ['费用类别', '款别', '本年应收', '本年已开票', '本年未开票', '占比', '', '', '', '', ''],
        fill_color=C_RPT, height=26)
F0 = HF + 1
for i in range(N_LIST):
    r, s = F0 + i, B0 + i
    put(ws, f'A{r}', f'=IF(基础资料!$M{s}="","",基础资料!$M{s})', font=F_IN, fill=FILL_AUTO, align=CL_)
    put(ws, f'B{r}', f'=IF($A{r}="","",基础资料!$N{s})', font=F_IN, fill=FILL_AUTO)
    put(ws, f'C{r}', f'=IF($A{r}="","",ROUND(SUMIFS(应收金额,应收类别,$A{r},应收年度,管理年度),2))',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'D{r}', f'=IF($A{r}="","",ROUND(SUMIFS(应收开票额,应收类别,$A{r},应收开票年度,管理年度),2))',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'E{r}', f'=IF($A{r}="","",ROUND($C{r}-$D{r},2))', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",IFERROR($C{r}/$C${F0+N_LIST+1},""))', font=F_IN, fill=FILL_AUTO, fmt=PCT)
FX = F0 + N_LIST
put(ws, f'A{FX}', '不在清单里的（没选下拉直接手打的）', font=F_WARN, fill=FILL_WARN, align=CL_)
ws.merge_cells(f'A{FX}:B{FX}')
put(ws, f'C{FX}', f'=ROUND(应收对账!$I${HDR-1}-SUM($C${F0}:$C${F0+N_LIST-1}),2)',
    font=F_WARN, fill=FILL_WARN, fmt=MONEY)
put(ws, f'D{FX}', f'=ROUND(应收对账!$L${HDR-1}-SUM($D${F0}:$D${F0+N_LIST-1}),2)',
    font=F_WARN, fill=FILL_WARN, fmt=MONEY)
put(ws, f'E{FX}', f'=ROUND($C{FX}-$D{FX},2)', font=F_WARN, fill=FILL_WARN, fmt=MONEY)
put(ws, f'F{FX}', None, font=F_WARN, fill=FILL_WARN)
FT = FX + 1
put(ws, f'A{FT}', '合计', font=F_TOT, fill=FILL_TOT)
put(ws, f'B{FT}', None, font=F_TOT, fill=FILL_TOT)
for col in 'CDE':
    put(ws, f'{col}{FT}', f'=ROUND(SUM(${col}${F0}:${col}${FX}),2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'F{FT}', f'=IF($C{FT}=0,"",1)', font=F_TOT, fill=FILL_TOT, fmt=PCT)

r_ = FT + 2
MSEC(r_, '③ 按片区：哪个片区台量多、收得回来多少')
HA = r_ + 1
headers(ws, HA, ['片区', '项目数', '在册台量', '合同年费合计', '本年应收', '本年已收', '欠款余额',
                 '', '', '', ''], fill_color=C_RPT, height=26)
P0 = HA + 1
for i in range(N_LIST):
    r, s = P0 + i, B0 + i
    put(ws, f'A{r}', f'=IF(基础资料!$A{s}="","",基础资料!$A{s})', font=F_IN, fill=FILL_AUTO, align=CL_)
    put(ws, f'B{r}', f'=IF($A{r}="","",COUNTIF(合同片区,$A{r}))', font=F_IN, fill=FILL_AUTO, fmt=NUM)
    put(ws, f'C{r}', f'=IF($A{r}="","",SUMIF(合同片区,$A{r},合同台量))', font=F_IN, fill=FILL_AUTO, fmt=NUM)
    put(ws, f'D{r}', f'=IF($A{r}="","",ROUND(SUMIF(合同片区,$A{r},维保合同!$L${K0}:$L${K1}),2))',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'E{r}', f'=IF($A{r}="","",ROUND(SUMIF(合同片区,$A{r},维保合同!$T${K0}:$T${K1}),2))',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",ROUND(SUMIF(合同片区,$A{r},维保合同!$U${K0}:$U${K1}),2))',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'G{r}', f'=IF($A{r}="","",ROUND(SUMIF(合同片区,$A{r},维保合同!$V${K0}:$V${K1}),2))',
        font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
PX = P0 + N_LIST
put(ws, f'A{PX}', '没分片区的（电梯资料里还没填片区）', font=F_WARN, fill=FILL_WARN, align=CL_)
put(ws, f'B{PX}', f'=COUNTIF(合同键,"?*")-SUM($B${P0}:$B${P0+N_LIST-1})',
    font=F_WARN, fill=FILL_WARN, fmt=NUM)
for col, sc in [('C', 'K'), ('D', 'L'), ('E', 'T'), ('F', 'U'), ('G', 'V')]:
    put(ws, f'{col}{PX}', f'=ROUND(维保合同!${sc}${HDR-1}-SUM(${col}${P0}:${col}${P0+N_LIST-1}),2)',
        font=F_WARN, fill=FILL_WARN, fmt=NUM if col == 'C' else MONEY)
PT = PX + 1
put(ws, f'A{PT}', '合计', font=F_TOT, fill=FILL_TOT)
for col in 'BCDEFG':
    put(ws, f'{col}{PT}', f'=ROUND(SUM(${col}${P0}:${col}${PX}),2)', font=F_TOT, fill=FILL_TOT,
        fmt=NUM if col in 'BC' else MONEY)

r_ = PT + 2
MSEC(r_, '④ 按票据类型收款：专票 / 普票 / 电子票 / 收据 / 无票')
HB = r_ + 1
headers(ws, HB, ['票据类型', '归类', '本年收款', '占比', '', '', '', '', '', '', ''],
        fill_color=C_RPT, height=26)
Q0 = HB + 1
for i in range(N_LIST):
    r, s = Q0 + i, B0 + i
    put(ws, f'A{r}', f'=IF(基础资料!$P{s}="","",基础资料!$P{s})', font=F_IN, fill=FILL_AUTO, align=CL_)
    put(ws, f'B{r}', f'=IF($A{r}="","",基础资料!$Q{s})', font=F_IN, fill=FILL_AUTO)
    put(ws, f'C{r}', f'=IF($A{r}="","",ROUND(SUMIFS(收款金额,收款票据,$A{r},收款年度,管理年度),2))',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'D{r}', f'=IF($A{r}="","",IFERROR($C{r}/$C${Q0+N_LIST+1},""))', font=F_IN, fill=FILL_AUTO, fmt=PCT)
QX = Q0 + N_LIST
put(ws, f'A{QX}', '不在清单里的', font=F_WARN, fill=FILL_WARN, align=CL_)
put(ws, f'B{QX}', None, font=F_WARN, fill=FILL_WARN)
put(ws, f'C{QX}', f'=ROUND(应收对账!$N${HDR-1}-SUM($C${Q0}:$C${Q0+N_LIST-1}),2)',
    font=F_WARN, fill=FILL_WARN, fmt=MONEY)
put(ws, f'D{QX}', None, font=F_WARN, fill=FILL_WARN)
QT = QX + 1
put(ws, f'A{QT}', '合计', font=F_TOT, fill=FILL_TOT)
put(ws, f'B{QT}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'C{QT}', f'=ROUND(SUM($C${Q0}:$C${QX}),2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'D{QT}', f'=IF($C{QT}=0,"",1)', font=F_TOT, fill=FILL_TOT, fmt=PCT)

r_ = QT + 2
MSEC(r_, '⑤ 按收款方式：钱是怎么进来的')
HW = r_ + 1
headers(ws, HW, ['收款方式', '本年收款', '占比', '', '', '', '', '', '', '', ''],
        fill_color=C_RPT, height=26)
W0 = HW + 1
for i in range(N_LIST):
    r, s = W0 + i, B0 + i
    put(ws, f'A{r}', f'=IF(基础资料!$S{s}="","",基础资料!$S{s})', font=F_IN, fill=FILL_AUTO, align=CL_)
    put(ws, f'B{r}', f'=IF($A{r}="","",ROUND(SUMIFS(收款金额,收款方式,$A{r},收款年度,管理年度),2))',
        font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'C{r}', f'=IF($A{r}="","",IFERROR($B{r}/$B${W0+N_LIST+1},""))', font=F_IN, fill=FILL_AUTO, fmt=PCT)
WX = W0 + N_LIST
put(ws, f'A{WX}', '不在清单里的', font=F_WARN, fill=FILL_WARN, align=CL_)
put(ws, f'B{WX}', f'=ROUND(应收对账!$N${HDR-1}-SUM($B${W0}:$B${W0+N_LIST-1}),2)',
    font=F_WARN, fill=FILL_WARN, fmt=MONEY)
put(ws, f'C{WX}', None, font=F_WARN, fill=FILL_WARN)
WT = WX + 1
put(ws, f'A{WT}', '合计', font=F_TOT, fill=FILL_TOT)
put(ws, f'B{WT}', f'=ROUND(SUM($B${W0}:$B${WX}),2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'C{WT}', f'=IF($B{WT}=0,"",1)', font=F_TOT, fill=FILL_TOT, fmt=PCT)

widths(ws, {'A': 30, 'B': 15, 'C': 15, 'D': 15, 'E': 15, 'F': 15, 'G': 15, 'H': 15, 'I': 16,
            'J': 16, 'K': 16})
ws.freeze_panes = 'B6'
page(ws)
MON_TOT, FEE_TOT, ARE_TOT, BIL_TOT, WAY_TOT, MON_LAST = MT_, FT, PT, QT, WT, M0 + 11
print('  ✓ 月度汇总')

# ══════════════════════════════════════════════════════════════
# ⑩ 核对表
# ══════════════════════════════════════════════════════════════
H3_ = HDR - 1   # 各表的合计行
DZ = f'ROUND(SUM(应收对账!$Z${K0}:$Z${K1}),2)'
CHECKS = [
    ('本年应收：《应收对账》合计 ＝《应收登记》本年度合计',
     f'应收对账!$I${H3_}', f'ROUND(SUMIF(应收年度,管理年度,应收金额),2)',
     '对账表是按项目汇的，登记表是按笔记的，两头必须一样多。不等就是有应收挂到了《维保合同》里没有的项目上。'),
    ('本年已收：《应收对账》合计 ＝《收款登记》本年度合计',
     f'应收对账!$N${H3_}', f'ROUND(SUMIF(收款年度,管理年度,收款金额),2)',
     '同上，收款这一路。'),
    ('上年结转欠款 ＝ 上年及以前「应收 － 收款」',
     f'应收对账!$H${H3_}', f'ROUND(SUMIF(应收年度,"<"&管理年度,应收金额)-SUMIF(收款年度,"<"&管理年度,收款金额),2)',
     '★ 这一条就是「跨年衔接」：上年没收完的钱自动变成本年期初，不用手工搬数。'),
    ('欠款余额 ＝ 全部「应收 － 收款」（不分年度）',
     f'应收对账!$T${H3_}', f'ROUND(SUM(应收金额)-SUM(收款金额),2)',
     '最终还欠多少，以这一条为准。'),
    ('期初 ＋ 本年应收 － 本年已收 ＋ 以后年度 ＝ 欠款余额',
     f'ROUND(应收对账!$H${H3_}+应收对账!$I${H3_}-应收对账!$N${H3_}+{DZ},2)', f'应收对账!$T${H3_}',
     '三年口径串起来，串不上说明有哪一行的年度判错了。'),
    ('已开票应收 ＋ 未开票应收 ＝ 还欠的钱',
     f'ROUND(应收对账!$R${H3_}+应收对账!$S${H3_},2)', f'ROUND(SUMIF(对账欠款,">0"),2)',
     '★ 欠款拆两半：票开了钱没到 / 票还没开。两半加起来得等于欠款（多收的不算在里面）。'),
    ('合同款 ＋ 合同外款 ＝ 本年应收',
     f'ROUND(应收对账!$J${H3_}+应收对账!$K${H3_},2)', f'应收对账!$I${H3_}',
     '★ 合同内外分开登记、合并收款，这一条保证分开之后没漏没重。'),
    ('发票收款 ＋ 收据收款 ＋ 现金无票 ＝ 本年已收',
     f'ROUND(应收对账!$O${H3_}+应收对账!$P${H3_}+应收对账!$Q${H3_},2)', f'应收对账!$N${H3_}',
     '★ 收款乱不乱，就看这一条：三路加起来必须等于总收款。'),
    ('逐单核销：《应收登记》未收余额合计 ＝ 各项目欠款合计',
     f'应收登记!$Q${H3_}', f'ROUND(SUMIF(对账欠款,">0"),2)',
     '按单核销（指定冲销 + 先挂先冲）的结果，跟按项目汇总的结果要一致。'),
    ('《应收对账》有没有不平的行',
     f'COUNTIF(应收对账!$X${K0}:$X${K1},"★*")', '0', '有★就点进去看那一行。'),
    ('《应收登记》有没有要改的行',
     f'COUNTIF(应收登记!$T${A0}:$T${A1},"★*")', '0', '没选客户项目、金额空着、开票大于应收，都会报。'),
    ('《收款登记》有没有要改的行',
     f'COUNTIF(收款登记!$O${R0}:$O${R1},"★*")', '0', '关联单号填错、指定冲销超额，都会报。'),
    ('《电梯资料》有没有要改的行',
     f'COUNTIF(电梯资料!$AE${L0}:$AE${L1},"★*")', '0', '免保没填到期日、台量没填、还没建合同行，都会报。'),
    ('《维保合同》有没有要改的行',
     f'COUNTIF(维保合同!$X${K0}:$X${K1},"★*")', '0', '一个项目建了两行、合同起止反了，都会报。'),
    ('《月度汇总》12 个月应收合计 ＝ 本年应收',
     f'月度汇总!$B${MON_TOT}', f'应收对账!$I${H3_}', '按日期切的月，跟按年度汇的总数要对上。'),
    ('《月度汇总》12 个月收款合计 ＝ 本年已收',
     f'月度汇总!$F${MON_TOT}', f'应收对账!$N${H3_}', '同上。'),
    ('《月度汇总》按费用类别合计 ＝ 本年应收',
     f'月度汇总!$C${FEE_TOT}', f'应收对账!$I${H3_}', '有人不走下拉直接手打费用类别，会掉进「不在清单里的」那一行，总数照样平。'),
    ('《月度汇总》按票据类型合计 ＝ 本年已收',
     f'月度汇总!$C${BIL_TOT}', f'应收对账!$N${H3_}', '同上。'),
    ('《月度汇总》按收款方式合计 ＝ 本年已收',
     f'月度汇总!$B${WAY_TOT}', f'应收对账!$N${H3_}', '同上。'),
    ('《月度汇总》12 月累计欠款 ＝ 欠款余额 － 以后年度',
     f'月度汇总!$K${MON_LAST}', f'ROUND(应收对账!$T${H3_}-{DZ},2)', '月份表的滚存跟对账表的余额要合得上。'),
    ('《维保合同》本年应收合计 ＝《应收对账》本年应收合计',
     f'维保合同!$T${H3_}', f'应收对账!$I${H3_}', '合同表和对账表是两套公式算出来的，互相印证。'),
    ('《维保合同》欠款合计 ＝《应收对账》欠款合计',
     f'维保合同!$V${H3_}', f'应收对账!$T${H3_}', '同上。'),
    ('《月度汇总》按片区欠款合计 ＝《维保合同》欠款合计',
     f'月度汇总!$G${ARE_TOT}', f'维保合同!$V${H3_}', '片区没填的会掉到「没分片区的」那一行，总数照样平。'),
    ('电梯台量：《电梯资料》合计 ＝《维保合同》在册台量合计',
     f'电梯资料!$P${H3_}', f'维保合同!$K${H3_}',
     '不等 ＝ 有电梯的「客户项目」在《维保合同》里没有对应行，去《电梯资料》AE 列找「★还没建合同行」。'),
    ('《分期收款计划》各合同「各期计划金额合计 ＝ 合同总额」',
     f'COUNTIF(分期收款计划!$V${T0}:$V${T1},"★*")', '0',
     '★ 分期付款的底线：把一份合同拆成几期，几期加起来必须正好等于合同总额，一分不能多也不能少。'),
    ('有没有合同的期数排不下（超过 %d 期）' % MAX_TERM,
     f'SUMPRODUCT((维保合同!$C${K0}:$C${K1}<>"")*'
     f'(维保合同!$Z${K0}:$Z${K1}*维保合同!$Y${K0}:$Y${K1}/12>{MAX_TERM}))', '0',
     '比如两年期的半月付合同要 48 期，表里一份合同只排 %d 期。真碰上了就把合同按年拆成两行。' % MAX_TERM),
    ('《分期收款计划》合同总额合计 ＝《维保合同》按月折算的合同总额合计',
     f'分期收款计划!$J${HDR-1}', f'ROUND(SUM(维保合同!$AA${K0}:$AA${K1}),2)',
     '合同总额 ＝ 年费 × 合同月数 ÷ 12，两张表要一致。'),
]
OVER = [('《到期与提醒》① 合同提醒', TOT_C, N_B1),
        ('《到期与提醒》② 年检/校验提醒', TOT_J, N_B2),
        ('《到期与提醒》③ 免保到期提醒', TOT_M, N_B3),
        ('《到期与提醒》④ 欠款账龄预警', TOT_Q, N_B4)]

ws = sheet('核对表')
title(ws, '核 对 表 · 一眼看出这份表还平不平', 'G',
      '★ F 列全是「✔ 平」就说明数据是自洽的；出现「★ 差」就点开 B 列写的那两张表去对。\n'
      '★ 这张表不用填，改完数据按 Ctrl+Alt+F9 重算一次再看。', color=C_DASH)
headers(ws, HDR, ['序号', '核 对 项', '口径一', '口径二', '差额', '结论', '这一条是在防什么'],
        fill_color=C_DASH, height=28)
CK0 = HDR + 1
for i, (nm, lf, rt, note) in enumerate(CHECKS):
    r = CK0 + i
    put(ws, f'A{r}', i + 1, font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'B{r}', nm, font=F_TXT, fill=FILL_AUTO, align=CL_)
    put(ws, f'C{r}', f'={lf}', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'D{r}', f'={rt}', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'E{r}', f'=ROUND($C{r}-$D{r},2)', font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'F{r}', f'=IF(ABS($E{r})<0.005,"✔ 平","★ 差 "&TEXT($E{r},"#,##0.00"))',
        font=F_TOT, fill=FILL_KPI)
    put(ws, f'G{r}', note, font=F_NOTE, fill=FILL_AUTO, align=CL_)
CK1 = CK0 + len(CHECKS) - 1
for j, (nm, tot, cap) in enumerate(OVER):
    r = CK1 + 1 + j
    put(ws, f'A{r}', len(CHECKS) + j + 1, font=F_AUTO, fill=FILL_AUTO, fmt='0')
    put(ws, f'B{r}', nm + ' 的条数有没有超出显示上限', font=F_TXT, fill=FILL_AUTO, align=CL_)
    put(ws, f'C{r}', f'={tot}', font=F_IN, fill=FILL_AUTO, fmt=NUM)
    put(ws, f'D{r}', cap, font=F_IN, fill=FILL_AUTO, fmt=NUM)
    put(ws, f'E{r}', f'=$D{r}-$C{r}', font=F_TOT, fill=FILL_AUTO, fmt=NUM)
    put(ws, f'F{r}', f'=IF($C{r}<=$D{r},"✔ 装得下","★ 超出 "&TEXT($C{r}-$D{r},"0")&" 条")',
        font=F_TOT, fill=FILL_KPI)
    put(ws, f'G{r}', '超了就去源表用筛选看全量，或者把提醒天数调小一点。', font=F_NOTE, fill=FILL_AUTO, align=CL_)
CKE = CK1 + len(OVER)
r = CKE + 1
put(ws, f'A{r}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'B{r}', f'=" 合计 {len(CHECKS)+len(OVER)} 项，通过 "&COUNTIF($F${CK0}:$F${CKE},"✔*")&" 项，'
                 f'不平 "&COUNTIF($F${CK0}:$F${CKE},"★*")&" 项 "', font=F_BIG, fill=FILL_TOT, align=CL_)
ws.merge_cells(f'B{r}:E{r}')
put(ws, f'C{r}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'D{r}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'E{r}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'F{r}', f'=IF(COUNTIF($F${CK0}:$F${CKE},"★*")=0,"✔ 全平","★ 有不平")',
    font=F_BIG, fill=FILL_TOT)
put(ws, f'G{r}', None, font=F_TOT, fill=FILL_TOT)
ws.row_dimensions[r].height = 26
CK_ALL, CK_ROW = f'核对表!$F${CK0}:$F${CKE}', r
widths(ws, {'A': 5, 'B': 48, 'C': 16, 'D': 16, 'E': 14, 'F': 18, 'G': 62})
ws.freeze_panes = 'C5'
page(ws)
print('  ✓ 核对表')

# ══════════════════════════════════════════════════════════════
# ⑪ 主页
# ══════════════════════════════════════════════════════════════
from openpyxl.worksheet.hyperlink import Hyperlink

ws = sheet('主页')
title(ws, '电 梯 维 保 管 理 系 统 · 主 页', 'H',
      '★ 蓝底格子是可以点进去的表名。所有数字都是自动算的，只在《基础资料》《电梯资料》《维保合同》'
      '《应收登记》《收款登记》这五张表的淡黄格里填东西。')


def card(rl, c0, c1, label, formula, fmt=MONEY, fill=FILL_KPI, font=F_BIG):
    put(ws, f'{c0}{rl}', label, font=F_HDR2, fill=FILL_SEC)
    if c0 != c1:
        ws.merge_cells(f'{c0}{rl}:{c1}{rl}')
        ws.merge_cells(f'{c0}{rl+1}:{c1}{rl+1}')
    put(ws, f'{c0}{rl+1}', formula, font=font, fill=fill, fmt=fmt)
    for cc in range(ord(c0), ord(c1) + 1):
        put(ws, f'{chr(cc)}{rl}', None, font=F_HDR2, fill=FILL_SEC)
        put(ws, f'{chr(cc)}{rl+1}', None, font=font, fill=fill)
    put(ws, f'{c0}{rl}', label, font=F_HDR2, fill=FILL_SEC)
    put(ws, f'{c0}{rl+1}', formula, font=font, fill=fill, fmt=fmt)
    ws.row_dimensions[rl + 1].height = 24


def sec(rr, txt, color=C_MAIN):
    put(ws, f'A{rr}', txt, font=F_HDR, fill=PatternFill('solid', fgColor=color), align=CL_)
    for cc in 'BCDEFGH':
        put(ws, f'{cc}{rr}', None, font=F_HDR, fill=PatternFill('solid', fgColor=color))
    ws.merge_cells(f'A{rr}:H{rr}')
    ws.row_dimensions[rr] = ws.row_dimensions[rr]
    ws.row_dimensions[rr].height = 22


sec(4, ' ① 本年概览（年度＝《基础资料》D3，改那一格整套表换年）')
card(5, 'A', 'B', '管理年度', '=管理年度&" 年"', fmt=None)
card(5, 'C', 'D', '在册台量', f'=电梯资料!$P${H3_}', fmt=NUM)
card(5, 'E', 'F', '在保项目数', f'=COUNTIF(合同键,"?*")', fmt=NUM)
card(5, 'G', 'H', '合同没回来的份数', f'=COUNTIF(维保合同!$R${K0}:$R${K1},"▲*")', fmt=NUM, fill=FILL_WARN, font=F_WARN)
card(8, 'A', 'B', '本年应收', f'=应收对账!$I${H3_}')
card(8, 'C', 'D', '本年已收', f'=应收对账!$N${H3_}')
card(8, 'E', 'F', '本年回款率', f'=IFERROR(应收对账!$N${H3_}/(应收对账!$H${H3_}+应收对账!$I${H3_}),"")', fmt=PCT)
card(8, 'G', 'H', '欠款余额(累计)', f'=应收对账!$T${H3_}', fill=FILL_WARN, font=F_WARN)
card(11, 'A', 'B', '已开票应收（票开了钱没到）', f'=应收对账!$R${H3_}')
card(11, 'C', 'D', '未开票应收（票还没开）', f'=应收对账!$S${H3_}')
card(11, 'E', 'F', '本年现金/无票收款', f'=应收对账!$Q${H3_}')
card(11, 'G', 'H', '全表核对', f'=核对表!$F${CK_ROW}', fmt=None)

sec(14, ' ② 分期付款回款情况（明细点《分期收款计划》）', C_DASH)
card(15, 'A', 'B', '截至今天该回款(按分期计划)', f'=分期收款计划!$N${PLAN_TOT}')
card(15, 'C', 'D', '实际已回款', f'=分期收款计划!$P${PLAN_TOT}')
card(15, 'E', 'F', '回款缺口', f'=分期收款计划!$Q${PLAN_TOT}', fill=FILL_WARN, font=F_WARN)
card(15, 'G', 'H', '没按期收齐的合同', f'=COUNTIF(分期收款计划!$U${T0}:$U${T1},"▲*")&" 个"',
     fmt=None, fill=FILL_WARN, font=F_WARN)

sec(18, ' ③ 提醒速览（详细清单点《到期与提醒》）', C_WARN)
card(19, 'A', 'B', '合同到期 / 没回签', f'={TOT_C}&" 条"', fmt=None, fill=FILL_WARN, font=F_WARN)
card(19, 'C', 'D', '年检 / 校验到期', f'={TOT_J}&" 台"', fmt=None, fill=FILL_WARN, font=F_WARN)
card(19, 'E', 'F', '免保到期该转收费', f'={TOT_M}&" 台"', fmt=None, fill=FILL_WARN, font=F_WARN)
card(19, 'G', 'H', '还有欠款的项目', f'={TOT_Q}&" 个"', fmt=None, fill=FILL_WARN, font=F_WARN)

sec(22, ' ④ 各表导航（点表名直接跳过去）')
NAV = [('使用说明', '怎么用、怎么跨年、坑在哪，先看这个', '看'),
       ('基础资料', '片区、梯型、维保状态、费用类别、票据类型…所有下拉都在这里改', '填'),
       ('电梯资料', '一台梯子一行：注册代码、年检日期、维保状态、免保到期', '填'),
       ('维保合同', '一个项目一行：合同起止、年费、合同状态（没回来也先建行）', '填'),
       ('应收登记', '该收多少钱，一笔一行（合同款 / 合同外款都在这）', '填'),
       ('收款登记', '钱几时到、开的什么票，一笔一行', '填'),
       ('应收对账', '★ 每个项目欠多少：上年结转 / 本年应收 / 已收 / 已开票应收 / 未开票应收 / 欠款', '看'),
       ('分期收款计划', '★ 按季 / 半月 / 半年付款的合同，每期该回多少、回了没有、欠几期', '看'),
       ('到期与提醒', '该催的合同、该约的年检、该转收费的免保、该要的钱', '看'),
       ('月度汇总', '按月 / 费用类别 / 片区 / 票据类型 / 收款方式 五个口径', '看'),
       ('核对表', '这份表还平不平，一眼看出来', '看')]
NV = 23
headers(ws, NV, ['表 名', '填/看', '这张表是干什么的', '', '', '', '', ''], height=24)
ws.merge_cells(f'C{NV}:H{NV}')
for i, (nm, desc, kind) in enumerate(NAV):
    r = NV + 1 + i
    c = put(ws, f'A{r}', nm, font=F_LINK, fill=FILL_SEC, align=C)
    c.hyperlink = Hyperlink(ref=f'A{r}', location=f"'{nm}'!A1")
    put(ws, f'B{r}', kind, font=F_TOT, fill=FILL_AUTO)
    put(ws, f'C{r}', desc, font=F_TXT, fill=FILL_AUTO, align=CL_)
    for cc in 'DEFGH':
        put(ws, f'{cc}{r}', None, font=F_TXT, fill=FILL_AUTO)
    ws.merge_cells(f'C{r}:H{r}')
    ws.row_dimensions[r].height = 20

widths(ws, {'A': 18, 'B': 10, 'C': 22, 'D': 18, 'E': 16, 'F': 18, 'G': 18, 'H': 20})
page(ws, landscape=False)
print('  ✓ 主页')

# ══════════════════════════════════════════════════════════════
# ⑫ 使用说明
# ══════════════════════════════════════════════════════════════
ws = sheet('使用说明')
title(ws, '使 用 说 明 · 花十分钟看完，后面省很多事', 'H',
      '★ 按 F11 或点左下角表名切换。这张表只是说明，不影响任何计算。')
DOC = [
    ('sec', '一、这套表是怎么搭的'),
    ('p', '原来那两张表（电梯资料 + 合同应收款）最大的毛病是：合同、开票、收款全挤在一行。'
          '一个客户分三次收、有的开发票有的开收据，一行根本写不下；合同没回来更没地方挂。'),
    ('p', '这一版拆成「三张流水 + 一张对账」：'),
    ('p', '　　《维保合同》—— 一个「客户项目」一行，记合同起止、年费、合同状态。台量多也不怕，一行管一个项目。'),
    ('p', '　　《应收登记》—— 该收人家多少钱，一笔一行。维保费、维修费、配件费、报检费…都往这里记。'),
    ('p', '　　《收款登记》—— 钱几时到、开的什么票、怎么付的，一笔一行。'),
    ('p', '　　《应收对账》—— 全自动。每个项目的：上年结转 → 本年应收 → 已开票 / 未开票 → 已收（发票/收据/现金三路）→ 最后欠多少、拖了多少天。'),
    ('p', '串起这四张表的是一把钥匙：E 列「客户项目」＝ 使用单位 ＋ " ｜ " ＋ 项目名称。同一个单位有几个小区/几栋楼，就按项目分开建行，各算各的账。'),
    ('sec', '二、第一次用，照这个顺序填'),
    ('p', '① 《基础资料》：把片区、梯型、维保性质、维保状态、合同状态、收费周期、费用类别、票据类型、收款方式这几张小清单改成你们自己的说法；'
          '再把第 3 行四个参数定一下（管理年度、合同提前提醒天数、年检提前提醒天数、欠款账龄预警天数、分期应收款日账期天数）。'),
    ('p', '② 《电梯资料》：一台梯子一行。使用单位、项目名称必须填（这两列拼出钥匙）；台量一般填 1。'
          '维保状态选了「安装免保 / 技术免保 / 质保期内」的，一定要填「免保到期日」。'),
    ('p', '③ 《维保合同》：一个项目一行。合同没回来也先建行 —— 合同编号先空着，合同状态选「已签未回」或「未签」。'),
    ('p', '④ 《应收登记》《收款登记》：往下记流水就行，不用分年度、不用分表。'),
    ('p', '⑤ 填完去《核对表》扫一眼，F 列全是「✔ 平」就对了。'),
    ('sec', '三、合同没及时回来，只记得住单位名称 —— 怎么办'),
    ('p', '《维保合同》里照样建行：C 列使用单位、D 列项目名称先填上，F 列合同编号空着，G 列合同状态选「已签未回」「未签」或「续签中」。'),
    ('p', '这样做之后：R 列会一直挂着「▲合同未回（已签未回）」；《到期与提醒》① 里单独列一块，专门给你催合同用；'
          '《应收对账》V 列那个项目后面也会跟一句「▲合同未回」。合同回来了，把编号补上、状态改成「已签已回」，提示自己就没了。'),
    ('p', '重点是：合同没回来不影响挂应收、不影响收款、不影响对账。该收的钱一分不少地记着。'),
    ('sec', '四、合同款和合同外的款：分开登记，合并收款'),
    ('p', '登记的时候分开 —— 《应收登记》G 列选「费用类别」，右边 H 列自动判成「合同款」还是「合同外款」：'
          '维保费 / 合同维保费(补收) 算合同款；维修费、配件费、报检费、年检费、限速器校验费、载重试验费、改造费、急修…算合同外款。'
          '这个对应关系在《基础资料》M、N 两列，想改就在那儿改。'),
    ('p', '收款的时候合并 —— 客户一笔钱把维保费和维修费一起付了，《收款登记》就记一行：K 列「款别」选「合并(不细分)」，'
          'L 列关联单号留空。系统会按「先挂先冲」自动去冲这个项目最早那几笔没收完的应收。'),
    ('p', '如果客户特意交代「这笔是付去年那张维修费的」，就在 L 列把那张应收单号选上 —— 指定的优先冲，剩下的再按先挂先冲走。'),
    ('p', '最后在《应收对账》里，本年应收会自动拆成 J「其中：合同款」和 K「其中：合同外款」两列，合起来一分不差。'),
    ('sec', '五、合同到期提醒'),
    ('p', '《维保合同》Q 列自动算：合同止期到了没有、还剩几天。提前多少天开始提醒，在《基础资料》O3 改（默认 60 天）。'
          '已经过期的会显示「已到期 XX 天」，红字。'),
    ('p', '《到期与提醒》① 把「快到期」和「没回签」的合同压成一张连续清单，直接照着打电话就行。'),
    ('sec', '六、开票应收、未开票应收、现金收款、最后欠款'),
    ('p', '这四个数在《应收对账》里分别是 R、S、Q、T 四列：'),
    ('p', '　　T 欠款余额 ＝ 这个项目从头到尾「应收合计 － 收款合计」，跨多少年都算进去，这是最终口径。'),
    ('p', '　　R 已开票应收 ＝ 票已经开出去了、钱还没到的那部分（累计开票额 － 累计收款，最多不超过 T）。这部分最好催，客户赖不掉。'),
    ('p', '　　S 未开票应收 ＝ T － R，票都还没开的那部分。要么是客户没要票，要么是还没走开票流程。'),
    ('p', '　　Q 现金/无票收款 ＝ 本年收进来的、没开任何票的那部分（票据类型选「无票」的）。'),
    ('p', '《核对表》第 6 条专门盯这个：R ＋ S 必须等于 T，错不了。'),
    ('sec', '七、收款乱：有的开发票、有的开收据、有的现金没票'),
    ('p', '《收款登记》H 列「票据类型」按实际开的选：增值税专用发票 / 增值税普通发票 / 电子发票 / 收据 / 无票。'
          'I 列会自动归成三路：前三种 → 「发票」；收据 → 「收据」；无票 → 「现金无票」。'),
    ('p', '《应收对账》O、P、Q 三列就是这三路各收了多少，加起来等于 N 列本年已收（《核对表》第 8 条在盯）。'
          '《月度汇总》④ 还能看到五种票分别收了多少、各占几成。'),
    ('p', '另外 J 列「收款方式」记的是钱怎么进来的（银行转账 / 现金 / 微信 / 支付宝 / 承兑汇票 / 抵扣冲账），跟开什么票是两回事，别混。'),
    ('sec', '八、26 年的应收款怎么接到 27 年 —— 这表能连续用很多年'),
    ('p', '《应收登记》《收款登记》两张流水不分年度，一直往下记就行，过年不用新建表、不用搬数。'),
    ('p', '到了 27 年，只做一个动作：把《基础资料》D3 的「管理年度」从 2026 改成 2027。'),
    ('p', '然后《应收对账》H 列「上年结转欠款」会自动算成「2027 年以前的应收 － 2027 年以前的收款」，'
          '也就是 26 年（以及更早）没收完的钱，自动变成 27 年的期初。I 列本年应收、N 列本年已收自动只看 2027 年的。'
          'T 列欠款余额本来就是全累计，跨几年都准。'),
    ('p', '《核对表》第 3、第 5 条就是专门验这个衔接的，改完年度去看一眼是不是还「✔ 平」。'),
    ('p', '想回头看 26 年的账？把管理年度改回 2026，整套表立刻变回 26 年口径，看完再改回来，数据一个字都不会变。'),
    ('sec', '九、维保状态：安装免保、技术免保怎么记'),
    ('p', '《电梯资料》Y 列选：正常维保 / 安装免保 / 技术免保 / 质保期内 / 暂停维保 / 已解约。'
          '选了前三种里的免保类型，Z 列「免保到期日」必须填 —— 不填 AE 列会报「★免保到期日没填」。'),
    ('p', '免保到期前后，AC 列会提示「免保还剩 XX 天」→「▲免保已到期 该转收费了」，《到期与提醒》③ 也单列一块，'
          '免得免保期过了还在白干。'),
    ('p', '《维保合同》O 列「免保台量」、P 列「收费台量」会自动按电梯的维保状态算出来，一眼看出这个项目实际该按几台收钱。'),
    ('sec', '十、年检、限速器校验、载重试验提醒'),
    ('p', '《电梯资料》T、U、V 三列填下次到期日。AA 列盯年检，AB 列盯限速器校验和载重试验（取两者里较早的那个）。'
          '提前多少天提醒在《基础资料》D4 改（默认 30 天）。'),
    ('p', '《到期与提醒》② 把所有该约检的梯子压成一张清单，「剩余天数」是红色负数就是已经过期了。'),
    ('sec', '十一、金额大的合同分季度 / 半月 / 半年付款，回款怎么算'),
    ('p', '先在《维保合同》把三样填对：H/I 合同起止、L 合同年费（含税）、M 收费周期。'
          '收费周期就是「一年分几期收」——《基础资料》K、L 两列定的：一次性 1 期、半年一次 2 期、'
          '季度一次 4 期、月度 12 期、半月一次 24 期、按次不排期。要加别的周期就在那两列加一行。'),
    ('p', '填完《分期收款计划》自动排出来，不用手工算：'),
    ('p', '　　合同总额 ＝ 合同年费 × 合同月数 ÷ 12（所以半年合同、一年半合同都算得对）；'
          '每期金额 ＝ 合同总额 ÷ 期数，除不尽的尾差进最后一期，几期加起来正好等于合同总额。'),
    ('p', '　　每期的区间也自动排：季度付就是合同起算起每 3 个月一期，半年付每 6 个月一期，'
          '半月付就是每月切成 1—15 号、16 号—月底两期。'),
    ('p', '　　「应收款日」＝ 每期期初 ＋《基础资料》D5 的账期天数。合同写「每期期初付」就填 0，'
          '写「期初 15 天内付清」就填 15，写「期末付」就填每期的天数。过了这一天还没收够就开始算逾期天数。'),
    ('p', '这张表分两块看：'),
    ('p', '　　① 按合同看回款进度 —— 一个合同一行：已到期几期、累计该回多少、实际回了多少、'
          '差多少（还差约几期）、回款进度百分比、下一期什么时候该收多少。催款、开会汇报看这一块。'),
    ('p', '　　② 分期明细排期 —— 一个合同一期一行：这一期的区间、该收多少、应收款日、'
          '有没有挂应收、收到多少、状态、逾期多少天。点表头的筛选按「本期状态」筛「▲」，'
          '一眼看出哪几期没收到、拖了多久，直接照着催。'),
    ('p', '「累计已回款」用的是核销口径 —— 客户一笔钱把两三期一起付了，系统按「先挂先冲」'
          '自动分摊到各期上，所以不会因为收款和期次对不齐就算错。'),
    ('p', '状态会分三种提醒，意思不一样，别搞混：'),
    ('p', '　　「▲还没挂应收 XXX」＝ 这一期该收的钱，你在《应收登记》里还没挂上去，先去补登记；'),
    ('p', '　　「▲欠 XXX」＝ 应收挂了、票也可能开了，就是钱没回来，该催款了；'),
    ('p', '　　「✔ 已收清」＝ 到这一期为止累计收够了。'),
    ('p', '例子：建业森林半岛 24 台，年费 86,400，合同约定半月付一次 → 24 期，每期 3,600。'
          '到今天已经过了 18 期，累计该回 64,800；实际收到 50,400（只付到 7 月），'
          '缺口 14,400，约等于欠 4 期 —— ① 里直接显示「▲欠 14,400.00（约 4.0 期）」，'
          '② 里能看到是 8 月上半月那一期开始断的、拖了多少天。'),
    ('p', '注意两点：一是一份合同最多排 24 期，两年期的半月付合同排不下，把合同按年拆成两行；'
          '二是排期是「计划」，《应收登记》是「实际」，两边对不上不是错 —— 正是要靠这个差额'
          '发现「该挂的应收忘了挂」。《核对表》只保证各期金额加起来等于合同总额。'),
    ('sec', '十二、日常动作清单'),
    ('p', '每天：有维修、有换件、有报检 —— 《应收登记》加一行；有钱到账 —— 《收款登记》加一行。'),
    ('p', '每周：打开《到期与提醒》，① 催合同、② 约年检、③ 转收费、④ 催款，四块顺着做一遍；'
          '再看《分期收款计划》② 筛出「▲」的期次，分期付款的合同照着催。'),
    ('p', '每月：看《月度汇总》①，本月应收多少、收回来多少、累计还欠多少；看《应收对账》，按欠款余额从大到小排一排（用筛选）。'),
    ('p', '每年：《基础资料》D3 改年度；《维保合同》里到期的续签、改合同起止和年费。'),
    ('sec', '十三、几个容易踩的坑'),
    ('p', '① 淡黄格＝手工填，浅灰格＝公式自动算。往灰格里打字会把公式覆盖掉，那一行就废了。'
          '不小心覆盖了，按 Ctrl+Z 撤销，或者从上下相邻行往下拉一下把公式补回来。'),
    ('p', '② 「客户项目」一定要从下拉里选，别手打。手打多一个空格、少一个字，就对不上号了，核对列会报「★《维保合同》里没这个项目」。'),
    ('p', '③ 《维保合同》里一个「客户项目」只能建一行，建重了 X 列报「★键重复」，对账数会翻倍。'),
    ('p', '④ 金额全部按含税价填，跟开票金额、到账金额一致就行，这套表不单独拆税。'),
    ('p', '⑤ 关联应收单号是选填的，不填不影响对账总数，只影响《应收登记》P、Q 两列那一行显示得准不准。'),
    ('p', '⑥ 改完数据数字没变，按 Ctrl＋Alt＋F9 强制重算一次。'),
    ('p', '⑦ 行不够了：选中最后一行整行，往下拖填充柄，公式和下拉都会跟着复制。'
          '《电梯资料》留了 600 行、《维保合同》《应收对账》《分期收款计划》300 个合同、'
          '《应收登记》1500 行、《收款登记》2000 行，一般够用好几年。'),
]
dr = 4
for kind, txt in DOC:
    if kind == 'sec':
        put(ws, f'A{dr}', txt, font=F_SEC, fill=FILL_SEC, align=CL_)
        for cc in 'BCDEFGH':
            put(ws, f'{cc}{dr}', None, font=F_SEC, fill=FILL_SEC)
        ws.merge_cells(f'A{dr}:H{dr}')
        ws.row_dimensions[dr].height = 24
    else:
        put(ws, f'A{dr}', txt, font=F_TXT, fill=None, align=CT, border=None)
        ws.merge_cells(f'A{dr}:H{dr}')
        ws.row_dimensions[dr].height = max(18, 17 * (1 + len(txt) // 58))
    dr += 1
widths(ws, {c: 18 for c in 'ABCDEFGH'})
ws.column_dimensions['A'].width = 26
page(ws, landscape=False)
print('  ✓ 使用说明')

# ══════════════════════════════════════════════════════════════
# 收尾：表序、页签颜色、保存
# ══════════════════════════════════════════════════════════════
ORDER = ['主页', '使用说明', '基础资料', '电梯资料', '维保合同', '应收登记', '收款登记',
         '应收对账', '分期收款计划', '到期与提醒', '月度汇总', '核对表', '_自动清单']
wb._sheets = [SH[n] for n in ORDER]
TABC = {'主页': C_MAIN, '使用说明': '808080', '基础资料': C_BASE, '电梯资料': C_BASE,
        '维保合同': C_BASE, '应收登记': C_BIZ, '收款登记': C_BIZ, '应收对账': C_RPT,
        '分期收款计划': C_DASH, '到期与提醒': C_WARN, '月度汇总': C_RPT,
        '核对表': C_DASH, '_自动清单': 'BFBFBF'}
for n, c in TABC.items():
    SH[n].sheet_properties.tabColor = c
for n in ORDER:
    SH[n].sheet_view.showGridLines = False
    SH[n].sheet_view.tabSelected = (n == '主页')
wb.active = 0
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
n_f = sum(1 for s in wb.worksheets for row in s.iter_rows()
          for c in row if isinstance(c.value, str) and c.value.startswith('='))
print(f'\n✅ 已生成：{OUT}')
print(f'   {len(wb.worksheets)} 张表 / {n_f:,} 条公式 / {os.path.getsize(OUT)/1048576:.2f} MB')
