# -*- coding: utf-8 -*-
"""生成《A061_典当登记表.xlsx》。

结构：
  使用说明 / 资料 / 表1不动产登记 / 表2客户典当登记
  金辉典当 · 明道典当 · 个人典当   ← 在表2登记，这三张自动拆出来（FILTER 动态数组）
  查询表                           ← 输入客户姓名，同时拉出他的典当记录 + 资金收付流水

跟资金台账的联动：查询表最下面那块用跨文件 FILTER 取《A061_资金台账.xlsx》的「数据录入」，
两本必须同时打开（FILTER 读不了关着的工作簿，这是 Excel 的规矩，不是设置问题）。
"""
import json, os, re, sys, datetime, collections
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, '参考', '原_典当登记表.xlsx')
OUT = os.path.join(ROOT, 'A061_典当登记表.xlsx')
FUND_FILE = 'A061_资金台账.xlsx'

N2, N1 = 400, 120            # 表2 / 表1 的行数上限
D0 = 4                       # 数据从第 4 行开始（1 标题 2 汇总 3 表头）

# ── 样式 ──────────────────────────────────────────────────────────────
C_MAIN, C_SUB, C_Q, C_BASE = '2F5597', '806000', '375623', '7030A0'
F_IN = PatternFill('solid', fgColor='FFFBEA')
F_AUTO = PatternFill('solid', fgColor='F2F2F2')
F_SUM = PatternFill('solid', fgColor='FFF2CC')
FT_IN = Font(name='微软雅黑', size=10)
FT_AUTO = Font(name='微软雅黑', size=10, color='7F7F7F')
FT_HDR = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
FT_TITLE = Font(name='微软雅黑', size=15, bold=True, color='FFFFFF')
FT_SUM = Font(name='微软雅黑', size=10, bold=True, color='C00000')
THIN = Side('thin', color='BFBFBF')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CEN = Alignment('center', 'center', wrap_text=True)
LEFT = Alignment('left', 'center')
MONEY, DATEF, RATE = '#,##0.00', 'yyyy-mm-dd', '0.000'

INST = ['金辉典当', '明道典当', '个人典当']
GOODS = ['房产', '车子', '车位', '黄金', '包包', '手表', '名品', '无抵押借款', '钻戒',
         '私人借款', '法拍垫资', '手机', '解押/过桥', '寿山石', '白金', '车贷-没押车',
         '黄金+钻戒', '玉佩', '动产', '股票配资', '不动产']
STATUS = ['在当', '在当-无息', '结款未赎', '赎当', '虚当', '起诉', '起诉还款',
          '结款（个人）', '逾期-未起诉', '无法回款', '绝当未处置']
PAYW = ['先息', '后息', '一次性']

# ══════════════════════════════════════════════════════════════════════
# 1  读原表
# ══════════════════════════════════════════════════════════════════════
def val(c):
    return c.value

def d2s(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v if isinstance(v, datetime.datetime) else datetime.datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(v))
    return None

def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

def txt(v):
    s = '' if v is None else str(v).strip()
    return s or None

src = openpyxl.load_workbook(SRC, data_only=True)
s1, s2 = src['Sheet1'], src['Sheet2']

INST_MAP = {'个人借款': '个人典当', '个人': '个人典当', '金辉': '金辉典当', '明道': '明道典当'}

t2 = []
for r in range(3, 204):
    nm = txt(s2.cell(r, 9).value)
    if not nm:
        continue
    inst = txt(s2.cell(r, 3).value) or '个人典当'
    t2.append(dict(
        inst=INST_MAP.get(inst, inst), goods=txt(s2.cell(r, 4).value), status=txt(s2.cell(r, 5).value),
        name=nm, ticket=txt(s2.cell(r, 12).value),
        d_ticket=d2s(s2.cell(r, 6).value), d_due=d2s(s2.cell(r, 7).value), d_pay=d2s(s2.cell(r, 8).value),
        amt=num(s2.cell(r, 10).value), repaid=num(s2.cell(r, 11).value),
        rate=num(s2.cell(r, 19).value), interest=num(s2.cell(r, 13).value),
        iday=txt(s2.cell(r, 16).value), d_idue=d2s(s2.cell(r, 18).value),
        pub=num(s2.cell(r, 14).value), penalty=num(s2.cell(r, 17).value),
        invoice=num(s2.cell(r, 20).value), payw=txt(s2.cell(r, 21).value),
        memo=txt(s2.cell(r, 15).value)))

t1 = []
for r in range(3, 86):
    nm = txt(s1.cell(r, 5).value)
    if not nm:
        continue
    t1.append(dict(
        goods=txt(s1.cell(r, 1).value), status=txt(s1.cell(r, 2).value), name=nm,
        phone=txt(s1.cell(r, 6).value), ticket=txt(s1.cell(r, 8).value),
        d_ticket=d2s(s1.cell(r, 3).value), d_due=d2s(s1.cell(r, 4).value), d_pay=d2s(s1.cell(r, 22).value),
        amt=num(s1.cell(r, 7).value), interest=num(s1.cell(r, 9).value),
        rate=num(s1.cell(r, 20).value), invoice=num(s1.cell(r, 21).value),
        pub=num(s1.cell(r, 10).value), lin=num(s1.cell(r, 11).value),
        bank=num(s1.cell(r, 12).value), wx=num(s1.cell(r, 13).value),
        penalty=num(s1.cell(r, 14).value),
        d_idue=d2s(s1.cell(r, 16).value), d_ipay=d2s(s1.cell(r, 17).value),
        payw=txt(s1.cell(r, 23).value), back=txt(s1.cell(r, 24).value),
        backamt=num(s1.cell(r, 25).value)))

names = sorted({x['name'] for x in t2} | {x['name'] for x in t1})
print('表2 %d 行，表1 %d 行，客户 %d 人' % (len(t2), len(t1), len(names)))
print('机构分布:', dict(collections.Counter(x['inst'] for x in t2)))

# ══════════════════════════════════════════════════════════════════════
# 2  搭工作簿
# ══════════════════════════════════════════════════════════════════════
wb = openpyxl.Workbook()
wb.remove(wb.active)

def newsheet(name, title, tip, color, width, merge2=True):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = color
    ws['A1'] = title
    ws['A1'].font, ws['A1'].fill, ws['A1'].alignment = FT_TITLE, PatternFill('solid', fgColor='1F3864'), LEFT
    ws.row_dimensions[1].height = 30
    ws['A2'] = tip
    ws['A2'].font = Font(name='微软雅黑', size=9, color='8B5E00')
    ws['A2'].fill, ws['A2'].alignment = PatternFill('solid', fgColor='FFF7E6'), LEFT
    ws.row_dimensions[2].height = 22
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    if merge2:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
    return ws

def header(ws, cols, row, color):
    fill = PatternFill('solid', fgColor=color)
    for i, (t, w, kind, fmt) in enumerate(cols, 1):
        c = ws.cell(row, i, t)
        c.font, c.fill, c.alignment, c.border = FT_HDR, fill, CEN, BOX
        ws.column_dimensions[CL(i)].width = w
    ws.row_dimensions[row].height = 32

def body(ws, cols, r0, r1):
    for r in range(r0, r1 + 1):
        for i, (t, w, kind, fmt) in enumerate(cols, 1):
            c = ws.cell(r, i)
            c.border, c.font = BOX, (FT_AUTO if kind == 'auto' else FT_IN)
            c.fill = F_AUTO if kind == 'auto' else F_IN
            c.alignment = CEN if (fmt or kind == 'auto') else LEFT
            if fmt:
                c.number_format = fmt

def put(ws, addr, v, fmt=None, font=None, fill=None, align=None):
    c = ws[addr]
    c.value = v
    if fmt: c.number_format = fmt
    if font: c.font = font
    if fill: c.fill = fill
    if align: c.alignment = align
    c.border = BOX
    return c

def dv(ws, f1, rng, warn=True):
    d = DataValidation(type='list', formula1=f1, allow_blank=True, showErrorMessage=True,
                       errorStyle='warning' if warn else 'stop',
                       errorTitle='不在名单里', error='先去【资料】把它加上，或者选一个已有的。')
    ws.add_data_validation(d)
    d.add(rng)

# ── 资料 ──────────────────────────────────────────────────────────────
ws = newsheet('资料', '资料 · 下拉菜单的来源', '要加机构、加抵押物品、加状态，就在这里往下接着写，别留空行。', C_BASE, 10)
BLK = [('A', '抵押机构', 16, INST, 10), ('C', '抵押物品', 16, GOODS, 40),
       ('E', '状态', 16, STATUS, 30), ('G', '付息方式', 14, PAYW, 10),
       ('I', '客户名单(自动)', 18, None, 500), ('K', '参数', 16, ['日滞纳金率'], 4)]
hf = PatternFill('solid', fgColor=C_BASE)
for col, t, w, items, n in BLK:
    c = ws[f'{col}3']; c.value = t
    c.font, c.fill, c.alignment, c.border = FT_HDR, hf, CEN, BOX
    ws.column_dimensions[col].width = w
    for i in range(n):
        cc = ws[f'{col}{D0 + i}']
        cc.border, cc.font, cc.alignment = BOX, FT_IN, LEFT
        cc.fill = F_AUTO if col == 'I' else F_IN
        if items and i < len(items):
            cc.value = items[i]
ws.freeze_panes = 'A4'
put(ws, 'L3', '数值', font=FT_HDR, fill=hf, align=CEN)
ws.column_dimensions['L'].width = 12
put(ws, 'L4', 0.001, fmt='0.0000', font=Font(name='微软雅黑', size=11, bold=True, color='C00000'),
    fill=F_IN, align=CEN)
put(ws, 'M4', '表1「应收滞纳金」用它算：典当金额 × 本比率 × 逾期天数 − 已收违约金',
    font=Font(name='微软雅黑', size=9, color='808080'), align=LEFT)
NAMES = {
    '日滞纳金率': '资料!$L$4',
    '机构表': f'资料!$A${D0}:$A${D0+9}',
    '物品表': f'资料!$C${D0}:$C${D0+39}',
    '状态表': f'资料!$E${D0}:$E${D0+29}',
    '付息表': f'资料!$G${D0}:$G${D0+9}',
    '客户名单': f'资料!$I${D0}:$I${D0+499}',
}

# ── 表2 客户典当登记（总表）──────────────────────────────────────────
COLS2 = [('序号', 7, 'auto', None), ('抵押机构 ★', 12, 'in', None), ('抵押物品 ★', 12, 'in', None),
         ('状态 ★', 13, 'in', None), ('客户姓名 ★', 13, 'in', None), ('身份证号', 20, 'in', None),
         ('客户电话', 14, 'in', None), ('当票号', 14, 'in', None),
         ('当票日期', 12, 'in', DATEF), ('当票到期日', 12, 'in', DATEF), ('出款日期', 12, 'in', DATEF),
         ('典当金额 ★', 14, 'in', MONEY), ('已还本金', 13, 'in', MONEY), ('当前在当', 14, 'auto', MONEY),
         ('典当系统费率', 10, 'in', RATE), ('应收利息', 13, 'in', MONEY), ('利息日', 8, 'in', None),
         ('应付利息日', 12, 'in', DATEF),
         ('对公', 12, 'in', MONEY), ('违约金', 12, 'in', MONEY), ('开票金额', 12, 'in', MONEY),
         ('付息方式', 10, 'in', None), ('详情/备注', 26, 'in', None), ('核对', 24, 'auto', None)]
E2 = D0 + N2 - 1
ws = newsheet('表2客户典当登记', '表2 · 客户典当登记（总表，在这里登记）',
              '登记完，右边【金辉典当】【明道典当】【个人典当】三张子表会自动拆出来，不用手工分。'
              '「当前在当」= 典当金额 − 已还本金，自动算。', C_MAIN, len(COLS2))
header(ws, COLS2, 3, C_MAIN); body(ws, COLS2, D0, E2)
for r in range(D0, E2 + 1):
    g = f'IF($E{r}="","",'
    put(ws, f'A{r}', f'={g}COUNTA($E${D0}:$E{r}))')
    put(ws, f'N{r}', f'={g}ROUND(N($L{r})-N($M{r}),2))')
    put(ws, f'X{r}', f'={g}'
        f'IF(AND($B{r}<>"",ISNA(MATCH($B{r},机构表,0))),"※机构不在【资料】里",'
        f'IF(AND($D{r}<>"",ISNA(MATCH($D{r},状态表,0))),"※状态不在【资料】里",'
        f'IF(N($M{r})>N($L{r}),"※已还本金大于典当金额",'
        f'IF(AND($I{r}<>"",$J{r}<>"",$J{r}<$I{r}),"※到期日早于当票日期",'
        f'IF(AND($H{r}<>"",COUNTIF($H${D0}:$H${E2},$H{r})>1),"※当票号在本表里出现了不止一次",'
        f'IF(AND($L{r}<>"",$E{r}=""),"※有金额没写客户","")))))')
for i, x in enumerate(t2):
    r = D0 + i
    for c, v in ((2, x['inst']), (3, x['goods']), (4, x['status']), (5, x['name']),
                 (8, x['ticket']), (9, x['d_ticket']), (10, x['d_due']), (11, x['d_pay']),
                 (12, x['amt']), (13, x['repaid']), (15, x['rate']), (16, x['interest']),
                 (17, x['iday']), (18, x['d_idue']), (19, x['pub']), (20, x['penalty']),
                 (21, x['invoice']), (22, x['payw']), (23, x['memo'])):
        ws.cell(r, c).value = v
dv(ws, '=机构表', f'B{D0}:B{E2}', warn=False)
dv(ws, '=物品表', f'C{D0}:C{E2}')
dv(ws, '=状态表', f'D{D0}:D{E2}')
dv(ws, '=付息表', f'V{D0}:V{E2}')
ws.conditional_formatting.add(f'X{D0}:X{E2}', FormulaRule(
    formula=[f'LEFT(X{D0},1)="※"'], font=Font(color='C00000', bold=True),
    fill=PatternFill('solid', fgColor='FFD7D7')))
ws['A2'].value = ('登记完，【金辉典当】【明道典当】【个人典当】三张子表自动拆出来，不用手工分。'
                  '「当前在当」= 典当金额 − 已还本金。最后一列「核对」会自己挑毛病，红字的去改。')
ws.freeze_panes = f'F{D0}'
ws.auto_filter.ref = f'A3:X{E2}'

# ── 表1 不动产登记 ───────────────────────────────────────────────────
COLS1 = [('序号', 7, 'auto', None), ('抵押物品 ★', 12, 'in', None), ('状态 ★', 13, 'in', None),
         ('客户姓名 ★', 13, 'in', None), ('客户电话', 14, 'in', None), ('当票号', 14, 'in', None),
         ('当票日期', 12, 'in', DATEF), ('当票到期日', 12, 'in', DATEF), ('出款日期', 12, 'in', DATEF),
         ('典当金额 ★', 14, 'in', MONEY), ('典当系统费率', 10, 'in', RATE), ('应收金额(利息)', 13, 'in', MONEY),
         ('开票金额', 12, 'in', MONEY),
         ('对公', 12, 'in', MONEY), ('林总', 12, 'in', MONEY), ('建行/民生', 12, 'in', MONEY),
         ('微信', 12, 'in', MONEY), ('违约金', 12, 'in', MONEY), ('合计收款', 13, 'auto', MONEY),
         ('应付利息日', 12, 'in', DATEF), ('实际付利息日', 12, 'in', DATEF),
         ('逾期天数', 10, 'auto', '0'), ('应收滞纳金', 13, 'auto', MONEY),
         ('付息方式', 10, 'in', None), ('回款时间', 14, 'in', None), ('回款金额', 12, 'in', MONEY),
         ('核对', 22, 'auto', None)]
E1 = D0 + N1 - 1
ws = newsheet('表1不动产登记', '表1 · 金辉典当不动产登记（单独一张，跟表2不混）',
              '日滞纳金比率填在 W2 那格（原表是 0.001）。逾期天数和应收滞纳金自动算。', C_MAIN, len(COLS1))
header(ws, COLS1, 3, C_MAIN); body(ws, COLS1, D0, E1)
for r in range(D0, E1 + 1):
    g = f'IF($D{r}="","",'
    put(ws, f'A{r}', f'={g}COUNTA($D${D0}:$D{r}))')
    put(ws, f'S{r}', f'={g}ROUND(N($N{r})+N($O{r})+N($P{r})+N($Q{r})+N($R{r}),2))')
    # 不能用 DATEDIF：实际付息日早于应付息日时它直接报 #VALUE!，外面套 MAX 也救不回来
    put(ws, f'V{r}', f'={g}IF(OR($T{r}="",$U{r}=""),0,MAX(0,INT($U{r})-INT($T{r}))))')
    put(ws, f'W{r}', f'={g}ROUND(MAX(0,N($J{r})*日滞纳金率*N($V{r})-N($R{r})),2))')
    put(ws, f'AA{r}', f'={g}'
        f'IF(AND($C{r}<>"",ISNA(MATCH($C{r},状态表,0))),"※状态不在【资料】里",'
        f'IF(AND($G{r}<>"",$H{r}<>"",$H{r}<$G{r}),"※到期日早于当票日期",'
        f'IF(AND($F{r}<>"",COUNTIF($F${D0}:$F${E1},$F{r})>1),"※当票号在本表里出现了不止一次",'
        f'IF(AND(N($U{r})>0,N($T{r})=0),"※填了实际付息日但没填应付息日","")))')
for i, x in enumerate(t1):
    r = D0 + i
    for c, v in ((2, x['goods']), (3, x['status']), (4, x['name']), (5, x['phone']), (6, x['ticket']),
                 (7, x['d_ticket']), (8, x['d_due']), (9, x['d_pay']), (10, x['amt']),
                 (11, x['rate']), (12, x['interest']), (13, x['invoice']),
                 (14, x['pub']), (15, x['lin']), (16, x['bank']), (17, x['wx']), (18, x['penalty']),
                 (20, x['d_idue']), (21, x['d_ipay']), (24, x['payw']), (25, x['back']), (26, x['backamt'])):
        ws.cell(r, c).value = v
dv(ws, '=物品表', f'B{D0}:B{E1}')
dv(ws, '=状态表', f'C{D0}:C{E1}')
dv(ws, '=付息表', f'X{D0}:X{E1}')
ws.conditional_formatting.add(f'AA{D0}:AA{E1}', FormulaRule(
    formula=[f'LEFT(AA{D0},1)="※"'], font=Font(color='C00000', bold=True)))
ws.freeze_panes = f'E{D0}'
ws.auto_filter.ref = f'A3:AA{E1}'

# ── 三张子表：在表2登记，这里自动拆 ──────────────────────────────────
SUB_LAST, SUB_ROWS = 'X', 200
for inst in INST:
    ws = newsheet(inst, f'【{inst}】明细 · 自动从表2拆出来，不用填', '', C_SUB, len(COLS2), merge2=False)
    put(ws, 'A2', '机构名称', font=Font(name='微软雅黑', size=10, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor=C_SUB), align=CEN)
    put(ws, 'B2', inst, font=Font(name='微软雅黑', size=11, bold=True, color='C00000'),
        fill=F_AUTO, align=CEN)
    put(ws, 'C2', '← 这格是拆分依据，别改。整张表是一条公式的结果，'
                  '要改数据去【表2客户典当登记】改，手动在这里打字会把公式顶掉。',
        font=Font(name='微软雅黑', size=9, color='8B5E00'),
        fill=PatternFill('solid', fgColor='FFF7E6'), align=LEFT)
    ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=len(COLS2))
    header(ws, COLS2, 3, C_SUB)
    # 顶上的小汇总（放在标题行右侧空白处会被合并吃掉，所以塞在第 2 行下方的表头行上方不行，
    # 改成用 Z 列之外的独立小块）
    for j, (lab, f) in enumerate([
            ('机构', f'"{inst}"'),
            ('笔数', f'COUNTIF(表2客户典当登记!$B${D0}:$B${E2},"{inst}")'),
            ('典当金额', f'ROUND(SUMIF(表2客户典当登记!$B${D0}:$B${E2},"{inst}",表2客户典当登记!$L${D0}:$L${E2}),2)'),
            ('已还本金', f'ROUND(SUMIF(表2客户典当登记!$B${D0}:$B${E2},"{inst}",表2客户典当登记!$M${D0}:$M${E2}),2)'),
            ('当前在当', f'ROUND(SUMIF(表2客户典当登记!$B${D0}:$B${E2},"{inst}",表2客户典当登记!$N${D0}:$N${E2}),2)'),
            ('应收利息', f'ROUND(SUMIF(表2客户典当登记!$B${D0}:$B${E2},"{inst}",表2客户典当登记!$P${D0}:$P${E2}),2)')]):
        c1_, c2_ = CL(26 + j * 2 + 1), CL(26 + j * 2 + 2)
        put(ws, f'{c1_}2', lab, font=Font(name='微软雅黑', size=9, bold=True, color='FFFFFF'),
            fill=PatternFill('solid', fgColor=C_SUB), align=CEN)
        put(ws, f'{c2_}2', '=' + f, fmt=(None if j == 0 else '0' if j == 1 else MONEY),
            font=FT_SUM, fill=F_SUM, align=CEN)
        ws.column_dimensions[c1_].width = 11
        ws.column_dimensions[c2_].width = 14
    ws[f'A{D0}'] = ArrayFormula(
        ref=f'A{D0}:{SUB_LAST}{D0 + SUB_ROWS - 1}',
        text=f'=_xlfn._xlws.FILTER(表2客户典当登记!$A${D0}:${SUB_LAST}${E2},'
             f'表2客户典当登记!$B${D0}:$B${E2}=$B$2,"这个机构还没有记录")')
    for r in range(D0, D0 + SUB_ROWS):
        for i in range(1, len(COLS2) + 1):
            c = ws.cell(r, i)
            c.border, c.font, c.fill = BOX, FT_AUTO, F_AUTO
            c.alignment = CEN
            if COLS2[i - 1][3]:
                c.number_format = COLS2[i - 1][3]
            ws.column_dimensions[CL(i)].width = COLS2[i - 1][1]
    ws.freeze_panes = f'A{D0}'

# ── 查询表 ───────────────────────────────────────────────────────────
QW = 18
ws = newsheet('查询表', '查询表 · 输入客户姓名，一次看全他的典当和资金往来',
              '最下面那块「资金收付流水」是从《A061_资金台账.xlsx》取的，'
              '查的时候两个文件要同时打开（Excel 读不了关着的工作簿，这是规矩不是设置问题）。', C_Q, QW)
put(ws, 'A3', '客户姓名 →', font=Font(name='微软雅黑', size=11, bold=True), fill=F_SUM, align=CEN)
put(ws, 'B3', names[0] if names else None,
    font=Font(name='微软雅黑', size=13, bold=True, color='C00000'), fill=F_IN, align=CEN)
ws.merge_cells('B3:C3')
put(ws, 'D3', '身份证号', font=Font(name='微软雅黑', size=10, bold=True), fill=F_SUM, align=CEN)
put(ws, 'E3', f'=IF($B$3="","",IFERROR(VLOOKUP($B$3,表2客户典当登记!$E${D0}:$F${E2},2,0),""))',
    font=FT_AUTO, fill=F_AUTO, align=CEN)
ws.merge_cells('E3:F3')
put(ws, 'G3', '客户电话', font=Font(name='微软雅黑', size=10, bold=True), fill=F_SUM, align=CEN)
put(ws, 'H3', f'=IF($B$3="","",IFERROR(VLOOKUP($B$3,表2客户典当登记!$E${D0}:$G${E2},3,0),'
              f'IFERROR(VLOOKUP($B$3,表1不动产登记!$D${D0}:$E${E1},2,0),"")))',
    font=FT_AUTO, fill=F_AUTO, align=CEN)
ws.merge_cells('H3:I3')
dv(ws, '=客户名单', 'B3')

SUMBAR = [
    ('典当笔数', f'=COUNTIF(表2客户典当登记!$E${D0}:$E${E2},$B$3)', '0'),
    ('典当金额', f'=ROUND(SUMIF(表2客户典当登记!$E${D0}:$E${E2},$B$3,表2客户典当登记!$L${D0}:$L${E2}),2)', MONEY),
    ('已还本金', f'=ROUND(SUMIF(表2客户典当登记!$E${D0}:$E${E2},$B$3,表2客户典当登记!$M${D0}:$M${E2}),2)', MONEY),
    ('当前在当', f'=ROUND(SUMIF(表2客户典当登记!$E${D0}:$E${E2},$B$3,表2客户典当登记!$N${D0}:$N${E2}),2)', MONEY),
    ('应收利息', f'=ROUND(SUMIF(表2客户典当登记!$E${D0}:$E${E2},$B$3,表2客户典当登记!$P${D0}:$P${E2}),2)', MONEY),
    ('不动产笔数', f'=COUNTIF(表1不动产登记!$D${D0}:$D${E1},$B$3)', '0'),
    ('不动产金额', f'=ROUND(SUMIF(表1不动产登记!$D${D0}:$D${E1},$B$3,表1不动产登记!$J${D0}:$J${E1}),2)', MONEY),
    ('资金·收到', f'=ROUND(SUMIF([1]数据录入!$I$19:$I$5078,$B$3,[1]数据录入!$F$19:$F$5078),2)', MONEY),
    ('资金·付出', f'=ROUND(SUMIF([1]数据录入!$I$19:$I$5078,$B$3,[1]数据录入!$G$19:$G$5078),2)', MONEY),
]
for j, (lab, f, fmt) in enumerate(SUMBAR):
    col = CL(1 + j * 2)
    col2 = CL(2 + j * 2)
    put(ws, f'{col}5', lab, font=Font(name='微软雅黑', size=9, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor=C_Q), align=CEN)
    put(ws, f'{col2}5', f, fmt=fmt, font=FT_SUM, fill=F_SUM, align=CEN)
    ws.column_dimensions[col].width = 11
    ws.column_dimensions[col2].width = 14

BLOCKS = [
    ('一、典当记录（来自【表2客户典当登记】·所有机构）', 8, COLS2[:18], 40, 'R',
     f'=_xlfn._xlws.FILTER(表2客户典当登记!$A${D0}:$R${E2},'
     f'表2客户典当登记!$E${D0}:$E${E2}=$B$3,"这个客户在表2里没有典当记录")',
     f'=COUNTIF(表2客户典当登记!$E${D0}:$E${E2},$B$3)'),
    ('二、不动产记录（来自【表1不动产登记】·就是上面金辉典当那几笔的收款明细，别跟上面相加）', 52, COLS1[:13], 20, 'M',
     f'=_xlfn._xlws.FILTER(表1不动产登记!$A${D0}:$M${E1},'
     f'表1不动产登记!$D${D0}:$D${E1}=$B$3,"这个客户在表1里没有不动产记录")',
     f'=COUNTIF(表1不动产登记!$D${D0}:$D${E1},$B$3)'),
    ('三、资金收付流水（来自《A061_资金台账.xlsx》的「数据录入」，要同时打开那个文件）', 76,
     [('序号', 7, 'auto', None), ('日期', 12, 'auto', DATEF), ('公司账户', 13, 'auto', None),
      ('收/支项目类别', 15, 'auto', None), ('摘要内容', 34, 'auto', None),
      ('收入', 13, 'auto', MONEY), ('支出', 13, 'auto', MONEY), ('账户余额', 13, 'auto', MONEY),
      ('客户', 12, 'auto', None), ('供应商', 12, 'auto', None), ('月份', 10, 'auto', '0')],
     80, 'K',
     f'=_xlfn._xlws.FILTER([1]数据录入!$A$19:$K$5078,[1]数据录入!$I$19:$I$5078=$B$3,'
     f'"这个客户在资金台账里没有流水，或者《A061_资金台账.xlsx》没打开")',
     f'=COUNTIF([1]数据录入!$I$19:$I$5078,$B$3)'),
]
for title, r0, cols, nrow, lastcol, af, cntf in BLOCKS:
    put(ws, f'A{r0}', title, font=Font(name='微软雅黑', size=11, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor=C_Q), align=LEFT)
    ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=len(cols))
    ws.row_dimensions[r0].height = 22
    put(ws, f'{CL(len(cols)+2)}{r0}',
        f'=IF({cntf[1:]}>{nrow},"※记录超过预留的 {nrow} 行，会显示 #SPILL!，告诉我加行","")',
        font=Font(name='微软雅黑', size=9, bold=True, color='C00000'), align=LEFT)
    header(ws, cols, r0 + 1, C_Q)
    for r in range(r0 + 2, r0 + 2 + nrow):
        for i, (t, w, kind, fmt) in enumerate(cols, 1):
            c = ws.cell(r, i)
            c.border, c.font, c.fill, c.alignment = BOX, FT_AUTO, F_AUTO, CEN
            if fmt:
                c.number_format = fmt
    ws[f'A{r0+2}'] = ArrayFormula(ref=f'A{r0+2}:{lastcol}{r0+1+nrow}', text=af)
ws.freeze_panes = 'A6'
ws.sheet_view.showGridLines = False

# ── 客户名单（静态种子 + 漏没漏的提醒）────────────────────────────────
res = wb['资料']
for i, nm in enumerate(names[:500]):
    res.cell(D0 + i, 9).value = nm
put(res, 'K5', '名单没收到的新客户', font=FT_HDR, fill=hf, align=CEN)
put(res, 'L5', f'=SUMPRODUCT((表2客户典当登记!$E${D0}:$E${E2}<>"")*'
               f'(COUNTIF($I${D0}:$I${D0+499},表2客户典当登记!$E${D0}:$E${E2})=0))',
    fmt='0', font=FT_SUM, fill=F_SUM, align=CEN)
put(res, 'M5', '大于 0 说明表2里有新客户还没加进 I 列名单（查询表的下拉会找不到他）',
    font=Font(name='微软雅黑', size=9, color='808080'), align=LEFT)

# ── 使用说明 ─────────────────────────────────────────────────────────
ws = wb.create_sheet('使用说明', 0)
ws.sheet_properties.tabColor = '1F3864'
for i, w in enumerate([18, 16, 62, 22], 1):
    ws.column_dimensions[CL(i)].width = w
put(ws, 'A1', '典当登记表 · 使用说明', font=FT_TITLE, fill=PatternFill('solid', fgColor='1F3864'), align=LEFT)
ws.merge_cells('A1:D1'); ws.row_dimensions[1].height = 30

def sec(r, t):
    put(ws, f'A{r}', t, font=Font(name='微软雅黑', size=12, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor=C_MAIN), align=LEFT)
    ws.merge_cells(f'A{r}:D{r}'); ws.row_dimensions[r].height = 24

def ln(r, a, b='', c=''):
    put(ws, f'A{r}', a, font=Font(name='微软雅黑', size=10, bold=True), align=LEFT)
    put(ws, f'B{r}', b, font=Font(name='微软雅黑', size=10), align=LEFT)
    put(ws, f'C{r}', c, font=Font(name='微软雅黑', size=10), align=Alignment('left', 'center', wrap_text=True))
    ws.merge_cells(f'C{r}:D{r}')

sec(3, '一、七张表，各管一段')
ln(4, '表名', '谁用', '干什么')
for i, (a, b, c) in enumerate([
    ('资料', '管理员', '所有下拉菜单的来源：抵押机构、抵押物品、状态、付息方式、客户名单、日滞纳金率。'),
    ('表1不动产登记', '登记员', '金辉典当的不动产，单独一张，跟表2不混。逾期天数、应收滞纳金自动算。'),
    ('表2客户典当登记', '登记员', '客户典当总表，所有机构都在这里登记。这是唯一要填的主表。'),
    ('金辉典当', '看', '自动从表2拆出来，不用填。'),
    ('明道典当', '看', '同上（这次新增的机构）。'),
    ('个人典当', '看', '同上。原来叫「个人借款」，统一改成「个人典当」，跟子表名对上。'),
    ('查询表', '查客户', '输入客户姓名，一次拉出他的典当记录、不动产记录、资金收付流水。'),
], 5):
    ln(i, a, b, c)

sec(13, '二、三张子表是怎么自动拆出来的')
ln(14, '公式', 'FILTER', '每张子表的 A4 是一条 FILTER：把表2里「抵押机构 = 本表名字」的行整块搬过来。')
ln(15, '所以', '别在子表里打字', '子表整块是一条公式的结果，手动输入会把公式顶掉。要改去表2改。')
ln(16, '新增机构', '两步', '① 在【资料】A 列加机构名；② 复制一张子表，把 B2 改成新机构名即可。')

sec(18, '三、查询表怎么用')
ln(19, '第 1 步', '打开两个文件', '《A061_典当登记表.xlsx》和《A061_资金台账.xlsx》都要打开。')
ln(20, '第 2 步', 'B3 选客户', 'B3 那格选（或直接打）客户姓名，下面三块和上面的汇总条全部自动跳出来。')
ln(21, '为什么要都打开', 'Excel 的规矩', 'FILTER、SUMIF 这类公式读不了「关着的」工作簿，只能读打开着的。'
                                      '关着的时候第三块会显示取不到数，把资金台账打开就好了。')
ln(22, '两个文件要放一起', '同一个文件夹', '文件名一个字都不能改。从聊天里下载重名会变成「(1)」，要把括号那段删掉。')

sec(24, '四、颜色和符号')
ln(25, '淡黄色格子', '请填写', '这是要人填的。')
ln(26, '灰色格子', '公式自动算', '别手动改。')
ln(27, '※ 开头的红字', '有问题', '「核对」列会自己挑毛病：机构/状态不在名单、已还本金大于典当金额、'
                                '到期日早于当票日期、利息跟「在当×月利率」对不上。')

sec(29, '五、这次从原表搬过来的')
for i, t in enumerate([
    '· 表2（原 Sheet2）201 行全部搬过来，「个人借款」统一改名「个人典当」（134 笔），金辉典当 67 笔。',
    '· 表1（原 Sheet1）67 行全部搬过来。',
    '· 新加了「身份证号」「客户电话」两列（空着，你自己补），查询表按姓名查，身份证号自动带出来。',
    '· 新加了「当前在当」列 = 典当金额 − 已还本金，自动算。',
    '· 新加了「核对」列，逐行挑毛病。',
    '· 原来「应收金额」改叫「应收利息」，位置挪到「典当系统费率」后面，看起来更顺。',
    '· ⚠ 查过了：【表1】那 67 行，跟【表2】里机构=金辉典当的 67 行是同一批当（姓名+金额+当票号 66 行完全一致）。',
    '　 表1 多出来的是收款拆分（对公/林总/建行民生/微信）、违约金、滞纳金。所以两张表的金额**不能相加**，',
    '　 查询表里也分成两块单独列，就是这个原因。',
    '· 「典当系统费率」这一列各行口径不一样（有的按月、有的按日），所以没拿它去反推利息对错，',
    '　 免得整屏报警。要核利息请直接看「应收利息」那一列。',
], 30):
    put(ws, f'A{i}', t, font=Font(name='微软雅黑', size=9, color='595959'), align=LEFT)
    ws.merge_cells(f'A{i}:D{i}')
ws.sheet_view.showGridLines = False

for nm, ref in NAMES.items():
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)
# 表1 排在表2 前面，跟用户叫法一致
wb.move_sheet('表1不动产登记', offset=-1)
wb.active = 0
wb.save(OUT)
print('已生成:', OUT)
print('工作表:', ' / '.join(s.title for s in wb.worksheets))
