# -*- coding: utf-8 -*-
"""生成《资金日报表·混合录入版.xlsx》—— 多账户资金日报（单币种：人民币）。

一张【数据录入】混合录入，其余全部自动：
  · 账户子表按账户名自动拆（FILTER 动态数组）
  · 原来那三块汇总界面（资金日报统计表 / 货币资金日报表 / 多帐户资金汇总表）原样保留，
    但账户清单改成从【基础资料】自动长，加账户不用再改三个地方
  · 新增【月度汇报表】和【收支分类报表】，按外汇资金台账那套格式做：
    顶上一条「收入合计 / 支出合计 / 内部转账净额 / 经营净额」，月度那张分出内部转入和内部转出
  · 内部转账（账户之间调头寸）单列一类，不计入收入/支出合计 ——
    原模板没有这个概念，账户互转会被同时算成收入和支出，两头虚增

跑法：python3 build_cash.py
"""
import os, sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, '资金日报表_混合录入版.xlsx')

YEAR = 2026
N_ACC, N_ROW, N_CAT, N_CUS, N_SUP, N_PRJ = 30, 5000, 40, 200, 200, 50
N_ITEM, SUB = 40, 600            # 收支报表每侧的项目行数 / 子表预留行数
B0, R0 = 5, 5                    # 基础资料各清单起始行 / 数据录入起始行
AE = B0 + N_ACC - 1
RE_ = R0 + N_ROW - 1
HDR = 4                          # 基础资料 / 数据录入 的表头行

C_MAIN, C_SUB, C_RPT, C_BASE, C_CHK = '1F4E79', '806000', '375623', '7030A0', 'C00000'
F_IN = PatternFill('solid', fgColor='FFFBEA')
F_AUTO = PatternFill('solid', fgColor='F2F2F2')
F_SUM = PatternFill('solid', fgColor='FFF2CC')
FT_IN = Font(name='微软雅黑', size=10)
FT_AUTO = Font(name='微软雅黑', size=10, color='7F7F7F')
FT_HDR = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
FT_TITLE = Font(name='微软雅黑', size=15, bold=True, color='FFFFFF')
FT_SUM = Font(name='微软雅黑', size=10, bold=True, color='C00000')
FT_TIP = Font(name='微软雅黑', size=9, color='8B5E00')
THIN = Side('thin', color='BFBFBF')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CEN = Alignment('center', 'center', wrap_text=True)
LEFT = Alignment('left', 'center')
MONEY, NUM0, DATEF, PCT, YM = '#,##0.00', '#,##0', 'yyyy-mm-dd', '0.0%', '000000'
DATEQ = 'yyyy-mm-dd;;;@'         # 自动区日期：0 不显示，免得画成 1900-01-00

# ══════════════════════════════════════════════════════════════════════
# 原模板里的资料，一条不少地搬过来（账户 14 个、支出类别 12 项、客户 13、供应商 13）
# ══════════════════════════════════════════════════════════════════════
ACCOUNTS = [('现金', '现金', 5000), ('微信', '微信', 6000), ('工行', '银行', 7000),
            ('民生', '银行', 8000), ('建行', '银行', 9000), ('农行', '银行', 10000),
            ('中行', '银行', 11000), ('支付宝', '第三方', 12000), ('农商行', '银行', 13000),
            ('备用', '备用', 14000), ('备用1', '备用', 15000), ('备用2', '备用', 16000),
            ('备用3', '备用', 17000), ('备用4', '备用', 18000)]

# 类别 → 属性。原模板只有「支出业务类别」，收入一侧一个类别都没有，
# 所以收入只能按客户分析；这里把收入类别补上，并单列「内部转账」。
CATS = ([('销售回款', '收入'), ('其他业务收入', '收入'), ('投资款', '收入'), ('借款', '收入'),
         ('利息收入', '收入'), ('退款收回', '收入'), ('政府补助', '收入')]
        + [(x, '支出') for x in ['供应商', '办公费', '水费', '电费', '伙食费', '工资', '租赁费',
                                 '差旅费', '业务招待费', '车辆费用', '银行手续费', '税费',
                                 '还款', '资产购买', '利息支出']]
        + [('内部转账', '内部转账'), ('备用金调拨', '内部转账')])

CUSTOMERS = ['客户%d' % i for i in range(1, 14)]
SUPPLIERS = ['供应商%d' % i for i in range(1, 14)]
PROJECTS = ['主营业务', '其他业务', '行政', '融资', '内部调拨']

# 原模板【数据录入】第 19~41 行那 23 笔流水，一笔不少地搬过来，
# 再补 2 笔内部转账（工行 → 现金），好让新加的那几块有东西可看。
FLOWS = [
    ('2026-06-02', '现金', '销售回款', '', 2000, None, '客户1', '', '主营业务', ''),
    ('2026-06-03', '微信', '办公费', '', None, 200, '', '', '行政', ''),
    ('2026-06-05', '工行', '电费', '', None, 300, '', '', '行政', ''),
    ('2026-06-06', '民生', '差旅费', '', None, 500, '', '', '行政', ''),
    ('2026-06-10', '建行', '销售回款', '', 2600, None, '客户2', '', '主营业务', ''),
    ('2026-06-12', '农行', '供应商', '', None, 5000, '', '供应商1', '主营业务', ''),
    ('2026-06-20', '中行', '伙食费', '', None, 200, '', '', '行政', ''),
    ('2026-06-23', '支付宝', '办公费', '', None, 10, '', '', '行政', ''),
    ('2026-06-24', '农商行', '车辆费用', '', None, 50, '', '', '行政', ''),
    ('2026-06-25', '备用', '销售回款', '', 2100, None, '客户3', '', '主营业务', ''),
    ('2026-06-25', '备用1', '租赁费', '', None, 2000, '', '', '行政', ''),
    ('2026-06-25', '现金', '业务招待费', '', None, 200, '', '', '行政', ''),
    ('2026-06-26', '备用2', '销售回款', '', 3000, None, '客户4', '', '主营业务', ''),
    ('2026-06-26', '备用3', '工资', '', None, 8000, '', '', '行政', ''),
    ('2026-06-27', '备用4', '销售回款', '', 1500, None, '客户5', '', '主营业务', ''),
    ('2026-06-28', '工行', '供应商', '', None, 3000, '', '供应商2', '主营业务', ''),
    ('2026-06-28', '微信', '销售回款', '', 800, None, '客户6', '', '主营业务', ''),
    ('2026-06-29', '民生', '水费', '', None, 120, '', '', '行政', ''),
    ('2026-06-29', '建行', '银行手续费', '', None, 30, '', '', '行政', ''),
    ('2026-06-30', '农行', '销售回款', '', 4200, None, '客户7', '', '主营业务', ''),
    ('2026-06-30', '中行', '供应商', '', None, 2500, '', '供应商3', '主营业务', ''),
    ('2026-06-30', '支付宝', '销售回款', '', 600, None, '客户8', '', '主营业务', ''),
    ('2026-06-30', '农商行', '办公费', '', None, 80, '', '', '行政', ''),
    ('2026-06-30', '工行', '内部转账', '工行转现金备用', None, 2000, '', '', '内部调拨', '现金'),
    ('2026-06-30', '现金', '内部转账', '收到工行转入', 2000, None, '', '', '内部调拨', '工行'),
]

# ══════════════════════════════════════════════════════════════════════
wb = openpyxl.Workbook()
wb.remove(wb.active)
NAMES = {}


def newsheet(name, title, tip, color, width, merge2=True):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = color
    ws['A1'] = title
    ws['A1'].font, ws['A1'].fill, ws['A1'].alignment = FT_TITLE, PatternFill('solid', fgColor='1F3864'), LEFT
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    ws.row_dimensions[2].height = 24
    if merge2:
        ws['A2'] = tip
        ws['A2'].font, ws['A2'].fill = FT_TIP, PatternFill('solid', fgColor='FFF7E6')
        ws['A2'].alignment = LEFT
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
    ws.sheet_view.showGridLines = False
    return ws


def put(ws, addr, v, fmt=None, font=None, fill=None, align=None, border=True):
    c = ws[addr]
    c.value = v
    if fmt: c.number_format = fmt
    c.font = font or FT_IN
    if fill: c.fill = fill
    c.alignment = align or CEN
    if border: c.border = BOX
    return c


def hdr(ws, row, col0, cols, color, h=30):
    fill = PatternFill('solid', fgColor=color)
    for i, (t, w, fmt) in enumerate(cols, col0):
        c = ws.cell(row, i, t)
        c.font, c.fill, c.alignment, c.border = FT_HDR, fill, CEN, BOX
        if w: ws.column_dimensions[CL(i)].width = w
    ws.row_dimensions[row].height = h


def band(ws, r0, r1, col0, cols, auto=False):
    for r in range(r0, r1 + 1):
        for i, (t, w, fmt) in enumerate(cols, col0):
            c = ws.cell(r, i)
            c.border = BOX
            c.font = FT_AUTO if auto else FT_IN
            c.fill = F_AUTO if auto else F_IN
            c.alignment = CEN
            if fmt: c.number_format = fmt


def bar(ws, row, col0, ncol, text, color=None, h=22):
    put(ws, f'{CL(col0)}{row}', text,
        font=Font(name='微软雅黑', size=11, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor=color or C_RPT), align=LEFT)
    ws.merge_cells(start_row=row, start_column=col0, end_row=row, end_column=col0 + ncol - 1)
    ws.row_dimensions[row].height = h


def dv(ws, f1, rng, stop=False):
    d = DataValidation(type='list', formula1=f1, allow_blank=True, showErrorMessage=True,
                       errorStyle='stop' if stop else 'warning', errorTitle='不在清单里',
                       error='先去【基础资料】把它加上，或者选一个已有的。')
    ws.add_data_validation(d)
    d.add(rng)


def redflag(ws, rng):
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'LEFT({rng.split(":")[0]},1)="※"'], font=Font(color='C00000', bold=True),
        fill=PatternFill('solid', fgColor='FFD7D7')))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'LEFT({rng.split(":")[0]},1)="△"'], font=Font(color='8B5E00')))


def nb(rng):
    """FILTER 取数区先过一道：源表空格子会被当成 0 带回来，日期列就成了 1900-01-00。"""
    return f'IF({rng}="","",{rng})'


# ══════════════════════════════════════════════════════════════════════
# 1  基础资料
# ══════════════════════════════════════════════════════════════════════
ws = newsheet('基础资料', '基础资料 · 账户、收支类别、客户、供应商都在这里维护',
              '淡黄色格子是要你填的，灰色是公式算的。账户名称必须唯一。'
              '「收支类别」比原表多一列「属性」：收入 / 支出 / 内部转账 —— '
              '属性填「内部转账」的，不计入收入合计和支出合计，账户之间调头寸就靠它。', C_BASE, 16)

ACOL = [('序号', 6, NUM0), ('账户名称 ★（唯一）', 16, None), ('账户类型', 10, None),
        ('期初余额', 14, MONEY), ('启用日期', 12, DATEF), ('核对', 26, None)]
bar(ws, 3, 1, len(ACOL), '一、账户档案　——　加一个账户：这里接着写一行，再复制一张账户子表改 B2', C_MAIN)
hdr(ws, HDR, 1, ACOL, C_MAIN)
band(ws, B0, AE, 1, ACOL)
for r in range(B0, AE + 1):
    for c in (1, 6):
        cc = ws.cell(r, c); cc.font, cc.fill = FT_AUTO, F_AUTO
    g = f'IF($B{r}="","",'
    put(ws, f'A{r}', f'={g}COUNTA($B${B0}:$B{r}))', fmt=NUM0, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'F{r}', f'={g}'
        f'IF(COUNTIF($B${B0}:$B${AE},$B{r})>1,"※账户名称重复了，后面按名字取数会串",'
        f'IF($D{r}="","※期初余额没填（没有就填 0）",'
        f'IF(AND($E{r}<>"",COUNTIFS(流水账户,$B{r},流水日期,"<"&$E{r})>0),'
        f'"※有流水的日期早于启用日期，期初余额会重复计算",""))))',
        font=FT_AUTO, fill=F_AUTO, align=LEFT)
for i, (nm, ty, op) in enumerate(ACCOUNTS):
    r = B0 + i
    ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value = nm, ty, op
    ws.cell(r, 5).value = datetime.datetime(YEAR, 6, 1)
redflag(ws, f'F{B0}:F{AE}')


def list_block(col0, cols, n, title, seed, color=C_BASE):
    bar(ws, 3, col0, len(cols), title, color)
    hdr(ws, HDR, col0, cols, color)
    band(ws, B0, B0 + n - 1, col0, cols)
    for i, row in enumerate(seed):
        for j, v in enumerate(row if isinstance(row, (list, tuple)) else [row]):
            ws.cell(B0 + i, col0 + j).value = v


list_block(8, [('收/支类别 ★', 15, None), ('属性 ★', 11, None)], N_CAT,
           '二、收支类别（属性 = 内部转账的不计入收支合计）', CATS, C_MAIN)
list_block(11, [('客户', 18, None)], N_CUS, '三、客户', CUSTOMERS)
list_block(13, [('供应商', 18, None)], N_SUP, '四、供应商', SUPPLIERS)
list_block(15, [('项目 / 部门', 15, None)], N_PRJ, '五、项目 / 部门', PROJECTS)
dv(ws, '"收入,支出,内部转账"', f'I{B0}:I{B0 + N_CAT - 1}', stop=True)
ws.column_dimensions['G'].width = 2.2
ws.column_dimensions['J'].width = 2.2
ws.column_dimensions['L'].width = 2.2
ws.column_dimensions['N'].width = 2.2
ws.freeze_panes = f'A{B0}'

NAMES.update({
    '账户表': f'基础资料!$B${B0}:$B${AE}', '账户类型表': f'基础资料!$C${B0}:$C${AE}',
    '账户期初': f'基础资料!$D${B0}:$D${AE}', '账户启用': f'基础资料!$E${B0}:$E${AE}',
    '类别表': f'基础资料!$H${B0}:$H${B0 + N_CAT - 1}',
    '类别属性': f'基础资料!$I${B0}:$I${B0 + N_CAT - 1}',
    '客户表': f'基础资料!$K${B0}:$K${B0 + N_CUS - 1}',
    '供应商表': f'基础资料!$M${B0}:$M${B0 + N_SUP - 1}',
    '项目表': f'基础资料!$O${B0}:$O${B0 + N_PRJ - 1}',
})

# ══════════════════════════════════════════════════════════════════════
# 2  数据录入（全套只有这一张要填）
# ══════════════════════════════════════════════════════════════════════
# A~K 跟原模板一模一样，新增的列一律往 L 后面接，原来的用法一点不变。
DCOL = [('序号', 6, NUM0), ('日期 ★', 12, DATEF), ('公司账户 ★', 13, None),
        ('收/支类别 ★', 13, None), ('摘要内容', 26, None),
        ('收入', 13, MONEY), ('支出', 13, MONEY), ('账户余额', 14, MONEY),
        ('客户', 14, None), ('供应商', 14, None), ('月份', 9, YM),
        ('类别属性', 10, None), ('对方账户\n(内部转账填)', 13, None), ('项目/部门', 12, None),
        ('备注', 16, None), ('核对', 32, None)]
AUTO_D = {1, 8, 11, 12, 16}

ws = newsheet('数据录入', '数据录入 · 全套表只有这一张要填',
              '选账户 → 余额自动结；收入和支出分两列，只填一个。'
              '账户之间调头寸：转出账户记一笔支出、转入账户记一笔收入，类别都选「内部转账」，'
              '在「对方账户」写上另一头 —— 这样它就不会被算进收入和支出合计。'
              '期初余额挪到【基础资料】了，这里只放真实流水。最后一列「核对」会自己挑毛病。', C_MAIN, len(DCOL))
# 顶上一条汇总。标签短一点，格子窄，写长了会被切掉。
put(ws, 'A3', '收入总额', font=Font(name='微软雅黑', size=9, bold=True), fill=F_SUM)
# 用整列 $F:$F 会把第 3 行自己算进去 —— 经营净额又反过来引用 B3，转成死循环。
# 一律用限定到数据区的名称（流水收入 = 数据录入!$F$5:$F$5004），第 3 行在范围外。
put(ws, 'B3', '=ROUND(SUMIFS(流水收入,流水属性,"<>内部转账"),2)', fmt=MONEY, font=FT_SUM, fill=F_AUTO)
put(ws, 'C3', '支出总额', font=Font(name='微软雅黑', size=9, bold=True), fill=F_SUM)
put(ws, 'D3', '=ROUND(SUMIFS(流水支出,流水属性,"<>内部转账"),2)', fmt=MONEY, font=FT_SUM, fill=F_AUTO)
put(ws, 'E3', '经营净额', font=Font(name='微软雅黑', size=9, bold=True), fill=F_SUM)
put(ws, 'F3', '=ROUND($B$3-$D$3,2)', fmt=MONEY, font=FT_SUM, fill=F_AUTO)
put(ws, 'G3', '内部转账净额', font=Font(name='微软雅黑', size=9, bold=True), fill=F_SUM)
put(ws, 'H3', '=ROUND(SUMIFS(流水收入,流水属性,"内部转账")'
              '-SUMIFS(流水支出,流水属性,"内部转账"),2)',
    fmt=MONEY, font=FT_SUM, fill=F_AUTO)
put(ws, 'I3', '资金总额', font=Font(name='微软雅黑', size=9, bold=True), fill=F_SUM)
put(ws, 'J3', '=ROUND(SUM(账户期初)+SUM(流水收入)-SUM(流水支出),2)',
    fmt=MONEY, font=FT_SUM, fill=F_AUTO)
put(ws, 'K3', '已用行数', font=Font(name='微软雅黑', size=9, bold=True), fill=F_SUM)
put(ws, 'L3', f'=COUNTA($C${R0}:$C${RE_})&" / {N_ROW}"', font=FT_SUM, fill=F_AUTO)
put(ws, 'M3', '收入总额、支出总额都不含内部转账；内部转账净额应该是 0，不为 0 就是有一头没录',
    font=FT_TIP, align=LEFT)
ws.merge_cells(start_row=3, start_column=13, end_row=3, end_column=len(DCOL))
ws.row_dimensions[3].height = 20
hdr(ws, HDR, 1, DCOL, C_MAIN)
band(ws, R0, RE_, 1, DCOL)
for r in range(R0, RE_ + 1):
    for c in AUTO_D:
        cc = ws.cell(r, c); cc.font, cc.fill = FT_AUTO, F_AUTO
    for c in (5, 15, 16):
        ws.cell(r, c).alignment = LEFT
    g = f'IF($C{r}="","",'
    put(ws, f'A{r}', f'={g}COUNTA($C${R0}:$C{r}))', fmt=NUM0, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'H{r}', f'={g}ROUND(SUMIF(账户表,$C{r},账户期初)'
                     f'+SUMIFS($F${R0}:F{r},$C${R0}:C{r},$C{r})'
                     f'-SUMIFS($G${R0}:G{r},$C${R0}:C{r},$C{r}),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'K{r}', f'=IF($B{r}="","",--TEXT($B{r},"yyyymm"))', fmt=YM, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'L{r}', f'=IF($D{r}="","",IFERROR(INDEX(类别属性,MATCH($D{r},类别表,0)),""))',
        font=FT_AUTO, fill=F_AUTO)
    put(ws, f'P{r}', f'={g}'
        f'IF(ISNA(MATCH($C{r},账户表,0)),"※这个账户不在【基础资料】的账户档案里",'
        f'IF($B{r}="","※没填日期，这一笔不会进任何月份的报表",'
        f'IF(AND($F{r}<>"",$G{r}<>""),"※收入和支出不能同时填",'
        f'IF(AND(N($F{r})=0,N($G{r})=0),"※收入和支出都没填",'
        f'IF(OR(N($F{r})<0,N($G{r})<0),"※金额不能填负数，方向反了就换一列填",'
        f'IF($D{r}="","※没填收支类别，分类报表会漏掉这一笔",'
        f'IF(ISNA(MATCH($D{r},类别表,0)),"※收支类别不在【基础资料】的清单里",'
        f'IF(AND($L{r}="内部转账",$M{r}=""),"△内部转账要在「对方账户」写上另一头",'
        f'IF(AND($L{r}="内部转账",$M{r}=$C{r}),"※对方账户跟本账户是同一个",'
        f'IF(AND($L{r}<>"内部转账",$M{r}<>""),"△填了对方账户，但类别不是内部转账",'
        f'IF($H{r}<0,"△这一笔之后该账户余额成负数了，核一下",""))))))))))))',
        font=FT_AUTO, fill=F_AUTO, align=LEFT)
for i, f in enumerate(FLOWS):
    r = R0 + i
    d = datetime.datetime.strptime(f[0], '%Y-%m-%d')
    for c, v in ((2, d), (3, f[1]), (4, f[2]), (5, f[3] or None), (6, f[4]), (7, f[5]),
                 (9, f[6] or None), (10, f[7] or None), (13, f[9] or None), (14, f[8] or None)):
        ws.cell(r, c).value = v
dv(ws, '=账户表', f'C{R0}:C{RE_}', stop=True)
dv(ws, '=类别表', f'D{R0}:D{RE_}', stop=True)
dv(ws, '=客户表', f'I{R0}:I{RE_}')
dv(ws, '=供应商表', f'J{R0}:J{RE_}')
dv(ws, '=账户表', f'M{R0}:M{RE_}')
dv(ws, '=项目表', f'N{R0}:N{RE_}')
redflag(ws, f'P{R0}:P{RE_}')
ws.conditional_formatting.add(f'H{R0}:H{RE_}', FormulaRule(
    formula=[f'N(H{R0})<0'], font=Font(color='C00000', bold=True)))
# 录了数就自动出边框（跟资金台账一个做法）
ws.conditional_formatting.add(f'A{R0}:P{RE_}', FormulaRule(
    formula=[f'$C{R0}<>""'],
    border=Border(left=Side('thin', color='808080'), right=Side('thin', color='808080'),
                  top=Side('thin', color='808080'), bottom=Side('thin', color='808080'))))
ws.freeze_panes = f'C{R0}'
ws.auto_filter.ref = f'A{HDR}:P{RE_}'
NAMES.update({
    '流水日期': f'数据录入!$B${R0}:$B${RE_}', '流水账户': f'数据录入!$C${R0}:$C${RE_}',
    '流水类别': f'数据录入!$D${R0}:$D${RE_}', '流水收入': f'数据录入!$F${R0}:$F${RE_}',
    '流水支出': f'数据录入!$G${R0}:$G${RE_}', '流水月份': f'数据录入!$K${R0}:$K${RE_}',
    '流水属性': f'数据录入!$L${R0}:$L${RE_}', '流水对方': f'数据录入!$M${R0}:$M${RE_}',
    '流水客户': f'数据录入!$I${R0}:$I${RE_}', '流水供应商': f'数据录入!$J${R0}:$J${RE_}',
    '流水项目': f'数据录入!$N${R0}:$N${RE_}', '流水核对': f'数据录入!$P${R0}:$P${RE_}',
})


def acc(i):
    """账户档案第 i 个账户那一格的**直接地址**。

    不用 INDEX(账户表,i)：INDEX 落到空格子返回的是数字 0 而不是空文本，
    外面 IF(...="","",...) 判不出来，没用上的那几十行会跟着算出一堆 0。"""
    return f'基础资料!$B${B0 + i - 1}'


# ══════════════════════════════════════════════════════════════════════
# 3  汇报表（原来那三块，界面原样保留；账户清单改成自动长）
# ══════════════════════════════════════════════════════════════════════
NDAY = 31
LASTD = CL(2 + NDAY * 2)          # BL
C_RCV, C_PAY, C_END = CL(3 + NDAY * 2), CL(4 + NDAY * 2), CL(5 + NDAY * 2)   # BM BN BO
A8, A_E = 8, 8 + N_ACC - 1
ATOT = A_E + 1
ws = newsheet('汇报表', '汇报表 · 三块老界面原样保留（账户清单改成自动长，加账户不用再改三处）',
              '① 资金日报统计表：改上面两个日期就行，最多 31 天。这一块是「账户口径」 ——'
              '收款/付款按钱进出账户算，内部划转也算在里面（因为账户余额确实动了）。'
              '② 货币资金日报表：指定某一天。③ 多帐户资金汇总表：按月，'
              '这一块是「经营口径」，内部转入/转出单独两列，不混进本月收入和本月支出。',
              C_RPT, 5 + NDAY * 2)
for lab, addr, v in (('起始日期', 'B3', datetime.datetime(YEAR, 6, 1)),
                     ('截止日期', 'E3', datetime.datetime(YEAR, 6, 30))):
    put(ws, f'{CL(ord(addr[0]) - 65)}3', lab, font=Font(name='微软雅黑', size=10, bold=True), fill=F_SUM)
    put(ws, addr, v, fmt=DATEF, font=Font(name='微软雅黑', size=11, bold=True, color='C00000'), fill=F_IN)
put(ws, 'G3', '查询天数', font=Font(name='微软雅黑', size=10, bold=True), fill=F_SUM)
put(ws, 'H3', '=IF(OR($B$3="",$E$3=""),"",IF($E$3<$B$3,"※截止早于起始",'
              'IF($E$3-$B$3+1>31,"※超过 31 天，请分段查",$E$3-$B$3+1)))',
    font=FT_SUM, fill=F_AUTO)
DOK = 'ISNUMBER($H$3)'
for j, (lab, f, col) in enumerate([
        ('公司期初余额', f'=ROUND(SUM(B{A8}:B{A_E}),2)', 'B'),
        ('期间收款合计', f'=ROUND(SUM({C_RCV}{A8}:{C_RCV}{A_E}),2)', 'D'),
        ('期间付款合计', f'=ROUND(SUM({C_PAY}{A8}:{C_PAY}{A_E}),2)', 'F'),
        ('期末结余金额', f'=ROUND(SUM({C_END}{A8}:{C_END}{A_E}),2)', 'H')]):
    lc = CL(ord(col) - 65)
    put(ws, f'{lc}4', lab, font=Font(name='微软雅黑', size=10, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor=C_RPT))
    put(ws, f'{col}5', f, fmt=MONEY, font=FT_SUM, fill=F_SUM)
    ws.column_dimensions[lc].width = 14
    ws.column_dimensions[col].width = 15
bar(ws, 6, 1, 2, '', C_RPT, h=18)
put(ws, 'A6', '① 资金日报统计表（账户口径，含内部划转）',
    font=Font(name='微软雅黑', size=10, bold=True, color='FFFFFF'),
    fill=PatternFill('solid', fgColor=C_RPT), align=LEFT)
ws.merge_cells('A6:B6')
for d in range(NDAY):
    for k in range(2):
        c = CL(3 + d * 2 + k)
        put(ws, f'{c}6', f'=IF(NOT({DOK}),"",IF({d + 1}<=MIN($H$3,{NDAY}),$B$3+{d},""))',
            fmt=DATEQ, font=FT_AUTO, fill=F_AUTO)
        put(ws, f'{c}7', f'=IF({c}$6="","","{"收款" if k == 0 else "付款"}")',
            font=FT_HDR, fill=PatternFill('solid', fgColor=C_RPT))
        ws.column_dimensions[c].width = 12
for c, t, w in (('A', '账户', 14), ('B', '期初余额', 14), (C_RCV, '期间收款', 14),
                (C_PAY, '期间付款', 14), (C_END, '期末余额', 14)):
    put(ws, f'{c}7', t, font=FT_HDR, fill=PatternFill('solid', fgColor=C_RPT))
    ws.column_dimensions[c].width = w
ws.row_dimensions[7].height = 28
for i in range(1, N_ACC + 1):
    r = A8 + i - 1
    g = f'IF(OR($A{r}="",NOT({DOK})),"",'
    put(ws, f'A{r}', f'=IF({acc(i)}="","",{acc(i)})', font=FT_AUTO, fill=F_AUTO)
    put(ws, f'B{r}', f'={g}ROUND(SUMIF(账户表,$A{r},账户期初)'
                     f'+SUMIFS(流水收入,流水账户,$A{r},流水日期,"<"&$B$3)'
                     f'-SUMIFS(流水支出,流水账户,$A{r},流水日期,"<"&$B$3),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    for d in range(NDAY):
        for k in range(2):
            c = CL(3 + d * 2 + k)
            src = '流水收入' if k == 0 else '流水支出'
            put(ws, f'{c}{r}', f'=IF(OR($A{r}="",{c}$6=""),"",'
                               f'ROUND(SUMIFS({src},流水账户,$A{r},流水日期,{c}$6),2))',
                fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'{C_RCV}{r}', f'=IF($A{r}="","",ROUND(SUMIF($C$7:${LASTD}$7,"收款",C{r}:{LASTD}{r}),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'{C_PAY}{r}', f'=IF($A{r}="","",ROUND(SUMIF($C$7:${LASTD}$7,"付款",C{r}:{LASTD}{r}),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'{C_END}{r}', f'=IF($A{r}="","",ROUND(N($B{r})+N({C_RCV}{r})-N({C_PAY}{r}),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
put(ws, f'A{ATOT}', '合计', font=FT_HDR, fill=PatternFill('solid', fgColor=C_RPT))
for c in ['B'] + [CL(3 + x) for x in range(NDAY * 2)] + [C_RCV, C_PAY, C_END]:
    put(ws, f'{c}{ATOT}', f'=ROUND(SUM({c}{A8}:{c}{A_E}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)

# ── ② 货币资金日报表 / ③ 多帐户资金汇总表（月度）──────────────────────
D2R = ATOT + 2
DAY_COL = [('序号', 6, NUM0), ('账户', 14, None), ('昨日余额', 14, MONEY),
           ('本日收入', 14, MONEY), ('本日支出', 14, MONEY), ('本日余额', 14, MONEY)]
MON_COL = [('序号', 6, NUM0), ('账号名称', 14, None), ('期初余额', 14, MONEY),
           ('本月收入', 14, MONEY), ('本月支出', 14, MONEY), ('内部转入', 13, MONEY),
           ('内部转出', 13, MONEY), ('结余', 14, MONEY)]
bar(ws, D2R, 1, len(DAY_COL), '② 货币资金日报表（账户口径，含内部划转）', C_MAIN)
bar(ws, D2R, 8, len(MON_COL), '③ 多帐户资金汇总表·月度（经营口径，内部转入/转出单列）', C_SUB)
put(ws, f'A{D2R+1}', '日期：', font=Font(name='微软雅黑', size=10, bold=True), fill=F_SUM)
put(ws, f'B{D2R+1}', '=IF($E$3="","",$E$3)', fmt=DATEF,
    font=Font(name='微软雅黑', size=11, bold=True, color='C00000'), fill=F_AUTO)
put(ws, f'C{D2R+1}', '默认取①的截止日期；要看别的天，直接把这格改成那一天。',
    font=FT_TIP, align=LEFT)
ws.merge_cells(start_row=D2R + 1, start_column=3, end_row=D2R + 1, end_column=6)
put(ws, f'H{D2R+1}', '月份：', font=Font(name='微软雅黑', size=10, bold=True), fill=F_SUM)
put(ws, f'I{D2R+1}', datetime.datetime(YEAR, 6, 1), fmt='yyyy年m月',
    font=Font(name='微软雅黑', size=11, bold=True, color='C00000'), fill=F_IN)
put(ws, f'J{D2R+1}', '填该月 1 号。内部转入/转出是账户之间调头寸，不算收入也不算支出。',
    font=FT_TIP, align=LEFT)
ws.merge_cells(start_row=D2R + 1, start_column=10, end_row=D2R + 1, end_column=15)
hdr(ws, D2R + 2, 1, DAY_COL, C_MAIN)
hdr(ws, D2R + 2, 8, MON_COL, C_SUB)
band(ws, D2R + 3, D2R + 2 + N_ACC, 1, DAY_COL, auto=True)
band(ws, D2R + 3, D2R + 2 + N_ACC, 8, MON_COL, auto=True)
DAY, MS, ME = f'$B${D2R+1}', f'$I${D2R+1}', f'EOMONTH($I${D2R+1},0)'
for i in range(1, N_ACC + 1):
    r = D2R + 2 + i
    put(ws, f'B{r}', f'=IF({acc(i)}="","",{acc(i)})', font=FT_AUTO, fill=F_AUTO)
    put(ws, f'A{r}', f'=IF($B{r}="","",{i})', fmt=NUM0, font=FT_AUTO, fill=F_AUTO)
    g2 = f'IF(OR($B{r}="",{DAY}=""),"",'
    put(ws, f'C{r}', f'={g2}ROUND(SUMIF(账户表,$B{r},账户期初)'
                     f'+SUMIFS(流水收入,流水账户,$B{r},流水日期,"<"&{DAY})'
                     f'-SUMIFS(流水支出,流水账户,$B{r},流水日期,"<"&{DAY}),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'D{r}', f'={g2}ROUND(SUMIFS(流水收入,流水账户,$B{r},流水日期,{DAY}),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'E{r}', f'={g2}ROUND(SUMIFS(流水支出,流水账户,$B{r},流水日期,{DAY}),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'F{r}', f'={g2}ROUND($C{r}+$D{r}-$E{r},2))', fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'I{r}', f'=IF({acc(i)}="","",{acc(i)})', font=FT_AUTO, fill=F_AUTO)
    put(ws, f'H{r}', f'=IF($I{r}="","",{i})', fmt=NUM0, font=FT_AUTO, fill=F_AUTO)
    g3 = f'IF(OR($I{r}="",{MS}=""),"",'
    IN_M = f'流水日期,">="&{MS},流水日期,"<="&{ME}'
    put(ws, f'J{r}', f'={g3}ROUND(SUMIF(账户表,$I{r},账户期初)'
                     f'+SUMIFS(流水收入,流水账户,$I{r},流水日期,"<"&{MS})'
                     f'-SUMIFS(流水支出,流水账户,$I{r},流水日期,"<"&{MS}),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'K{r}', f'={g3}ROUND(SUMIFS(流水收入,流水账户,$I{r},{IN_M},流水属性,"<>内部转账"),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'L{r}', f'={g3}ROUND(SUMIFS(流水支出,流水账户,$I{r},{IN_M},流水属性,"<>内部转账"),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'M{r}', f'={g3}ROUND(SUMIFS(流水收入,流水账户,$I{r},{IN_M},流水属性,"内部转账"),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'N{r}', f'={g3}ROUND(SUMIFS(流水支出,流水账户,$I{r},{IN_M},流水属性,"内部转账"),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'O{r}', f'={g3}ROUND($J{r}+$K{r}-$L{r}+$M{r}-$N{r},2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
DT = D2R + 3 + N_ACC
put(ws, f'A{DT}', '合计', font=FT_HDR, fill=PatternFill('solid', fgColor=C_MAIN))
for c in 'CDEF':
    put(ws, f'{c}{DT}', f'=ROUND(SUM({c}{D2R+3}:{c}{DT-1}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
put(ws, f'B{DT}', '', fill=PatternFill('solid', fgColor=C_MAIN))
put(ws, f'H{DT}', '合计', font=FT_HDR, fill=PatternFill('solid', fgColor=C_SUB))
put(ws, f'I{DT}', '', fill=PatternFill('solid', fgColor=C_SUB))
for c in 'JKLMNO':
    put(ws, f'{c}{DT}', f'=ROUND(SUM({c}{D2R+3}:{c}{DT-1}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
ws.freeze_panes = 'C8'
NAMES['日报期末合计'] = f'汇报表!${C_END}${ATOT}'


def MN(m): return f'{YEAR}{m:02d}'
def MEOM(m): return f'EOMONTH(DATE({YEAR},{m},1),0)'


NOTT = '流水属性,"<>内部转账"'
IST = '流水属性,"内部转账"'

# ══════════════════════════════════════════════════════════════════════
# 4  月度汇报表（新增，按外汇资金台账那张的格式做）
# ══════════════════════════════════════════════════════════════════════
MCOL = [('月份', 11, None), ('收入', 15, MONEY), ('支出', 15, MONEY), ('净流入', 15, MONEY),
        ('内部转入', 14, MONEY), ('内部转出', 14, MONEY), ('月末资金总额', 17, MONEY),
        ('说明', 34, None)]
ws = newsheet('月度汇报表', '月度汇报表 · 按月看，内部转入和内部转出单独两列',
              '收入和支出都「不含内部转账」（账户之间调头寸不是收入也不是支出）；'
              '内部转入和内部转出单列两列，两边正常应该相等，不等就是有一头没录。'
              '「月末资金总额」= 所有账户期初 + 截至当月末的全部收支，就是公司手上的钱。', C_RPT, len(MCOL))
bar(ws, 3, 1, len(MCOL), '一、按月汇总', C_RPT)
hdr(ws, 4, 1, MCOL, C_RPT)
band(ws, 5, 16, 1, MCOL, auto=True)
for m in range(1, 13):
    r = 4 + m
    put(ws, f'A{r}', f'=DATE({YEAR},{m},1)', fmt='yyyy年m月', font=FT_AUTO, fill=F_AUTO)
    put(ws, f'B{r}', f'=ROUND(SUMIFS(流水收入,流水月份,{MN(m)},{NOTT}),2)',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'C{r}', f'=ROUND(SUMIFS(流水支出,流水月份,{MN(m)},{NOTT}),2)',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'D{r}', f'=ROUND($B{r}-$C{r},2)', fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'E{r}', f'=ROUND(SUMIFS(流水收入,流水月份,{MN(m)},{IST}),2)',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'F{r}', f'=ROUND(SUMIFS(流水支出,流水月份,{MN(m)},{IST}),2)',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'G{r}', f'=ROUND(SUM(账户期初)+SUMIFS(流水收入,流水日期,"<="&{MEOM(m)})'
                     f'-SUMIFS(流水支出,流水日期,"<="&{MEOM(m)}),2)',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'H{r}', f'=IF(ROUND($E{r}-$F{r},2)=0,"",'
                     f'"※内部转入比转出多 "&TEXT($E{r}-$F{r},"#,##0.00")&"，有一头没录")',
        font=FT_TIP, align=LEFT)
MT = 17
put(ws, f'A{MT}', '全年合计', font=FT_HDR, fill=PatternFill('solid', fgColor=C_RPT))
for c in 'BCDEF':
    put(ws, f'{c}{MT}', f'=ROUND(SUM({c}5:{c}16),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
put(ws, f'G{MT}', '=G16', fmt=MONEY, font=FT_SUM, fill=F_SUM)
put(ws, f'H{MT}', '年末资金总额 = 12 月那一行', font=FT_TIP, align=LEFT)
redflag(ws, f'H5:H{MT}')

M2 = MT + 2
ACOL2 = [('账户', 14, None), ('月初余额', 14, MONEY), ('本月收入', 14, MONEY),
         ('本月支出', 14, MONEY), ('内部转入', 13, MONEY), ('内部转出', 13, MONEY),
         ('月末余额', 14, MONEY), ('占资金比重', 11, PCT)]
bar(ws, M2, 1, len(ACOL2), '二、指定月份 · 各账户明细', C_MAIN)
put(ws, f'A{M2+1}', '报表月份', font=Font(name='微软雅黑', size=10, bold=True), fill=F_SUM)
put(ws, f'B{M2+1}', datetime.datetime(YEAR, 6, 1), fmt='yyyy年m月',
    font=Font(name='微软雅黑', size=11, bold=True, color='C00000'), fill=F_IN)
put(ws, f'C{M2+1}', '填该月 1 号', font=FT_TIP, align=LEFT)
ws.merge_cells(start_row=M2 + 1, start_column=3, end_row=M2 + 1, end_column=len(ACOL2))
hdr(ws, M2 + 2, 1, ACOL2, C_MAIN)
band(ws, M2 + 3, M2 + 2 + N_ACC, 1, ACOL2, auto=True)
RM, RE2 = f'$B${M2+1}', f'EOMONTH($B${M2+1},0)'
RNG_M = f'流水日期,">="&{RM},流水日期,"<="&{RE2}'
for i in range(1, N_ACC + 1):
    r = M2 + 2 + i
    g = f'IF(OR($A{r}="",{RM}=""),"",'
    put(ws, f'A{r}', f'=IF({acc(i)}="","",{acc(i)})', font=FT_AUTO, fill=F_AUTO)
    put(ws, f'B{r}', f'={g}ROUND(SUMIF(账户表,$A{r},账户期初)'
                     f'+SUMIFS(流水收入,流水账户,$A{r},流水日期,"<"&{RM})'
                     f'-SUMIFS(流水支出,流水账户,$A{r},流水日期,"<"&{RM}),2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    for col, src, cond in (('C', '流水收入', NOTT), ('D', '流水支出', NOTT),
                           ('E', '流水收入', IST), ('F', '流水支出', IST)):
        put(ws, f'{col}{r}', f'={g}ROUND(SUMIFS({src},流水账户,$A{r},{RNG_M},{cond}),2))',
            fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'G{r}', f'={g}ROUND($B{r}+$C{r}-$D{r}+$E{r}-$F{r},2))',
        fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'H{r}', f'=IF(OR($A{r}="",N($G${M2+3+N_ACC})=0),"",$G{r}/$G${M2+3+N_ACC})',
        fmt=PCT, font=FT_AUTO, fill=F_AUTO)
M2T = M2 + 3 + N_ACC
put(ws, f'A{M2T}', '合计', font=FT_HDR, fill=PatternFill('solid', fgColor=C_MAIN))
for c in 'BCDEFG':
    put(ws, f'{c}{M2T}', f'=ROUND(SUM({c}{M2+3}:{c}{M2T-1}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
put(ws, f'H{M2T}', f'=IF(N($G${M2T})=0,"",SUM(H{M2+3}:H{M2T-1}))', fmt=PCT, font=FT_SUM, fill=F_SUM)
ws.freeze_panes = 'A5'
NAMES['月报收入合计'] = f'月度汇报表!$B${MT}'
NAMES['月报支出合计'] = f'月度汇报表!$C${MT}'

# ══════════════════════════════════════════════════════════════════════
# 5  收支分类报表（新增，顶上那条「收入合计/支出合计/内部转账」按外汇台账的做）
# ══════════════════════════════════════════════════════════════════════
KCOL = ([('收/支类别', 15, None), ('属性', 10, None)]
        + [(f'{m}月', 12, MONEY) for m in range(1, 13)]
        + [('全年合计', 15, MONEY), ('占比', 9, PCT)])
K0 = 12
KE = K0 + N_CAT - 1
ws = newsheet('收支分类报表', '收支分类报表 · 按收支类别 × 12 个月',
              '支出类别显示成正数（好读）。顶上四行是汇总：收入合计、支出合计、内部转账净额、经营净额 —— '
              '内部转账「不进」收入和支出合计，它的净额正常应该是 0，不是 0 就说明有一头没录。'
              '下面每个类别一行，类别清单在【基础资料】维护，这里自动跟着长。', C_RPT, len(KCOL))
hdr(ws, 4, 1, KCOL, C_RPT)
for j, (lab, attr, color) in enumerate([('收入合计', '收入', C_RPT), ('支出合计', '支出', C_RPT),
                                        ('内部转账净额', '内部转账', C_SUB)]):
    r = 5 + j
    put(ws, f'A{r}', lab, font=FT_HDR, fill=PatternFill('solid', fgColor=color), align=LEFT)
    put(ws, f'B{r}', attr, font=FT_HDR, fill=PatternFill('solid', fgColor=color))
    for m in range(1, 13):
        cc = CL(2 + m)
        put(ws, f'{cc}{r}', f'=ROUND(SUMIF($B${K0}:$B${KE},"{attr}",{cc}${K0}:{cc}${KE}),2)',
            fmt=MONEY, font=FT_SUM, fill=F_SUM)
    put(ws, f'O{r}', f'=ROUND(SUM(C{r}:N{r}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
    put(ws, f'P{r}', '', fill=F_SUM)
put(ws, 'A8', '经营净额', font=FT_HDR, fill=PatternFill('solid', fgColor=C_MAIN), align=LEFT)
put(ws, 'B8', '收入−支出', font=FT_HDR, fill=PatternFill('solid', fgColor=C_MAIN))
for m in range(1, 13):
    cc = CL(2 + m)
    put(ws, f'{cc}8', f'=ROUND({cc}5-{cc}6,2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
put(ws, 'O8', '=ROUND(SUM(C8:N8),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
put(ws, 'P8', '', fill=F_SUM)
put(ws, 'A9', '内部转账净额不为 0 就是有一头没录，去【数据录入】核对列看。',
    font=FT_TIP, align=LEFT, border=False)
ws.merge_cells(start_row=9, start_column=1, end_row=9, end_column=len(KCOL))
bar(ws, 10, 1, len(KCOL), '按类别明细　——　类别清单在【基础资料】维护，这里自动跟着长', C_SUB)
hdr(ws, 11, 1, KCOL, C_SUB)
band(ws, K0, KE, 1, KCOL, auto=True)
for i in range(1, N_CAT + 1):
    r = K0 + i - 1
    cat = f'基础资料!$H${B0 + i - 1}'
    att = f'基础资料!$I${B0 + i - 1}'
    g = f'IF($A{r}="","",'
    put(ws, f'A{r}', f'=IF({cat}="","",{cat})', font=FT_AUTO, fill=F_AUTO, align=LEFT)
    put(ws, f'B{r}', f'={g}{att})', font=FT_AUTO, fill=F_AUTO)
    for m in range(1, 13):
        cc = CL(2 + m)
        put(ws, f'{cc}{r}', f'={g}ROUND(IF($B{r}="支出",-1,1)*'
                            f'(SUMIFS(流水收入,流水类别,$A{r},流水月份,{MN(m)})'
                            f'-SUMIFS(流水支出,流水类别,$A{r},流水月份,{MN(m)})),2))',
            fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'O{r}', f'={g}ROUND(SUM(C{r}:N{r}),2))', fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'P{r}', f'=IF(OR($A{r}="",$B{r}="内部转账"),"",'
                     f'IFERROR($O{r}/IF($B{r}="收入",$O$5,$O$6),""))',
        fmt=PCT, font=FT_AUTO, fill=F_AUTO)
ws.freeze_panes = 'C12'
NAMES['分类收入合计'] = '收支分类报表!$O$5'
NAMES['分类支出合计'] = '收支分类报表!$O$6'
NAMES['分类转账净额'] = '收支分类报表!$O$7'


INC_SEED = CUSTOMERS
EXP_SEED = [c for c, a in CATS if a == '支出']

# ══════════════════════════════════════════════════════════════════════
# 6  收支报表（保留原来的收支分列，每侧 40 行，后面空着随便加）
# ══════════════════════════════════════════════════════════════════════
SCOL = ([('序', 5, NUM0), ('项目', 18, None)]
        + [(f'{m}月', 11, MONEY) for m in range(1, 13)]
        + [('全年合计', 14, MONEY), ('占比', 9, PCT)])
IN_H, IN_0 = 5, 6
IN_E = IN_0 + N_ITEM - 1          # 45
IN_X = IN_E + 1                   # 46 未列示
EX_T = IN_X + 2                   # 48 支出总额
EX_H, EX_0 = EX_T + 1, EX_T + 2   # 49 / 50
EX_E = EX_0 + N_ITEM - 1          # 89
EX_X = EX_E + 1                   # 90

ws = newsheet('收支报表', '收支报表 · 收入、支出上下两块分开列，× 12 个月',
              '收入按「客户」归集、支出按「收支类别」归集 —— 跟原表一样。'
              f'每一块都留了 {N_ITEM} 行，空着的行直接往下打字就行，合计范围已经把整块都算进去了，'
              '不用改公式。最后一行「未列示」是自动补的差额，所以合计永远等于总额，不会漏数。'
              '收入总额和支出总额都「不含内部转账」。', C_RPT, len(SCOL))
put(ws, 'B3', '收支盈亏', font=FT_HDR, fill=PatternFill('solid', fgColor=C_MAIN), align=LEFT)
put(ws, 'B4', '收入总额（不含内部转账）', font=FT_HDR,
    fill=PatternFill('solid', fgColor=C_RPT), align=LEFT)
put(ws, f'B{EX_T}', '支出总额（不含内部转账）', font=FT_HDR,
    fill=PatternFill('solid', fgColor=C_RPT), align=LEFT)
for m in range(1, 13):
    cc = CL(2 + m)
    put(ws, f'{cc}4', f'=ROUND(SUMIFS(流水收入,流水月份,{MN(m)},{NOTT}),2)',
        fmt=MONEY, font=FT_SUM, fill=F_SUM)
    put(ws, f'{cc}{EX_T}', f'=ROUND(SUMIFS(流水支出,流水月份,{MN(m)},{NOTT}),2)',
        fmt=MONEY, font=FT_SUM, fill=F_SUM)
    put(ws, f'{cc}3', f'=ROUND({cc}4-{cc}{EX_T},2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
for r in (3, 4, EX_T):
    put(ws, f'A{r}', '', fill=PatternFill('solid', fgColor=C_MAIN if r == 3 else C_RPT))
    put(ws, f'O{r}', f'=ROUND(SUM(C{r}:N{r}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
    put(ws, f'P{r}', '', fill=F_SUM)

for h0, r0, r1, rx, seed, src, key, tot in (
        (IN_H, IN_0, IN_E, IN_X, INC_SEED, '流水收入', '流水客户', 4),
        (EX_H, EX_0, EX_E, EX_X, EXP_SEED, '流水支出', '流水类别', EX_T)):
    lab = '资金收入（按客户）' if key == '流水客户' else '资金支出（按收支类别）'
    cols = [('序', 5, NUM0), (lab, 18, None)] + SCOL[2:]
    hdr(ws, h0, 1, cols, C_RPT if key == '流水客户' else C_SUB)
    band(ws, r0, r1, 1, cols)
    for r in range(r0, r1 + 1):
        for c in [1] + list(range(3, 17)):
            cc = ws.cell(r, c); cc.font, cc.fill = FT_AUTO, F_AUTO
        ws.cell(r, 2).alignment = LEFT
        g = f'IF($B{r}="","",'
        put(ws, f'A{r}', f'={g}COUNTA($B${r0}:$B{r}))', fmt=NUM0, font=FT_AUTO, fill=F_AUTO)
        for m in range(1, 13):
            cc = CL(2 + m)
            put(ws, f'{cc}{r}', f'={g}ROUND(SUMIFS({src},{key},$B{r},流水月份,{MN(m)},{NOTT}),2))',
                fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
        put(ws, f'O{r}', f'={g}ROUND(SUM(C{r}:N{r}),2))', fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
        put(ws, f'P{r}', f'=IF(OR($B{r}="",N($O${tot})=0),"",$O{r}/$O${tot})',
            fmt=PCT, font=FT_AUTO, fill=F_AUTO)
    for i, v in enumerate(seed):
        ws.cell(r0 + i, 2).value = v
    put(ws, f'B{rx}', '未列示（自动补差，别删）', font=FT_SUM, fill=F_SUM, align=LEFT)
    put(ws, f'A{rx}', '', fill=F_SUM)
    for m in range(1, 13):
        cc = CL(2 + m)
        put(ws, f'{cc}{rx}', f'=ROUND({cc}{tot}-SUM({cc}{r0}:{cc}{r1}),2)',
            fmt=MONEY, font=FT_SUM, fill=F_SUM)
    put(ws, f'O{rx}', f'=ROUND(SUM(C{rx}:N{rx}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
    put(ws, f'P{rx}', f'=IF(N($O${tot})=0,"",$O{rx}/$O${tot})', fmt=PCT, font=FT_SUM, fill=F_SUM)
ws.freeze_panes = 'C6'
NAMES['报表收入总额'] = '收支报表!$O$4'
NAMES['报表支出总额'] = f'收支报表!$O${EX_T}'

# ══════════════════════════════════════════════════════════════════════
# 7  任意时间段收支报表（保留原界面，左右两栏收支分列）
# ══════════════════════════════════════════════════════════════════════
QCOL = [('序', 5, NUM0), ('资金收入（按客户）', 20, None), ('金额', 15, MONEY), ('收入占比', 10, PCT),
        ('序', 5, NUM0), ('资金支出（按收支类别）', 20, None), ('金额', 15, MONEY), ('支出占比', 10, PCT)]
Q0 = 7
QE = Q0 + N_ITEM - 1              # 46
QX = QE + 1                       # 47
ws = newsheet('任意时间段收支报表', '任意时间段收支报表 · 改上面两个日期就行',
              f'左边收入按客户、右边支出按收支类别，各留了 {N_ITEM} 行，'
              '空行直接往下打字，合计已经把整块算进去了。最后一行「未列示」自动补差，合计永远等于总额。'
              '收入总额和支出总额都不含内部转账，内部转账净额单独显示。', C_RPT, len(QCOL))
for lab, addr, v in (('开始日期', 'B3', datetime.datetime(YEAR, 6, 1)),
                     ('结束日期', 'B4', datetime.datetime(YEAR, 6, 30))):
    put(ws, f'A{addr[1:]}', lab, font=Font(name='微软雅黑', size=10, bold=True), fill=F_SUM)
    put(ws, addr, v, fmt=DATEF, font=Font(name='微软雅黑', size=11, bold=True, color='C00000'), fill=F_IN)
QR = '流水日期,">="&$B$3,流水日期,"<="&$B$4'
for j, (lab, f, col) in enumerate([
        ('收入总额', f'=ROUND(SUMIFS(流水收入,{QR},{NOTT}),2)', 'D'),
        ('支出总额', f'=ROUND(SUMIFS(流水支出,{QR},{NOTT}),2)', 'F'),
        ('经营盈亏', '=ROUND($D$3-$F$3,2)', 'H'),
        ('内部转账净额', f'=ROUND(SUMIFS(流水收入,{QR},{IST})-SUMIFS(流水支出,{QR},{IST}),2)', 'D')]):
    r = 3 if j < 3 else 4
    lc = CL(ord(col) - 65)
    put(ws, f'{lc}{r}', lab, font=Font(name='微软雅黑', size=10, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor=C_RPT if j < 3 else C_SUB))
    put(ws, f'{col}{r}', f, fmt=MONEY, font=FT_SUM, fill=F_SUM)
put(ws, 'F4', '=IF(OR($B$3="",$B$4=""),"",IF($B$4<$B$3,"※结束日期早于开始日期",'
              '"共 "&($B$4-$B$3+1)&" 天，"&COUNTIFS(' + QR + ')&" 笔"))',
    font=FT_TIP, align=LEFT)
ws.merge_cells('F4:H4')
hdr(ws, 6, 1, QCOL, C_RPT)
band(ws, Q0, QE, 1, QCOL)
for r in range(Q0, QE + 1):
    for c in (1, 3, 4, 5, 7, 8):
        cc = ws.cell(r, c); cc.font, cc.fill = FT_AUTO, F_AUTO
    for c in (2, 6):
        ws.cell(r, c).alignment = LEFT
    for key, nm, src, tot, sn, vn, pn in (('流水客户', 'B', '流水收入', '$D$3', 'A', 'C', 'D'),
                                          ('流水类别', 'F', '流水支出', '$F$3', 'E', 'G', 'H')):
        g = f'IF(${nm}{r}="","",'
        put(ws, f'{sn}{r}', f'={g}COUNTA(${nm}${Q0}:${nm}{r}))', fmt=NUM0, font=FT_AUTO, fill=F_AUTO)
        put(ws, f'{vn}{r}', f'={g}ROUND(SUMIFS({src},{key},${nm}{r},{QR},{NOTT}),2))',
            fmt=MONEY, font=FT_AUTO, fill=F_AUTO)
        put(ws, f'{pn}{r}', f'=IF(OR(${nm}{r}="",N({tot})=0),"",{vn}{r}/{tot})',
            fmt=PCT, font=FT_AUTO, fill=F_AUTO)
for i, v in enumerate(INC_SEED):
    ws.cell(Q0 + i, 2).value = v
for i, v in enumerate(EXP_SEED):
    ws.cell(Q0 + i, 6).value = v
put(ws, f'B{QX}', '未列示（自动补差，别删）', font=FT_SUM, fill=F_SUM, align=LEFT)
put(ws, f'F{QX}', '未列示（自动补差，别删）', font=FT_SUM, fill=F_SUM, align=LEFT)
for cn, tot in (('C', '$D$3'), ('G', '$F$3')):
    put(ws, f'{cn}{QX}', f'=ROUND({tot}-SUM({cn}{Q0}:{cn}{QE}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
for cn, vn, tot in (('D', 'C', '$D$3'), ('H', 'G', '$F$3')):
    put(ws, f'{cn}{QX}', f'=IF(N({tot})=0,"",{vn}{QX}/{tot})', fmt=PCT, font=FT_SUM, fill=F_SUM)
for cn in ('A', 'E'):
    put(ws, f'{cn}{QX}', '', fill=F_SUM)
QT = QX + 1
put(ws, f'B{QT}', '合计', font=FT_HDR, fill=PatternFill('solid', fgColor=C_RPT), align=LEFT)
put(ws, f'F{QT}', '合计', font=FT_HDR, fill=PatternFill('solid', fgColor=C_RPT), align=LEFT)
for cn in ('C', 'G'):
    put(ws, f'{cn}{QT}', f'=ROUND(SUM({cn}{Q0}:{cn}{QX}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM)
for cn in ('A', 'D', 'E', 'H'):
    put(ws, f'{cn}{QT}', '', fill=PatternFill('solid', fgColor=C_RPT))
ws.freeze_panes = 'A7'


# ══════════════════════════════════════════════════════════════════════
# 8  账户子表（每个账户一张，自动拆）
# ══════════════════════════════════════════════════════════════════════
SUBSUM = [('期初余额', 'ROUND(SUMIF(账户表,$B$2,账户期初),2)'),
          ('收入合计', f'ROUND(SUMIFS(流水收入,流水账户,$B$2,{NOTT}),2)'),
          ('支出合计', f'ROUND(SUMIFS(流水支出,流水账户,$B$2,{NOTT}),2)'),
          ('内部转入', f'ROUND(SUMIFS(流水收入,流水账户,$B$2,{IST}),2)'),
          ('内部转出', f'ROUND(SUMIFS(流水支出,流水账户,$B$2,{IST}),2)'),
          ('当前余额', 'ROUND(SUMIF(账户表,$B$2,账户期初)'
                    '+SUMIF(流水账户,$B$2,流水收入)-SUMIF(流水账户,$B$2,流水支出),2)')]
SUB_SHEETS = [a[0] for a in ACCOUNTS]
for name in SUB_SHEETS:
    ws = newsheet(name, f'【{name}】账户流水明细 · 自动从【数据录入】拆出来，不用填', '',
                  C_SUB, len(DCOL), merge2=False)
    put(ws, 'A2', '账户名称', font=Font(name='微软雅黑', size=10, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor=C_SUB))
    put(ws, 'B2', name, font=Font(name='微软雅黑', size=11, bold=True, color='C00000'), fill=F_IN)
    put(ws, 'C2', '← 这格是拆分依据（可以改成别的账户名）。下面整块是一条公式的结果，'
                  '在里面打字会把公式顶掉；要改数据去【数据录入】改。',
        font=FT_TIP, fill=PatternFill('solid', fgColor='FFF7E6'), align=LEFT)
    ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=len(DCOL))
    dv(ws, '=账户表', 'B2', stop=True)
    for j, (lab, f) in enumerate(SUBSUM):
        c1, c2 = CL(1 + j * 2), CL(2 + j * 2)
        put(ws, f'{c1}3', lab, font=Font(name='微软雅黑', size=9, bold=True, color='FFFFFF'),
            fill=PatternFill('solid', fgColor=C_SUB))
        put(ws, f'{c2}3', f'=IF($B$2="","",{f})', fmt=MONEY, font=FT_SUM, fill=F_SUM)
    ws.row_dimensions[3].height = 20
    hdr(ws, HDR, 1, DCOL, C_SUB)
    band(ws, R0, R0 + SUB - 1, 1, [(t, w, DATEQ if fmt == DATEF else fmt) for t, w, fmt in DCOL],
         auto=True)
    ws[f'A{R0}'] = ArrayFormula(
        ref=f'A{R0}:P{R0 + SUB - 1}',
        text=f'=_xlfn._xlws.FILTER({nb(f"数据录入!$A${R0}:$P${RE_}")},'
             f'数据录入!$C${R0}:$C${RE_}=$B$2,"这个账户还没有流水")')
    ws.freeze_panes = f'C{R0}'

# ══════════════════════════════════════════════════════════════════════
# 9  核对表
# ══════════════════════════════════════════════════════════════════════
NSUB = len(SUB_SHEETS)
ALLIN = 'SUM(流水收入)'
ALLOUT = 'SUM(流水支出)'
CHECKS = [
    ('勾稽', '分类报表：收入合计 − 支出合计 + 内部转账净额 = 全部流水的收 − 支',
     f'=ROUND((分类收入合计-分类支出合计+分类转账净额)-({ALLIN}-{ALLOUT}),2)', MONEY,
     '不为 0 说明有流水的「收支类别」没填、或者填了【基础资料】清单里没有的类别，那几笔进不了分类报表。'),
    ('勾稽', '收支报表全年收入总额 = 月度汇报表全年收入合计',
     '=ROUND(报表收入总额-月报收入合计,2)', MONEY, '两张表算法不同，必须一致。'),
    ('勾稽', '收支报表全年支出总额 = 月度汇报表全年支出合计',
     '=ROUND(报表支出总额-月报支出合计,2)', MONEY, '同上。'),
    ('勾稽', '汇报表①「期末结余金额」= 期初合计 + 截至截止日的收 − 支',
     f'=ROUND(日报期末合计-(SUM(账户期初)'
     f'+SUMIFS(流水收入,流水日期,"<="&汇报表!$E$3)-SUMIFS(流水支出,流水日期,"<="&汇报表!$E$3)),2)',
     MONEY, '31 天那块横着加出来的期末，跟直接从流水算的期末必须一致。'),
    ('勾稽', '汇报表③ 各账户「结余」合计 = 期初合计 + 截至该月末的收 − 支',
     f'=ROUND(汇报表!$O${DT}-(SUM(账户期初)'
     f'+SUMIFS(流水收入,流水日期,"<="&EOMONTH(汇报表!$I${D2R+1},0))'
     f'-SUMIFS(流水支出,流水日期,"<="&EOMONTH(汇报表!$I${D2R+1},0))),2)', MONEY,
     '月度那块的「期初+收−支+转入−转出」必须还原成同一个数。'),
    ('勾稽', '账户子表张数 = 账户档案里的账户数',
     f'={NSUB}-COUNTA(账户表)', NUM0,
     f'现在有 {NSUB} 张子表。加了账户没加子表，这里就不是 0 —— 复制一张子表改 B2 即可。'),
    ('待办', '流水里的红字（※，是错，必须改）', '=COUNTIF(流水核对,"※*")', NUM0,
     '去【数据录入】最后一列看，红的都要改。'),
    ('待办', '流水里的琥珀字（△，是待补的数据）', '=COUNTIF(流水核对,"△*")', NUM0,
     '不影响公式，但补齐了数才准。'),
    ('待办', '账户档案里的红字（※）', f'=COUNTIF(基础资料!$F${B0}:$F${AE},"※*")', NUM0,
     '账户名重复、期初没填、流水早于启用日期。'),
    ('待办', '内部转账净额（收 − 支，应该是 0）',
     f'=ROUND(SUMIFS(流水收入,{IST})-SUMIFS(流水支出,{IST}),2)', MONEY,
     '账户之间调头寸要两头都录：转出账户记支出、转入账户记收入。不为 0 就是漏了一头。'),
    ('待办', '内部转账「收」「支」笔数差',
     f'=COUNTIFS({IST},流水收入,">0")-COUNTIFS({IST},流水支出,">0")', NUM0, '同上，按笔数再查一遍。'),
    ('待办', '没填收支类别的流水笔数', '=COUNTIFS(流水账户,"<>",流水类别,"")', NUM0,
     '没类别的笔数进不了【收支分类报表】，也进不了【收支报表】的支出那一块。'),
    ('待办', '收支报表 · 收入「未列示」全年金额',
     f'=ROUND(收支报表!$O${IN_X},2)', MONEY,
     '不是 0 说明有客户的钱没被上面那几行列到。把客户名补进【收支报表】收入区的空行里就归位了。'),
    ('待办', '收支报表 · 支出「未列示」全年金额',
     f'=ROUND(收支报表!$O${EX_X},2)', MONEY, '同上，补进支出区的空行。'),
    ('待办', '有流水、但不在账户档案里的账户数',
     f'=SUMPRODUCT((流水账户<>"")*(COUNTIF(账户表,流水账户&"")=0))', NUM0,
     '这些笔在按账户的报表里一行都看不到。'),
    ('待办', '账户名 / 类别 / 客户 清单里的重复项',
     '=SUMPRODUCT((账户表<>"")*(COUNTIF(账户表,账户表&"")>1))'
     '+SUMPRODUCT((类别表<>"")*(COUNTIF(类别表,类别表&"")>1))'
     '+SUMPRODUCT((客户表<>"")*(COUNTIF(客户表,客户表&"")>1))', NUM0,
     '重复会让报表里出现两行一模一样的，金额算两遍。'),
    ('提示', '流水已用行数 / 预留行数',
     f'=COUNTA(流水账户)&" / {N_ROW}"', None, '快满了就告诉我加行。'),
    ('提示', '资金总额（所有账户当前余额合计）',
     f'=ROUND(SUM(账户期初)+{ALLIN}-{ALLOUT},2)', MONEY, '公司手上一共有多少钱。只是报个数，不是错。'),
    ('提示', '本年经营净额（收入 − 支出，不含内部转账）',
     '=ROUND(分类收入合计-分类支出合计,2)', MONEY, '同上。'),
]
CKC = [('序号', 6, NUM0), ('类别', 8, None), ('检查项', 50, None), ('结果', 16, None),
       ('判断', 12, None), ('说明', 62, None)]
ws = newsheet('核对表', '核对表 · 每次录完看一眼这张就够了',
              '【勾稽】是算法自检，必须全是 0，不是 0 就是表算错了要告诉我；'
              '【待办】是数据没填齐，你们自己补。', C_CHK, len(CKC))
hdr(ws, 4, 1, CKC, C_CHK)
band(ws, 5, 4 + len(CHECKS), 1, CKC, auto=True)
for i, (kind, nm, f, fmt, memo) in enumerate(CHECKS, 1):
    r = 4 + i
    put(ws, f'A{r}', i, fmt=NUM0, font=FT_AUTO, fill=F_AUTO)
    put(ws, f'B{r}', kind, font=Font(name='微软雅黑', size=10, bold=True,
                                     color={'勾稽': 'C00000', '待办': '8B5E00'}.get(kind, '7F7F7F')),
        fill=F_AUTO)
    put(ws, f'C{r}', nm, font=FT_AUTO, fill=F_AUTO, align=LEFT)
    put(ws, f'D{r}', f, fmt=fmt, font=FT_SUM, fill=F_SUM)
    if fmt is None or kind == '提示':
        put(ws, f'E{r}', '—', font=FT_AUTO, fill=F_AUTO)
    elif kind == '勾稽':
        put(ws, f'E{r}', f'=IF(ROUND(N($D{r}),2)=0,"✓ 通过","✗ 算不平")', font=FT_SUM, fill=F_AUTO)
    else:
        put(ws, f'E{r}', f'=IF(ROUND(N($D{r}),2)=0,"✓ 没有","△ 有")', font=FT_SUM, fill=F_AUTO)
    put(ws, f'F{r}', memo, font=FT_TIP, align=LEFT)
CKN = 4 + len(CHECKS)
NG = sum(1 for c in CHECKS if c[0] == '勾稽')
put(ws, f'A{CKN+2}', '勾稽通过', font=Font(name='微软雅黑', size=11, bold=True, color='FFFFFF'),
    fill=PatternFill('solid', fgColor=C_CHK), align=LEFT)
ws.merge_cells(f'A{CKN+2}:C{CKN+2}')
put(ws, f'D{CKN+2}', f'=COUNTIF($E$5:$E${CKN},"✓ 通过")&" / {NG}"', font=FT_SUM, fill=F_SUM)
put(ws, f'E{CKN+2}', '待办', font=Font(name='微软雅黑', size=11, bold=True, color='FFFFFF'),
    fill=PatternFill('solid', fgColor=C_SUB))
put(ws, f'F{CKN+2}', f'=COUNTIF($E$5:$E${CKN},"△ 有")&" 项没填齐"', font=FT_SUM, fill=F_SUM, align=LEFT)
ws.conditional_formatting.add(f'E5:E{CKN}', FormulaRule(
    formula=['LEFT(E5,1)="✗"'], font=Font(color='C00000', bold=True),
    fill=PatternFill('solid', fgColor='FFD7D7')))
ws.conditional_formatting.add(f'E5:E{CKN}', FormulaRule(
    formula=['LEFT(E5,1)="△"'], font=Font(color='8B5E00', bold=True),
    fill=PatternFill('solid', fgColor='FFF2CC')))
ws.freeze_panes = 'A5'


# ══════════════════════════════════════════════════════════════════════
# 10  使用说明（放最前面）
# ══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet('使用说明', 0)
ws.sheet_properties.tabColor = '1F3864'
ws.sheet_view.showGridLines = False
for i, w in enumerate([20, 16, 76, 20], 1):
    ws.column_dimensions[CL(i)].width = w
put(ws, 'A1', '资金日报表 · 混合录入版 ｜ 使用说明（单币种：人民币）', font=FT_TITLE,
    fill=PatternFill('solid', fgColor='1F3864'), align=LEFT)
ws.merge_cells('A1:D1'); ws.row_dimensions[1].height = 30


def sec(r, t, color=C_MAIN):
    put(ws, f'A{r}', t, font=Font(name='微软雅黑', size=12, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor=color), align=LEFT)
    ws.merge_cells(f'A{r}:D{r}'); ws.row_dimensions[r].height = 24


def ln(r, a, b='', c=''):
    put(ws, f'A{r}', a, font=Font(name='微软雅黑', size=10, bold=True), align=LEFT)
    put(ws, f'B{r}', b, font=Font(name='微软雅黑', size=10), align=LEFT)
    put(ws, f'C{r}', c, font=Font(name='微软雅黑', size=10),
        align=Alignment('left', 'center', wrap_text=True))
    ws.merge_cells(f'C{r}:D{r}')


def para(r, t, color='595959'):
    put(ws, f'A{r}', t, font=Font(name='微软雅黑', size=9, color=color), align=LEFT)
    ws.merge_cells(f'A{r}:D{r}')


r = 3
sec(r, '一、两步就能用起来'); r += 1
for a, b, c in [('第 1 步', '【基础资料】', '把账户档案改成你们真实的账户：账户名称、类型、期初余额。'
                                      '收支类别、客户、供应商、项目也在这一页维护。'),
                ('第 2 步', '【数据录入】', '每天在这一张录。选账户 → 余额自动结；收入和支出分两列，只填一个。'
                                      '其余所有表全自动。')]:
    ln(r, a, b, c); r += 1
r += 1

sec(r, '二、这次跟原来那份比，改了什么'); r += 1
ln(r, '改动', '为什么', '影响'); r += 1
for a, b, c in [
    ('加了「内部转账」', '账户互转原来两头都算进收支',
     '【基础资料】收支类别多了一列「属性」：收入 / 支出 / 内部转账。'
     '账户之间调头寸选「内部转账」，它就不会被算进收入合计和支出合计 —— '
     '原来把 5 万从工行转到现金，收入 +5 万、支出 +5 万，两头一起虚增。'),
    ('收入也有类别了', '原表只有「支出业务类别」',
     '原来收入一侧一个类别都没有，只能按客户分析。现在收入、支出、内部转账统一用一张「收支类别」清单。'
     '【收支报表】的收入区「仍然按客户」归集，跟原来一样，没动。'),
    ('期初余额挪了个地方', '原来放在【数据录入】第 5~18 行',
     '挪到【基础资料】账户档案的「期初余额」列。原来的写法有个坑：'
     '【汇报表】取期初用的是「期初余额行的日期 <= 查询起始日」，'
     '一旦查询起始日早于期初行的日期（比如期初记的是 6 月 1 日，你查 5 月），期初余额直接取不到，整列变成 0。'),
    ('账户清单自动长', '原来写死在【汇报表】三个地方',
     '【汇报表】那三块的账户列（原 A8:A21 / B27:B40 / I27:I40）现在都是公式，'
     '在【基础资料】加一个账户，三块同时多一行。以前要手工改三处，漏一处就少一个账户。'),
    ('收支报表加了空行 + 未列示', '原来收入区只有 13 行，且合计范围正好卡死',
     f'原表 C4=SUM(C6:C18) 正好等于 13 个客户，加第 14 个客户「不会进合计」。'
     f'现在每一块留 {N_ITEM} 行，合计范围把整块都算进去；'
     '最后再加一行「未列示」自动补差 —— 所以合计永远等于总额，漏列谁都不会少数。'),
    ('新增两张汇总表', '按你说的参照外汇资金台账',
     '【月度汇报表】：按月的收入 / 支出 / 净流入 / 「内部转入」 / 「内部转出」 / 月末资金总额，'
     '下半张是指定月份的各账户明细。'
     '【收支分类报表】：收支类别 × 12 个月，顶上四行是「收入合计 / 支出合计 / 内部转账净额 / 经营净额」。'),
    ('新增【核对表】', '原来没有自检', '勾稽 6 项必须全是 0；待办 12 项是数据没填齐。录完看一眼这张就够了。'),
]:
    ln(r, a, b, c); r += 1
r += 1

sec(r, '三、账户之间调头寸怎么录', C_SUB); r += 1
ln(r, '录两行', '转出一行、转入一行', '转出账户记一笔「支出」，转入账户记一笔「收入」，两行的「收/支类别」都选「内部转账」。'); r += 1
ln(r, '写上对方', '「对方账户」列', '转出那行写转入账户名，转入那行写转出账户名。核对列会盯着你有没有写。'); r += 1
ln(r, '报表里怎么显示', '单列，不进合计',
   '【月度汇报表】和【汇报表】③ 各有「内部转入 / 内部转出」两列；'
   '【收支分类报表】单列一行「内部转账净额」。收入合计、支出合计里都不含它。'); r += 1
ln(r, '净额应该是 0', '不是 0 = 漏了一头', '同一笔转账两头金额相等，净额自然是 0。'
                                 '【核对表】第 10、11 项专门盯这个。'); r += 1
r += 1

sec(r, '四、三块老界面都还在'); r += 1
ln(r, '表名', '哪一块', '说明'); r += 1
for a, b, c in [
    ('汇报表 ①', '资金日报统计表', '改上面两个日期，最多 31 天，每天两列「收款 / 付款」。'
                              '这一块是「账户口径」：内部划转也算进收款付款，因为账户余额确实动了。'),
    ('汇报表 ②', '货币资金日报表', '指定某一天，各账户的昨日余额 / 本日收入 / 本日支出 / 本日余额。也是账户口径。'),
    ('汇报表 ③', '多帐户资金汇总表', '按月，这一块是「经营口径」：本月收入 / 本月支出不含内部划转，'
                               '内部转入 / 内部转出另开两列。'),
    ('收支报表', '12 个月', '收入按客户、支出按收支类别，上下两块分开列 —— 跟原来一样。'),
    ('任意时间段收支报表', '任意起止日期', '左右两栏收支分列 —— 跟原来一样。'),
    ('账户子表', f'每个账户一张（现在 {NSUB} 张）', '自动从【数据录入】拆，不用填。B2 是拆分依据。'),
]:
    ln(r, a, b, c); r += 1
r += 1

sec(r, '五、加东西怎么加', C_SUB); r += 1
ln(r, '加一个账户', '两步', '① 【基础资料】账户档案接着写一行（记得填期初余额）；'
                       '② 右键复制任意一张账户子表，改表名，把 B2 改成新账户名。'
                       '【汇报表】三块会自动多出这一行。核对表第 6 项盯着你有没有漏第 ② 步。'); r += 1
ln(r, '加一个收支类别', '一步', '【基础资料】收支类别接着写，「属性一定要填」（收入 / 支出 / 内部转账）。'
                          '【收支分类报表】自动多一行。'); r += 1
ln(r, '加客户 / 供应商 / 项目', '一步', '【基础资料】对应清单接着写。'
                                '【收支报表】和【任意时间段收支报表】的清单是手填的，'
                                '记得把新客户也补到那两张表的空行里（不补也不会少数，会落到「未列示」那一行）。'); r += 1
ln(r, '行不够了', '告诉我', f'流水预留 {N_ROW} 行、账户 {N_ACC} 个、收支类别 {N_CAT} 个、'
                        f'客户和供应商各 {N_CUS} 个、子表每张 {SUB} 行。核对表第 17 项显示用了多少。'); r += 1
r += 1

sec(r, '六、颜色和符号', C_BASE); r += 1
ln(r, '淡黄色格子', '要你填', '这是人工输入区。'); r += 1
ln(r, '灰色格子', '公式自动算', '别手动改，改了就把公式顶掉了。'); r += 1
ln(r, '※ 红字', '是错，必须改', '账户不在档案、收支同时填、金额填了负数、类别不在清单……'); r += 1
ln(r, '△ 琥珀字', '待补的数据', '内部转账没写对方账户、余额成负数……不影响公式，补齐了数才准。'); r += 1
r += 1

sec(r, '七、示例数据', C_BASE); r += 1
para(r, f'· 【数据录入】第 {R0} ~ {R0 + len(FLOWS) - 1} 行是示例（原表那 23 笔 + 2 笔内部转账），'
        '正式用之前整段选中删掉即可，所有报表自动归零。'); r += 1
para(r, '· 账户、客户、供应商、支出类别全部按原表搬过来，一个不少；收入类别和内部转账类别是这次新加的。'); r += 1
para(r, '· 期初余额也按原表那 14 个账户的数搬过来了（现金 5000、微信 6000、工行 7000 …… 备用4 18000）。'); r += 1

# ══════════════════════════════════════════════════════════════════════
for nm, ref in NAMES.items():
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)
for s in wb.worksheets:
    s.page_setup.orientation = 'landscape'
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 0
    s.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = 0
wb.save(OUT)
print('已生成:', OUT)
print('工作表 %d 张:' % len(wb.worksheets), ' / '.join(s.title for s in wb.worksheets))
n = sum(1 for s in wb.worksheets for row in s.iter_rows()
        for c in row if isinstance(c.value, str) and c.value.startswith('='))
print('公式格子 %d 个' % n)

sys.path.insert(0, os.path.join(os.path.dirname(ROOT), '工具'))
import dyn_array, fix_sheet_selection, check_formula                  # noqa: E402
dyn_array.install(OUT)
fix_sheet_selection.fix(OUT)
if check_formula.scan(OUT):
    raise SystemExit('公式括号不配对，先修了再交付')
