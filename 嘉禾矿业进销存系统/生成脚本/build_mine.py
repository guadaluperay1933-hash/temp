# -*- coding: utf-8 -*-
"""生成《嘉禾矿业进销存系统.xlsx》。

12 张表：使用说明 / 参数表 / 物料档案 / 费用台账 / 入库单 / 出库单 /
        到货批次 / 库存台账 / 月度盘点 / 领导月报 / 应收对账 / 核对表

几条定死的口径：
· 本位币先令。人民币按参数表的汇率折（370），全表不许再出现 370 这个数字。
· 不做单位换算：数量、单位、单价三者一致，金额 = 数量 × 单价。
  （原表「100米」那种单位，单价就是每 100 米的价，一换算就差 100 倍。）
· 海运费按**批次**分摊，批次 = 集装箱号 + 到港年月（柜号会被船公司重复使用）。
  分摊基数是本批物料的先令货值，费用先折先令再摊，只折一次。
  用累计差额法，保证每批分摊合计 = 费用合计，一分不差。
  本批只要还有没填单价的行，运费就**挂账不摊** —— 摊了会全压到有价的那几行上。
· 期末 = 期初 + 采购入库 + 盘盈 + 退库 − 自用消耗 − 对外销售 − 盘亏 − 退货给供应商
  这一句在库存台账、月报、核对表三处用的是同一个式子。
· 售价：国内货 = 人民币进价 × 2 × 汇率（不含运费）；当地货 = 先令进价 × 1.03。
"""
import json, os, datetime, collections
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DATA = json.load(open(os.path.join(HERE, 'mine_data.json'), encoding='utf-8'))
OUT = os.path.join(ROOT, '嘉禾矿业进销存系统.xlsx')

N_ITEM, N_IN, N_OUT, N_FEE, N_BAT = 1200, 2500, 4000, 500, 200
HDR, D0 = 3, 4
S0 = 5                                   # 自动表：4 行合计，5 行起数据

C_IN, C_AUTO, C_BASE, C_BOSS = '2F5597', '375623', '806000', '833C00'
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
TZS, CNY, QTY, DATEF, PCT = '#,##0', '#,##0.00', '#,##0.###', 'yyyy-mm-dd', '0.0%'

MINES = ['一矿区', '二矿区', '三矿区', '六矿区']
OWN, SELL = ['一矿区', '二矿区'], ['三矿区', '六矿区']
OUTK = ['自用消耗', '对外销售', '盘亏', '退货给供应商']
INK = ['采购入库', '盘盈', '矿区退库']
FEEK = ['国内海运费', '国内内陆运费', '装柜费', '报关服务费', '坦桑到港费',
        '坦桑清关费', '坦桑内陆运输费', '卸车费', '其他费用']
SRC_ = ['A国内', 'B当地']
PSTAT = ['正常', '待补价', '暂估', '含在主件', '赠品']
CATS = sorted({i['cat'] for i in DATA['items']})
UNITS = sorted({i['unit'] for i in DATA['items']})[:60]

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
    ws.row_dimensions[2].height = 24
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    if merge2:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
    return ws


def header(ws, cols, color, row=HDR):
    fill = PatternFill('solid', fgColor=color)
    for i, (t, w, kind, fmt) in enumerate(cols, 1):
        c = ws.cell(row, i, t)
        c.font, c.fill, c.alignment, c.border = FT_HDR, fill, CEN, BOX
        ws.column_dimensions[CL(i)].width = w
    ws.row_dimensions[row].height = 34


def body(ws, cols, r0, r1):
    for r in range(r0, r1 + 1):
        for i, (t, w, kind, fmt) in enumerate(cols, 1):
            c = ws.cell(r, i)
            c.border, c.font = BOX, (FT_AUTO if kind == 'auto' else FT_IN)
            c.fill = F_AUTO if kind == 'auto' else F_IN
            c.alignment = CEN if (fmt or kind == 'auto') else LEFT
            if fmt:
                c.number_format = fmt


def put(ws, addr, v, fmt=None, font=None, fill=None, align=None, box=True):
    c = ws[addr]
    c.value = v
    if fmt: c.number_format = fmt
    if font: c.font = font
    if fill: c.fill = fill
    if align: c.alignment = align
    if box: c.border = BOX
    return c


def dv(ws, f1, rng, warn=True):
    d = DataValidation(type='list', formula1=f1, allow_blank=True, showErrorMessage=True,
                       errorStyle='warning' if warn else 'stop',
                       errorTitle='不在名单里', error='先去【参数表】或【物料档案】把它加上。')
    ws.add_data_validation(d)
    d.add(rng)


def redflag(ws, rng):
    """※ = 真的错了（红）；△ = 只是还没补齐（琥珀），别把提醒和错误混成一种颜色"""
    a = rng.split(':')[0].replace('$', '')
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'LEFT({a},1)="※"'], font=Font(color='C00000', bold=True),
        fill=PatternFill('solid', fgColor='FFD7D7')))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'LEFT({a},1)="△"'], font=Font(color='BF8F00', bold=True),
        fill=PatternFill('solid', fgColor='FFF2CC')))


# ══════════════════════════════════════════════════════════════════════
# 参数表
# ══════════════════════════════════════════════════════════════════════
ws = newsheet('参数表', '参数表 · 全表唯一能改规则的地方',
              '汇率、加价倍数、账期只在这里改一次，所有表跟着变。'
              '全表别的地方不许再出现 370、2、1.03 这几个数字。', C_BASE, 20)
PARAM = [
    ('汇率_人民币兑先令', 370, '0', '国内采购是人民币，除这个数就是先令'),
    ('国内货加价倍数', 2, '0.00', '售价 = 人民币进价 × 本倍数 × 汇率（不含运费）'),
    ('当地货加价系数', 1.03, '0.000', '售价 = 先令进价 × 本系数（加卸车运输补偿）'),
    ('本期账期起', datetime.datetime(2026, 1, 1), DATEF,
     '月报、库存台账都按这一段算。正常就填当月 1 号；现在填 1/1 是为了把 1~8 月的历史一次并进来'),
    ('本期账期止', datetime.datetime(2026, 8, 31), DATEF,
     '正常填当月最后一天（右边 D 列给了参考值）；首期填 8/31，把搬过来的 867 行全部收进来'),
    ('已结账截止日', datetime.datetime(2025, 12, 31), DATEF, '早于这个日期的单据不许再录'),
    ('本月一矿处理矿石量(吨)', None, '#,##0.##', '选填，填了月报才算单位成本'),
    ('本月二矿处理矿石量(吨)', None, '#,##0.##', '选填'),
]
put(ws, 'A3', '参数', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BASE), align=CEN)
put(ws, 'B3', '数值', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BASE), align=CEN)
put(ws, 'C3', '说明', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BASE), align=CEN)
for c, w in (('A', 26), ('B', 16), ('C', 46)):
    ws.column_dimensions[c].width = w
for i, (lab, v, fmt, note) in enumerate(PARAM):
    r = D0 + i
    put(ws, f'A{r}', lab, font=Font(name='微软雅黑', size=10, bold=True), align=LEFT)
    put(ws, f'B{r}', v, fmt=fmt,
        font=Font(name='微软雅黑', size=12, bold=True, color='C00000'),
        fill=(F_AUTO if isinstance(v, str) and str(v).startswith('=') else F_IN), align=CEN)
    put(ws, f'C{r}', note, font=Font(name='微软雅黑', size=9, color='808080'), align=LEFT)
put(ws, 'D5', '=EOMONTH($B$4,0)', fmt=DATEF,
    font=Font(name='微软雅黑', size=10, color='7F7F7F'), fill=F_AUTO, align=CEN)
put(ws, 'D3', '本月末参考', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BASE), align=CEN)
ws.column_dimensions['D'].width = 14
LISTS = [('E', '矿区', MINES, 12), ('G', '出库去向', OUTK, 16), ('I', '入库类型', INK, 14),
         ('K', '费用类型', FEEK, 18), ('M', '货源', SRC_, 12), ('O', '价格状态', PSTAT, 14),
         ('Q', '物料大类', CATS, 14), ('S', '计量单位', UNITS, 12), ('U', '人员', [], 14)]
for col, t, items, w in LISTS:
    put(ws, f'{col}3', t, font=FT_HDR, fill=PatternFill('solid', fgColor=C_BASE), align=CEN)
    ws.column_dimensions[col].width = w
    for i in range(max(len(items), 30)):
        c = ws[f'{col}{D0+i}']
        c.border, c.font, c.fill, c.alignment = BOX, FT_IN, F_IN, LEFT
        if i < len(items):
            c.value = items[i]
NAMES = {
    '汇率': '参数表!$B$4', '倍数A': '参数表!$B$5', '系数B': '参数表!$B$6',
    '账期起': '参数表!$B$7', '账期止': '参数表!$B$8', '结账锁': '参数表!$B$9',
    '一矿处理量': '参数表!$B$10', '二矿处理量': '参数表!$B$11',
    '矿区表': f'参数表!$E${D0}:$E${D0+29}', '去向表': f'参数表!$G${D0}:$G${D0+29}',
    '入库类型表': f'参数表!$I${D0}:$I${D0+29}', '费用类型表': f'参数表!$K${D0}:$K${D0+29}',
    '货源表': f'参数表!$M${D0}:$M${D0+29}', '价格状态表': f'参数表!$O${D0}:$O${D0+29}',
    '大类表': f'参数表!$Q${D0}:$Q${D0+29}', '单位表': f'参数表!$S${D0}:$S${D0+59}',
    '人员表': f'参数表!$U${D0}:$U${D0+29}',
}
ws.freeze_panes = 'A4'

# ══════════════════════════════════════════════════════════════════════
# 物料档案
# ══════════════════════════════════════════════════════════════════════
I_END = D0 + N_ITEM - 1
COLS_I = [('物料编码', 11, 'in', None), ('大类 ★', 13, 'in', None), ('品名 ★', 26, 'in', None),
          ('规格型号', 26, 'in', None), ('单位 ★', 9, 'in', None),
          ('重点物资', 10, 'in', None), ('安全库存', 10, 'in', QTY),
          ('最近货源', 10, 'auto', None), ('最近原币进价', 13, 'auto', CNY),
          ('建议售价·先令', 15, 'auto', TZS), ('价格状态', 12, 'auto', None),
          ('现有库存数量', 12, 'auto', QTY), ('现有库存金额·先令', 15, 'auto', TZS),
          ('最后出库日', 12, 'auto', DATEF), ('呆滞天数', 10, 'auto', '0'),
          ('备注', 20, 'in', None)]
ws = newsheet('物料档案', '物料档案 · 每样东西一行，全表靠「物料编码」认人',
              '跨表只用编码关联，不用品名 ——「钢球 Φ30」和「钢球-30」是同一样东西的两种写法，'
              '用名字关联必错。新物料在这里开一行，编码接着往下编。', C_BASE, len(COLS_I))
header(ws, COLS_I, C_BASE); body(ws, COLS_I, D0, I_END)
IN_R = lambda c: f'入库单!${c}${D0}:${c}${D0+N_IN-1}'
OUT_R = lambda c: f'出库单!${c}${D0}:${c}${D0+N_OUT-1}'
ST_R = lambda c: f'库存台账!${c}${S0}:${c}${S0+N_ITEM-1}'
for r in range(D0, I_END + 1):
    g = f'IF($A{r}="","",'
    put(ws, f'H{r}', f'={g}IFERROR(LOOKUP(2,1/(({IN_R("G")}=$A{r})*({IN_R("Y")}="正常")'
                     f'*({IN_R("L")}>0)),{IN_R("D")}),""))')
    put(ws, f'I{r}', f'={g}IFERROR(LOOKUP(2,1/(({IN_R("G")}=$A{r})*({IN_R("Y")}="正常")'
                     f'*({IN_R("L")}>0)),{IN_R("L")}),0))')
    put(ws, f'J{r}', f'={g}IF($H{r}="{SRC_[0]}",ROUND($I{r}*倍数A*汇率,0),'
                     f'IF($H{r}="{SRC_[1]}",ROUND($I{r}*系数B,0),0)))')
    put(ws, f'K{r}', f'={g}IF(COUNTIFS({IN_R("G")},$A{r},{IN_R("X")},"待补价")>0,"待补价",'
                     f'IF(COUNTIFS({IN_R("G")},$A{r},{IN_R("X")},"暂估")>0,"暂估",'
                     f'IF($I{r}=0,"没有进价","正常"))))')
    put(ws, f'L{r}', f'={g}IFERROR(VLOOKUP($A{r},库存台账!$A${S0}:$Y${S0+N_ITEM-1},24,0),0))')
    put(ws, f'M{r}', f'={g}IFERROR(VLOOKUP($A{r},库存台账!$A${S0}:$Y${S0+N_ITEM-1},25,0),0))')
    put(ws, f'N{r}', f'={g}IFERROR(LOOKUP(2,1/(({OUT_R("D")}=$A{r})*({OUT_R("T")}="正常")),'
                     f'{OUT_R("C")}),""))')
    put(ws, f'O{r}', f'={g}IF($L{r}<=0,"",IF($N{r}="",'
                     f'IFERROR(账期止-LOOKUP(2,1/({IN_R("G")}=$A{r}),{IN_R("C")}),""),'
                     f'账期止-$N{r})))')
for i, it in enumerate(DATA['items']):
    r = D0 + i
    for c, v in ((1, it['code']), (2, it['cat']), (3, it['name']), (4, it['spec']),
                 (5, it['unit']), (6, it['key'])):
        ws.cell(r, c).value = v
dv(ws, '=大类表', f'B{D0}:B{I_END}')
dv(ws, '=单位表', f'E{D0}:E{I_END}')
dv(ws, '"是,否"', f'F{D0}:F{I_END}')
ws.conditional_formatting.add(f'O{D0}:O{I_END}', CellIsRule(
    operator='greaterThan', formula=['180'], font=Font(color='C00000', bold=True)))
ws.conditional_formatting.add(f'F{D0}:F{I_END}', CellIsRule(
    operator='equal', formula=['"是"'], font=Font(color='C00000', bold=True)))
ws.freeze_panes = f'C{D0}'
ws.auto_filter.ref = f'A{HDR}:P{I_END}'

# ══════════════════════════════════════════════════════════════════════
# 费用台账
# ══════════════════════════════════════════════════════════════════════
F_END = D0 + N_FEE - 1
COLS_F = [('序号', 7, 'auto', None), ('费用日期 ★', 12, 'in', DATEF), ('批次号 ★', 24, 'in', None),
          ('费用类型 ★', 15, 'in', None), ('供应商', 22, 'in', None),
          ('币种 ★', 9, 'in', None), ('原币金额 ★', 14, 'in', CNY),
          ('汇率', 9, 'auto', '0'), ('金额·先令', 15, 'auto', TZS),
          ('参与分摊', 10, 'in', None), ('状态', 9, 'in', None), ('备注', 30, 'in', None)]
ws = newsheet('费用台账', '费用台账 · 海运费、装柜费、到港费都记这里（财务填）',
              '费用**绝不能**混进【入库单】—— 混在一起分摊基数会把费用自己算进去，越算越错。'
              '国内那段填 CNY、坦桑那段填 TZS，系统自动折先令再按批次分摊。', C_IN, len(COLS_F))
header(ws, COLS_F, C_IN); body(ws, COLS_F, D0, F_END)
for r in range(D0, F_END + 1):
    g = f'IF($C{r}="","",'
    put(ws, f'A{r}', f'={g}COUNTA($C${D0}:$C{r}))')
    put(ws, f'H{r}', f'={g}IF($F{r}="CNY",汇率,1))')
    put(ws, f'I{r}', f'={g}ROUND(N($G{r})*N($H{r}),0))')
for i, x in enumerate(DATA['fees']):
    r = D0 + i
    cat = ('国内海运费' if '海运' in x['name'] else '报关服务费' if '服务费' in x['name']
           else '装柜费' if x['hint'] == '装柜费用' else '国内内陆运费' if x['hint'] == '运费'
           else '其他费用')
    for c, v in ((2, datetime.datetime.strptime(x['date'], '%Y-%m-%d') if x['date'] else None),
                 (3, x['batch']), (4, cat), (5, x['supplier'] or None), (6, 'CNY'),
                 (7, x['amount']), (10, '是'), (11, '正常'),
                 (12, f"原表第 {x['row']} 行：{x['name']}")):
        ws.cell(r, c).value = v
dv(ws, '=费用类型表', f'D{D0}:D{F_END}')
dv(ws, '"CNY,TZS"', f'F{D0}:F{F_END}', warn=False)
dv(ws, '"是,否"', f'J{D0}:J{F_END}', warn=False)
dv(ws, '"正常,作废"', f'K{D0}:K{F_END}', warn=False)
ws.freeze_panes = f'C{D0}'
ws.auto_filter.ref = f'A{HDR}:L{F_END}'

# ══════════════════════════════════════════════════════════════════════
# 入库单
# ══════════════════════════════════════════════════════════════════════
IN_END = D0 + N_IN - 1
COLS_IN = [('序号', 7, 'auto', None), ('入库单号', 14, 'in', None), ('入库日期 ★', 12, 'in', DATEF),
           ('货源 ★', 10, 'in', None), ('批次号 ★', 24, 'in', None), ('供应商', 22, 'in', None),
           ('物料编码 ★', 11, 'in', None), ('品名', 24, 'auto', None), ('规格型号', 24, 'auto', None),
           ('单位', 9, 'auto', None), ('入库数量 ★', 12, 'in', QTY), ('原币单价', 13, 'in', CNY),
           ('币种', 8, 'auto', None), ('汇率', 8, 'auto', '0'), ('原币金额', 14, 'auto', CNY),
           ('货值·先令', 15, 'auto', TZS), ('本批货值合计·先令', 16, 'auto', TZS),
           ('本批可摊费用·先令', 16, 'auto', TZS), ('本批无价行数', 12, 'auto', '0'),
           ('累计货值·先令', 16, 'auto', TZS), ('分摊运费·先令', 15, 'auto', TZS),
           ('入库成本·先令', 16, 'auto', TZS), ('单位成本·先令', 15, 'auto', TZS),
           ('入库类型', 12, 'in', None), ('价格状态', 12, 'in', None), ('状态', 9, 'in', None),
           ('录入人', 10, 'in', None), ('备注', 26, 'in', None), ('核对', 26, 'auto', None)]
ws = newsheet('入库单', '入库单 · 库管录数量，财务补单价',
              '库管只填：日期、批次号、物料编码、入库数量、录入人。单价财务后补，不耽误收货。'
              '批次号 = 集装箱号-到港年月（柜号会被船公司重复使用，光写柜号两年后会摊错）。'
              '当地采购批次号写「本地-年月」。', C_IN, len(COLS_IN))
header(ws, COLS_IN, C_IN); body(ws, COLS_IN, D0, IN_END)
for r in range(D0, IN_END + 1):
    g = f'IF($G{r}="","",'
    put(ws, f'A{r}', f'={g}COUNTA($G${D0}:$G{r}))')
    for col, idx in (('H', 3), ('I', 4), ('J', 5)):
        put(ws, f'{col}{r}', f'={g}IFERROR(VLOOKUP($G{r},物料档案!$A${D0}:$P${I_END},{idx},0),"※编码不在物料档案里"))')
    put(ws, f'M{r}', f'={g}IF($D{r}="{SRC_[0]}","CNY","TZS"))')
    put(ws, f'N{r}', f'={g}IF($D{r}="{SRC_[0]}",汇率,1))')
    put(ws, f'O{r}', f'={g}ROUND(N($K{r})*N($L{r}),2))')
    put(ws, f'P{r}', f'={g}ROUND($O{r}*$N{r},0))')
    put(ws, f'Q{r}', f'={g}SUMIFS($P${D0}:$P${IN_END},$E${D0}:$E${IN_END},$E{r},'
                     f'$Z${D0}:$Z${IN_END},"正常"))')
    put(ws, f'R{r}', f'={g}SUMIFS(费用台账!$I${D0}:$I${F_END},费用台账!$C${D0}:$C${F_END},$E{r},'
                     f'费用台账!$J${D0}:$J${F_END},"是",费用台账!$K${D0}:$K${F_END},"正常"))')
    put(ws, f'S{r}', f'={g}COUNTIFS($E${D0}:$E${IN_END},$E{r},$Z${D0}:$Z${IN_END},"正常",'
                     f'$Y${D0}:$Y${IN_END},"待补价"))')
    put(ws, f'T{r}', f'={g}SUMIFS($P${D0}:$P${IN_END},$E${D0}:$E${IN_END},$E{r},'
                     f'$Z${D0}:$Z${IN_END},"正常",$A${D0}:$A${IN_END},"<="&$A{r}))')
    # 累计差额法：保证本批各行分摊之和 = 本批费用，一分不差。
    # 本批只要还有没填单价的行就整批挂账不摊 —— 摊了会把运费全压到有价的那几行上。
    put(ws, f'U{r}', f'={g}IF(OR($S{r}>0,$Q{r}=0,$R{r}=0,$Z{r}<>"正常"),0,'
                     f'ROUND($T{r}/$Q{r}*$R{r},0)-ROUND(($T{r}-$P{r})/$Q{r}*$R{r},0)))')
    put(ws, f'V{r}', f'={g}ROUND($P{r}+$U{r},0))')
    put(ws, f'W{r}', f'={g}IF(N($K{r})=0,0,ROUND($V{r}/$K{r},2)))')
    put(ws, f'AC{r}', f'={g}'
        f'IF($H{r}="※编码不在物料档案里","※物料编码不在档案里",'
        f'IF(N($K{r})<=0,"※入库数量要大于 0",'
        f'IF(AND($C{r}<>"",$C{r}<=结账锁),"※日期早于已结账截止日，不能再录",'
        f'IF(AND($L{r}=0,$X{r}="正常"),"※没填单价，价格状态请选「待补价」",'
        f'IF(AND($S{r}>0,$R{r}>0),"△本批有 "&$S{r}&" 行没单价，整批运费先挂账没摊",'
        f'IF(AND($B{r}<>"",COUNTIF($B${D0}:$B${IN_END},$B{r})>1),"※入库单号重复","")))))))')
for i, x in enumerate(DATA['inbound']):
    r = D0 + i
    for c, v in ((3, datetime.datetime.strptime(x['date'], '%Y-%m-%d') if x['date'] else None),
                 (4, x['source']), (5, x['batch']), (6, x['supplier'] or None), (7, x['code']),
                 (11, x['qty']), (12, x['price'] or None), (24, '采购入库'),
                 (25, '正常' if x['price'] else '待补价'), (26, '正常'),
                 (28, (f"原表第 {x['row']} 行" + (('　' + x['note']) if x['note'] else '')))):
        ws.cell(r, c).value = v
dv(ws, '=货源表', f'D{D0}:D{IN_END}', warn=False)
dv(ws, '=物料档案表', f'G{D0}:G{IN_END}')
dv(ws, '=入库类型表', f'X{D0}:X{IN_END}'.replace('X', 'X'), warn=False)
dv(ws, '=价格状态表', f'Y{D0}:Y{IN_END}', warn=False)
dv(ws, '"正常,作废"', f'Z{D0}:Z{IN_END}', warn=False)
dv(ws, '=人员表', f'AA{D0}:AA{IN_END}')
redflag(ws, f'AC{D0}:AC{IN_END}'); redflag(ws, f'H{D0}:H{IN_END}')
ws.freeze_panes = f'H{D0}'
ws.auto_filter.ref = f'A{HDR}:AC{IN_END}'

# ══════════════════════════════════════════════════════════════════════
# 出库单
# ══════════════════════════════════════════════════════════════════════
O_END = D0 + N_OUT - 1
COLS_O = [('序号', 7, 'auto', None), ('出库单号', 14, 'in', None), ('出库日期 ★', 12, 'in', DATEF),
          ('物料编码 ★', 11, 'in', None), ('品名', 24, 'auto', None), ('规格型号', 24, 'auto', None),
          ('单位', 9, 'auto', None), ('大类', 13, 'auto', None), ('重点物资', 10, 'auto', None),
          ('出库数量 ★', 12, 'in', QTY), ('去向类型 ★', 14, 'in', None), ('去向矿区 ★', 12, 'in', None),
          ('领用人/设备号', 15, 'in', None), ('库管 ★', 10, 'in', None),
          ('建议售价·先令', 14, 'auto', TZS), ('实际售价·先令', 14, 'in', TZS),
          ('销售金额·先令', 15, 'auto', TZS), ('单位成本·先令', 14, 'auto', TZS),
          ('出库成本·先令', 15, 'auto', TZS), ('毛利·先令', 15, 'auto', TZS),
          ('状态', 9, 'in', None), ('备注', 24, 'in', None), ('核对', 28, 'auto', None)]
ws = newsheet('出库单', '出库单 · 库管每天就开这一张',
              '一矿区/二矿区 = 自用消耗（只出成本）；三矿区/六矿区 = 对外销售（出成本也出收入）。'
              '「去向类型」这一列是老板月报能不能分清「自己耗的」和「卖出去的」的总开关，必须填对。'
              '盘盈不在这里录，去【入库单】选「盘盈」。', C_IN, len(COLS_O))
header(ws, COLS_O, C_IN); body(ws, COLS_O, D0, O_END)
for r in range(D0, O_END + 1):
    g = f'IF($D{r}="","",'
    put(ws, f'A{r}', f'={g}COUNTA($D${D0}:$D{r}))')
    for col, idx in (('E', 3), ('F', 4), ('G', 5), ('H', 2), ('I', 6)):
        put(ws, f'{col}{r}', f'={g}IFERROR(VLOOKUP($D{r},物料档案!$A${D0}:$P${I_END},{idx},0),'
                             f'"※编码不在物料档案里"))')
    put(ws, f'O{r}', f'={g}IFERROR(VLOOKUP($D{r},物料档案!$A${D0}:$P${I_END},10,0),0))')
    put(ws, f'Q{r}', f'={g}IF($K{r}<>"{OUTK[1]}",0,ROUND(N($J{r})*IF($P{r}="",N($O{r}),N($P{r})),0)))')
    put(ws, f'R{r}', f'={g}IFERROR(VLOOKUP($D{r},库存台账!$A${S0}:$Y${S0+N_ITEM-1},11,0),0))')
    put(ws, f'S{r}', f'={g}ROUND(N($J{r})*N($R{r}),0))')
    put(ws, f'T{r}', f'={g}IF($K{r}<>"{OUTK[1]}","",ROUND($Q{r}-$S{r},0)))')
    put(ws, f'W{r}', f'={g}'
        f'IF($E{r}="※编码不在物料档案里","※物料编码不在档案里",'
        f'IF(N($J{r})<=0,"※出库数量要大于 0",'
        f'IF($K{r}="","※没填去向类型",'
        f'IF($L{r}="","※没填去向矿区",'
        f'IF(AND($C{r}<>"",$C{r}<=结账锁),"※日期早于已结账截止日，不能再录",'
        f'IF(AND($K{r}="{OUTK[0]}",COUNTIF({{"{OWN[0]}";"{OWN[1]}"}},$L{r})=0),'
        f'"※自用消耗的去向只能是一矿区或二矿区",'
        f'IF(AND($K{r}="{OUTK[1]}",COUNTIF({{"{SELL[0]}";"{SELL[1]}"}},$L{r})=0),'
        f'"※对外销售的去向只能是三矿区或六矿区",'
        f'IF(AND($K{r}="{OUTK[1]}",ROUND($Q{r},0)=0),"※外销却没有售价，请填实际售价",'
        f'IF(AND($K{r}="{OUTK[1]}",$T{r}<>"",$T{r}<0),"※这一笔外销是亏的，请核对售价",'
        f'IF(IFERROR(VLOOKUP($D{r},库存台账!$A${S0}:$Y${S0+N_ITEM-1},24,0),0)<-0.001,'
        f'"※这个物料账面已经变成负数了，查一下是不是漏记入库","")))))))))))')
for i, x in enumerate(DATA.get('outbound', [])):
    r = D0 + i
    for c, v in ((3, datetime.datetime.strptime(x['date'], '%Y-%m-%d') if x['date'] else None),
                 (4, x['code']), (10, x['qout']), (11, OUTK[1]), (12, SELL[0]),
                 (21, '正常'), (22, f"原表第 {x['row']} 行：{x['supplier'] or x['name']}")):
        ws.cell(r, c).value = v
dv(ws, '=物料档案表', f'D{D0}:D{O_END}')
dv(ws, '=去向表', f'K{D0}:K{O_END}', warn=False)
dv(ws, '=矿区表', f'L{D0}:L{O_END}', warn=False)
dv(ws, '=人员表', f'N{D0}:N{O_END}')
dv(ws, '"正常,作废"', f'U{D0}:U{O_END}', warn=False)
redflag(ws, f'W{D0}:W{O_END}'); redflag(ws, f'E{D0}:E{O_END}')
ws.conditional_formatting.add(f'I{D0}:I{O_END}', CellIsRule(
    operator='equal', formula=['"是"'], font=Font(color='C00000', bold=True)))
ws.conditional_formatting.add(f'T{D0}:T{O_END}', CellIsRule(
    operator='lessThan', formula=['0'], font=Font(color='C00000', bold=True)))
ws.freeze_panes = f'E{D0}'
ws.auto_filter.ref = f'A{HDR}:W{O_END}'

# ══════════════════════════════════════════════════════════════════════
# 到货批次（自动，看每个柜摊得对不对）
# ══════════════════════════════════════════════════════════════════════
B_END = S0 + N_BAT - 1
COLS_B = [('批次号', 26, 'auto', None), ('集装箱号', 18, 'in', None), ('到港日期', 12, 'in', DATEF),
          ('货源', 10, 'auto', None), ('物料行数', 10, 'auto', '0'), ('没填单价行数', 12, 'auto', '0'),
          ('货值·先令', 16, 'auto', TZS), ('费用·先令', 15, 'auto', TZS),
          ('运费率', 10, 'auto', PCT), ('已摊运费·先令', 16, 'auto', TZS),
          ('未摊(挂账)·先令', 16, 'auto', TZS), ('入库总成本·先令', 16, 'auto', TZS),
          ('状态', 30, 'auto', None)]
ws = newsheet('到货批次', '到货批次 · 自动，一个柜一行，看运费摊得对不对',
              '「运费率」差得越多越说明必须按柜摊：真实数据里 TIIU4204331 柜 22.28%、'
              'DRYU9381826 柜 5.75%，差快 4 倍，用一个统一费率去摊必错。', C_AUTO, len(COLS_B))
header(ws, COLS_B, C_AUTO); body(ws, COLS_B, S0 - 1, B_END)
for i in range(N_BAT):
    r = S0 + i
    g = f'IF($A{r}="","",'
    put(ws, f'D{r}', f'={g}IFERROR(LOOKUP(2,1/(入库单!$E${D0}:$E${IN_END}=$A{r}),'
                     f'入库单!$D${D0}:$D${IN_END}),""))')
    put(ws, f'E{r}', f'={g}COUNTIFS(入库单!$E${D0}:$E${IN_END},$A{r},入库单!$Z${D0}:$Z${IN_END},"正常"))')
    put(ws, f'F{r}', f'={g}COUNTIFS(入库单!$E${D0}:$E${IN_END},$A{r},入库单!$Z${D0}:$Z${IN_END},"正常",'
                     f'入库单!$Y${D0}:$Y${IN_END},"待补价"))')
    put(ws, f'G{r}', f'={g}ROUND(SUMIFS(入库单!$P${D0}:$P${IN_END},入库单!$E${D0}:$E${IN_END},$A{r},'
                     f'入库单!$Z${D0}:$Z${IN_END},"正常"),0))')
    put(ws, f'H{r}', f'={g}ROUND(SUMIFS(费用台账!$I${D0}:$I${F_END},费用台账!$C${D0}:$C${F_END},$A{r},'
                     f'费用台账!$J${D0}:$J${F_END},"是",费用台账!$K${D0}:$K${F_END},"正常"),0))')
    put(ws, f'I{r}', f'={g}IF($G{r}=0,"",ROUND($H{r}/$G{r},4)))')
    put(ws, f'J{r}', f'={g}ROUND(SUMIFS(入库单!$U${D0}:$U${IN_END},入库单!$E${D0}:$E${IN_END},$A{r},'
                     f'入库单!$Z${D0}:$Z${IN_END},"正常"),0))')
    put(ws, f'K{r}', f'={g}ROUND($H{r}-$J{r},0))')
    put(ws, f'L{r}', f'={g}ROUND(SUMIFS(入库单!$V${D0}:$V${IN_END},入库单!$E${D0}:$E${IN_END},$A{r},'
                     f'入库单!$Z${D0}:$Z${IN_END},"正常"),0))')
    put(ws, f'M{r}', f'={g}IF($E{r}=0,"※这个批次号在入库单里找不到",'
                     f'IF(AND($H{r}>0,$F{r}>0),"△有 "&$F{r}&" 行没单价，运费整批挂账没摊 —— 补完价自动摊",'
                     f'IF($H{r}=0,"△还没录运费（27 个柜里只有 2 个录了，其余的要补）",'
                     f'IF(ABS($K{r})>1,"※摊完还差 "&TEXT($K{r},"#,##0")&" 先令，查一下","已摊平")))))')
for i, b in enumerate(DATA['batches']):
    r = S0 + i
    ws.cell(r, 1).value = b['batch']
    ws.cell(r, 2).value = b['container'] or None
    ws.cell(r, 3).value = datetime.datetime.strptime(b['date'], '%Y-%m-%d') if b['date'] else None
put(ws, f'A{S0-1}', '合计', font=FT_SUM, fill=F_SUM, align=CEN)
for c in 'EFGHJKL':
    put(ws, f'{c}{S0-1}', f'=ROUND(SUM({c}{S0}:{c}{B_END}),0)',
        fmt=('0' if c in 'EF' else TZS), font=FT_SUM, fill=F_SUM, align=CEN)
redflag(ws, f'M{S0}:M{B_END}')
ws.freeze_panes = f'B{S0}'
ws.auto_filter.ref = f'A{HDR}:M{B_END}'

# ══════════════════════════════════════════════════════════════════════
# 库存台账
# ══════════════════════════════════════════════════════════════════════
ST_END = S0 + N_ITEM - 1
COLS_S = [('物料编码', 11, 'auto', None), ('品名', 24, 'auto', None), ('规格型号', 22, 'auto', None),
          ('单位', 8, 'auto', None), ('大类', 13, 'auto', None), ('重点', 8, 'auto', None),
          ('期初数量', 11, 'in', QTY), ('期初金额·先令', 15, 'in', TZS),
          ('本月入库数量', 12, 'auto', QTY), ('本月入库金额·先令', 16, 'auto', TZS),
          ('月加权单位成本·先令', 17, 'auto', TZS),
          ('盘盈数量', 10, 'auto', QTY), ('盘盈金额', 13, 'auto', TZS),
          ('矿区退库数量', 12, 'auto', QTY), ('退库金额', 13, 'auto', TZS),
          ('自用消耗数量', 12, 'auto', QTY), ('自用消耗金额', 14, 'auto', TZS),
          ('对外销售数量', 12, 'auto', QTY), ('对外销售成本', 14, 'auto', TZS),
          ('盘亏数量', 10, 'auto', QTY), ('盘亏金额', 13, 'auto', TZS),
          ('退货数量', 10, 'auto', QTY), ('退货金额', 13, 'auto', TZS),
          ('期末数量', 12, 'auto', QTY), ('期末金额·先令', 16, 'auto', TZS),
          ('成本可信度', 12, 'auto', None), ('异常', 30, 'auto', None)]
ws = newsheet('库存台账', '库存台账 · 一样东西一行，全套报表都从这里取数',
              '期末 = 期初 + 采购入库 + 盘盈 + 退库 − 自用消耗 − 对外销售 − 盘亏 − 退货。'
              '这一句在本表、月报、核对表三处用的是同一个式子。'
              '月加权单位成本只看「期初 + 本月入库」，不看出库 —— 这样才不会循环引用。'
              '【月初开账】把 X、Y 两列复制→选择性粘贴为数值→贴到 G、H 两列，再改参数表的账期。',
              C_AUTO, len(COLS_S))
header(ws, COLS_S, C_AUTO); body(ws, COLS_S, S0 - 1, ST_END)
PER = f'入库单!$C${D0}:$C${IN_END},">="&账期起,入库单!$C${D0}:$C${IN_END},"<="&账期止'
PERO = f'出库单!$C${D0}:$C${O_END},">="&账期起,出库单!$C${D0}:$C${O_END},"<="&账期止'
def sin_(valcol, kind):
    return (f'SUMIFS(入库单!${valcol}${D0}:${valcol}${IN_END},入库单!$G${D0}:$G${IN_END},$A{{r}},'
            f'入库单!$X${D0}:$X${IN_END},"{kind}",入库单!$Z${D0}:$Z${IN_END},"正常",{PER})')
def sout_(valcol, kind):
    return (f'SUMIFS(出库单!${valcol}${D0}:${valcol}${O_END},出库单!$D${D0}:$D${O_END},$A{{r}},'
            f'出库单!$K${D0}:$K${O_END},"{kind}",出库单!$U${D0}:$U${O_END},"正常",{PERO})')
for i in range(N_ITEM):
    r = S0 + i
    g = f'IF($A{r}="","",'
    put(ws, f'A{r}', f'=IF(物料档案!$A{D0+i}="","",物料档案!$A{D0+i})')
    for col, idx in (('B', 3), ('C', 4), ('D', 5), ('E', 2), ('F', 6)):
        put(ws, f'{col}{r}', f'={g}IFERROR(VLOOKUP($A{r},物料档案!$A${D0}:$P${I_END},{idx},0),""))')
    put(ws, f'I{r}', f'={g}' + sin_('K', '采购入库').format(r=r) + ')')
    put(ws, f'J{r}', f'={g}' + sin_('V', '采购入库').format(r=r) + ')')
    put(ws, f'K{r}', f'={g}IFERROR(ROUND((N($H{r})+N($J{r}))/(N($G{r})+N($I{r})),2),0))')
    put(ws, f'L{r}', f'={g}' + sin_('K', '盘盈').format(r=r) + ')')
    put(ws, f'M{r}', f'={g}ROUND($L{r}*$K{r},0))')
    put(ws, f'N{r}', f'={g}' + sin_('K', '矿区退库').format(r=r) + ')')
    put(ws, f'O{r}', f'={g}ROUND($N{r}*$K{r},0))')
    put(ws, f'P{r}', f'={g}' + sout_('J', '自用消耗').format(r=r) + ')')
    put(ws, f'Q{r}', f'={g}ROUND(' + sout_('S', '自用消耗').format(r=r) + ',0))')
    put(ws, f'R{r}', f'={g}' + sout_('J', '对外销售').format(r=r) + ')')
    put(ws, f'S{r}', f'={g}ROUND(' + sout_('S', '对外销售').format(r=r) + ',0))')
    put(ws, f'T{r}', f'={g}' + sout_('J', '盘亏').format(r=r) + ')')
    put(ws, f'U{r}', f'={g}ROUND(' + sout_('S', '盘亏').format(r=r) + ',0))')
    put(ws, f'V{r}', f'={g}' + sout_('J', '退货给供应商').format(r=r) + ')')
    put(ws, f'W{r}', f'={g}ROUND(' + sout_('S', '退货给供应商').format(r=r) + ',0))')
    put(ws, f'X{r}', f'={g}ROUND(N($G{r})+$I{r}+$L{r}+$N{r}-$P{r}-$R{r}-$T{r}-$V{r},3))')
    put(ws, f'Y{r}', f'={g}ROUND(N($H{r})+$J{r}+$M{r}+$O{r}-$Q{r}-$S{r}-$U{r}-$W{r},0))')
    put(ws, f'Z{r}', f'={g}IF(COUNTIFS(入库单!$G${D0}:$G${IN_END},$A{r},'
                     f'入库单!$Y${D0}:$Y${IN_END},"待补价")+COUNTIFS(入库单!$G${D0}:$G${IN_END},$A{r},'
                     f'入库单!$Y${D0}:$Y${IN_END},"暂估")>0,"不可信","可信"))')
    put(ws, f'AA{r}', f'={g}'
        f'IF($X{r}<-0.001,"※账面负库存，查一下是不是漏记入库",'
        f'IF(AND(ROUND($X{r},3)=0,ABS($Y{r})>1),"※数量已经是 0，金额还挂着 "&TEXT($Y{r},"#,##0")&" 先令",'
        f'IF(AND($X{r}>0,$Y{r}<=0,$K{r}=0),"※有数量没金额（这个物料一直没填过进价）",'
        f'IF($Z{r}="不可信","△成本含待补价/暂估，毛利仅供参考","")))))')
put(ws, f'A{S0-1}', '合计', font=FT_SUM, fill=F_SUM, align=CEN)
for c in ('G', 'H', 'I', 'J', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y'):
    put(ws, f'{c}{S0-1}', f'=ROUND(SUM({c}{S0}:{c}{ST_END}),2)', fmt=TZS,
        font=FT_SUM, fill=F_SUM, align=CEN)
redflag(ws, f'AA{S0}:AA{ST_END}')
ws.conditional_formatting.add(f'X{S0}:X{ST_END}', CellIsRule(
    operator='lessThan', formula=['0'], font=Font(color='C00000', bold=True)))
ws.freeze_panes = f'C{S0}'
ws.auto_filter.ref = f'A{HDR}:AA{ST_END}'

# ══════════════════════════════════════════════════════════════════════
# 月度盘点（盲盘）
# ══════════════════════════════════════════════════════════════════════
P_END = S0 + N_ITEM - 1
COLS_P = [('物料编码', 11, 'auto', None), ('品名', 26, 'auto', None), ('规格型号', 24, 'auto', None),
          ('单位', 8, 'auto', None), ('大类', 13, 'auto', None), ('重点', 8, 'auto', None),
          ('存放位置', 12, 'in', None), ('实盘数量', 12, 'in', QTY),
          ('账面数量', 12, 'auto', QTY), ('差异数量', 12, 'auto', QTY),
          ('单位成本·先令', 14, 'auto', TZS), ('差异金额·先令', 15, 'auto', TZS),
          ('差异率', 10, 'auto', PCT), ('差异原因', 14, 'in', None), ('处理方式', 12, 'in', None),
          ('是否已调账', 14, 'auto', None), ('盘点人', 10, 'in', None), ('复核人', 10, 'in', None)]
ws = newsheet('月度盘点', '月度盘点 · 盲盘（打印给库管的那一版只印到 H 列，看不到账面数）',
              '炸药、柴油、药剂、钢球这些重点物资每月必须全盘、两个人签字；其余可以按大类轮盘，'
              '但一年内要盘一遍。盘出差异**不许直接改库存台账** —— 盘盈去【入库单】记一行「盘盈」，'
              '盘亏去【出库单】记一行「盘亏」，库存永远是算出来的，不是手改的。', C_AUTO, len(COLS_P))
header(ws, COLS_P, C_AUTO); body(ws, COLS_P, S0, P_END)
REASON = ['盘盈', '盘亏', '记录错误', '自然损耗', '被盗', '单位记错', '未查明']
HANDLE = ['调账', '追责', '待查']
for i in range(N_ITEM):
    r, sr = S0 + i, S0 + i
    g = f'IF($A{r}="","",'
    put(ws, f'A{r}', f'=IF(库存台账!$A{sr}="","",库存台账!$A{sr})')
    for col, sc in (('B', 'B'), ('C', 'C'), ('D', 'D'), ('E', 'E'), ('F', 'F')):
        put(ws, f'{col}{r}', f'={g}库存台账!${sc}{sr})')
    put(ws, f'I{r}', f'={g}库存台账!$X{sr})')
    put(ws, f'J{r}', f'=IF(OR($A{r}="",$H{r}=""),"",ROUND($H{r}-$I{r},3))')
    put(ws, f'K{r}', f'={g}库存台账!$K{sr})')
    put(ws, f'L{r}', f'=IF($J{r}="","",ROUND($J{r}*$K{r},0))')
    put(ws, f'M{r}', f'=IF(OR($J{r}="",$I{r}=0),"",ROUND($J{r}/$I{r},4))')
    put(ws, f'P{r}', f'=IF($J{r}="","未盘",IF(ROUND($J{r},3)=0,"账实相符",'
                     f'IF(COUNTIFS(入库单!$G${D0}:$G${IN_END},$A{r},入库单!$X${D0}:$X${IN_END},"盘盈",'
                     f'{PER})+COUNTIFS(出库单!$D${D0}:$D${O_END},$A{r},'
                     f'出库单!$K${D0}:$K${O_END},"盘亏",{PERO})>0,"已调账","※差异还没调账")))')
dv(ws, f'"{",".join(REASON)}"', f'N{S0}:N{P_END}', warn=False)
dv(ws, f'"{",".join(HANDLE)}"', f'O{S0}:O{P_END}', warn=False)
redflag(ws, f'P{S0}:P{P_END}')
ws.conditional_formatting.add(f'J{S0}:J{P_END}', CellIsRule(
    operator='notEqual', formula=['0'], font=Font(color='C00000', bold=True),
    fill=PatternFill('solid', fgColor='FFE2E2')))
ws.print_area = f'A1:H{P_END}'
ws.freeze_panes = f'C{S0}'
ws.auto_filter.ref = f'A{HDR}:R{P_END}'

# ══════════════════════════════════════════════════════════════════════
# 应收对账
# ══════════════════════════════════════════════════════════════════════
AR_END = S0 + 19
COLS_AR = [('矿区/客户', 14, 'auto', None), ('期初欠款·先令', 16, 'in', TZS),
           ('本月销售额·先令', 16, 'auto', TZS), ('本月收款·先令', 16, 'in', TZS),
           ('期末欠款·先令', 16, 'auto', TZS), ('本月销售成本·先令', 16, 'auto', TZS),
           ('本月毛利·先令', 15, 'auto', TZS), ('毛利率', 10, 'auto', PCT),
           ('本月笔数', 10, 'auto', '0'), ('备注', 26, 'in', None)]
ws = newsheet('应收对账', '应收对账 · 卖给三矿六矿的货款收没收回来',
              '货出去了钱没回来，比库管拿点东西金额大得多。期初欠款和本月收款要人填，其余自动。'
              '月初开账时把「期末欠款」抄到下个月的「期初欠款」。', C_AUTO, len(COLS_AR))
header(ws, COLS_AR, C_AUTO); body(ws, COLS_AR, S0 - 1, AR_END)
for i in range(20):
    r = S0 + i
    g = f'IF($A{r}="","",'
    put(ws, f'A{r}', f'=IF(参数表!$E{D0+i}="","",参数表!$E{D0+i})')
    put(ws, f'C{r}', f'={g}ROUND(SUMIFS(出库单!$Q${D0}:$Q${O_END},出库单!$L${D0}:$L${O_END},$A{r},'
                     f'出库单!$K${D0}:$K${O_END},"{OUTK[1]}",出库单!$U${D0}:$U${O_END},"正常",{PERO}),0))')
    put(ws, f'E{r}', f'={g}ROUND(N($B{r})+$C{r}-N($D{r}),0))')
    put(ws, f'F{r}', f'={g}ROUND(SUMIFS(出库单!$S${D0}:$S${O_END},出库单!$L${D0}:$L${O_END},$A{r},'
                     f'出库单!$K${D0}:$K${O_END},"{OUTK[1]}",出库单!$U${D0}:$U${O_END},"正常",{PERO}),0))')
    put(ws, f'G{r}', f'={g}ROUND($C{r}-$F{r},0))')
    put(ws, f'H{r}', f'={g}IF($C{r}=0,"",ROUND($G{r}/$C{r},4)))')
    put(ws, f'I{r}', f'={g}COUNTIFS(出库单!$L${D0}:$L${O_END},$A{r},'
                     f'出库单!$K${D0}:$K${O_END},"{OUTK[1]}",出库单!$U${D0}:$U${O_END},"正常",{PERO}))')
put(ws, f'A{S0-1}', '合计', font=FT_SUM, fill=F_SUM, align=CEN)
for c in 'BCDEFG':
    put(ws, f'{c}{S0-1}', f'=ROUND(SUM({c}{S0}:{c}{AR_END}),0)', fmt=TZS, font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'H{S0-1}', f'=IF($C{S0-1}=0,"",ROUND($G{S0-1}/$C{S0-1},4))', fmt=PCT,
    font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'I{S0-1}', f'=SUM(I{S0}:I{AR_END})', fmt='0', font=FT_SUM, fill=F_SUM, align=CEN)
ws.conditional_formatting.add(f'E{S0}:E{AR_END}', CellIsRule(
    operator='greaterThan', formula=['0'], font=Font(color='C00000', bold=True)))
ws.freeze_panes = f'B{S0}'

# ══════════════════════════════════════════════════════════════════════
# 领导月报
# ══════════════════════════════════════════════════════════════════════
ws = newsheet('领导月报', '领导月报 · 老板看这一张就够（自动，不用填）',
              '改【参数表】的「本期账期起」就能翻月。金额全是先令，括号里是折人民币。', C_BOSS, 8)
for c, w in (('A', 34), ('B', 18), ('C', 18), ('D', 18), ('E', 18), ('F', 18), ('G', 16), ('H', 30)):
    ws.column_dimensions[c].width = w
ST = lambda c: f'库存台账!${c}${S0}:${c}${ST_END}'
STT = lambda c: f'库存台账!${c}${S0-1}'
OUTQ = lambda c: f'出库单!${c}${D0}:${c}${O_END}'
PERO2 = PERO

def sec(r, t):
    put(ws, f'A{r}', t, font=Font(name='微软雅黑', size=12, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor=C_BOSS), align=LEFT)
    ws.merge_cells(f'A{r}:H{r}')
    ws.row_dimensions[r].height = 24

def row(r, lab, *vals, fmt=TZS, bold=False, note=''):
    f = Font(name='微软雅黑', size=11 if bold else 10, bold=bold,
             color='C00000' if bold else '000000')
    put(ws, f'A{r}', lab, font=Font(name='微软雅黑', size=10, bold=bold), align=LEFT)
    for j, v in enumerate(vals):
        put(ws, f'{CL(2+j)}{r}', v, fmt=fmt, font=f, fill=(F_SUM if bold else F_AUTO), align=CEN)
    if note:
        put(ws, f'{CL(2+len(vals))}{r}', note,
            font=Font(name='微软雅黑', size=9, color='808080'), align=LEFT)

put(ws, 'A3', '账期', font=Font(name='微软雅黑', size=10, bold=True), align=CEN)
put(ws, 'B3', '=账期起', fmt=DATEF, font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, 'C3', '至', font=Font(name='微软雅黑', size=10), align=CEN)
put(ws, 'D3', '=账期止', fmt=DATEF, font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, 'E3', '汇率', font=Font(name='微软雅黑', size=10, bold=True), align=CEN)
put(ws, 'F3', '=汇率', fmt='0', font=FT_SUM, fill=F_SUM, align=CEN)

sec(5, '一、这个月花了多少钱')
row(6, '本月到货批次数', f'=COUNTIFS(到货批次!$C${S0}:$C${B_END},">="&账期起,'
                        f'到货批次!$C${S0}:$C${B_END},"<="&账期止)', fmt='0')
row(7, '本月采购入库·货值（不含运费）',
    f'=ROUND(SUMIFS(入库单!$P${D0}:$P${IN_END},入库单!$X${D0}:$X${IN_END},"采购入库",'
    f'入库单!$Z${D0}:$Z${IN_END},"正常",{PER}),0)')
row(8, '　其中：国内货柜', f'=ROUND(SUMIFS(入库单!$P${D0}:$P${IN_END},入库单!$D${D0}:$D${IN_END},'
                         f'"{SRC_[0]}",入库单!$X${D0}:$X${IN_END},"采购入库",'
                         f'入库单!$Z${D0}:$Z${IN_END},"正常",{PER}),0)')
row(9, '　其中：坦桑当地采购', f'=ROUND(SUMIFS(入库单!$P${D0}:$P${IN_END},入库单!$D${D0}:$D${IN_END},'
                            f'"{SRC_[1]}",入库单!$X${D0}:$X${IN_END},"采购入库",'
                            f'入库单!$Z${D0}:$Z${IN_END},"正常",{PER}),0)')
row(10, '本月运杂费（已摊进成本）', f'=ROUND(SUMIFS(入库单!$U${D0}:$U${IN_END},'
                                 f'入库单!$X${D0}:$X${IN_END},"采购入库",'
                                 f'入库单!$Z${D0}:$Z${IN_END},"正常",{PER}),0)')
row(11, '本月运杂费（还挂着没摊）', f'=ROUND(到货批次!$K${S0-1},0)',
    note='本批还有没填单价的行，摊了会算错，所以先挂着')
row(12, '本月采购总成本', '=ROUND($B$7+$B$10,0)', bold=True)
row(13, '　折人民币（约）', '=ROUND($B$12/汇率,2)', fmt=CNY)

sec(15, '二、炸药·柴油·药剂这些命门物资')
put(ws, 'B16', '本月耗用金额', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
put(ws, 'C16', '一矿区', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
put(ws, 'D16', '二矿区', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
put(ws, 'E16', '卖给三矿六矿', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
put(ws, 'F16', '期末库存金额', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
put(ws, 'G16', '期末数量品种数', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
put(ws, 'A16', '大类', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
KEYCATS = ['炸药火工', '油料', '选矿药剂', '磨矿介质']
for i, cat in enumerate(KEYCATS):
    r = 17 + i
    put(ws, f'A{r}', cat, font=Font(name='微软雅黑', size=10, bold=True), align=LEFT)
    put(ws, f'B{r}', f'=ROUND(SUMIF({ST("E")},$A{r},{ST("Q")})+SUMIF({ST("E")},$A{r},{ST("S")}),0)',
        fmt=TZS, font=FT_SUM, fill=F_AUTO, align=CEN)
    for j, mine in enumerate(OWN):
        put(ws, f'{CL(3+j)}{r}', f'=ROUND(SUMIFS({OUTQ("S")},{OUTQ("H")},$A{r},{OUTQ("L")},"{mine}",'
                                 f'{OUTQ("K")},"{OUTK[0]}",{OUTQ("U")},"正常",{PERO2}),0)',
            fmt=TZS, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'E{r}', f'=ROUND(SUMIFS({OUTQ("Q")},{OUTQ("H")},$A{r},{OUTQ("K")},"{OUTK[1]}",'
                     f'{OUTQ("U")},"正常",{PERO2}),0)', fmt=TZS, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'F{r}', f'=ROUND(SUMIF({ST("E")},$A{r},{ST("Y")}),0)', fmt=TZS, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'G{r}', f'=COUNTIFS({ST("E")},$A{r},{ST("X")},">0")', fmt='0', font=FT_AUTO, fill=F_AUTO, align=CEN)
put(ws, 'A21', '⚠ 现有 867 行货柜数据里，炸药火工一行都没有，油料只有 2 台柴油发电机。'
               '这两样是你说的最主要成本，但它们是在坦桑当地买的，没进过这张货柜清单 —— '
               '要盯住就得从现在起把当地采购也录进【入库单】（货源选 B当地）。',
    font=Font(name='微软雅黑', size=9, color='C00000'), align=LEFT)
ws.merge_cells('A21:H21')
ws.row_dimensions[21].height = 30

sec(23, '三、一矿二矿自己耗了多少')
put(ws, 'A24', '矿区', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
for j, t in enumerate(['本月自用消耗·先令', '占两矿合计', '折人民币', '处理矿石量(吨)', '吨耗·先令']):
    put(ws, f'{CL(2+j)}24', t, font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
for i, mine in enumerate(OWN):
    r = 25 + i
    put(ws, f'A{r}', mine, font=Font(name='微软雅黑', size=10, bold=True), align=LEFT)
    put(ws, f'B{r}', f'=ROUND(SUMIFS({OUTQ("S")},{OUTQ("L")},"{mine}",{OUTQ("K")},"{OUTK[0]}",'
                     f'{OUTQ("U")},"正常",{PERO2}),0)', fmt=TZS, font=FT_SUM, fill=F_AUTO, align=CEN)
    put(ws, f'C{r}', f'=IF($B$27=0,"",ROUND($B{r}/$B$27,4))', fmt=PCT, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'D{r}', f'=ROUND($B{r}/汇率,2)', fmt=CNY, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'E{r}', f'={"一矿处理量" if i == 0 else "二矿处理量"}', fmt='#,##0.##',
        font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'F{r}', f'=IF(N($E{r})=0,"—",ROUND($B{r}/$E{r},0))', fmt=TZS,
        font=FT_AUTO, fill=F_AUTO, align=CEN)
row(27, '两矿合计自用消耗', f'=ROUND(SUM($B$25:$B$26),0)', bold=True)

sec(29, '四、卖给三矿六矿，赚了多少')
put(ws, 'A30', '矿区', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
for j, t in enumerate(['本月销售额·先令', '本月销售成本', '本月毛利', '毛利率', '期末还欠我们', '本月笔数']):
    put(ws, f'{CL(2+j)}30', t, font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
for i, mine in enumerate(SELL):
    r = 31 + i
    ar = S0 + MINES.index(mine)
    put(ws, f'A{r}', mine, font=Font(name='微软雅黑', size=10, bold=True), align=LEFT)
    for j, sc in enumerate(['C', 'F', 'G', 'H', 'E', 'I']):
        put(ws, f'{CL(2+j)}{r}', f'=应收对账!${sc}{ar}',
            fmt=(PCT if sc == 'H' else '0' if sc == 'I' else TZS),
            font=(FT_SUM if sc in 'CG' else FT_AUTO), fill=F_AUTO, align=CEN)
row(33, '两个矿区合计', '=ROUND($B$31+$B$32,0)', '=ROUND($C$31+$C$32,0)',
    '=ROUND($D$31+$D$32,0)', '=IF($B$33=0,"",ROUND($D$33/$B$33,4))',
    '=ROUND($F$31+$F$32,0)', bold=True)
ws['E33'].number_format = PCT
row(34, '　折人民币（销售额）', '=ROUND($B$33/汇率,2)', fmt=CNY)

sec(36, '五、库存还剩多少（金额勾稽，一行一行看得见）')
LINES = [('期初金额', f'={STT("H")}'), ('＋ 本月采购入库', f'={STT("J")}'),
         ('＋ 盘盈', f'={STT("M")}'), ('＋ 矿区退库', f'={STT("O")}'),
         ('－ 一矿二矿自用消耗', f'={STT("Q")}'), ('－ 卖给三矿六矿（成本）', f'={STT("S")}'),
         ('－ 盘亏', f'={STT("U")}'), ('－ 退货给供应商', f'={STT("W")}')]
for i, (lab, f) in enumerate(LINES):
    row(37 + i, lab, f)
row(45, '＝ 期末金额·先令', f'={STT("Y")}', bold=True)
row(46, '　勾稽差额（应该是 0）',
    f'=ROUND({STT("H")}+{STT("J")}+{STT("M")}+{STT("O")}-{STT("Q")}-{STT("S")}'
    f'-{STT("U")}-{STT("W")}-{STT("Y")},0)', bold=True, note='不是 0 就别信这个月的数')
row(47, '期末有库存的品种数', f'=COUNTIF({ST("X")},">0")', fmt='0')
row(48, '呆滞超过 180 天的品种数',
    f'=COUNTIFS(物料档案!$O${D0}:$O${I_END},">180")', fmt='0')
row(49, '成本不可信（含待补价/暂估）的品种数', f'=COUNTIF({ST("Z")},"不可信")', fmt='0')
row(50, '　这些品种占期末库存金额',
    f'=IF({STT("Y")}=0,"",ROUND(SUMIF({ST("Z")},"不可信",{ST("Y")})/{STT("Y")},4))', fmt=PCT)

sec(52, '六、这个月的账对不对')
row(53, '核对表通过项 / 总项数', '=核对表!$B$3&" / "&核对表!$C$3', fmt=None, bold=True)
row(54, '没通过的项', '=核对表!$D$3', fmt=None, note='去【核对表】看红的那几行')
ws.sheet_view.showGridLines = False

# ══════════════════════════════════════════════════════════════════════
# 核对表
# ══════════════════════════════════════════════════════════════════════
ws = newsheet('核对表', '核对表 · 结账前必看，全绿才能把月报发给老板', '', C_AUTO, 6)
for c, w in (('A', 40), ('B', 18), ('C', 18), ('D', 16), ('E', 12), ('F', 46)):
    ws.column_dimensions[c].width = w
put(ws, 'A3', '通过 / 总数 →', font=Font(name='微软雅黑', size=11, bold=True), align=CEN)
CHK_R0 = 6
CHECKS = [
    ('1a 金额勾稽：期初+入库+盘盈+退库−自用−外销−盘亏−退货 = 期末',
     f'=ROUND({STT("H")}+{STT("J")}+{STT("M")}+{STT("O")}-{STT("Q")}-{STT("S")}-{STT("U")}-{STT("W")},0)',
     f'=ROUND({STT("Y")},0)', '两边不等说明有公式被改坏了'),
    ('1b 数量勾稽：同样的式子用数量再对一遍',
     f'=ROUND({STT("G")}+{STT("I")}+{STT("L")}+{STT("N")}-{STT("P")}-{STT("R")}-{STT("T")}-{STT("V")},3)',
     f'=ROUND({STT("X")},3)', '金额对不代表数量对，要分开查'),
    ('2 费用分摊：可摊费用 = 已摊 + 挂账未摊',
     f'=ROUND(SUMIFS(费用台账!$I${D0}:$I${F_END},费用台账!$J${D0}:$J${F_END},"是",'
     f'费用台账!$K${D0}:$K${F_END},"正常"),0)',
     f'=ROUND(到货批次!$J${S0-1}+到货批次!$K${S0-1},0)', '一分钱都不能漏在外面'),
    ('3 有费用但还有没填单价的行的批次数', '=0',
     f'=COUNTIFS(到货批次!$H${S0}:$H${B_END},">0",到货批次!$F${S0}:$F${B_END},">0")',
     '这些批次运费整批挂账没摊，补完单价会自动摊进去', '待办'),
    ('4 还没录运费的批次数', '=0',
     f'=COUNTIFS(到货批次!$E${S0}:$E${B_END},">0",到货批次!$H${S0}:$H${B_END},"<=0")',
     '27 个柜里只有 2 个录了运费，其余的成本是裸货值，毛利率会虚高', '待办'),
    ('5 账面负库存的品种数', '=0', f'=COUNTIF({ST("AA")},"※账面负库存*")', '领的比账上有的多，或者漏记入库'),
    ('6 还没填单价的入库行数', '=0',
     f'=COUNTIFS(入库单!$Y${D0}:$Y${IN_END},"待补价",入库单!$Z${D0}:$Z${IN_END},"正常")',
     '价格没补，这些货的成本是 0', '待办'),
    ('7 成本不可信的品种占期末库存金额', '=0',
     f'=IF({STT("Y")}=0,0,ROUND(SUMIF({ST("Z")},"不可信",{ST("Y")})/{STT("Y")},4))',
     '超过 10% 这个月的毛利率就别当真', '待办'),
    ('8 数量已经是 0 金额还挂着的品种数', '=0',
     f'=COUNTIF({ST("AA")},"※数量已经是 0*")', '尾差要清掉，不然永远挂在库存里'),
    ('9 外销却没有售价的笔数', '=0',
     f'=COUNTIFS({OUTQ("K")},"{OUTK[1]}",{OUTQ("U")},"正常",{OUTQ("Q")},0)',
     '老板最怕的就是这个 —— 货出去了没开价'),
    ('10 外销是亏的笔数', '=0',
     f'=COUNTIFS({OUTQ("K")},"{OUTK[1]}",{OUTQ("U")},"正常",{OUTQ("T")},"<0")',
     '卖价低于成本，要么定价错要么成本错'),
    ('11 出库单号重复的笔数', '=0',
     f'=SUMPRODUCT(({OUTQ("B")}<>"")*(COUNTIF({OUTQ("B")},{OUTQ("B")}&"")>1)*1)', '一单两记'),
    ('12 出库没填去向类型或矿区的笔数', '=0',
     f'=COUNTIFS({OUTQ("D")},"?*",{OUTQ("U")},"正常",{OUTQ("K")},"")'
     f'+COUNTIFS({OUTQ("D")},"?*",{OUTQ("U")},"正常",{OUTQ("L")},"")'
     f'-COUNTIFS({OUTQ("D")},"?*",{OUTQ("U")},"正常",{OUTQ("K")},"",{OUTQ("L")},"")',
     '不填就分不清是自己耗的还是卖的'),
    ('13 出入库里物料编码不在档案里的笔数', '=0',
     f'=COUNTIF(入库单!$H${D0}:$H${IN_END},"※*")+COUNTIF(出库单!$E${D0}:$E${O_END},"※*")', ''),
    ('14 本月盘点差异金额·先令', '=0',
     f'=ROUND(SUMIF(月度盘点!$H${S0}:$H${P_END},">0",月度盘点!$L${S0}:$L${P_END})'
     f'+SUMIF(月度盘点!$H${S0}:$H${P_END},0,月度盘点!$L${S0}:$L${P_END}),0)',
     '差异不为 0 要在【月度盘点】填原因并调账', '待办'),
    ('15 盘出差异但还没调账的品种数', '=0',
     f'=COUNTIF(月度盘点!$P${S0}:$P${P_END},"※*")', '盘盈去入库单记、盘亏去出库单记'),
    ('16 入库/出库有红字提示的笔数', '=0',
     f'=COUNTIF(入库单!$AC${D0}:$AC${IN_END},"※*")+COUNTIF(出库单!$W${D0}:$W${O_END},"※*")', ''),
]
for j, t in enumerate(['检查项', '应该是', '实际是', '差', '结果', '说明']):
    put(ws, f'{CL(1+j)}5', t, font=FT_HDR, fill=PatternFill('solid', fgColor=C_AUTO), align=CEN)
for i, item in enumerate(CHECKS):
    lab, want, got, note = item[0], item[1], item[2], item[3]
    kind = item[4] if len(item) > 4 else '勾稽'
    r = CHK_R0 + i
    put(ws, f'A{r}', lab, font=Font(name='微软雅黑', size=10), align=LEFT)
    put(ws, f'B{r}', want, fmt='#,##0.###', font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'C{r}', got, fmt='#,##0.###', font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'D{r}', f'=ROUND(N($C{r})-N($B{r}),3)', fmt='#,##0.###',
        font=Font(name='微软雅黑', size=10, bold=True), fill=F_AUTO, align=CEN)
    put(ws, f'E{r}', (f'=IF(ROUND(N($D{r}),3)=0,"OK","※异常")' if kind == '勾稽'
                      else f'=IF(ROUND(N($C{r}),3)=0,"OK","△待补")'),
        font=Font(name='微软雅黑', size=10, bold=True), fill=F_AUTO, align=CEN)
    put(ws, f'F{r}', ('【勾稽】' if kind == '勾稽' else '【待办】') + note,
        font=Font(name='微软雅黑', size=9, color='808080'), align=LEFT)
CHK_END = CHK_R0 + len(CHECKS) - 1
N_JG = sum(1 for it in CHECKS if len(it) < 5 or it[4] == '勾稽')
put(ws, 'B3', f'=COUNTIF($E${CHK_R0}:$E${CHK_END},"OK")', fmt='0',
    font=Font(name='微软雅黑', size=14, bold=True, color='006100'), fill=F_SUM, align=CEN)
put(ws, 'C3', f'={N_JG}', fmt='0', font=Font(name='微软雅黑', size=14, bold=True), fill=F_SUM, align=CEN)
put(ws, 'D3', f'=COUNTIF($E${CHK_R0}:$E${CHK_END},"※*")', fmt='0',
    font=Font(name='微软雅黑', size=14, bold=True, color='C00000'), fill=F_SUM, align=CEN)
put(ws, 'E3', '项没过（只数【勾稽】类；△待补是还没补齐的数据，不是算错）',
    font=Font(name='微软雅黑', size=10), align=LEFT)
put(ws, 'A4', '【勾稽】类必须全 OK，有一项红就说明账算错了，这个月的数不能信。'
              '【待办】类是数据还没补齐（缺单价、缺运费），补一项少一项，不影响已有数字的正确性。',
    font=Font(name='微软雅黑', size=9, color='C00000'), align=LEFT)
ws.merge_cells('A4:F4')
ws.conditional_formatting.add(f'E{CHK_R0}:E{CHK_END}', FormulaRule(
    formula=[f'LEFT(E{CHK_R0},1)="※"'], font=Font(color='C00000', bold=True),
    fill=PatternFill('solid', fgColor='FFD7D7')))
ws.conditional_formatting.add(f'E{CHK_R0}:E{CHK_END}', FormulaRule(
    formula=[f'LEFT(E{CHK_R0},1)="△"'], font=Font(color='BF8F00', bold=True),
    fill=PatternFill('solid', fgColor='FFF2CC')))
ws.conditional_formatting.add(f'E{CHK_R0}:E{CHK_END}', CellIsRule(
    operator='equal', formula=['"OK"'], font=Font(color='006100', bold=True),
    fill=PatternFill('solid', fgColor='E2EFDA')))
ws.sheet_view.showGridLines = False

# ══════════════════════════════════════════════════════════════════════
# 使用说明
# ══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet('使用说明', 0)
ws.sheet_properties.tabColor = '1F3864'
for i, w in enumerate([20, 16, 64, 24], 1):
    ws.column_dimensions[CL(i)].width = w
put(ws, 'A1', '嘉禾矿业进销存系统 · 使用说明', font=FT_TITLE,
    fill=PatternFill('solid', fgColor='1F3864'), align=LEFT)
ws.merge_cells('A1:D1'); ws.row_dimensions[1].height = 30

def sc2(r, t):
    put(ws, f'A{r}', t, font=Font(name='微软雅黑', size=12, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor=C_IN), align=LEFT)
    ws.merge_cells(f'A{r}:D{r}'); ws.row_dimensions[r].height = 24

def ln2(r, a, b='', c=''):
    put(ws, f'A{r}', a, font=Font(name='微软雅黑', size=10, bold=True), align=LEFT)
    put(ws, f'B{r}', b, font=Font(name='微软雅黑', size=10), align=LEFT)
    put(ws, f'C{r}', c, font=Font(name='微软雅黑', size=10),
        align=Alignment('left', 'center', wrap_text=True))
    ws.merge_cells(f'C{r}:D{r}')

sc2(3, '一、谁填哪张表')
ln2(4, '岗位', '表', '填什么')
for i, (a, b, c) in enumerate([
    ('库管（每天）', '出库单', '发出去一笔记一行：日期、物料编码、数量、去向类型、去向矿区、领用人、自己签名。'
                              '一矿二矿选「自用消耗」，三矿六矿选「对外销售」。'),
    ('库管（到货时）', '入库单', '只填 5 样：日期、批次号、物料编码、入库数量、录入人。单价财务后面补，不耽误收货。'),
    ('库管（每月）', '月度盘点', '打印出来（只印到 H 列，看不到账面数），去仓库一样一样数，把实盘数量填回去。'),
    ('财务', '费用台账', '海运费、装柜费、到港费、清关费。国内那段填 CNY，坦桑那段填 TZS。'),
    ('财务', '入库单', '补单价（原币：国内货填人民币，当地货填先令）、补价格状态。'),
    ('财务', '应收对账', '期初欠款、本月收款两列。'),
    ('财务', '参数表', '汇率、加价倍数、账期。月初开账要改这里。'),
    ('老板', '领导月报', '只看这一张。'),
], 5):
    ln2(i, a, b, c)

sc2(14, '二、几条定死的口径')
ln2(15, '本位币', '先令', '所有金额列都是先令。人民币按【参数表】的汇率折，全表别处不许再出现 370。')
ln2(16, '不做单位换算', '数量·单位·单价三者一致',
     '金额 = 数量 × 单价。原表「100米」那种单位，单价就是每 100 米的价，'
     '一换算就差 100 倍 —— 所以干脆不换算，跟原表口径一样。')
ln2(17, '海运费按批次摊', '批次 = 柜号-到港年月',
     '柜号会被船公司重复使用，光写柜号两年后会把运费摊到上一批货上。'
     '费用先折先令，再按本批物料的先令货值分摊（累计差额法，保证摊完一分不差）。')
ln2(18, '缺单价就先不摊', '整批挂账',
     '本批只要还有没填单价的行，运费整批挂账不摊 —— 摊了会把运费全压到有价的那几行上，成本严重失真。'
     '补完单价自动摊进去。')
ln2(19, '售价怎么来', '两套规则',
     '国内货：人民币进价 × 2 × 汇率（不含运费）；当地货：先令进价 × 1.03。'
     '出库单会自动带出建议售价，可以改。')
ln2(20, '库存怎么算', '一句话',
     '期末 = 期初 + 采购入库 + 盘盈 + 矿区退库 − 自用消耗 − 对外销售 − 盘亏 − 退货给供应商。'
     '这一句在库存台账、月报、核对表三处用的是同一个式子。')
ln2(21, '成本怎么算', '月加权平均',
     '月加权单位成本 =（期初金额 + 本月入库金额）÷（期初数量 + 本月入库数量）。'
     '只看入库不看出库，所以不会循环引用。库管不用挑批次。')

sc2(23, '三、月初开账三步（财务做，一个月一次）')
ln2(24, '第 1 步', '结上个月', '看【核对表】，16 项全绿才算结得掉。有红的先处理。')
ln2(25, '第 2 步', '结转期初', '【库存台账】X、Y 两列 → 复制 → 选择性粘贴为**数值** → 贴到 G、H 两列。')
ln2(26, '第 3 步', '改账期', '【参数表】改「本期账期起」和「已结账截止日」。整套表就换月了。')
ln2(27, '⚠ 顺序不能反', '先结转再改账期', '反了的话 X、Y 已经按新月份重算了，贴回去就是错的。')

sc2(29, '四、防丢账、防改账')
ln2(30, '盘点用盲盘', '打印只到 H 列', '库管拿到的单子上没有账面数量，数完填上来才对账。')
ln2(31, '盘盈盘亏走单据', '不许改库存台账',
     '盘盈去【入库单】记一行「盘盈」，盘亏去【出库单】记一行「盘亏」。'
     '库存台账永远是算出来的，不是手改的 —— 这是防篡改的根。')
ln2(32, '错单不删', '改成「作废」', '所有汇总都带「状态=正常」的条件，作废行自动不进数，但永远看得见。')
ln2(33, '结账锁', '参数表', '出库/入库日期早于「已结账截止日」会红字拦住，改不了历史月份。')
ln2(34, '每月封存', '另存只读副本', '结完账另存一份带账期的副本存两个地方。这一条比表里任何公式都实用。')

sc2(36, '五、这次从原表搬过来的 + 你要知道的三件事')
for i, t in enumerate([
    '· 原表 877 行全部处理完：867 行物料 → 【入库单】，10 行费用 → 【费用台账】（两张表物理分开，混在一起分摊必错）。',
    '· 按（品名+规格）去重发了 795 个物料编码，跨表只用编码关联。',
    '· 集装箱号和供货商列有 394 行写「同上」，已经向下填充还原。批次号 = 柜号-到港年月。',
    '· 单位里 333 行是把数量重复写进了单位（「10个」+数量10），已删掉数字；另有 9 行的数字是包装规格（「100米」），原样保留。',
    '',
    '⚠ 一、炸药和柴油，这张表里一行都没有。',
    '　 全表搜过：炸药/雷管/导爆/乳化/硝铵 一条没有，油料只有 2 台柴油发电机。',
    '　 你说这两样是最主要的生产成本 —— 它们是在坦桑当地买的，从来没进过这份货柜清单。',
    '　 要盯住，就得从现在起把当地采购也录进【入库单】（货源选 B当地，批次号写「本地-年月」）。',
    '',
    '⚠ 二、27 个柜里只有 2 个柜录了运费。',
    '　 TIIU4204331 柜 25,835 元、DRYU9381826 柜 29,534.02 元，其余 25 个柜一分运费都没有。',
    '　 这两个柜的运费率是 22.28% 和 5.75%，差快 4 倍 —— 所以必须一个柜一个柜地摊，不能用统一费率。',
    '　 没录运费的柜，成本是裸货值，毛利率会虚高。【核对表】第 4 项会一直提醒你补。',
    '',
    '⚠ 三、867 行里有 534 行没填单价，占 61%。',
    '　 货值合计 3,192,166.05 元只是有价那 333 行算出来的。这 534 行现在价格状态是「待补价」，',
    '　 它们所在批次的运费整批挂账没摊。补完单价，运费会自动摊进去，库存金额也会自动补上。',
], 37):
    put(ws, f'A{i}', t, font=Font(name='微软雅黑', size=9,
        color='C00000' if t.startswith('⚠') or t.startswith('　') else '595959'), align=LEFT)
    ws.merge_cells(f'A{i}:D{i}')

sc2(59, '六、体检（活的）')
HEALTH = [
    ('核对表通过项', '=核对表!$B$3&" / "&核对表!$C$3', ''),
    ('还没填单价的入库行数', f'=COUNTIFS(入库单!$Y${D0}:$Y${IN_END},"待补价",入库单!$Z${D0}:$Z${IN_END},"正常")', ''),
    ('还没录运费的批次数', f'=SUMPRODUCT((到货批次!$A${S0}:$A${B_END}<>"")*(到货批次!$E${S0}:$E${B_END}>0)*(到货批次!$H${S0}:$H${B_END}=0))', ''),
    ('账面负库存的品种数', f'=COUNTIF({ST("X")},"<-0.001")', ''),
    ('入库/出库红字提示笔数', f'=COUNTIF(入库单!$AC${D0}:$AC${IN_END},"※*")+COUNTIF(出库单!$W${D0}:$W${O_END},"※*")', ''),
    ('期末库存金额·先令', f'={STT("Y")}', '#,##0'),
    ('本月自用消耗·先令', f'={STT("Q")}', '#,##0'),
    ('本月对外销售成本·先令', f'={STT("S")}', '#,##0'),
]
for i, (lab, f, fmt) in enumerate(HEALTH, 60):
    put(ws, f'A{i}', lab, font=Font(name='微软雅黑', size=10), align=LEFT)
    ws.merge_cells(f'A{i}:B{i}')
    put(ws, f'C{i}', f, fmt=fmt or None,
        font=Font(name='微软雅黑', size=11, bold=True, color='C00000'), align=CEN)
ws.sheet_view.showGridLines = False

NAMES['物料档案表'] = f'物料档案!$A${D0}:$A${I_END}'
for nm, ref in NAMES.items():
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)
for s in wb.worksheets:
    s.page_setup.orientation = 'landscape'
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 0
    s.sheet_properties.pageSetUpPr.fitToPage = True
wb.move_sheet('领导月报', offset=-1)
wb.active = 0
wb.save(OUT)
print('已生成:', OUT)
print('工作表:', ' / '.join(s.title for s in wb.worksheets))
n = sum(1 for s in wb.worksheets for row in s.iter_rows()
        for c in row if isinstance(c.value, str) and c.value.startswith('='))
print('公式格子 %d 个' % n)
