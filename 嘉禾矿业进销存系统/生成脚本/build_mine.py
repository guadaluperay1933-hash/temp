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
import json, os, sys, datetime, collections
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

N_ITEM, N_IN, N_OUT, N_FEE, N_BAT = 1800, 4000, 5000, 800, 200
HDR, D0 = 3, 4
S0 = 5                                   # 自动表：4 行合计，5 行起数据

C_IN, C_AUTO, C_BASE, C_BOSS = '2F5597', '375623', '806000', '833C00'
DATEQ = 'yyyy-mm-dd;;;@'          # 自动区日期：0 不显示，免得画成 1900-01-00
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

MINES = DATA.get('mines', ['一矿区', '二矿区', '三矿区', '六矿区', '选厂', '维修部', '沙金'])
BELONG = {'一矿区': '自己', '二矿区': '自己', '选厂': '自己', '维修部': '自己', '沙金': '自己',
          '三矿区': '外部', '六矿区': '外部'}
OWN = [m for m in MINES if BELONG.get(m, '自己') == '自己']
SELL = [m for m in MINES if BELONG.get(m, '自己') == '外部']
OUTK = ['自用消耗', '对外销售', '矿区借调', '盘亏', '退货给供应商']
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
put(ws, 'F3', '归属', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BASE), align=CEN)
put(ws, 'D5', '=EOMONTH($B$4,0)', fmt=DATEF,
    font=Font(name='微软雅黑', size=10, color='7F7F7F'), fill=F_AUTO, align=CEN)
put(ws, 'D3', '本月末参考', font=FT_HDR, fill=PatternFill('solid', fgColor=C_BASE), align=CEN)
ws.column_dimensions['D'].width = 14
LISTS = [('E', '矿区', MINES, 12), ('F', '归属', [BELONG.get(m, '自己') for m in MINES], 10),
         ('G', '出库去向', OUTK, 16), ('I', '入库类型', INK, 14),
         ('K', '费用类型', FEEK, 18), ('M', '货源', SRC_, 12), ('O', '价格状态', PSTAT, 14),
         ('Q', '物料大类', CATS, 14), ('S', '计量单位', UNITS, 12), ('U', '人员', [], 14),
         ('W', '重复嫌疑物料编码', DATA.get('dup_codes', []), 200)]
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
    '矿区表': f'参数表!$E${D0}:$E${D0+29}', '矿区归属': f'参数表!$E${D0}:$F${D0+29}',
    '去向表': f'参数表!$G${D0}:$G${D0+29}',
    '入库类型表': f'参数表!$I${D0}:$I${D0+29}', '费用类型表': f'参数表!$K${D0}:$K${D0+29}',
    '货源表': f'参数表!$M${D0}:$M${D0+29}', '价格状态表': f'参数表!$O${D0}:$O${D0+29}',
    '大类表': f'参数表!$Q${D0}:$Q${D0+29}', '单位表': f'参数表!$S${D0}:$S${D0+59}',
    '人员表': f'参数表!$U${D0}:$U${D0+29}',
    '重复嫌疑表': f'参数表!$W${D0}:$W${D0+199}',
}
ws.freeze_panes = 'A4'


def helper_col(ws, col, r0, r1, title, memo):
    """在正表右边挂一根辅助列：灰头灰底，一眼看出是公式用的、不能删。"""
    idx = ws[f'{col}1'].column
    c = ws[f'{col}{HDR}']
    c.value, c.font, c.alignment, c.border = title, FT_HDR, CEN, BOX
    c.fill = PatternFill('solid', fgColor='808080')
    ws.column_dimensions[col].width = 13
    m = ws[f'{col}{HDR - 1}']
    m.value = memo
    m.font, m.alignment = Font(name='微软雅黑', size=9, color='C00000'), LEFT
    for r in range(r0, r1 + 1):
        cc = ws.cell(r, idx)
        cc.border, cc.font, cc.fill, cc.alignment = BOX, FT_AUTO, F_AUTO, CEN

# ══════════════════════════════════════════════════════════════════════
# 物料档案
# ══════════════════════════════════════════════════════════════════════
I_END = D0 + N_ITEM - 1
F_END0 = D0 + N_FEE - 1
IN_END = D0 + N_IN - 1
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
    put(ws, f'N{r}', f'={g}IFERROR(LOOKUP(2,1/(({OUT_R("D")}=$A{r})*({OUT_R("W")}="正常")),'
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
F_END = F_END0
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
    # 只给「入库单里还没有」的批次编号：费用先到、货还没到的柜子，
    # 【到货批次】也要能看见它，不然那笔运费就挂在没人管的地方。
    put(ws, f'M{r}', f'={g}IF(COUNTIF(入库单!$E${D0}:$E${IN_END},$C{r})>0,"",'
                     f'IF(COUNTIF($C${D0}:$C{r},$C{r})>1,"",COUNT($M${HDR}:M{r - 1})+1)))', fmt='0')
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
nf = len(DATA['fees'])
for i, x in enumerate(DATA.get('new_fees', [])):
    r = D0 + nf + i
    for c, v in ((2, datetime.datetime.strptime(x['date'], '%Y-%m-%d')), (3, x['batch']),
                 (4, x['cat']), (5, '福建全力供应链管理有限公司'), (6, 'CNY'), (7, x['cny']),
                 (10, '是'), (11, '正常'), (12, x['note'])):
        ws.cell(r, c).value = v
nf += len(DATA.get('new_fees', []))
LOCAL_CAT = {'TBS费用': '坦桑清关费', '船公司费用': '坦桑到港费', '港口费': '坦桑到港费',
             'GALCO海关操作费': '坦桑清关费', '滞箱延误费': '坦桑到港费', '堆场费': '坦桑到港费',
             '集装箱运费': '坦桑内陆运输费', '海关税': '坦桑清关费', '免税税费': '坦桑清关费',
             '放行费及代理费': '坦桑清关费'}
for i, x in enumerate(DATA.get('local_fees', [])):
    r = D0 + nf + i
    for c, v in ((2, datetime.datetime.strptime(x['date'], '%Y-%m-%d') if x['date'] else None),
                 (3, x['batch']), (4, LOCAL_CAT.get(x['cat'], '其他费用')),
                 (5, 'EDAN/GALCO/港口等', ), (6, 'TZS'), (7, x['tzs']),
                 (10, '是'), (11, '正常'), (12, f"{x['cat']}　{x['note']}")):
        ws.cell(r, c).value = v
helper_col(ws, 'M', D0, F_END, '批次首现(自动)',
           '← 这一列是【到货批次】自动长清单用的，别删、别手填')
dv(ws, '=费用类型表', f'D{D0}:D{F_END}')
dv(ws, '"CNY,TZS"', f'F{D0}:F{F_END}', warn=False)
dv(ws, '"是,否"', f'J{D0}:J{F_END}', warn=False)
dv(ws, '"正常,作废"', f'K{D0}:K{F_END}', warn=False)
ws.freeze_panes = f'C{D0}'
ws.auto_filter.ref = f'A{HDR}:L{F_END}'

# ══════════════════════════════════════════════════════════════════════
# 入库单
# ══════════════════════════════════════════════════════════════════════
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
    # 批次首现序号：一个批次号**第一次**出现在哪一行，就在那一行编个号 1,2,3…
    # 【到货批次】靠它自动长出批次清单，不用再手工把柜号誊一遍。
    put(ws, f'AD{r}', f'=IF($E{r}="","",IF(COUNTIF($E${D0}:$E{r},$E{r})>1,"",'
                      f'COUNT($AD${HDR}:AD{r - 1})+1))', fmt='0')
    put(ws, f'AC{r}', f'={g}'
        f'IF($H{r}="※编码不在物料档案里","※物料编码不在档案里",'
        f'IF(N($K{r})<=0,"※入库数量要大于 0",'
        f'IF(AND($C{r}<>"",$C{r}<=结账锁),"※日期早于已结账截止日，不能再录",'
        f'IF(AND($L{r}=0,$X{r}="正常"),"※没填单价，价格状态请选「待补价」",'
        f'IF(AND($S{r}>0,$R{r}>0),"△本批有 "&$S{r}&" 行没单价，整批运费先挂账没摊",'
        f'IF(AND($B{r}<>"",COUNTIF($B${D0}:$B${IN_END},$B{r})>1),"※入库单号重复",'
        f'IF($C{r}="","△这一笔没填日期，不会进任何月份的报表",'
        f'IF(COUNTIF(重复嫌疑表,$G{r})>0,'
        f'"△这个物料货柜和本地采购都登过，两边数量对不上，等你定以哪份为准","")))))))))')
for i, x in enumerate(DATA['inbound']):
    r = D0 + i
    for c, v in ((3, datetime.datetime.strptime(x['date'], '%Y-%m-%d') if x['date'] else None),
                 (4, x['source']), (5, x['batch']), (6, x['supplier'] or None), (7, x['code']),
                 (11, x['qty']), (12, x['price'] or None), (24, '采购入库'),
                 (25, '正常' if x['price'] else '待补价'), (26, '正常'),
                 (28, (f"原表第 {x['row']} 行" + (('　' + x['note']) if x['note'] else '')))):
        ws.cell(r, c).value = v
ni = len(DATA['inbound'])
for i, x in enumerate(DATA.get('new_inbound', [])):
    r = D0 + ni + i
    for c, v in ((3, datetime.datetime.strptime(x['date'], '%Y-%m-%d')), (4, 'A国内'),
                 (5, x['batch']), (6, x['supplier']), (7, x['code']), (11, x['qty']),
                 (12, x['price'] or None), (24, '采购入库'),
                 (25, '正常' if x['price'] else '待补价'), (26, '正常'), (28, x['note'])):
        ws.cell(r, c).value = v
ni += len(DATA.get('new_inbound', []))
for i, x in enumerate(DATA.get('buy', [])):
    r = D0 + ni + i
    for c, v in ((3, datetime.datetime.strptime(x['date'], '%Y-%m-%d') if x['date'] else None),
                 (4, 'B当地'), (5, x['batch']), (7, x['code']), (11, x['qty']),
                 (24, '采购入库'), (25, '待补价'), (26, '正常'), (28, x['note'])):
        ws.cell(r, c).value = v
dv(ws, '=货源表', f'D{D0}:D{IN_END}', warn=False)
dv(ws, '=物料档案表', f'G{D0}:G{IN_END}')
dv(ws, '=入库类型表', f'X{D0}:X{IN_END}'.replace('X', 'X'), warn=False)
dv(ws, '=价格状态表', f'Y{D0}:Y{IN_END}', warn=False)
dv(ws, '"正常,作废"', f'Z{D0}:Z{IN_END}', warn=False)
dv(ws, '=人员表', f'AA{D0}:AA{IN_END}')
redflag(ws, f'AC{D0}:AC{IN_END}'); redflag(ws, f'H{D0}:H{IN_END}')
helper_col(ws, 'AD', D0, IN_END, '批次首现(自动)',
           '← 这一列是【到货批次】自动长清单用的，别删、别手填')
ws.freeze_panes = f'H{D0}'
ws.auto_filter.ref = f'A{HDR}:AC{IN_END}'

# ══════════════════════════════════════════════════════════════════════
# 出库单
# ══════════════════════════════════════════════════════════════════════
O_END = D0 + N_OUT - 1
COLS_O = [('序号', 7, 'auto', None), ('出库单号', 14, 'in', None), ('出库日期 ★', 12, 'in', DATEF),
          ('物料编码 ★', 11, 'in', None), ('品名', 24, 'auto', None), ('规格型号', 22, 'auto', None),
          ('单位', 9, 'auto', None), ('大类', 13, 'auto', None), ('重点物资', 10, 'auto', None),
          ('出库数量 ★', 12, 'in', QTY), ('去向类型 ★', 14, 'in', None),
          ('去向矿区 ★', 12, 'in', None), ('调出矿区\n(只有借调填)', 13, 'in', None),
          ('矿区归属', 10, 'auto', None),
          ('领用人/设备号', 15, 'in', None), ('库管 ★', 10, 'in', None),
          ('建议售价·先令', 14, 'auto', TZS), ('实际售价·先令', 14, 'in', TZS),
          ('销售金额·先令', 15, 'auto', TZS), ('单位成本·先令', 14, 'auto', TZS),
          ('出库成本·先令', 15, 'auto', TZS), ('毛利·先令', 15, 'auto', TZS),
          ('状态', 9, 'in', None), ('备注', 24, 'in', None), ('核对', 30, 'auto', None)]
ws = newsheet('出库单', '出库单 · 库管每天就开这一张',
              '「去向类型」是总开关：自用消耗 = 自己矿上耗掉（只出成本）；对外销售 = 卖给外面的矿区（出成本也出收入）；'
              '矿区借调 = 矿区之间调货（东西还在公司里，不减总库存，但要留痕）。'
              '哪个矿区算「自己」、哪个算「外部」，在【参数表】E/F 两列改。盘盈不在这里录，去【入库单】选「盘盈」。',
              C_IN, len(COLS_O))
header(ws, COLS_O, C_IN); body(ws, COLS_O, D0, O_END)
for r in range(D0, O_END + 1):
    g = f'IF($D{r}="","",'
    put(ws, f'A{r}', f'={g}COUNTA($D${D0}:$D{r}))')
    for col, idx in (('E', 3), ('F', 4), ('G', 5), ('H', 2), ('I', 6)):
        put(ws, f'{col}{r}', f'={g}IFERROR(VLOOKUP($D{r},物料档案!$A${D0}:$P${I_END},{idx},0),'
                             f'"※编码不在物料档案里"))')
    put(ws, f'N{r}', f'={g}IFERROR(VLOOKUP($L{r},矿区归属,2,0),""))')
    put(ws, f'Q{r}', f'={g}IFERROR(VLOOKUP($D{r},物料档案!$A${D0}:$P${I_END},10,0),0))')
    put(ws, f'S{r}', f'={g}IF($K{r}<>"{OUTK[1]}",0,ROUND(N($J{r})*IF($R{r}="",N($Q{r}),N($R{r})),0)))')
    put(ws, f'T{r}', f'={g}IFERROR(VLOOKUP($D{r},库存台账!$A${S0}:$Y${S0+N_ITEM-1},11,0),0))')
    put(ws, f'U{r}', f'={g}IF($K{r}="{OUTK[2]}",0,ROUND(N($J{r})*N($T{r}),0)))')
    put(ws, f'V{r}', f'={g}IF($K{r}<>"{OUTK[1]}","",ROUND($S{r}-$U{r},0)))')
    put(ws, f'Y{r}', f'={g}'
        f'IF($E{r}="※编码不在物料档案里","※物料编码不在档案里",'
        f'IF(N($J{r})<=0,"※出库数量要大于 0",'
        f'IF($K{r}="","※没填去向类型",'
        f'IF($L{r}="","※没填去向矿区",'
        f'IF(AND($C{r}<>"",$C{r}<=结账锁),"※日期早于已结账截止日，不能再录",'
        f'IF(AND($K{r}="{OUTK[0]}",$N{r}="外部"),"※这个矿区在【参数表】里归「外部」，应该选对外销售",'
        f'IF(AND($K{r}="{OUTK[1]}",$N{r}="自己"),"※这个矿区在【参数表】里归「自己」，应该选自用消耗",'
        f'IF(AND($K{r}="{OUTK[2]}",$M{r}=""),"※借调要填调出矿区",'
        f'IF(AND($K{r}="{OUTK[2]}",$M{r}=$L{r}),"※调出和调入是同一个矿区",'
        f'IF(AND($K{r}="{OUTK[1]}",ROUND($S{r},0)=0),"※外销却没有售价，请填实际售价",'
        f'IF(AND($K{r}="{OUTK[1]}",$V{r}<>"",$V{r}<0),"※这一笔外销是亏的，请核对售价",'
        f'IF(IFERROR(VLOOKUP($D{r},库存台账!$A${S0}:$Y${S0+N_ITEM-1},24,0),0)<-0.001,'
        f'"※这个物料账面已经变成负数了，查一下是不是漏记入库",'
        f'IF($C{r}="","△这一笔没填日期，不会进任何月份的报表",""))))))))))))))')
for i, x in enumerate(DATA.get('outbound', [])):
    r = D0 + i
    for c, v in ((3, datetime.datetime.strptime(x['date'], '%Y-%m-%d') if x['date'] else None),
                 (4, x['code']), (10, x['qout']), (11, OUTK[1]), (12, SELL[0]),
                 (23, '正常'), (24, f"原表第 {x['row']} 行：{x['supplier'] or x['name']}")):
        ws.cell(r, c).value = v
n0 = len(DATA.get('outbound', []))
for i, x in enumerate(DATA.get('use', [])):
    r = D0 + n0 + i
    mine = x['mine'] or ''
    kind = OUTK[1] if BELONG.get(mine) == '外部' else OUTK[0]
    for c, v in ((3, datetime.datetime.strptime(x['date'], '%Y-%m-%d') if x['date'] else None),
                 (4, x['code']), (10, x['qty']), (11, kind), (12, mine or None),
                 (23, '正常'), (24, x['note'])):
        ws.cell(r, c).value = v
n1 = n0 + len(DATA.get('use', []))
for i, x in enumerate(DATA.get('lend', [])):
    r = D0 + n1 + i
    for c, v in ((3, datetime.datetime.strptime(x['date'], '%Y-%m-%d') if x['date'] else None),
                 (4, x['code']), (10, x['qty']), (11, OUTK[2]), (12, x['to_mine'] or None),
                 (13, x['from_mine'] or None), (23, '正常'), (24, x['note'])):
        ws.cell(r, c).value = v
dv(ws, '=物料档案表', f'D{D0}:D{O_END}')
dv(ws, '=去向表', f'K{D0}:K{O_END}', warn=False)
dv(ws, '=矿区表', f'L{D0}:L{O_END}', warn=False)
dv(ws, '=矿区表', f'M{D0}:M{O_END}', warn=False)
dv(ws, '=人员表', f'P{D0}:P{O_END}')
dv(ws, '"正常,作废"', f'W{D0}:W{O_END}', warn=False)
redflag(ws, f'Y{D0}:Y{O_END}'); redflag(ws, f'E{D0}:E{O_END}')
ws.conditional_formatting.add(f'I{D0}:I{O_END}', CellIsRule(
    operator='equal', formula=['"是"'], font=Font(color='C00000', bold=True)))
ws.conditional_formatting.add(f'V{D0}:V{O_END}', CellIsRule(
    operator='lessThan', formula=['0'], font=Font(color='C00000', bold=True)))
ws.freeze_panes = f'E{D0}'
ws.auto_filter.ref = f'A{HDR}:Y{O_END}'

# ══════════════════════════════════════════════════════════════════════
# 到货批次（自动，看每个柜摊得对不对）
# ══════════════════════════════════════════════════════════════════════
B_END = S0 + N_BAT - 1
COLS_B = [('批次号', 26, 'auto', None), ('集装箱号', 18, 'auto', None), ('到港日期', 12, 'auto', DATEQ),
          ('货源', 10, 'auto', None), ('物料行数', 10, 'auto', '0'), ('没填单价行数', 12, 'auto', '0'),
          ('货值·先令', 16, 'auto', TZS), ('费用·先令', 15, 'auto', TZS),
          ('运费率', 10, 'auto', PCT), ('已摊运费·先令', 16, 'auto', TZS),
          ('未摊(挂账)·先令', 16, 'auto', TZS), ('入库总成本·先令', 16, 'auto', TZS),
          ('状态', 30, 'auto', None)]
ws = newsheet('到货批次', '到货批次 · 整张表都是公式，一个批次一行，不用手工登',
              '批次清单从【入库单】自动长出来：入库单里出现一个新批次号，这里就自动多一行；'
              '费用先到、货还没到的批次，从【费用台账】也能捞出来。柜号从批次号里截，'
              '到港日期取这个批次第一笔入库的日期。「运费率」差得越多越说明必须按柜摊：'
              'TIIU4204331 柜 22.28%、DRYU9381826 柜 5.75%，差快 4 倍，用统一费率去摊必错。',
              C_AUTO, len(COLS_B))
header(ws, COLS_B, C_AUTO); body(ws, COLS_B, S0 - 1, B_END)
# 批次清单的两个来源：入库单（主）和费用台账（费用先到、货还没到的）
NB_IN, NB_AD = f'入库单!$E${D0}:$E${IN_END}', f'入库单!$AD${D0}:$AD${IN_END}'
NB_DT, NB_SRC = f'入库单!$C${D0}:$C${IN_END}', f'入库单!$D${D0}:$D${IN_END}'
NF_BT, NF_M = f'费用台账!$C${D0}:$C${F_END}', f'费用台账!$M${D0}:$M${F_END}'
NF_DT = f'费用台账!$B${D0}:$B${F_END}'
put(ws, 'N2', '辅助·入库单里的批次数', font=Font(name='微软雅黑', size=9, color='C00000'), align=LEFT)
put(ws, 'O2', f'=MAX({NB_AD})', fmt='0', font=FT_SUM, fill=F_SUM, align=CEN)
ws.column_dimensions['N'].width = 20
ws.column_dimensions['O'].width = 8
for i in range(N_BAT):
    r = S0 + i
    g = f'IF($A{r}="","",'
    # 批次号：先把入库单里第 i 个批次捞出来；捞不到再去费用台账接着往下捞
    put(ws, f'A{r}', f'=IFERROR(INDEX({NB_IN},MATCH({i + 1},{NB_AD},0)),'
                     f'IFERROR(INDEX({NF_BT},MATCH({i + 1}-$O$2,{NF_M},0)),""))')
    put(ws, f'B{r}', f'={g}IF($D{r}="{SRC_[1]}","",IFERROR(LEFT($A{r},FIND("-",$A{r})-1),"")))')
    # 取到的那一行日期本身是空的时候 INDEX 会返回 0，不挡住就画成 1900-01-00
    put(ws, f'C{r}', f'={g}IF(N(IFERROR(INDEX({NB_DT},MATCH($A{r},{NB_IN},0)),'
                     f'IFERROR(INDEX({NF_DT},MATCH($A{r},{NF_BT},0)),0)))=0,"",'
                     f'IFERROR(INDEX({NB_DT},MATCH($A{r},{NB_IN},0)),'
                     f'IFERROR(INDEX({NF_DT},MATCH($A{r},{NF_BT},0)),""))))', fmt=DATEQ)
    put(ws, f'D{r}', f'={g}IFERROR(LOOKUP(2,1/({NB_IN}=$A{r}),{NB_SRC}),'
                     f'IF(LEFT($A{r},2)="本地","{SRC_[1]}","{SRC_[0]}")))')
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
    put(ws, f'M{r}', f'={g}IF($E{r}=0,"△这个批次只有费用、还没有入库单（货还没到？）",'
                     f'IF(AND($H{r}>0,$F{r}>0),"△有 "&$F{r}&" 行没单价，运费整批挂账没摊 —— 补完价自动摊",'
                     f'IF($H{r}=0,"△这个批次还没录运费",'
                     f'IF(ABS($K{r})>1,"※摊完还差 "&TEXT($K{r},"#,##0")&" 先令，查一下","已摊平")))))')
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
            f'出库单!$K${D0}:$K${O_END},"{kind}",出库单!$W${D0}:$W${O_END},"正常",{PERO})')
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
    put(ws, f'Q{r}', f'={g}ROUND(' + sout_('U', '自用消耗').format(r=r) + ',0))')
    put(ws, f'R{r}', f'={g}' + sout_('J', '对外销售').format(r=r) + ')')
    put(ws, f'S{r}', f'={g}ROUND(' + sout_('U', '对外销售').format(r=r) + ',0))')
    put(ws, f'T{r}', f'={g}' + sout_('J', '盘亏').format(r=r) + ')')
    put(ws, f'U{r}', f'={g}ROUND(' + sout_('U', '盘亏').format(r=r) + ',0))')
    put(ws, f'V{r}', f'={g}' + sout_('J', '退货给供应商').format(r=r) + ')')
    put(ws, f'W{r}', f'={g}ROUND(' + sout_('U', '退货给供应商').format(r=r) + ',0))')
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
COLS_AR = [('矿区/客户', 14, 'auto', None), ('归属', 10, 'auto', None),
           ('期初欠款·先令', 16, 'in', TZS), ('本月销售额·先令', 16, 'auto', TZS),
           ('本月收款·先令', 16, 'in', TZS), ('期末欠款·先令', 16, 'auto', TZS),
           ('本月销售成本·先令', 16, 'auto', TZS), ('本月毛利·先令', 15, 'auto', TZS),
           ('毛利率', 10, 'auto', PCT), ('本月笔数', 10, 'auto', '0'), ('备注', 26, 'in', None)]
ws = newsheet('应收对账', '应收对账 · 卖给外部矿区的货款收没收回来',
              '哪个矿区算「外部」看【参数表】E/F 两列。归属是「自己」的行不会有销售额。'
              '期初欠款和本月收款要人填，其余自动。月初开账把「期末欠款」抄到下个月的「期初欠款」。',
              C_AUTO, len(COLS_AR))
header(ws, COLS_AR, C_AUTO); body(ws, COLS_AR, S0 - 1, AR_END)
for i in range(20):
    r = S0 + i
    g = f'IF($A{r}="","",'
    put(ws, f'A{r}', f'=IF(参数表!$E{D0+i}="","",参数表!$E{D0+i})')
    put(ws, f'B{r}', f'={g}IFERROR(VLOOKUP($A{r},矿区归属,2,0),""))')
    put(ws, f'D{r}', f'={g}ROUND(SUMIFS(出库单!$S${D0}:$S${O_END},出库单!$L${D0}:$L${O_END},$A{r},'
                     f'出库单!$K${D0}:$K${O_END},"{OUTK[1]}",出库单!$W${D0}:$W${O_END},"正常",{PERO}),0))')
    put(ws, f'F{r}', f'={g}ROUND(N($C{r})+$D{r}-N($E{r}),0))')
    put(ws, f'G{r}', f'={g}ROUND(SUMIFS(出库单!$U${D0}:$U${O_END},出库单!$L${D0}:$L${O_END},$A{r},'
                     f'出库单!$K${D0}:$K${O_END},"{OUTK[1]}",出库单!$W${D0}:$W${O_END},"正常",{PERO}),0))')
    put(ws, f'H{r}', f'={g}ROUND($D{r}-$G{r},0))')
    put(ws, f'I{r}', f'={g}IF($D{r}=0,"",ROUND($H{r}/$D{r},4)))')
    put(ws, f'J{r}', f'={g}COUNTIFS(出库单!$L${D0}:$L${O_END},$A{r},'
                     f'出库单!$K${D0}:$K${O_END},"{OUTK[1]}",出库单!$W${D0}:$W${O_END},"正常",{PERO}))')
put(ws, f'A{S0-1}', '合计', font=FT_SUM, fill=F_SUM, align=CEN)
for c in 'CDEFGH':
    put(ws, f'{c}{S0-1}', f'=ROUND(SUM({c}{S0}:{c}{AR_END}),0)', fmt=TZS, font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'I{S0-1}', f'=IF($D{S0-1}=0,"",ROUND($H{S0-1}/$D{S0-1},4))', fmt=PCT,
    font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'J{S0-1}', f'=SUM(J{S0}:J{AR_END})', fmt='0', font=FT_SUM, fill=F_SUM, align=CEN)
ws.conditional_formatting.add(f'F{S0}:F{AR_END}', CellIsRule(
    operator='greaterThan', formula=['0'], font=Font(color='C00000', bold=True)))
ws.freeze_panes = f'C{S0}'

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
HD2 = ['大类', '本月耗用金额', '自己矿区耗掉', '卖给外部矿区', '期末库存金额', '期末品种数', '占期末库存']
for j, t in enumerate(HD2):
    put(ws, f'{CL(1+j)}16', t, font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
KEYCATS = ['炸药火工', '油料', '选矿药剂', '磨矿介质']
OWN_Q = '+'.join(f'SUMIFS({OUTQ("U")},{OUTQ("H")},$A{{r}},{OUTQ("L")},"{m}",'
                 f'{OUTQ("K")},"{OUTK[0]}",{OUTQ("W")},"正常",{PERO2})' for m in OWN)
for i, cat in enumerate(KEYCATS):
    r = 17 + i
    put(ws, f'A{r}', cat, font=Font(name='微软雅黑', size=10, bold=True), align=LEFT)
    put(ws, f'B{r}', f'=ROUND($C{r}+$D{r},0)', fmt=TZS, font=FT_SUM, fill=F_SUM, align=CEN)
    put(ws, f'C{r}', '=ROUND(' + OWN_Q.format(r=r) + ',0)', fmt=TZS, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'D{r}', f'=ROUND(SUMIFS({OUTQ("U")},{OUTQ("H")},$A{r},{OUTQ("K")},"{OUTK[1]}",'
                     f'{OUTQ("W")},"正常",{PERO2}),0)', fmt=TZS, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'E{r}', f'=ROUND(SUMIF({ST("E")},$A{r},{ST("Y")}),0)', fmt=TZS, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'F{r}', f'=COUNTIFS({ST("E")},$A{r},{ST("X")},">0")', fmt='0', font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'G{r}', f'=IF({STT("Y")}=0,"",ROUND($E{r}/{STT("Y")},4))', fmt=PCT, font=FT_AUTO, fill=F_AUTO, align=CEN)
put(ws, 'A21', '⚠ 这四行的数现在还很小，因为炸药柴油这些是在坦桑当地买的：'
               '货柜清单里一条炸药都没有，油料只有 2 台柴油发电机。物资流水帐补进来之后才有了炸药 / 导爆管 / 电管 / 电雷管 / 柴油，'
               '但它们**没有单价**（流水帐只记数量），所以金额还是 0。把单价补进【入库单】，这四行才有意义。',
    font=Font(name='微软雅黑', size=9, color='C00000'), align=LEFT)
ws.merge_cells('A21:H22')
ws.row_dimensions[21].height = 30

sec(24, '三、各矿区耗了多少（自己的算成本，外部的算销售）')
HD3 = ['矿区', '归属', '本月出库成本·先令', '占全部', '折人民币', '处理矿石量(吨)', '吨耗·先令', '借调进/出笔数']
for j, t in enumerate(HD3):
    put(ws, f'{CL(1+j)}25', t, font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
M_R0 = 26
for i, mine in enumerate(MINES):
    r = M_R0 + i
    own = BELONG.get(mine, '自己') == '自己'
    kind = OUTK[0] if own else OUTK[1]
    put(ws, f'A{r}', mine, font=Font(name='微软雅黑', size=10, bold=True), align=LEFT)
    put(ws, f'B{r}', f'=IFERROR(VLOOKUP($A{r},矿区归属,2,0),"")', font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'C{r}', f'=ROUND(SUMIFS({OUTQ("U")},{OUTQ("L")},$A{r},{OUTQ("K")},"{kind}",'
                     f'{OUTQ("W")},"正常",{PERO2}),0)', fmt=TZS, font=FT_SUM, fill=F_AUTO, align=CEN)
    put(ws, f'D{r}', f'=IF($C${M_R0+len(MINES)}=0,"",ROUND($C{r}/$C${M_R0+len(MINES)},4))',
        fmt=PCT, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'E{r}', f'=ROUND($C{r}/汇率,2)', fmt=CNY, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'F{r}', ('=一矿处理量' if mine == '一矿区' else '=二矿处理量' if mine == '二矿区' else None),
        fmt='#,##0.##', font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'G{r}', f'=IF(N($F{r})=0,"—",ROUND($C{r}/$F{r},0))', fmt=TZS, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'H{r}', f'=COUNTIFS({OUTQ("L")},$A{r},{OUTQ("K")},"{OUTK[2]}",{OUTQ("W")},"正常",{PERO2})'
                     f'&" / "&COUNTIFS({OUTQ("M")},$A{r},{OUTQ("K")},"{OUTK[2]}",{OUTQ("W")},"正常",{PERO2})',
        font=FT_AUTO, fill=F_AUTO, align=CEN)
MT = M_R0 + len(MINES)
put(ws, f'A{MT}', '合计', font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'C{MT}', f'=ROUND(SUM($C${M_R0}:$C${MT-1}),0)', fmt=TZS, font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'E{MT}', f'=ROUND($C{MT}/汇率,2)', fmt=CNY, font=FT_SUM, fill=F_SUM, align=CEN)

S4 = MT + 2
sec(S4, '四、卖给外部矿区，赚了多少')
HD4 = ['矿区', '本月销售额·先令', '本月销售成本', '本月毛利', '毛利率', '期末还欠我们', '本月笔数']
for j, t in enumerate(HD4):
    put(ws, f'{CL(1+j)}{S4+1}', t, font=FT_HDR, fill=PatternFill('solid', fgColor=C_BOSS), align=CEN)
for i, mine in enumerate(SELL):
    r = S4 + 2 + i
    ar = S0 + MINES.index(mine)
    put(ws, f'A{r}', mine, font=Font(name='微软雅黑', size=10, bold=True), align=LEFT)
    for j, sc in enumerate(['D', 'G', 'H', 'I', 'F', 'J']):
        put(ws, f'{CL(2+j)}{r}', f'=应收对账!${sc}{ar}',
            fmt=(PCT if sc == 'I' else '0' if sc == 'J' else TZS),
            font=(FT_SUM if sc in 'DH' else FT_AUTO), fill=F_AUTO, align=CEN)
ST4 = S4 + 2 + len(SELL)
put(ws, f'A{ST4}', '合计', font=FT_SUM, fill=F_SUM, align=CEN)
for c in 'BCDF':
    put(ws, f'{c}{ST4}', f'=ROUND(SUM({c}${S4+2}:{c}${ST4-1}),0)', fmt=TZS, font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'E{ST4}', f'=IF($B{ST4}=0,"",ROUND($D{ST4}/$B{ST4},4))', fmt=PCT, font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'G{ST4}', f'=SUM(G${S4+2}:G${ST4-1})', fmt='0', font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'A{ST4+1}', '　折人民币（销售额）', font=Font(name='微软雅黑', size=10), align=LEFT)
put(ws, f'B{ST4+1}', f'=ROUND($B{ST4}/汇率,2)', fmt=CNY, font=FT_AUTO, fill=F_AUTO, align=CEN)

S5 = ST4 + 3
sec(S5, '五、库存还剩多少（金额勾稽，一行一行看得见）')
LINES = [('期初金额', f'={STT("H")}'), ('＋ 本月采购入库', f'={STT("J")}'),
         ('＋ 盘盈', f'={STT("M")}'), ('＋ 矿区退库', f'={STT("O")}'),
         ('－ 一矿二矿自用消耗', f'={STT("Q")}'), ('－ 卖给三矿六矿（成本）', f'={STT("S")}'),
         ('－ 盘亏', f'={STT("U")}'), ('－ 退货给供应商', f'={STT("W")}')]
for i, (lab, f) in enumerate(LINES):
    row(S5 + 1 + i, lab, f)
row(S5 + 9, '＝ 期末金额·先令', f'={STT("Y")}', bold=True)
row(S5 + 10, '　勾稽差额（应该是 0）',
    f'=ROUND({STT("H")}+{STT("J")}+{STT("M")}+{STT("O")}-{STT("Q")}-{STT("S")}'
    f'-{STT("U")}-{STT("W")}-{STT("Y")},0)', bold=True, note='不是 0 就别信这个月的数')
row(S5 + 11, '期末有库存的品种数', f'=COUNTIF({ST("X")},">0")', fmt='0')
row(S5 + 12, '呆滞超过 180 天的品种数',
    f'=COUNTIFS(物料档案!$O${D0}:$O${I_END},">180")', fmt='0')
row(S5 + 13, '成本不可信（含待补价/暂估）的品种数', f'=COUNTIF({ST("Z")},"不可信")', fmt='0')
row(S5 + 14, '　这些品种占期末库存金额',
    f'=IF({STT("Y")}=0,"",ROUND(SUMIF({ST("Z")},"不可信",{ST("Y")})/{STT("Y")},4))', fmt=PCT)

S6 = S5 + 16
sec(S6, '六、这个月的账对不对')
row(S6 + 1, '核对表通过项 / 总项数', '=核对表!$B$3&" / "&核对表!$C$3', fmt=None, bold=True)
row(S6 + 2, '没通过的项', '=核对表!$D$3', fmt=None, note='去【核对表】看红的那几行')
ws.sheet_view.showGridLines = False

# ══════════════════════════════════════════════════════════════════════
# 核对表
# ══════════════════════════════════════════════════════════════════════
ws = newsheet('核对表', '核对表 · 结账前必看，全绿才能把月报发给老板', '', C_AUTO, 6)
for c, w in (('A', 40), ('B', 18), ('C', 18), ('D', 16), ('E', 12), ('F', 46)):
    ws.column_dimensions[c].width = w
put(ws, 'A3', '通过 / 总数 →', font=Font(name='微软雅黑', size=11, bold=True), align=CEN)
CHK_R0 = 6
LEND_OUT = '+'.join(
    f'COUNTIFS({OUTQ("K")},"{OUTK[2]}",{OUTQ("W")},"正常",{OUTQ("L")},"{m}")' for m in SELL) or '0'
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
    ('4 还没录运费的国内货柜数', '=0',
     f'=COUNTIFS(到货批次!$D${S0}:$D${B_END},"{SRC_[0]}",到货批次!$E${S0}:$E${B_END},">0",'
     f'到货批次!$H${S0}:$H${B_END},"<=0")',
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
     f'=COUNTIFS({OUTQ("K")},"{OUTK[1]}",{OUTQ("W")},"正常",{OUTQ("S")},0)',
     '货出去了没开价。搬过来的 241 笔是历史（流水帐只记数量不记价）；'
     '新录的一笔一笔会被出库单红字拦住', '待办'),
    ('10 外销是亏的笔数', '=0',
     f'=COUNTIFS({OUTQ("K")},"{OUTK[1]}",{OUTQ("W")},"正常",{OUTQ("V")},"<0")',
     '卖价低于成本，要么定价错要么成本错'),
    ('11 出库单号重复的笔数', '=0',
     f'=SUMPRODUCT(({OUTQ("B")}<>"")*(COUNTIF({OUTQ("B")},{OUTQ("B")}&"")>1)*1)', '一单两记'),
    ('12 出库没填去向类型或矿区的笔数', '=0',
     f'=COUNTIFS({OUTQ("D")},"?*",{OUTQ("W")},"正常",{OUTQ("K")},"")'
     f'+COUNTIFS({OUTQ("D")},"?*",{OUTQ("W")},"正常",{OUTQ("L")},"")'
     f'-COUNTIFS({OUTQ("D")},"?*",{OUTQ("W")},"正常",{OUTQ("K")},"",{OUTQ("L")},"")',
     '不填就分不清是自己耗的还是卖的'),
    ('13 出入库里物料编码不在档案里的笔数', '=0',
     f'=COUNTIF(入库单!$H${D0}:$H${IN_END},"※*")+COUNTIF(出库单!$E${D0}:$E${O_END},"※*")', ''),
    ('14 本月盘点差异金额·先令', '=0',
     f'=ROUND(SUMIF(月度盘点!$H${S0}:$H${P_END},">0",月度盘点!$L${S0}:$L${P_END})'
     f'+SUMIF(月度盘点!$H${S0}:$H${P_END},0,月度盘点!$L${S0}:$L${P_END}),0)',
     '差异不为 0 要在【月度盘点】填原因并调账', '待办'),
    ('15 盘出差异但还没调账的品种数', '=0',
     f'=COUNTIF(月度盘点!$P${S0}:$P${P_END},"※*")', '盘盈去入库单记、盘亏去出库单记'),
    ('16 借调没填调出矿区的笔数', '=0',
     f'=COUNTIFS({OUTQ("K")},"{OUTK[2]}",{OUTQ("W")},"正常",{OUTQ("M")},"")',
     '借调要写清楚从哪个矿调到哪个矿'),
    ('17 出入库没填日期的笔数', '=0',
     f'=COUNTIFS(入库单!$G${D0}:$G${IN_END},"?*",入库单!$C${D0}:$C${IN_END},"")'
     f'+COUNTIFS({OUTQ("D")},"?*",{OUTQ("C")},"")',
     '没日期的进不了任何月份的报表，物资流水帐搬过来时有 312 行是空的', '待办'),
    ('18 货柜和本地采购都登过的物料数', '=0',
     f'=COUNTIF(参数表!$W${D0}:$W${D0+199},"?*")',
     '这些品种两边数量对不上，库存偏高，要你定以哪份为准', '待办'),
    ('19 借调给「外部」矿区的笔数', '=0', '=' + LEND_OUT,
     '老板说三矿六矿是销售，但库管记的是「借调」。要不要改成对外销售，你定', '待办'),
    ('20 入库/出库有红字提示的笔数', '=0',
     f'=COUNTIF(入库单!$AC${D0}:$AC${IN_END},"※*")+COUNTIF(出库单!$Y${D0}:$Y${O_END},"※*")',
     '逐行提示的汇总，去两张单子最后一列看', '待办'),
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
put(ws, 'C3', f'=COUNTIF($E${CHK_R0}:$E${CHK_END},"OK")+COUNTIF($E${CHK_R0}:$E${CHK_END},"※*")',
    fmt='0', font=Font(name='微软雅黑', size=14, bold=True), fill=F_SUM, align=CEN)
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

sc2(36, '五、这次搬过来的 + 你要知道的几件事')
for i, t in enumerate([
    '搬进来的四份原始资料：',
    '· 集装箱库存明细表 877 行 → 866 行物料进【入库单】，10 行费用进【费用台账】，1 行（卖给三矿的空柜）进【出库单】',
    '· 达市到货费用统计表 → 27 个柜的坦桑当地费用 206 条，合计 10.11 亿先令（≈273 万元）',
    '· FJMZL-JH2026013 货柜（MSKU9402161）→ 145 行物料 42.6 万元 + 11 条国内段费用 22.9 万元',
    '· 物资流水帐 2677 行 → 采购 1412 行进【入库单】、使用 390 行和借调 873 行进【出库单】',
    '· 物料按品名去重后共 1547 个编码（货柜 795 + 流水帐新增 628，重名的沿用同一个编码）',
    '',
    '⚠ 一、两份帐对不上：93 个品种货柜和本地采购都登过，只有 10 种数量一样。',
    '　 半胶手套 货柜 2088 / 流水帐 3751，坑道服 货柜 200 / 流水帐 730，安全帽 货柜 200 / 流水帐 732……',
    '　 两边我都导进来了没有丢，但这 93 个品种的库存数**偏高**。入库单那一行会有琥珀色提示，',
    '　 【参数表】W 列列了全部 93 个编码。你定以哪一份为准，我再按你说的调。',
    '',
    '⚠ 二、12 个品种发出去了却没有入库记录（账面负库存）。',
    '　 钻杆 −10、铁轨 −10、240A铜鼻子 −6、2米钎杆 −6、铅头 −4、胶片手套 −3……',
    '　 这正是你说的「容易丢帐」：东西领走了，进来的那一笔没人记。【核对表】第 5 项盯着它。',
    '',
    '⚠ 三、炸药柴油现在有数量没金额。',
    '　 流水帐补进来之后，炸药 / 导爆管 / 电管 / 电雷管 / 柴油 终于有了，但流水帐只记数量不记价，',
    '　 所以月报第二块的金额还是 0。把这几样的单价补进【入库单】，那一块才有意义。',
    '',
    '⚠ 四、241 笔卖给三矿六矿的货没有售价。',
    '　 流水帐只记了数量。新录的外销只要不填价，出库单那一行就会红字拦住。历史这 241 笔要不要补价，你定。',
    '',
    '⚠ 五、库管记的是「借调」，你说的是「销售」。',
    '　 流水帐里给三矿 72 笔、给六矿 160 笔都记成了「矿区借调」（东西还在公司里，不减总库存、不算收入）。',
    '　 你说三矿六矿是卖出去的。要改成「对外销售」的话，把出库单那几行的去向类型改一下就行，',
    '　 【核对表】第 19 项给了笔数。哪个矿区算「自己」、哪个算「外部」，在【参数表】E/F 两列改。',
    '',
    '⚠ 六、汇率。你说 370，我按 370 记账。但你自己那份达市费用表里用的是 376 / 370.7 / 373.59 / 387，',
    '　 按到货批次分期不同。要不要改成按批次记历史汇率，你说一声。美金按你表里的 2650 / 2670 折的先令。',
    '',
    '另外两处原表数据错误已按你自己的「总运费」列修正并记录在案：',
    '· EMCU1461128 和 MSKV4646463 两行的「船公司费用」填成了先令（1,351,352.5 和 232,910.41），',
    '　 而公式按美金乘 2650 算。按先令处理之后，27 个柜跟你表里的总运费分毫不差。',
], 37):
    put(ws, f'A{i}', t, font=Font(name='微软雅黑', size=9,
        color='C00000' if t.startswith(('⚠', '　')) else '595959'), align=LEFT)
    ws.merge_cells(f'A{i}:D{i}')

sc2(76, '六、体检（活的）')
HEALTH = [
    ('核对表通过项', '=核对表!$B$3&" / "&核对表!$C$3', ''),
    ('还没填单价的入库行数', f'=COUNTIFS(入库单!$Y${D0}:$Y${IN_END},"待补价",入库单!$Z${D0}:$Z${IN_END},"正常")', ''),
    ('还没录运费的批次数', f'=SUMPRODUCT((到货批次!$A${S0}:$A${B_END}<>"")*(到货批次!$E${S0}:$E${B_END}>0)*(到货批次!$H${S0}:$H${B_END}=0))', ''),
    ('账面负库存的品种数', f'=COUNTIF({ST("X")},"<-0.001")', ''),
    ('入库/出库红字提示笔数', f'=COUNTIF(入库单!$AC${D0}:$AC${IN_END},"※*")+COUNTIF(出库单!$Y${D0}:$Y${O_END},"※*")', ''),
    ('期末库存金额·先令', f'={STT("Y")}', '#,##0'),
    ('本月自用消耗·先令', f'={STT("Q")}', '#,##0'),
    ('本月对外销售成本·先令', f'={STT("S")}', '#,##0'),
]
for i, (lab, f, fmt) in enumerate(HEALTH, 77):
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

# 括号少写一个，openpyxl 照存，Excel 打开却会把那一格**静默清空**，重算也查不出来
sys.path.insert(0, os.path.join(os.path.dirname(ROOT), '工具'))
import check_formula, fix_sheet_selection              # noqa: E402
fix_sheet_selection.fix(OUT)     # 一张表都没选中的话 WPS 会判成「工作组」，改一处等于改所有表
if check_formula.scan(OUT):
    raise SystemExit('公式括号不配对，先修了再交付')
