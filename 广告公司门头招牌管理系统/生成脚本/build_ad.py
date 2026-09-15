# -*- coding: utf-8 -*-
"""生成《广告公司项目管理系统.xlsx》。

一个工作簿，12 张表，各岗位各管一张录入表，核算表全是公式不用填。
公式只用 SUMIFS / SUMIF / COUNTIFS / VLOOKUP / INDEX / LOOKUP / EOMONTH，
不用数组公式、不用 FILTER/UNIQUE/XLOOKUP，WPS 在线协作文档里也能跑。
"""
import json, os, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = json.load(open(os.path.join(HERE, 'ad_data.json'), encoding='utf-8'))
OUT = os.path.join(os.path.dirname(HERE), '广告公司项目管理系统.xlsx')

# ── 行数上限 ──────────────────────────────────────────────────────────
N_PROJ, N_ORDER, N_BUY, N_LAB, N_JOUR = 400, 1500, 2000, 1200, 2500
N_ACC, N_CUST, N_SUPP, N_CAT, N_PERSON, N_UNIT = 10, 250, 150, 80, 50, 30
HDR = 3                      # 所有表第 3 行是表头
D0 = 4                       # 录入表从第 4 行开始
S0 = 5                       # 核算表第 4 行是合计，数据从第 5 行开始

# ── 配色 ──────────────────────────────────────────────────────────────
C_IN_HDR, C_AUTO_HDR, C_BASE_HDR = '2F5597', '375623', '806000'
F_IN = PatternFill('solid', fgColor='FFFBEA')      # 请填写
F_AUTO = PatternFill('solid', fgColor='F2F2F2')    # 自动算，别动
F_SUM = PatternFill('solid', fgColor='FFF2CC')     # 合计行
F_TITLE = PatternFill('solid', fgColor='1F3864')
F_NOTE = PatternFill('solid', fgColor='FFF7E6')
FT_AUTO = Font(name='微软雅黑', size=10, color='7F7F7F')
FT_IN = Font(name='微软雅黑', size=10)
FT_HDR = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
FT_TITLE = Font(name='微软雅黑', size=15, bold=True, color='FFFFFF')
FT_SUM = Font(name='微软雅黑', size=10, bold=True, color='C00000')
THIN = Side('thin', color='BFBFBF')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CEN = Alignment('center', 'center', wrap_text=True)
LEFT = Alignment('left', 'center')
MONEY, DATEF, PCT = '#,##0.00', 'yyyy-mm-dd', '0.0%'

wb = openpyxl.Workbook()
wb.remove(wb.active)

def newsheet(name, title, tip, tabcolor):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = tabcolor
    ws['A1'] = title
    ws['A1'].font, ws['A1'].fill, ws['A1'].alignment = FT_TITLE, F_TITLE, LEFT
    ws.row_dimensions[1].height = 30
    ws['A2'] = tip
    ws['A2'].font = Font(name='微软雅黑', size=9, color='8B5E00')
    ws['A2'].fill, ws['A2'].alignment = F_NOTE, LEFT
    ws.row_dimensions[2].height = 22
    return ws

def header(ws, cols, color):
    """cols = [(标题, 列宽, 'in'|'auto'|'key', 数字格式), ...]"""
    fill = PatternFill('solid', fgColor=color)
    for i, (t, w, kind, fmt) in enumerate(cols, 1):
        c = ws.cell(HDR, i, t)
        c.font, c.fill, c.alignment, c.border = FT_HDR, fill, CEN, BOX
        ws.column_dimensions[CL(i)].width = w
    ws.row_dimensions[HDR].height = 34
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(cols))

def body(ws, cols, r0, r1):
    for r in range(r0, r1 + 1):
        for i, (t, w, kind, fmt) in enumerate(cols, 1):
            c = ws.cell(r, i)
            c.border = BOX
            c.font = FT_AUTO if kind == 'auto' else FT_IN
            c.fill = F_AUTO if kind == 'auto' else F_IN
            c.alignment = LEFT if kind == 'in' and fmt is None else CEN
            if fmt:
                c.number_format = fmt

def put(ws, addr, value, fmt=None, font=None, fill=None, align=None):
    c = ws[addr]
    c.value = value
    if fmt: c.number_format = fmt
    if font: c.font = font
    if fill: c.fill = fill
    if align: c.alignment = align
    return c

def dv(ws, formula, rng, warn=False):
    d = DataValidation(type='list', formula1=formula, allow_blank=True,
                       showErrorMessage=True,
                       errorStyle='warning' if warn else 'stop',
                       error='不在名单里。先去【基础资料】把它加上，或者直接选一个已有的。',
                       errorTitle='这一项没在名单里')
    ws.add_data_validation(d)
    d.add(rng)

def redflag(ws, rng):
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=['LEFT(%s,1)="※"' % rng.split(':')[0].replace('$', '')],
        font=Font(color='C00000', bold=True),
        fill=PatternFill('solid', fgColor='FFD7D7')))

# ══════════════════════════════════════════════════════════════════════
# 0  使用说明
# ══════════════════════════════════════════════════════════════════════
ws = newsheet('使用说明', '广告公司项目管理系统 · 使用说明',
              '全公司一个文件。各人只管自己那张录入表，绿色标签的核算表全是公式，不用填也别删。',
              '1F3864')
for i, w in enumerate([16, 15, 46, 30, 16, 16], 1):
    ws.column_dimensions[CL(i)].width = w

def sec(r, txt):
    put(ws, f'A{r}', txt, font=Font(name='微软雅黑', size=12, bold=True, color='FFFFFF'),
        fill=PatternFill('solid', fgColor='2F5597'), align=LEFT)
    ws.merge_cells(f'A{r}:F{r}')
    ws.row_dimensions[r].height = 24

def line(r, a, b='', c='', bold=False):
    f = Font(name='微软雅黑', size=10, bold=bold)
    put(ws, f'A{r}', a, font=f, align=LEFT)
    put(ws, f'B{r}', b, font=f, align=LEFT)
    put(ws, f'C{r}', c, font=Font(name='微软雅黑', size=10), align=Alignment('left', 'center', wrap_text=True))
    ws.merge_cells(f'C{r}:F{r}')

sec(3, '一、谁填哪一张表')
line(4, '岗位', '表名', '填什么', bold=True)
for i, (who, sheet, what) in enumerate([
    ('业务 / 下单', '项目台账', '接到单就开一行：项目名、客户、合同额、下单日期。做完了把「完工日期」填上——利润表就是按这个日期确认收入的。'),
    ('设计 / 下单', '订单明细', '按件报价的客户（一次几十个标牌那种）在这里逐项登记，金额自动汇总进项目台账。整单包死的项目不用填这张。'),
    ('采购', '采购单', '买的每一笔材料：供应商、项目、主材还是辅材、数量单价。金额自动算。'),
    ('设计 / 制作 / 安装', '生产安装', '设计费、制作工时、安装工时、加班、外叫小工、吊车租车，按项目登记。'),
    ('财务 / 出纳', '资金日记账', '五个账户（支付宝、平安银行、中信银行、微信1、微信2）所有进出都记这一张，右边自动结出该账户的即时余额。'),
], 5):
    line(i, who, sheet, what)

sec(11, '二、三条规则，记住就不会乱')
line(12, '规则 1', '成本从哪来',
     '项目成本 = 采购单 + 生产安装 + 日记账里选了成本类别的支出。三个来源相加，所以同一笔钱只能记在其中一个地方。')
line(13, '规则 2', '别把一笔钱记两次',
     '登过采购单的料，出纳付钱时「费用类别」要选『付供应商货款』（它不计成本）；没登采购单的零星开支（配送费、运费、餐费），出纳直接选成本类别即可，不用再补采购单。')
line(14, '规则 3', '收入按完工确认',
     '利润表的营业收入 = 当月填了「完工日期」的项目合同额，成本也跟着那个月走，这样毛利才配得上。没完工的项目算「在制」，不进利润表。长期客户（比如东尼电子）建议按月建一个项目，月底填完工日期。')
line(15, '账户互转', '记两行',
     '从支付宝转到平安银行：支付宝记一行支出、平安银行记一行收入，两行的费用类别都选『账户互转』，它不进利润表。')

sec(17, '三、颜色和符号')
line(18, '淡黄色格子', '请填写', '这是要人填的。')
line(19, '灰色格子', '公式自动算', '别手动改，改了就断了。整列往下都有公式，新增数据直接往下填即可。')
line(20, '※ 开头的红字', '有问题', '「核对」列会自己挑毛病：项目不在台账、类别没选、收支金额对不上、可能和采购单重复。看到红的就去改。')

sec(22, '四、录入体检（活的，随时看这里）')
CHK = [
    ('日记账还没分类的笔数', '=COUNTIF(资金日记账!$E:$E,"待分类")', '0',
     '这些笔的费用类别请挑一个，否则不进利润表'),
    ('　　对应金额（收+支）', '=ROUND(SUMIF(资金日记账!$E:$E,"待分类",资金日记账!$J:$J)+SUMIF(资金日记账!$E:$E,"待分类",资金日记账!$K:$K),2)', '#,##0.00', ''),
    ('成本类支出没填项目的笔数', '=COUNTIF(资金日记账!$P:$P,"※成本类支出没填项目")', '0',
     '补上项目名，这些钱就能进对应项目的成本，利润表也就准了'),
    ('　　对应金额', '=ROUND(SUMIFS(资金日记账!$K:$K,资金日记账!$P:$P,"※成本类支出没填项目"),2)', '#,##0.00', ''),
    ('还没填完工日期的项目数', '=COUNTIFS(项目台账!$B$4:$B$%d,"<>",项目台账!$J$4:$J$%d,"")' % (D0 + N_PROJ - 1, D0 + N_PROJ - 1), '0',
     '这些项目的收入和成本都还没进利润表'),
    ('　　这些项目的合同额', '=ROUND(SUMIFS(项目台账!$M$4:$M$%d,项目台账!$B$4:$B$%d,"<>",项目台账!$J$4:$J$%d,""),2)' % (D0 + N_PROJ - 1, D0 + N_PROJ - 1, D0 + N_PROJ - 1), '#,##0.00', ''),
    ('日记账其它红字提示笔数', '=COUNTIF(资金日记账!$P:$P,"※*")-COUNTIF(资金日记账!$P:$P,"※成本类支出没填项目")', '0', '去【资金日记账】最后一列看'),
    ('采购单有红字提示的笔数', '=COUNTIF(采购单!$N:$N,"※*")', '0', ''),
    ('生产安装有红字提示的笔数', '=COUNTIF(生产安装!$M:$M,"※*")', '0', ''),
    ('项目台账有红字提示的条数', '=COUNTIF(项目台账!$S:$S,"※*")', '0', ''),
    ('订单明细项目对不上的条数', '=COUNTIF(订单明细!$J:$J,"※*")', '0', ''),
]
line(23, '检查项', '数值', '说明', bold=True)
for i, (lab, f, fmt, note) in enumerate(CHK, 24):
    put(ws, f'A{i}', lab, font=Font(name='微软雅黑', size=10), align=LEFT)
    ws.merge_cells(f'A{i}:B{i}')
    put(ws, f'C{i}', f, fmt=fmt, font=Font(name='微软雅黑', size=11, bold=True, color='C00000'), align=CEN)
    put(ws, f'D{i}', note, font=Font(name='微软雅黑', size=9, color='808080'), align=LEFT)
    ws.merge_cells(f'D{i}:F{i}')

sec(36, '五、一眼看家底')
SNAP = [
    ('五个账户余额合计', '=账户余额!$E$14'),
    ('客户还欠我们（应收）', '=项目核算!$K$4'),
    ('我们还欠供应商（应付）', '=供应商对账!$F$4'),
    ('在制项目已投入成本', '=利润表!$B$34'),
    ('本年累计毛利', '=利润表!$N$17'),
    ('本年累计净利', '=利润表!$N$27'),
]
for i, (lab, f) in enumerate(SNAP, 37):
    put(ws, f'A{i}', lab, font=Font(name='微软雅黑', size=10), align=LEFT)
    ws.merge_cells(f'A{i}:B{i}')
    put(ws, f'C{i}', f, fmt=MONEY, font=Font(name='微软雅黑', size=11, bold=True, color='1F3864'), align=CEN)

sec(44, '六、这份表是怎么从旧表搬过来的')
for i, t in enumerate([
    '· 旧文件 22 张表全部取过来了：940 行日记帐、17 张项目表、东尼电子和 Sheet1 两张制作清单、综合汇总表。',
    '· 17 张项目表的成本一分不差地拆进了【采购单】和【生产安装】，逐张核对过（详见 说明.md）。',
    '· 日记帐前半段（3~4 月）原来手工分过类，照搬；后半段是微信账单原样粘贴，只有摘要，',
    '　按关键字能认出来的自动归了类，认不出的一律标「待分类」，没有瞎猜。请按上面「体检」里的笔数去补。',
    '· 项目名做了归并：兰溪牛肉面=兰溪手擀面、津华钢铁=德清津华钢铁、巴比贝克=芭比贝克…… 完整对照表在 说明.md。',
    '· 「织里店-炳秀」这类写法拆成了 项目=织里店、往来单位=炳秀。',
    '· 完工日期只给「钱已经收齐」的项目补了，其余留空，请自己填——它直接决定利润表哪个月确认收入。',
], 45):
    put(ws, f'A{i}', t, font=Font(name='微软雅黑', size=9, color='595959'), align=LEFT)
    ws.merge_cells(f'A{i}:F{i}')
ws.sheet_view.showGridLines = False

# ══════════════════════════════════════════════════════════════════════
# 1  基础资料
# ══════════════════════════════════════════════════════════════════════
ws = newsheet('基础资料', '基础资料 · 所有下拉菜单的来源',
              '要加客户、加供应商、加人，就在这里往下接着写，别留空行。改了这里，所有表的下拉菜单立刻跟着变。',
              '806000')
ACCOUNTS = [('支付宝', 0), ('平安银行', 0), ('中信银行', 0), ('微信1', 0), ('微信2', 0), ('现金', 0)]
GUI = ['收款-项目', '收入-其他', '成本-主材', '成本-辅材', '成本-人工', '成本-车运', '成本-其他',
       '费用-管理', '费用-销售', '费用-财务', '其他-营业外收入', '其他-营业外支出',
       '往来-货款', '往来-借贷', '内部转账']
CATS = [
    ('项目收款', '收款-项目'), ('预收定金', '收款-项目'), ('退还客户款', '收款-项目'),
    ('零星物料销售', '收入-其他'), ('设计费收入', '收入-其他'), ('其他收入', '收入-其他'),
    ('主材采购', '成本-主材'), ('外发加工', '成本-主材'),
    ('辅材采购', '成本-辅材'), ('网购材料耗材', '成本-辅材'), ('五金耗材', '成本-辅材'),
    ('设计费', '成本-人工'), ('制作工资', '成本-人工'), ('安装工资', '成本-人工'),
    ('加班费', '成本-人工'), ('外叫小工', '成本-人工'),
    ('运费/快递费', '成本-车运'), ('车费/油费/过路费', '成本-车运'), ('吊车/租车费', '成本-车运'),
    ('项目餐费补助', '成本-其他'), ('其他直接费用', '成本-其他'),
    ('员工工资', '费用-管理'), ('社保公积金', '费用-管理'), ('房租水电', '费用-管理'),
    ('办公用品', '费用-管理'), ('通讯费', '费用-管理'), ('差旅费', '费用-管理'),
    ('业务招待/餐费', '费用-管理'), ('团建费', '费用-管理'), ('车辆费用(非项目)', '费用-管理'),
    ('税金', '费用-管理'), ('其他管理费用', '费用-管理'),
    ('广告推广', '费用-销售'), ('平台服务费', '费用-销售'), ('业务提成', '费用-销售'),
    ('银行手续费', '费用-财务'), ('利息支出', '费用-财务'), ('利息收入', '费用-财务'),
    ('付供应商货款', '往来-货款'), ('代收代付', '往来-货款'),
    ('借出款', '往来-借贷'), ('收回借出款', '往来-借贷'), ('借入款', '往来-借贷'),
    ('归还借入款', '往来-借贷'), ('股东投入', '往来-借贷'), ('股东取款', '往来-借贷'),
    ('员工借支', '往来-借贷'), ('收回员工借支', '往来-借贷'),
    ('账户互转', '内部转账'),
    ('营业外收入', '其他-营业外收入'), ('营业外支出', '其他-营业外支出'),
    ('待分类', ''),
]
PTYPES = ['门头招牌', '发光字/灯箱', '亮化工程', '室内标识', '广告物料', '车贴喷绘', '维修保养', '其他']
MCATS = [('主材', '成本-主材'), ('外发加工', '成本-主材'), ('辅材', '成本-辅材'),
         ('五金耗材', '成本-辅材'), ('其他材料', '成本-其他')]
WCATS = [('设计费', '成本-人工'), ('制作工时', '成本-人工'), ('安装工时', '成本-人工'),
         ('加班费', '成本-人工'), ('外叫小工', '成本-人工'), ('其他人工', '成本-人工'),
         ('吊车/租车', '成本-车运'), ('运费/搬运', '成本-车运')]
UNITS = ['个', '套', '块', '张', '米', '平方', '根', '条', '箱', '桶', '支', '卷', '工', '天',
         '次', '趟', '项', '批', '台', '组', 'm²', '千克']
BLOCKS = [
    ('A', '资金账户', 16, [a for a, _ in ACCOUNTS], N_ACC, 'B', '期初余额', 12, [b for _, b in ACCOUNTS]),
    ('E', '客户名称', 22, DATA['customers'], N_CUST, None, None, None, None),
    ('G', '供应商名称', 22, DATA['suppliers'], N_SUPP, None, None, None, None),
    ('I', '费用类别', 20, [a for a, _ in CATS], N_CAT, 'J', '归属大类', 16, [b for _, b in CATS]),
    ('L', '归属大类清单', 18, GUI, 20, None, None, None, None),
    ('N', '人员', 14, DATA['persons'], N_PERSON, None, None, None, None),
    ('P', '项目类型', 16, PTYPES, 12, None, None, None, None),
    ('R', '材料类别', 14, [a for a, _ in MCATS], 12, 'S', '成本归属', 14, [b for _, b in MCATS]),
    ('U', '工作类型', 14, [a for a, _ in WCATS], 14, 'V', '成本归属', 14, [b for _, b in WCATS]),
    ('X', '计量单位', 12, UNITS, N_UNIT, None, None, None, None),
    ('Z', '往来单位(自动合并)', 24, None, 400, None, None, None, None),
    ('AB', '收支方向', 12, ['收入', '支出'], 4, None, None, None, None),
    ('AD', '是否外叫', 12, ['自有员工', '外叫'], 4, None, None, None, None),
]
hf = PatternFill('solid', fgColor=C_BASE_HDR)
maxcol = 30
for col, title, width, items, n, col2, title2, width2, items2 in BLOCKS:
    for cc, tt, ww in ((col, title, width), (col2, title2, width2)):
        if cc:
            c = put(ws, f'{cc}{HDR}', tt, font=FT_HDR, fill=hf, align=CEN)
            c.border = BOX
            ws.column_dimensions[cc].width = ww
    for i in range(n):
        r = D0 + i
        for cc, lst in ((col, items), (col2, items2)):
            if not cc:
                continue
            c = ws[f'{cc}{r}']
            c.border, c.font, c.alignment = BOX, FT_IN, LEFT
            c.fill = F_AUTO if cc == 'Z' else F_IN
            if lst and i < len(lst):
                c.value = lst[i]
            if cc in ('B',):
                c.number_format = MONEY
                c.alignment = CEN
ws['A1'].alignment = LEFT
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=maxcol)
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=maxcol)
# Z 列：客户 + 供应商 自动接成一列，给日记账的「往来单位」当下拉
CE, CG = f'$E${D0}:$E${D0+N_CUST-1}', f'$G${D0}:$G${D0+N_SUPP-1}'
for i in range(400):
    r = D0 + i
    put(ws, f'Z{r}',
        f'=IF(ROW()-{D0-1}<=COUNTA({CE}),INDEX({CE},ROW()-{D0-1}),'
        f'IF(ROW()-{D0-1}-COUNTA({CE})<=COUNTA({CG}),INDEX({CG},ROW()-{D0-1}-COUNTA({CE})),""))',
        font=FT_AUTO)
ws.freeze_panes = f'A{D0}'

R_ACC, R_CUST, R_SUPP, R_CAT = D0 + N_ACC - 1, D0 + N_CUST - 1, D0 + N_SUPP - 1, D0 + N_CAT - 1
NAMES = {
    '账户表': f'基础资料!$A${D0}:$A${R_ACC}',
    '账户期初': f'基础资料!$A${D0}:$B${R_ACC}',
    '客户表': f'基础资料!$E${D0}:$E${R_CUST}',
    '供应商表': f'基础资料!$G${D0}:$G${R_SUPP}',
    '类别表': f'基础资料!$I${D0}:$I${R_CAT}',
    '类别归属': f'基础资料!$I${D0}:$J${R_CAT}',
    '归属表': f'基础资料!$L${D0}:$L${D0+19}',
    '人员表': f'基础资料!$N${D0}:$N${D0+N_PERSON-1}',
    '项目类型表': f'基础资料!$P${D0}:$P${D0+11}',
    '材料类别表': f'基础资料!$R${D0}:$R${D0+11}',
    '材料归属': f'基础资料!$R${D0}:$S${D0+11}',
    '工作类型表': f'基础资料!$U${D0}:$U${D0+13}',
    '工作归属': f'基础资料!$U${D0}:$V${D0+13}',
    '单位表': f'基础资料!$X${D0}:$X${D0+N_UNIT-1}',
    '往来单位表': f'基础资料!$Z${D0}:$Z${D0+399}',
    '收支表': f'基础资料!$AB${D0}:$AB${D0+3}',
    '外叫表': f'基础资料!$AD${D0}:$AD${D0+3}',
    '项目名表': f'项目台账!$B${D0}:$B${D0+N_PROJ-1}',
}

# ══════════════════════════════════════════════════════════════════════
# 2  项目台账
# ══════════════════════════════════════════════════════════════════════
P_END = D0 + N_PROJ - 1
ws = newsheet('项目台账', '项目台账 · 业务 / 下单岗位填',
              '接到单就开一行。做完了一定要填「完工日期」——利润表就是按这个日期把收入和成本算进当月的。',
              C_IN_HDR)
cols = [('项目编号', 10, 'auto', None), ('项目名称 ★', 20, 'in', None), ('客户名称 ★', 18, 'in', None),
        ('项目类型', 13, 'in', None), ('负责人', 10, 'in', None), ('下单日期 ★', 12, 'in', DATEF),
        ('设计完成', 11, 'in', DATEF), ('制作完成', 11, 'in', DATEF), ('安装完成', 11, 'in', DATEF),
        ('完工日期 ★\n(确认收入)', 12, 'in', DATEF), ('合同额\n(整单包死)', 13, 'in', MONEY),
        ('订单明细额\n(自动)', 13, 'auto', MONEY), ('合同总额', 13, 'auto', MONEY),
        ('期初已收\n(导入用)', 12, 'in', MONEY), ('已收款', 13, 'auto', MONEY),
        ('未收款', 13, 'auto', MONEY), ('状态', 14, 'auto', None), ('备注', 22, 'in', None),
        ('核对', 22, 'auto', None)]
header(ws, cols, C_IN_HDR); body(ws, cols, D0, P_END)
JR = f'资金日记账!$J${D0}:$J${D0+N_JOUR-1}'
KR = f'资金日记账!$K${D0}:$K${D0+N_JOUR-1}'
GR = f'资金日记账!$G${D0}:$G${D0+N_JOUR-1}'
FR = f'资金日记账!$F${D0}:$F${D0+N_JOUR-1}'
for r in range(D0, P_END + 1):
    g = f'IF($B{r}="","",'
    put(ws, f'A{r}', f'={g}"P"&TEXT(ROW()-{D0-1},"000"))')
    put(ws, f'L{r}', f'={g}ROUND(SUMIF(订单明细!$C:$C,$B{r},订单明细!$I:$I),2))')
    put(ws, f'M{r}', f'={g}ROUND(N($K{r})+N($L{r}),2))')
    put(ws, f'O{r}', f'={g}ROUND(N($N{r})+SUMIFS({JR},{GR},$B{r},{FR},"收款-项目")'
                     f'-SUMIFS({KR},{GR},$B{r},{FR},"收款-项目"),2))')
    put(ws, f'P{r}', f'={g}ROUND($M{r}-$O{r},2))')
    put(ws, f'Q{r}', f'={g}IF(AND($J{r}<>"",ROUND($P{r},2)<=0),"已完工·已结清",'
                     f'IF($J{r}<>"","已完工·还有尾款",IF($I{r}<>"","已安装",IF($H{r}<>"","已制作",'
                     f'IF($G{r}<>"","设计完成","已下单"))))))')
    put(ws, f'S{r}', f'={g}IF(COUNTIF($B${D0}:$B${P_END},$B{r})>1,"※项目名重复了",'
                     f'IF(AND($C{r}<>"",ISNA(MATCH($C{r},客户表,0))),"※客户不在基础资料里",'
                     f'IF(AND($J{r}<>"",$F{r}<>"",$J{r}<$F{r}),"※完工日期早于下单日期",""))))')
dv(ws, '=客户表', f'C{D0}:C{P_END}', warn=True)
dv(ws, '=项目类型表', f'D{D0}:D{P_END}')
dv(ws, '=人员表', f'E{D0}:E{P_END}', warn=True)
redflag(ws, f'S{D0}:S{P_END}')
ws.freeze_panes = f'C{D0}'
ws.auto_filter.ref = f'A{HDR}:S{P_END}'

# ══════════════════════════════════════════════════════════════════════
# 3  订单明细
# ══════════════════════════════════════════════════════════════════════
O_END = D0 + N_ORDER - 1
ws = newsheet('订单明细', '订单明细 · 按件报价的客户用（下单 / 设计岗位填）',
              '一次几十个标牌那种客户在这里逐项登记，金额会自动加进【项目台账】的「订单明细额」。整单包死的项目不用填这张表。',
              C_IN_HDR)
cols = [('序号', 7, 'auto', None), ('下单日期 ★', 12, 'in', DATEF), ('项目名称 ★', 20, 'in', None),
        ('品名 ★', 26, 'in', None), ('规格', 20, 'in', None), ('单位', 8, 'in', None),
        ('数量', 10, 'in', '#,##0.##'), ('单价', 10, 'in', MONEY), ('金额', 13, 'auto', MONEY),
        ('客户', 16, 'auto', None), ('备注', 24, 'in', None)]
header(ws, cols, C_IN_HDR); body(ws, cols, D0, O_END)
for r in range(D0, O_END + 1):
    g = f'IF($C{r}="","",'
    put(ws, f'A{r}', f'={g}ROW()-{D0-1})')
    put(ws, f'I{r}', f'={g}ROUND(IF($G{r}="",1,$G{r})*N($H{r}),2))')
    put(ws, f'J{r}', f'={g}IFERROR(VLOOKUP($C{r},项目台账!$B:$C,2,0),"※项目不在项目台账里"))')
dv(ws, '=项目名表', f'C{D0}:C{O_END}', warn=True)
dv(ws, '=单位表', f'F{D0}:F{O_END}', warn=True)
redflag(ws, f'J{D0}:J{O_END}')
ws.freeze_panes = f'D{D0}'
ws.auto_filter.ref = f'A{HDR}:K{O_END}'

# ══════════════════════════════════════════════════════════════════════
# 4  采购单
# ══════════════════════════════════════════════════════════════════════
B_END = D0 + N_BUY - 1
ws = newsheet('采购单', '采购单 · 采购岗位填',
              '买的每一笔材料都记一行。注意：登过这里的料，出纳付钱时费用类别要选『付供应商货款』，不然成本会算两遍。',
              C_IN_HDR)
cols = [('序号', 7, 'auto', None), ('采购日期 ★', 12, 'in', DATEF), ('供应商 ★', 18, 'in', None),
        ('项目名称 ★', 20, 'in', None), ('材料类别 ★', 12, 'in', None), ('成本归属', 12, 'auto', None),
        ('品名 ★', 24, 'in', None), ('规格', 18, 'in', None), ('单位', 8, 'in', None),
        ('数量', 10, 'in', '#,##0.##'), ('单价', 10, 'in', MONEY), ('金额', 13, 'auto', MONEY),
        ('备注', 20, 'in', None), ('核对', 24, 'auto', None)]
header(ws, cols, C_IN_HDR); body(ws, cols, D0, B_END)
for r in range(D0, B_END + 1):
    g = f'IF($D{r}="","",'
    put(ws, f'A{r}', f'={g}ROW()-{D0-1})')
    put(ws, f'F{r}', f'={g}IFERROR(VLOOKUP($E{r},材料归属,2,0),"※材料类别没选"))')
    put(ws, f'L{r}', f'={g}ROUND(IF($J{r}="",1,$J{r})*N($K{r}),2))')
    put(ws, f'N{r}', f'={g}IF(ISNA(MATCH($D{r},项目名表,0)),"※项目不在项目台账里",'
                     f'IF($F{r}="※材料类别没选","※材料类别没选",'
                     f'IF(AND($C{r}<>"",ISNA(MATCH($C{r},供应商表,0))),"※供应商不在基础资料里",""))))')
dv(ws, '=供应商表', f'C{D0}:C{B_END}', warn=True)
dv(ws, '=项目名表', f'D{D0}:D{B_END}', warn=True)
dv(ws, '=材料类别表', f'E{D0}:E{B_END}')
dv(ws, '=单位表', f'I{D0}:I{B_END}', warn=True)
redflag(ws, f'F{D0}:F{B_END}'); redflag(ws, f'N{D0}:N{B_END}')
ws.freeze_panes = f'D{D0}'
ws.auto_filter.ref = f'A{HDR}:N{B_END}'

# ══════════════════════════════════════════════════════════════════════
# 5  生产安装
# ══════════════════════════════════════════════════════════════════════
L_END = D0 + N_LAB - 1
ws = newsheet('生产安装', '生产安装 · 设计 / 制作 / 安装岗位填',
              '设计费、制作工时、安装工时、加班、外叫小工、吊车租车都记这里，按项目分开记，会自动进项目成本。',
              C_IN_HDR)
cols = [('序号', 7, 'auto', None), ('日期 ★', 12, 'in', DATEF), ('项目名称 ★', 20, 'in', None),
        ('工作类型 ★', 13, 'in', None), ('成本归属', 12, 'auto', None), ('人员 / 班组', 14, 'in', None),
        ('单位', 8, 'in', None), ('数量', 10, 'in', '#,##0.##'), ('单价', 10, 'in', MONEY),
        ('金额', 13, 'auto', MONEY), ('自有/外叫', 11, 'in', None), ('备注', 22, 'in', None),
        ('核对', 24, 'auto', None)]
header(ws, cols, C_IN_HDR); body(ws, cols, D0, L_END)
for r in range(D0, L_END + 1):
    g = f'IF($C{r}="","",'
    put(ws, f'A{r}', f'={g}ROW()-{D0-1})')
    put(ws, f'E{r}', f'={g}IFERROR(VLOOKUP($D{r},工作归属,2,0),"※工作类型没选"))')
    put(ws, f'J{r}', f'={g}ROUND(IF($H{r}="",1,$H{r})*N($I{r}),2))')
    put(ws, f'M{r}', f'={g}IF(ISNA(MATCH($C{r},项目名表,0)),"※项目不在项目台账里",'
                     f'IF($E{r}="※工作类型没选","※工作类型没选","")))')
dv(ws, '=项目名表', f'C{D0}:C{L_END}', warn=True)
dv(ws, '=工作类型表', f'D{D0}:D{L_END}')
dv(ws, '=人员表', f'F{D0}:F{L_END}', warn=True)
dv(ws, '=单位表', f'G{D0}:G{L_END}', warn=True)
dv(ws, '=外叫表', f'K{D0}:K{L_END}', warn=True)
redflag(ws, f'E{D0}:E{L_END}'); redflag(ws, f'M{D0}:M{L_END}')
ws.freeze_panes = f'D{D0}'
ws.auto_filter.ref = f'A{HDR}:M{L_END}'

# ══════════════════════════════════════════════════════════════════════
# 6  资金日记账
# ══════════════════════════════════════════════════════════════════════
J_END = D0 + N_JOUR - 1
ws = newsheet('资金日记账', '资金日记账 · 出纳填（五个账户混在一起记）',
              '五个账户的钱都记这一张，按日期顺着往下记就行。L 列自动结出「这个账户」当时的余额，M 列是全部账户合计。'
              '账户之间互转记两行（转出一行支出、转入一行收入），类别都选『账户互转』。',
              C_IN_HDR)
cols = [('序号', 7, 'auto', None), ('日期 ★', 12, 'in', DATEF), ('账户 ★', 12, 'in', None),
        ('收/支 ★', 9, 'in', None), ('费用类别 ★', 17, 'in', None), ('归属大类', 14, 'auto', None),
        ('项目名称', 18, 'in', None), ('往来单位', 18, 'in', None), ('摘要 ★', 30, 'in', None),
        ('收入金额', 13, 'in', MONEY), ('支出金额', 13, 'in', MONEY),
        ('本账户余额', 14, 'auto', MONEY), ('全账户合计', 14, 'auto', MONEY),
        ('经手人', 10, 'in', None), ('备注', 20, 'in', None), ('核对', 26, 'auto', None)]
header(ws, cols, C_IN_HDR); body(ws, cols, D0, J_END)
ACC_SUM = f'SUM(基础资料!$B${D0}:$B${R_ACC})'
for r in range(D0, J_END + 1):
    g = f'IF($B{r}="","",'
    put(ws, f'A{r}', f'={g}ROW()-{D0-1})')
    put(ws, f'F{r}', f'=IF($E{r}="","",IFERROR(VLOOKUP($E{r},类别归属,2,0),"※类别不在基础资料里"))')
    put(ws, f'L{r}', f'=IF($C{r}="","",ROUND(IFERROR(VLOOKUP($C{r},账户期初,2,0),0)'
                     f'+SUMIFS($J${D0}:J{r},$C${D0}:C{r},$C{r})-SUMIFS($K${D0}:K{r},$C${D0}:C{r},$C{r}),2))')
    put(ws, f'M{r}', f'=IF($C{r}="","",ROUND({ACC_SUM}+SUM($J${D0}:J{r})-SUM($K${D0}:K{r}),2))')
    put(ws, f'P{r}', f'={g}'
        f'IF($F{r}="※类别不在基础资料里","※费用类别不在基础资料里",'
        f'IF(AND(N($J{r})=0,N($K{r})=0),"※这行没填金额",'
        f'IF(AND(N($J{r})>0,N($K{r})>0),"※收入和支出不能同一行都填",'
        f'IF(AND($D{r}="收入",N($J{r})=0),"※收入行要填收入金额",'
        f'IF(AND($D{r}="支出",N($K{r})=0),"※支出行要填支出金额",'
        f'IF(AND($C{r}<>"",ISNA(MATCH($C{r},账户表,0))),"※账户不在基础资料里",'
        f'IF(AND($G{r}<>"",ISNA(MATCH($G{r},项目名表,0))),"※项目不在项目台账里",'
        f'IF(AND(LEFT($F{r},3)="成本-",$G{r}=""),"※成本类支出没填项目",'
        f'IF(AND(LEFT($F{r},3)="成本-",$H{r}<>"",N($K{r})>0,'
        f'COUNTIFS(采购单!$B:$B,$B{r},采购单!$C:$C,$H{r},采购单!$L:$L,$K{r})>0),'
        f'"※这笔可能和采购单重复了",""))))))))))')
dv(ws, '=账户表', f'C{D0}:C{J_END}')
dv(ws, '=收支表', f'D{D0}:D{J_END}')
dv(ws, '=类别表', f'E{D0}:E{J_END}')
dv(ws, '=项目名表', f'G{D0}:G{J_END}', warn=True)
dv(ws, '=往来单位表', f'H{D0}:H{J_END}', warn=True)
dv(ws, '=人员表', f'N{D0}:N{J_END}', warn=True)
redflag(ws, f'F{D0}:F{J_END}'); redflag(ws, f'P{D0}:P{J_END}')
ws.conditional_formatting.add(f'D{D0}:D{J_END}', CellIsRule(
    operator='equal', formula=['"收入"'], font=Font(color='006100', bold=True)))
ws.conditional_formatting.add(f'D{D0}:D{J_END}', CellIsRule(
    operator='equal', formula=['"支出"'], font=Font(color='9C0006')))
ws.freeze_panes = f'C{D0}'
ws.auto_filter.ref = f'A{HDR}:P{J_END}'

# ══════════════════════════════════════════════════════════════════════
# 7  账户余额
# ══════════════════════════════════════════════════════════════════════
ws = newsheet('账户余额', '账户余额 · 自动，不用填',
              '每天收工前用「实盘余额」那一栏跟手机上的真实余额对一下，差异不为零就说明日记账少记或多记了。',
              C_AUTO_HDR)
cols = [('账户', 16, 'auto', None), ('期初余额', 14, 'auto', MONEY), ('累计收入', 14, 'auto', MONEY),
        ('累计支出', 14, 'auto', MONEY), ('当前余额', 15, 'auto', MONEY),
        ('本月收入', 14, 'auto', MONEY), ('本月支出', 14, 'auto', MONEY),
        ('笔数', 9, 'auto', '0'), ('最后一笔', 13, 'auto', DATEF)]
header(ws, cols, C_AUTO_HDR); body(ws, cols, D0, D0 + N_ACC)
JB = f'资金日记账!$B${D0}:$B${J_END}'
JC = f'资金日记账!$C${D0}:$C${J_END}'
JJ = f'资金日记账!$J${D0}:$J${J_END}'
JK = f'资金日记账!$K${D0}:$K${J_END}'
JF = f'资金日记账!$F${D0}:$F${J_END}'
JG = f'资金日记账!$G${D0}:$G${J_END}'
JH = f'资金日记账!$H${D0}:$H${J_END}'
A_END = D0 + N_ACC - 1
for i in range(N_ACC):
    r, br = D0 + i, D0 + i
    g = f'IF($A{r}="","",'
    put(ws, f'A{r}', f'=IF(基础资料!$A{br}="","",基础资料!$A{br})')
    put(ws, f'B{r}', f'={g}N(基础资料!$B{br}))')
    put(ws, f'C{r}', f'={g}ROUND(SUMIF({JC},$A{r},{JJ}),2))')
    put(ws, f'D{r}', f'={g}ROUND(SUMIF({JC},$A{r},{JK}),2))')
    put(ws, f'E{r}', f'={g}ROUND($B{r}+$C{r}-$D{r},2))')
    put(ws, f'F{r}', f'={g}ROUND(SUMIFS({JJ},{JC},$A{r},{JB},">="&$D$18,{JB},"<="&$E$18),2))')
    put(ws, f'G{r}', f'={g}ROUND(SUMIFS({JK},{JC},$A{r},{JB},">="&$D$18,{JB},"<="&$E$18),2))')
    put(ws, f'H{r}', f'={g}COUNTIF({JC},$A{r}))')
    put(ws, f'I{r}', f'={g}IFERROR(LOOKUP(1,0/({JC}=$A{r}),{JB}),""))')
TOT = D0 + N_ACC
put(ws, f'A{TOT}', '合计', font=FT_SUM, fill=F_SUM, align=CEN)
for c in 'BCDEFG':
    put(ws, f'{c}{TOT}', f'=ROUND(SUM({c}{D0}:{c}{A_END}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'H{TOT}', f'=SUM(H{D0}:H{A_END})', fmt='0', font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'I{TOT}', '', fill=F_SUM)

put(ws, 'A17', '查询月份 →', font=Font(name='微软雅黑', size=10, bold=True), align=LEFT)
put(ws, 'B17', '年', font=Font(name='微软雅黑', size=9, color='808080'), align=CEN)
put(ws, 'C17', '月', font=Font(name='微软雅黑', size=9, color='808080'), align=CEN)
put(ws, 'D17', '本月第一天', font=Font(name='微软雅黑', size=9, color='808080'), align=CEN)
put(ws, 'E17', '本月最后一天', font=Font(name='微软雅黑', size=9, color='808080'), align=CEN)
put(ws, 'B18', 2026, fmt='0', font=Font(name='微软雅黑', size=11, bold=True, color='C00000'),
    fill=F_IN, align=CEN)
put(ws, 'C18', 9, fmt='0', font=Font(name='微软雅黑', size=11, bold=True, color='C00000'),
    fill=F_IN, align=CEN)
put(ws, 'D18', '=DATE($B$18,$C$18,1)', fmt=DATEF, font=FT_AUTO, fill=F_AUTO, align=CEN)
put(ws, 'E18', '=EOMONTH($D$18,0)', fmt=DATEF, font=FT_AUTO, fill=F_AUTO, align=CEN)
for a in ('B18', 'C18', 'D18', 'E18'):
    ws[a].border = BOX

put(ws, 'A20', '每天对一下账（实盘余额自己填）',
    font=Font(name='微软雅黑', size=11, bold=True, color='FFFFFF'),
    fill=PatternFill('solid', fgColor=C_AUTO_HDR), align=LEFT)
ws.merge_cells('A20:E20')
for j, t in enumerate(['账户', '账面余额', '实盘余额(手填)', '差异', '说明'], 1):
    c = ws.cell(21, j, t)
    c.font, c.fill, c.alignment, c.border = FT_HDR, PatternFill('solid', fgColor=C_AUTO_HDR), CEN, BOX
for i in range(N_ACC):
    r = 22 + i
    put(ws, f'A{r}', f'=IF($A{D0+i}="","",$A{D0+i})', font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'B{r}', f'=IF($A{r}="","",$E{D0+i})', fmt=MONEY, font=FT_AUTO, fill=F_AUTO, align=CEN)
    put(ws, f'C{r}', None, fmt=MONEY, font=FT_IN, fill=F_IN, align=CEN)
    put(ws, f'D{r}', f'=IF(OR($A{r}="",$C{r}=""),"",ROUND($C{r}-$B{r},2))', fmt=MONEY,
        font=Font(name='微软雅黑', size=10, bold=True), fill=F_AUTO, align=CEN)
    put(ws, f'E{r}', f'=IF($D{r}="","",IF(ROUND($D{r},2)=0,"对得上","※差 "&TEXT($D{r},"0.00")&"，查日记账"))',
        font=FT_AUTO, fill=F_AUTO, align=LEFT)
    for c in 'ABCDE':
        ws[f'{c}{r}'].border = BOX
redflag(ws, f'E22:E{21+N_ACC}')
ws.sheet_view.showGridLines = False

# ══════════════════════════════════════════════════════════════════════
# 8  项目核算
# ══════════════════════════════════════════════════════════════════════
K_END = S0 + N_PROJ - 1
ws = newsheet('项目核算', '项目核算 · 自动，不用填（一个项目一行）',
              '成本 = 采购单 + 生产安装 + 日记账里选了成本类别的支出，三个地方自动加起来。'
              '想看某个项目的明细，去对应的表按项目名筛选即可。',
              C_AUTO_HDR)
cols = [('项目编号', 10, 'auto', None), ('项目名称', 20, 'auto', None), ('客户', 16, 'auto', None),
        ('类型', 12, 'auto', None), ('负责人', 10, 'auto', None), ('下单日期', 12, 'auto', DATEF),
        ('完工日期', 12, 'auto', DATEF), ('状态', 14, 'auto', None),
        ('合同总额', 13, 'auto', MONEY), ('已收款', 13, 'auto', MONEY), ('未收款', 13, 'auto', MONEY),
        ('主材', 12, 'auto', MONEY), ('辅材', 12, 'auto', MONEY), ('人工\n(设计制作安装)', 13, 'auto', MONEY),
        ('运费车费', 12, 'auto', MONEY), ('其他直接', 12, 'auto', MONEY),
        ('成本合计', 13, 'auto', MONEY), ('毛利', 13, 'auto', MONEY), ('毛利率', 10, 'auto', PCT),
        ('完工年月', 10, 'auto', '0'), ('备注', 20, 'auto', None)]
header(ws, cols, C_AUTO_HDR); body(ws, cols, S0 - 1, K_END)
BUY_AMT, BUY_PROJ, BUY_GUI = f'采购单!$L${D0}:$L${B_END}', f'采购单!$D${D0}:$D${B_END}', f'采购单!$F${D0}:$F${B_END}'
LAB_AMT, LAB_PROJ, LAB_GUI = f'生产安装!$J${D0}:$J${L_END}', f'生产安装!$C${D0}:$C${L_END}', f'生产安装!$E${D0}:$E${L_END}'
COSTCOL = [('L', '成本-主材'), ('M', '成本-辅材'), ('N', '成本-人工'), ('O', '成本-车运'), ('P', '成本-其他')]
for i in range(N_PROJ):
    r, pr = S0 + i, D0 + i
    g = f'IF($B{r}="","",'
    put(ws, f'A{r}', f'=IF(项目台账!$B{pr}="","",项目台账!$A{pr})')
    put(ws, f'B{r}', f'=IF(项目台账!$B{pr}="","",项目台账!$B{pr})')
    for col, src in (('C', 'C'), ('D', 'D'), ('E', 'E'), ('F', 'F'), ('G', 'J'), ('H', 'Q'), ('U', 'R')):
        put(ws, f'{col}{r}', f'={g}IF(项目台账!${src}{pr}="","",项目台账!${src}{pr}))')
    put(ws, f'I{r}', f'={g}项目台账!$M{pr})')
    put(ws, f'J{r}', f'={g}项目台账!$O{pr})')
    put(ws, f'K{r}', f'={g}项目台账!$P{pr})')
    for col, gui in COSTCOL:
        put(ws, f'{col}{r}',
            f'={g}ROUND(SUMIFS({BUY_AMT},{BUY_PROJ},$B{r},{BUY_GUI},"{gui}")'
            f'+SUMIFS({LAB_AMT},{LAB_PROJ},$B{r},{LAB_GUI},"{gui}")'
            f'+SUMIFS({JK},{JG},$B{r},{JF},"{gui}")-SUMIFS({JJ},{JG},$B{r},{JF},"{gui}"),2))')
    put(ws, f'Q{r}', f'={g}ROUND(SUM($L{r}:$P{r}),2))')
    put(ws, f'R{r}', f'={g}ROUND($I{r}-$Q{r},2))')
    put(ws, f'S{r}', f'={g}IF(ROUND($I{r},2)=0,"",ROUND($R{r}/$I{r},4)))')
    put(ws, f'T{r}', f'={g}IF($G{r}="",0,YEAR($G{r})*100+MONTH($G{r})))')
put(ws, f'A{S0-1}', '合计', font=FT_SUM, fill=F_SUM, align=CEN)
for c in 'IJKLMNOPQR':
    put(ws, f'{c}{S0-1}', f'=ROUND(SUM({c}{S0}:{c}{K_END}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'S{S0-1}', f'=IF(ROUND($I{S0-1},2)=0,"",ROUND($R{S0-1}/$I{S0-1},4))',
    fmt=PCT, font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'B{S0-1}', f'=COUNTIF($B{S0}:$B{K_END},"?*")&" 个项目"', font=FT_SUM, fill=F_SUM, align=CEN)
ws.column_dimensions['T'].hidden = True
ws.conditional_formatting.add(f'R{S0}:R{K_END}', CellIsRule(
    operator='lessThan', formula=['0'], font=Font(color='C00000', bold=True)))
ws.conditional_formatting.add(f'K{S0}:K{K_END}', CellIsRule(
    operator='greaterThan', formula=['0.004'], font=Font(color='BF8F00', bold=True)))
ws.conditional_formatting.add(f'K{S0}:K{K_END}', CellIsRule(
    operator='lessThan', formula=['-0.004'], font=Font(color='0070C0', bold=True)))
ws.freeze_panes = f'C{S0}'
ws.auto_filter.ref = f'A{HDR}:U{K_END}'

# ══════════════════════════════════════════════════════════════════════
# 9  利润表
# ══════════════════════════════════════════════════════════════════════
ws = newsheet('利润表', '利润表 · 自动，不用填',
              '收入按项目「完工日期」所在月份确认，成本跟着同一个项目走，所以毛利是配得上的。'
              '改左上角的年份就能看别的年度。没填完工日期的项目算「在制」，不进这张表，在下面的补充信息里看。',
              C_AUTO_HDR)
ws.column_dimensions['A'].width = 34
for i in range(2, 15):
    ws.column_dimensions[CL(i)].width = 13
ws.merge_cells('A1:N1'); ws.merge_cells('A2:N2')
put(ws, 'A3', '年份 →', font=Font(name='微软雅黑', size=10, bold=True, color='FFFFFF'),
    fill=PatternFill('solid', fgColor=C_AUTO_HDR), align=CEN)
ws['A3'].border = BOX
hf2 = PatternFill('solid', fgColor=C_AUTO_HDR)
for m in range(1, 13):
    c = ws.cell(3, m + 1, f'{m}月')
    c.font, c.fill, c.alignment, c.border = FT_HDR, hf2, CEN, BOX
c = ws.cell(3, 14, '全年合计')
c.font, c.fill, c.alignment, c.border = FT_HDR, hf2, CEN, BOX
put(ws, 'A4', 2026, fmt='0', font=Font(name='微软雅黑', size=12, bold=True, color='C00000'),
    fill=F_IN, align=CEN)
ws['A4'].border = BOX
for m in range(1, 13):
    c = ws.cell(4, m + 1, f'=$A$4*100+{m}')
    c.number_format, c.font, c.fill, c.alignment, c.border = '0', FT_AUTO, F_AUTO, CEN, BOX
put(ws, 'N4', '—', font=FT_AUTO, fill=F_AUTO, align=CEN)
ws.row_dimensions[4].hidden = True

KT = f'项目核算!$T${S0}:$T${K_END}'
def kcol(col):
    return f'项目核算!${col}${S0}:${col}${K_END}'
def per(col):
    """本列的月初 / 月末，给按日期筛的 SUMIFS 用"""
    d1 = f'DATE($A$4,COLUMN()-1,1)'
    return f'"&{d1},{JB},"<="&EOMONTH({d1},0)'
def ksum(col, src):
    return f'SUMIF({KT},{col}$4,{kcol(src)})'
def jsum(gui, income=False, proj_blank=False):
    src = JJ if income else JK
    extra = f',{JG},""' if proj_blank else ''
    return (f'SUMIFS({src},{JF},"{gui}"{extra},{JB},">="&DATE($A$4,COLUMN()-1,1),'
            f'{JB},"<="&EOMONTH(DATE($A$4,COLUMN()-1,1),0))')

ROWS = [
    ('h', 5, '一、营业收入', None),
    ('f', 6, '　　1. 主营业务收入（完工项目）', lambda col: ksum(col, 'I')),
    ('f', 7, '　　2. 其他业务收入', lambda col: f'{jsum("收入-其他", True)}-{jsum("收入-其他")}'),
    ('s', 8, '　　营业收入合计', 'B6+B7'),
    ('h', 9, '二、营业成本', None),
    ('f', 10, '　　3. 主材', lambda col: ksum(col, 'L')),
    ('f', 11, '　　4. 辅材', lambda col: ksum(col, 'M')),
    ('f', 12, '　　5. 人工（设计·制作·安装）', lambda col: ksum(col, 'N')),
    ('f', 13, '　　6. 运费车费', lambda col: ksum(col, 'O')),
    ('f', 14, '　　7. 其他直接费用', lambda col: ksum(col, 'P')),
    ('f', 15, '　　8. 还没归到项目的直接成本', lambda col: '+'.join(
        f'({jsum(g, False, True)}-{jsum(g, True, True)})'
        for g in ('成本-主材', '成本-辅材', '成本-人工', '成本-车运', '成本-其他'))),
    ('s', 16, '　　营业成本合计', 'SUM(B10:B15)'),
    ('t', 17, '三、毛利', 'B8-B16'),
    ('p', 18, '　　毛利率', 'IF(ROUND(B8,2)=0,"",ROUND(B17/B8,4))'),
    ('h', 19, '四、期间费用', None),
    ('f', 20, '　　管理费用', lambda col: f'{jsum("费用-管理")}-{jsum("费用-管理", True)}'),
    ('f', 21, '　　销售费用', lambda col: f'{jsum("费用-销售")}-{jsum("费用-销售", True)}'),
    ('f', 22, '　　财务费用', lambda col: f'{jsum("费用-财务")}-{jsum("费用-财务", True)}'),
    ('s', 23, '　　期间费用合计', 'SUM(B20:B22)'),
    ('t', 24, '五、营业利润', 'B17-B23'),
    ('f', 25, '　　营业外收入', lambda col: f'{jsum("其他-营业外收入", True)}'),
    ('f', 26, '　　营业外支出', lambda col: f'{jsum("其他-营业外支出")}'),
    ('t', 27, '六、净利润（税前）', 'B24+B25-B26'),
]
STY = {
    'h': (Font(name='微软雅黑', size=10, bold=True, color='1F3864'), PatternFill('solid', fgColor='DEEAF6')),
    'f': (Font(name='微软雅黑', size=10), None),
    's': (Font(name='微软雅黑', size=10, bold=True), PatternFill('solid', fgColor='F2F2F2')),
    't': (Font(name='微软雅黑', size=11, bold=True, color='C00000'), PatternFill('solid', fgColor='FFF2CC')),
    'p': (Font(name='微软雅黑', size=10, bold=True, color='C00000'), None),
}
for kind, r, label, expr in ROWS:
    f, fill = STY[kind]
    put(ws, f'A{r}', label, font=f, align=LEFT)
    ws[f'A{r}'].border = BOX
    if fill:
        ws[f'A{r}'].fill = fill
    ws.row_dimensions[r].height = 19
    for m in range(1, 13):
        c = ws.cell(r, m + 1)
        c.border, c.font, c.alignment = BOX, f, CEN
        c.number_format = PCT if kind == 'p' else MONEY
        if fill:
            c.fill = fill
        if expr is None:
            continue
        col = CL(m + 1)
        if callable(expr):
            c.value = '=ROUND(' + expr(col) + ',2)'
        else:
            c.value = '=' + expr.replace('B', col) if kind != 'p' else '=' + expr.replace('B', col)
    c = ws.cell(r, 14)
    c.border, c.font, c.alignment = BOX, f, CEN
    c.number_format = PCT if kind == 'p' else MONEY
    if fill:
        c.fill = fill
    if expr is not None:
        c.value = ('=IF(ROUND(N8,2)=0,"",ROUND(N17/N8,4))' if kind == 'p'
                   else f'=ROUND(SUM(B{r}:M{r}),2)')
ws.cell(18, 14).value = '=IF(ROUND($N$8,2)=0,"",ROUND($N$17/$N$8,4))'

put(ws, 'A30', '补充信息（不进利润表，但老板要看）',
    font=Font(name='微软雅黑', size=11, bold=True, color='FFFFFF'),
    fill=PatternFill('solid', fgColor=C_AUTO_HDR), align=LEFT)
ws.merge_cells('A30:D30')
EXTRA = [
    (31, '在制项目个数（还没填完工日期）', f'=COUNTIFS(项目核算!$B${S0}:$B${K_END},"?*",项目核算!$T${S0}:$T${K_END},0)'), 
    (32, '在制项目合同额', f'=ROUND(SUMIFS({kcol("I")},项目核算!$B${S0}:$B${K_END},"?*",项目核算!$T${S0}:$T${K_END},0),2)'),
    (33, '在制项目已收款', f'=ROUND(SUMIFS({kcol("J")},项目核算!$B${S0}:$B${K_END},"?*",项目核算!$T${S0}:$T${K_END},0),2)'),
    (34, '在制项目已投入成本', f'=ROUND(SUMIFS({kcol("Q")},项目核算!$B${S0}:$B${K_END},"?*",项目核算!$T${S0}:$T${K_END},0),2)'),
    (35, '期末应收账款（客户欠我们）', '=项目核算!$K$4'),
    (36, '期末应付供应商（我们欠人家）', '=供应商对账!$F$4'),
    (37, '期末资金余额（五个账户）', '=账户余额!$E$14'),
    (38, '借出去还没收回的钱（正数=别人欠我们）', f'=ROUND(SUMIFS({JK},{JF},"往来-借贷")-SUMIFS({JJ},{JF},"往来-借贷"),2)'),
    (39, '日记账还没分类的金额', '=使用说明!$C$25'),
]
for r, lab, f in EXTRA:
    put(ws, f'A{r}', lab, font=Font(name='微软雅黑', size=10), align=LEFT)
    ws[f'A{r}'].border = BOX
    c = put(ws, f'B{r}', f, fmt='0' if r == 31 else MONEY,
            font=Font(name='微软雅黑', size=10, bold=True, color='1F3864'), align=CEN)
    c.border = BOX
    put(ws, f'C{r}', '', align=LEFT)
ws.freeze_panes = 'B5'
ws.sheet_view.showGridLines = False

# ══════════════════════════════════════════════════════════════════════
# 10  供应商对账
# ══════════════════════════════════════════════════════════════════════
SP_END = S0 + N_SUPP - 1
ws = newsheet('供应商对账', '供应商对账 · 自动，不用填（名单来自【基础资料】G 列）',
              '应付余额 = 采购单发生 + 生产安装发生 + 日记账里直接付的成本 − 付给这家的钱合计。'
              '大于 0 就是还欠人家，小于 0 就是多付了或者有预付款。',
              C_AUTO_HDR)
cols = [('供应商', 22, 'auto', None), ('采购单发生额', 14, 'auto', MONEY),
        ('生产安装发生额', 14, 'auto', MONEY), ('日记账直接付的成本', 16, 'auto', MONEY),
        ('付款合计(净额)', 15, 'auto', MONEY), ('应付余额', 14, 'auto', MONEY),
        ('借出/借入净额\n(正=对方欠我们)', 15, 'auto', MONEY),
        ('最后付款日', 13, 'auto', DATEF), ('笔数', 9, 'auto', '0'), ('备注', 22, 'in', None)]
header(ws, cols, C_AUTO_HDR); body(ws, cols, S0 - 1, SP_END)
for i in range(N_SUPP):
    r, br = S0 + i, D0 + i
    g = f'IF($A{r}="","",'
    put(ws, f'A{r}', f'=IF(基础资料!$G{br}="","",基础资料!$G{br})')
    put(ws, f'B{r}', f'={g}ROUND(SUMIF(采购单!$C${D0}:$C${B_END},$A{r},{BUY_AMT}),2))')
    put(ws, f'C{r}', f'={g}ROUND(SUMIF(生产安装!$F${D0}:$F${L_END},$A{r},{LAB_AMT}),2))')
    put(ws, f'D{r}', f'={g}ROUND(SUMIFS({JK},{JH},$A{r},{JF},"成本-*"),2))')
    put(ws, f'E{r}', f'={g}ROUND(SUMIFS({JK},{JH},$A{r},{JF},"<>往来-借贷")'
                     f'-SUMIFS({JJ},{JH},$A{r},{JF},"<>往来-借贷"),2))')
    put(ws, f'F{r}', f'={g}ROUND($B{r}+$C{r}+$D{r}-$E{r},2))')
    put(ws, f'G{r}', f'={g}ROUND(SUMIFS({JK},{JH},$A{r},{JF},"往来-借贷")'
                     f'-SUMIFS({JJ},{JH},$A{r},{JF},"往来-借贷"),2))')
    put(ws, f'H{r}', f'={g}IFERROR(LOOKUP(1,0/(({JH}=$A{r})*({JK}>0)),{JB}),""))')
    put(ws, f'I{r}', f'={g}COUNTIF(采购单!$C${D0}:$C${B_END},$A{r})+COUNTIF({JH},$A{r}))')
put(ws, f'A{S0-1}', '合计', font=FT_SUM, fill=F_SUM, align=CEN)
for c in 'BCDEFG':
    put(ws, f'{c}{S0-1}', f'=ROUND(SUM({c}{S0}:{c}{SP_END}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM, align=CEN)
put(ws, f'I{S0-1}', f'=SUM(I{S0}:I{SP_END})', fmt='0', font=FT_SUM, fill=F_SUM, align=CEN)
ws.conditional_formatting.add(f'F{S0}:F{SP_END}', CellIsRule(
    operator='greaterThan', formula=['0.004'], font=Font(color='C00000', bold=True)))
ws.conditional_formatting.add(f'F{S0}:F{SP_END}', CellIsRule(
    operator='lessThan', formula=['-0.004'], font=Font(color='0070C0', bold=True)))
ws.freeze_panes = f'B{S0}'
ws.auto_filter.ref = f'A{HDR}:J{SP_END}'

# ══════════════════════════════════════════════════════════════════════
# 11  客户应收
# ══════════════════════════════════════════════════════════════════════
CU_END = S0 + N_CUST - 1
ws = newsheet('客户应收', '客户应收 · 自动，不用填（名单来自【基础资料】E 列）',
              '按客户汇总所有项目的合同额和已收款。要看具体是哪个项目欠着，去【项目核算】按「未收款」排个序。',
              C_AUTO_HDR)
cols = [('客户名称', 22, 'auto', None), ('项目数', 9, 'auto', '0'), ('合同总额', 14, 'auto', MONEY),
        ('已收款', 14, 'auto', MONEY), ('未收款', 14, 'auto', MONEY),
        ('未完工项目数', 12, 'auto', '0'), ('最近一次收款', 13, 'auto', DATEF), ('备注', 22, 'in', None)]
header(ws, cols, C_AUTO_HDR); body(ws, cols, S0 - 1, CU_END)
PC, PM, PO, PJ = (f'项目台账!$C${D0}:$C${P_END}', f'项目台账!$M${D0}:$M${P_END}',
                  f'项目台账!$O${D0}:$O${P_END}', f'项目台账!$J${D0}:$J${P_END}')
for i in range(N_CUST):
    r, br = S0 + i, D0 + i
    g = f'IF($A{r}="","",'
    put(ws, f'A{r}', f'=IF(基础资料!$E{br}="","",基础资料!$E{br})')
    put(ws, f'B{r}', f'={g}COUNTIF({PC},$A{r}))')
    put(ws, f'C{r}', f'={g}ROUND(SUMIF({PC},$A{r},{PM}),2))')
    put(ws, f'D{r}', f'={g}ROUND(SUMIF({PC},$A{r},{PO}),2))')
    put(ws, f'E{r}', f'={g}ROUND($C{r}-$D{r},2))')
    put(ws, f'F{r}', f'={g}COUNTIFS({PC},$A{r},{PJ},""))')
    put(ws, f'G{r}', f'={g}IFERROR(LOOKUP(1,0/(({JH}=$A{r})*({JJ}>0)),{JB}),""))')
put(ws, f'A{S0-1}', '合计', font=FT_SUM, fill=F_SUM, align=CEN)
for c in 'CDE':
    put(ws, f'{c}{S0-1}', f'=ROUND(SUM({c}{S0}:{c}{CU_END}),2)', fmt=MONEY, font=FT_SUM, fill=F_SUM, align=CEN)
for c in 'BF':
    put(ws, f'{c}{S0-1}', f'=SUM({c}{S0}:{c}{CU_END})', fmt='0', font=FT_SUM, fill=F_SUM, align=CEN)
ws.conditional_formatting.add(f'E{S0}:E{CU_END}', CellIsRule(
    operator='greaterThan', formula=['0.004'], font=Font(color='C00000', bold=True)))
ws.freeze_panes = f'B{S0}'
ws.auto_filter.ref = f'A{HDR}:H{CU_END}'

# ══════════════════════════════════════════════════════════════════════
# 12  把旧表的数据搬进来
# ══════════════════════════════════════════════════════════════════════
def d2x(s):
    return datetime.datetime.strptime(s, '%Y-%m-%d') if s else None

pj = wb['项目台账']
for i, p in enumerate(DATA['projects']):
    r = D0 + i
    pj.cell(r, 2).value = p['name']
    pj.cell(r, 3).value = p['customer']
    pj.cell(r, 4).value = p['ptype']
    pj.cell(r, 6).value = d2x(p['order_date'])
    pj.cell(r, 10).value = d2x(p['done_date'])
    if p['contract']:
        pj.cell(r, 11).value = p['contract']
    if p['opening_recv']:
        pj.cell(r, 14).value = p['opening_recv']
    pj.cell(r, 18).value = (p['memo'] + ('　来源：' + p['src'] if p['src'] else '')).strip() or None

od = wb['订单明细']
for i, x in enumerate(DATA['orders']):
    r = D0 + i
    for c, v in ((2, d2x(x['date'])), (3, x['project']), (4, x['item']), (5, x['spec']),
                 (6, x['unit']), (7, x['qty'] or None), (8, x['price'] or None), (11, x['note'] or None)):
        od.cell(r, c).value = v

bu = wb['采购单']
for i, x in enumerate(DATA['purchases']):
    r = D0 + i
    for c, v in ((2, d2x(x['date'])), (3, x['supplier'] or None), (4, x['project']), (5, x['mcat']),
                 (7, x['memo']), (8, x['spec'] or None), (9, x['unit'] or None),
                 (10, x['qty'] or None), (11, x['price'] or None), (13, x['note'])):
        bu.cell(r, c).value = v
    if not x['qty'] or not x['price']:          # 原表只有金额没拆数量单价的，按 1 × 金额 摆平
        bu.cell(r, 10).value = 1
        bu.cell(r, 11).value = x['amount']

lb = wb['生产安装']
for i, x in enumerate(DATA['labour']):
    r = D0 + i
    for c, v in ((2, d2x(x['date'])), (3, x['project']), (4, x['wcat']), (6, x['person'] or None),
                 (7, x['unit'] or None), (8, x['qty'] or None), (9, x['price'] or None), (12, x['note'])):
        lb.cell(r, c).value = v
    if not x['qty'] or not x['price']:
        lb.cell(r, 8).value = 1
        lb.cell(r, 9).value = x['amount']

jn = wb['资金日记账']
for i, x in enumerate(DATA['journal']):
    r = D0 + i
    for c, v in ((2, d2x(x['date'])), (3, x['account']), (4, x['direction']), (5, x['cat']),
                 (7, x['project'] or None), (8, x['party'] or None), (9, x['memo']),
                 (10, x['income'] or None), (11, x['outgo'] or None), (15, x['note'] or None)):
        jn.cell(r, c).value = v

# ── 定义名称 ─────────────────────────────────────────────────────────
for nm, ref in NAMES.items():
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)

# ── 打印设置 ─────────────────────────────────────────────────────────
for s in wb.worksheets:
    s.page_setup.orientation = 'landscape'
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 0
    s.sheet_properties.pageSetUpPr.fitToPage = True
    if s.title not in ('使用说明', '利润表', '账户余额'):
        s.print_title_rows = f'{HDR}:{HDR}'

wb.active = 0
wb.save(OUT)
print('已生成:', OUT)
print('工作表:', ' / '.join(s.title for s in wb.worksheets))
n = sum(1 for s in wb.worksheets for row in s.iter_rows()
        for c in row if isinstance(c.value, str) and c.value.startswith('='))
print('公式格子 %d 个' % n)
