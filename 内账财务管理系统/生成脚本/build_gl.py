# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, '/tmp/claude-0/-home-user-temp/9101e7cb-4f82-53a7-bbe8-41b0f93b12b0/scratchpad')
from gl_common import *
from gl_data import ACCOUNTS, CATEGORIES
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule

wb = Workbook()
wb.remove(wb.active)

YEAR_CELL = f"{Q(SH_HOME)}!$B$3"     # 会计年度
MON_CELL  = f"{Q(SH_HOME)}!$B$4"     # 查询月份（日期）
YM_CELL   = f"{Q(SH_HOME)}!$B$5"     # 查询月份文本 yyyy-mm
BOM_CELL  = f"{Q(SH_HOME)}!$B$6"     # 月初
EOM_CELL  = f"{Q(SH_HOME)}!$B$7"     # 月末

def block(ws, col0, ncol, r_hdr, title_txt, names, r0, r1, note=None):
    """横向分区：标题带 + 表头 + 数据区边框"""
    c0, c1 = L(col0), L(col0 + ncol - 1)
    ws.merge_cells(f'{c0}{r_hdr-1}:{c1}{r_hdr-1}')
    put(ws, f'{c0}{r_hdr-1}', title_txt, font=F_HDR2, fill=FILL_HDR2, align=CL)
    m = headers(ws, r_hdr, col0, names)
    if note:
        put(ws, f'{c0}{r1+1}', note, font=F_NOTE, align=CL, border=None)
        ws.merge_cells(f'{c0}{r1+1}:{c1}{r1+1}')
    return m

# ============================================================ 科目参数
ws = wb.create_sheet(SH_PARAM)
title(ws, '科目参数', 'Q', '会计科目表决定报表怎么出数；收支科目表决定每一笔资金流水自动记到哪个科目。两张表都可以按自己行业增删。')
widths(ws, {'A':11,'B':18,'C':8,'D':7,'E':8,'F':17,'G':15,'H':11,'I':13,'J':3,
            'K':6,'L':9,'M':17,'N':11,'O':16,'P':30,'Q':22})
block(ws, 1, 9, 5, '一、会计科目表（报表按这里的映射自动归集，不要删除已用科目）',
      ['科目编码','科目名称','类别','余额\n方向','启用','资产负债表项目','利润表项目','现金流性质','辅助核算'],
      ACC_R0, ACC_R1, '提示：新增科目请接着往下填，并填好【资产负债表项目】或【利润表项目】，否则该科目不会出现在报表上。')
for i, a in enumerate(ACCOUNTS):
    r = ACC_R0 + i
    for j, v in enumerate([a[0], a[1], a[2], a[3], '启用', a[4], a[5], a[6], a[7]]):
        put(ws, f'{L(1+j)}{r}', v, font=F_IN, fill=FILL_IN, align=CL if j in (1,5,6) else C)
for r in range(ACC_R0 + len(ACCOUNTS), ACC_R1 + 1):
    for j in range(9):
        put(ws, f'{L(1+j)}{r}', None, font=F_IN, fill=FILL_IN, align=CL if j in (1,5,6) else C)
for r in range(ACC_R0, ACC_R1 + 1):
    ws.row_dimensions[r].height = 16

block(ws, 11, 7, 5, '二、收支科目表（资金流水的下拉来源，决定自动分录）',
      ['序号','收支\n类型','收支科目','对方科目\n编码','对方科目名称','现金流量表项目','说明'],
      CAT_R0, CAT_R1, '提示：「对方科目」是这笔钱的另一半。收入类＝钱从哪来，支出类＝钱花到哪个科目。')
for i in range(CAT_R1 - CAT_R0 + 1):
    r = CAT_R0 + i
    put(ws, f'K{r}', f'=IF($M{r}="","",ROW()-{CAT_R0-1})', font=F_TXT)
    for col in ['L','M','N','P','Q']:
        put(ws, f'{col}{r}', None, font=F_IN, fill=FILL_IN, align=CL if col in ('P','Q') else C)
    put(ws, f'O{r}', f'=IF($N{r}="","",IFERROR(INDEX($B${ACC_R0}:$B${ACC_R1},'
                     f'MATCH($N{r},$A${ACC_R0}:$A${ACC_R1},0)),"⚠科目不存在"))', font=F_LINK, align=CL)
    if i < len(CATEGORIES):
        c = CATEGORIES[i]
        ws[f'L{r}'], ws[f'M{r}'], ws[f'N{r}'], ws[f'P{r}'] = c[0], c[1], c[2], c[3]
    ws.row_dimensions[r].height = 16
dv_list(ws, f'L{CAT_R0}:L{CAT_R1}', '"收入,支出,转账"')
dv_list(ws, f'E{ACC_R0}:E{ACC_R1}', '"启用,停用"')
dv_list(ws, f'C{ACC_R0}:C{ACC_R1}', '"资产,负债,权益,成本,收入,费用"')
dv_list(ws, f'D{ACC_R0}:D{ACC_R1}', '"借,贷"')
ws.freeze_panes = 'A6'
page(ws, titles='5:5')

# ============================================================ 基础资料
ws = wb.create_sheet(SH_BASE)
title(ws, '基础资料', 'AI',
      '所有下拉都来自这里。要停用某个档案请把【状态】改成停用，不要删除整行——本表六个区共用行号，删行会同时删掉其他区的档案。')
widths(ws, {'A':16,'B':11,'C':11,'D':7,'E':14,'F':3, 'G':14,'H':7,'I':14,'J':3,
            'K':11,'L':13,'M':11,'N':7,'O':3, 'P':11,'Q':18,'R':12,'S':7,'T':11,'U':11,'V':7,'W':14,'X':3,
            'Y':11,'Z':18,'AA':10,'AB':9,'AC':7,'AD':3, 'AE':11,'AF':18,'AG':10,'AH':9,'AI':7})
block(ws, 1, 5, 5, '① 资金账户', ['账户名称','账户类型','资金科目\n编码','状态','备注'], ACT_R0, ACT_R1)
block(ws, 7, 3, 5, '② 部门', ['部门名称','状态','备注'], DEP_R0, DEP_R1)
block(ws, 11, 4, 5, '③ 员工', ['姓名','所属部门','岗位','状态'], EMP_R0, EMP_R1)
block(ws, 16, 8, 5, '④ 商品 / 材料 / 服务',
      ['商品编号','商品名称','规格型号','单位','参考售价','参考成本','是否\n存货','状态'], GDS_R0, GDS_R1)
block(ws, 25, 5, 5, '⑤ 客户', ['客户编号','客户名称','业务员','账期\n(天)','状态'], CUS_R0, CUS_R1)
block(ws, 31, 5, 5, '⑥ 供应商', ['供应商编号','供应商名称','采购员','账期\n(天)','状态'], SUP_R0, SUP_R1)

DEMO_ACT = [('库存现金','现金','1001','正常'),('工商银行','银行','1002','正常'),
            ('支付宝','第三方','1012','正常'),('微信','第三方','1012','正常')]
DEMO_DEP = [('销售部',),('采购部',),('财务部',),('行政部',)]
DEMO_EMP = [('张伟','销售部','销售'),('李娜','采购部','采购'),('王强','财务部','会计'),('赵敏','行政部','行政')]
DEMO_GDS = [('P001','A型配件','A-01','个',200,100,'是'),('P002','B型配件','B-02','个',350,220,'是'),
            ('P003','C型组件','C-10','套',1200,780,'是'),('F001','安装服务','—','次',800,0,'否')]
DEMO_CUS = [('C001','示例客户甲','张伟',30),('C002','示例客户乙','张伟',60),('C003','示例客户丙','张伟',0)]
DEMO_SUP = [('S001','示例供应商甲','李娜',30),('S002','示例供应商乙','李娜',45)]

def fill_block(col0, ncol, r0, r1, demo, text_cols=(), pad=None):
    for i in range(r1 - r0 + 1):
        r = r0 + i
        for j in range(ncol):
            put(ws, f'{L(col0+j)}{r}', None, font=F_IN, fill=FILL_IN,
                align=CL if j in text_cols else C)
        if i < len(demo):
            for j, v in enumerate(demo[i]):
                ws[f'{L(col0+j)}{r}'] = v
            if pad:
                for j, v in pad:
                    ws[f'{L(col0+j)}{r}'] = v
        ws.row_dimensions[r].height = 16

fill_block(1, 5, ACT_R0, ACT_R1, DEMO_ACT, text_cols=(0, 4))
fill_block(7, 3, DEP_R0, DEP_R1, DEMO_DEP, text_cols=(0, 2), pad=[(1, '正常')])
fill_block(11, 4, EMP_R0, EMP_R1, DEMO_EMP, text_cols=(0,), pad=[(3, '正常')])
fill_block(16, 8, GDS_R0, GDS_R1, DEMO_GDS, text_cols=(1,), pad=[(7, '正常')])
fill_block(25, 5, CUS_R0, CUS_R1, DEMO_CUS, text_cols=(1,), pad=[(4, '正常')])
fill_block(31, 5, SUP_R0, SUP_R1, DEMO_SUP, text_cols=(1,), pad=[(4, '正常')])

for rng in [f'D{ACT_R0}:D{ACT_R1}', f'H{DEP_R0}:H{DEP_R1}', f'N{EMP_R0}:N{EMP_R1}',
            f'W{GDS_R0}:W{GDS_R1}', f'AC{CUS_R0}:AC{CUS_R1}', f'AI{SUP_R0}:AI{SUP_R1}']:
    dv_list(ws, rng, '"正常,停用"')
dv_list(ws, f'B{ACT_R0}:B{ACT_R1}', '"现金,银行,第三方,其他"')
dv_list(ws, f'V{GDS_R0}:V{GDS_R1}', '"是,否"')
dv_list(ws, f'C{ACT_R0}:C{ACT_R1}', f"={Q(SH_PARAM)}!$A${ACC_R0}:$A${ACC_R1}")
dv_list(ws, f'L{EMP_R0}:L{EMP_R1}', f"={Q(SH_BASE)}!$G${DEP_R0}:$G${DEP_R1}")
dv_list(ws, f'AA{CUS_R0}:AA{CUS_R1}', f"={Q(SH_BASE)}!$K${EMP_R0}:$K${EMP_R1}")
dv_list(ws, f'AG{SUP_R0}:AG{SUP_R1}', f"={Q(SH_BASE)}!$K${EMP_R0}:$K${EMP_R1}")
ws.freeze_panes = 'A6'
page(ws, titles='5:5')
print('  ✓ 科目参数 / 基础资料')

# ============================================================ 期初数据
ws = wb.create_sheet(SH_OPEN)
title(ws, '期初数据（启用时录一次）', 'AB',
      '只录启用日的资产、负债、所有者权益余额。本年已发生的收入和费用请到【其他凭证】按月补录，不要录在这里，否则资产负债表会不平。')
widths(ws, {'A':16,'B':14,'C':14,'D':3, 'E':16,'F':12,'G':13,'H':13,'I':3,
            'J':16,'K':12,'L':13,'M':13,'N':3, 'O':14,'P':12,'Q':7,'R':11,'S':12,'T':13,'U':13,'V':3,
            'W':11,'X':18,'Y':7,'Z':14,'AA':14,'AB':14})
put(ws, 'A3', '平衡校验：', font=F_H2, align=CR, border=None)
CHK = [('明细合计·资金', f'=SUM($B${OC_R0}:$B${OC_R1})', 'B'),
       ('明细合计·应收', f'=SUM($G${OR_R0}:$G${OR_R1})', 'E'),
       ('明细合计·应付', f'=SUM($L${OP_R0}:$L${OP_R1})', 'J'),
       ('明细合计·存货', f'=SUM($T${OG_R0}:$T${OG_R1})', 'O'),
       ('科目借方合计', f'=SUM($Z${OA_R0}:$Z${OA_R1})', 'W')]
for i, (lab, f, c) in enumerate(CHK):
    put(ws, f'{c}3', lab, font=F_NOTE, align=CR, border=None)
    put(ws, f'{L(ord(c)-64+1)}3', f, font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, 'AA3', f'=SUM($AA${OA_R0}:$AA${OA_R1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, 'Z3', '科目贷方合计', font=F_NOTE, align=CR, border=None)
put(ws, 'AB3', f'=IF(ABS(SUM($Z${OA_R0}:$Z${OA_R1})-SUM($AA${OA_R0}:$AA${OA_R1}))<0.01,"✓ 借贷平衡","✗ 借贷不平，差额 "'
                f'&TEXT(SUM($Z${OA_R0}:$Z${OA_R1})-SUM($AA${OA_R0}:$AA${OA_R1}),"#,##0.00"))',
    font=F_TOT, fill=FILL_CHK, align=CL)

block(ws, 1, 3, 5, '① 资金期初', ['账户','期初余额','备注'], OC_R0, OC_R1)
block(ws, 5, 4, 5, '② 应收期初', ['客户','原业务日期','期初应收','备注'], OR_R0, OR_R1)
block(ws, 10, 4, 5, '③ 应付期初', ['供应商','原业务日期','期初应付','备注'], OP_R0, OP_R1)
block(ws, 15, 7, 5, '④ 存货期初',
      ['商品','规格','单位','期初数量','期初单位成本','期初金额','备注'], OG_R0, OG_R1)
block(ws, 23, 6, 5, '⑤ 会计科目期初（总账口径，必须借贷相等）',
      ['科目编码','科目名称','方向','期初借方','期初贷方','备注'], OA_R0, OA_R1,
      '损益类科目（6xxx / 5xxx）不要填期初。')

def in_cells(cols, r0, r1, fmts=None, texts=()):
    for r in range(r0, r1 + 1):
        for c in cols:
            put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN,
                align=CL if c in texts else C, fmt=(fmts or {}).get(c))
        ws.row_dimensions[r].height = 16

in_cells(['A','B','C'], OC_R0, OC_R1, {'B': MONEY}, texts=('C',))
in_cells(['E','F','G','H'], OR_R0, OR_R1, {'F': DATE, 'G': MONEY}, texts=('H',))
in_cells(['J','K','L','M'], OP_R0, OP_R1, {'K': DATE, 'L': MONEY}, texts=('M',))
in_cells(['O','R','S','U'], OG_R0, OG_R1, {'R': QTY, 'S': MONEY}, texts=('U',))
in_cells(['W','Z','AA','AB'], OA_R0, OA_R1, {'Z': MONEY, 'AA': MONEY}, texts=('AB',))
BASE, PARAM = Q(SH_BASE), Q(SH_PARAM)
for r in range(OG_R0, OG_R1 + 1):
    put(ws, f'P{r}', f'=IF($O{r}="","",IFERROR(INDEX({BASE}!$R${GDS_R0}:$R${GDS_R1},'
                     f'MATCH($O{r},{BASE}!$Q${GDS_R0}:$Q${GDS_R1},0)),""))', font=F_LINK)
    put(ws, f'Q{r}', f'=IF($O{r}="","",IFERROR(INDEX({BASE}!$S${GDS_R0}:$S${GDS_R1},'
                     f'MATCH($O{r},{BASE}!$Q${GDS_R0}:$Q${GDS_R1},0)),""))', font=F_LINK)
    put(ws, f'T{r}', f'=IF($O{r}="","",ROUND(N($R{r})*N($S{r}),2))', font=F_TXT, fmt=MONEY)
for r in range(OA_R0, OA_R1 + 1):
    put(ws, f'X{r}', f'=IF($W{r}="","",IFERROR(INDEX({PARAM}!$B${ACC_R0}:$B${ACC_R1},'
                     f'MATCH($W{r},{PARAM}!$A${ACC_R0}:$A${ACC_R1},0)),"⚠科目不存在"))', font=F_LINK, align=CL)
    put(ws, f'Y{r}', f'=IF($W{r}="","",IFERROR(INDEX({PARAM}!$D${ACC_R0}:$D${ACC_R1},'
                     f'MATCH($W{r},{PARAM}!$A${ACC_R0}:$A${ACC_R1},0)),""))', font=F_LINK)
dv_list(ws, f'A{OC_R0}:A{OC_R1}', f"={BASE}!$A${ACT_R0}:$A${ACT_R1}")
dv_list(ws, f'E{OR_R0}:E{OR_R1}', f"={BASE}!$Z${CUS_R0}:$Z${CUS_R1}")
dv_list(ws, f'J{OP_R0}:J{OP_R1}', f"={BASE}!$AF${SUP_R0}:$AF${SUP_R1}")
dv_list(ws, f'O{OG_R0}:O{OG_R1}', f"={BASE}!$Q${GDS_R0}:$Q${GDS_R1}")
dv_list(ws, f'W{OA_R0}:W{OA_R1}', f"={PARAM}!$A${ACC_R0}:$A${ACC_R1}")
for rng, fmt in [(f'B{OC_R0}:B{OC_R1}', None), (f'G{OR_R0}:G{OR_R1}', None), (f'L{OP_R0}:L{OP_R1}', None),
                 (f'Z{OA_R0}:Z{OA_R1}', None), (f'AA{OA_R0}:AA{OA_R1}', None)]:
    dv_num(ws, rng, 'greaterThanOrEqual', '-999999999')
ws.freeze_panes = 'A6'
page(ws, titles='5:5')
print('  ✓ 期初数据')

# ============================================================ 录入表公共
PAR_CAT_T = f"{PARAM}!$L${CAT_R0}:$L${CAT_R1}"      # 收支类型
PAR_CAT_N = f"{PARAM}!$M${CAT_R0}:$M${CAT_R1}"      # 收支科目
PAR_CAT_A = f"{PARAM}!$N${CAT_R0}:$N${CAT_R1}"      # 对方科目
PAR_CAT_C = f"{PARAM}!$P${CAT_R0}:$P${CAT_R1}"      # 现金流项目
PAR_ACC_C = f"{PARAM}!$A${ACC_R0}:$A${ACC_R1}"      # 科目编码
PAR_ACC_N = f"{PARAM}!$B${ACC_R0}:$B${ACC_R1}"      # 科目名称
B_ACT_N   = f"{BASE}!$A${ACT_R0}:$A${ACT_R1}"
B_ACT_S   = f"{BASE}!$C${ACT_R0}:$C${ACT_R1}"
B_GDS_C   = f"{BASE}!$P${GDS_R0}:$P${GDS_R1}"
B_GDS_N   = f"{BASE}!$Q${GDS_R0}:$Q${GDS_R1}"
B_CUS_N   = f"{BASE}!$Z${CUS_R0}:$Z${CUS_R1}"
B_SUP_N   = f"{BASE}!$AF${SUP_R0}:$AF${SUP_R1}"
B_DEP_N   = f"{BASE}!$G${DEP_R0}:$G${DEP_R1}"
B_EMP_N   = f"{BASE}!$K${EMP_R0}:$K${EMP_R1}"

def sysband(ws, c0, c1, row, text='系统自动计算区 · 请勿修改'):
    ws.merge_cells(f'{c0}{row}:{c1}{row}')
    put(ws, f'{c0}{row}', text, font=F_HDR2, fill=FILL_AUTO, align=C)

def inband(ws, c0, c1, row, text='录入区 · 黄色格子手工填写'):
    ws.merge_cells(f'{c0}{row}:{c1}{row}')
    put(ws, f'{c0}{row}', text, font=F_HDR2, fill=FILL_HDR2, align=C)

def lookup(val, src_key, src_val, notfound='""'):
    return f'IFERROR(INDEX({src_val},MATCH({val},{src_key},0)),{notfound})'

# ============================================================ 资金流水
ws = wb.create_sheet(SH_CASH)
R0, R1 = CASH_R0, CASH_R1
title(ws, '资金流水', 'U', '所有实际收到、付出的钱只在这里登记一次。每一行自动生成一借一贷，借贷天然相等。')
widths(ws, {'A':15,'B':11,'C':9,'D':8,'E':15,'F':13,'G':16,'H':10,'I':9,'J':24,'K':13,'L':13,'M':12,
            'N':13,'O':11,'P':11,'Q':24,'R':14,'S':11,'T':11,'U':11})
inband(ws, 'A', 'M', 3); sysband(ws, 'N', 'U', 3)
headers(ws, 4, 1, ['单据编号','日期','月份','收支\n类型','收支科目','账户','往来单位 /\n转入账户','部门','员工','摘要',
                   '收入金额','支出金额','校验'])
headers(ws, 4, 14, ['发生金额','借方科目','贷方科目','现金流量表项目','该账户余额','对方科目','本账户\n资金科目','转入账户\n资金科目'],
        fill=FILL_AUTO, font=F_HDR2)
for r in range(R0, R1 + 1):
    for c, fmt, tx in [('B', DATE, 0), ('E', None, 0), ('F', None, 0), ('G', None, 1), ('H', None, 0),
                       ('I', None, 0), ('J', None, 1), ('K', MONEY, 0), ('L', MONEY, 0)]:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if tx else C, fmt=fmt)
    put(ws, f'A{r}', f'=IF($B{r}="","","ZJ"&TEXT($B{r},"yyyymm")&"-"&TEXT(ROW()-{R0-1},"000"))', font=F_LINK)
    put(ws, f'C{r}', f'=IF($B{r}="","",TEXT($B{r},"yyyy-mm"))', font=F_LINK)
    put(ws, f'D{r}', f'=IF($E{r}="","",{lookup(f"$E{r}", PAR_CAT_N, PAR_CAT_T, chr(34)+"⚠"+chr(34))})', font=F_LINK)
    put(ws, f'S{r}', f'=IF($E{r}="","",{lookup(f"$E{r}", PAR_CAT_N, PAR_CAT_A)})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'T{r}', f'=IF($F{r}="","",{lookup(f"$F{r}", B_ACT_N, B_ACT_S)})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'U{r}', f'=IF($G{r}="","",{lookup(f"$G{r}", B_ACT_N, B_ACT_S)})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'N{r}', f'=IF($B{r}="","",IF($D{r}="收入",N($K{r}),N($L{r})))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'O{r}', f'=IF($B{r}="","",IF($D{r}="收入",$T{r},IF($D{r}="转账",$U{r},$S{r})))', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'P{r}', f'=IF($B{r}="","",IF($D{r}="收入",$S{r},$T{r}))', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'Q{r}', f'=IF($E{r}="","",IF($D{r}="转账","—",{lookup(f"$E{r}", PAR_CAT_N, PAR_CAT_C)}))',
        font=F_LINK, fill=FILL_AUTO, align=CL)
    put(ws, f'R{r}', f'=IF($B{r}="","",SUMIFS({Q(SH_OPEN)}!$B${OC_R0}:$B${OC_R1},{Q(SH_OPEN)}!$A${OC_R0}:$A${OC_R1},$F{r})'
                     f'+SUMIFS($N${R0}:$N{r},$F${R0}:$F{r},$F{r},$D${R0}:$D{r},"收入")'
                     f'-SUMIFS($N${R0}:$N{r},$F${R0}:$F{r},$F{r},$D${R0}:$D{r},"支出")'
                     f'-SUMIFS($N${R0}:$N{r},$F${R0}:$F{r},$F{r},$D${R0}:$D{r},"转账")'
                     f'+SUMIFS($N${R0}:$N{r},$G${R0}:$G{r},$F{r},$D${R0}:$D{r},"转账"))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'M{r}',
        f'=IF($B{r}="","",'
        f'IF(NOT(ISNUMBER($B{r})),"日期不是日期格式",'
        f'IF(YEAR($B{r})<>{YEAR_CELL},"不在会计年度",'
        f'IF($E{r}="","未选收支科目",'
        f'IF($D{r}="⚠","收支科目不存在",'
        f'IF($F{r}="","未选账户",'
        f'IF($T{r}="","账户未设资金科目",'
        f'IF(AND($D{r}="转账",$U{r}=""),"转入账户无效",'
        f'IF(AND(N($K{r})>0,N($L{r})>0),"收支不能同时填",'
        f'IF(AND($D{r}="收入",N($K{r})<=0),"收入金额须大于0",'
        f'IF(AND($D{r}<>"收入",N($L{r})<=0),"支出金额须大于0",'
        f'IF($O{r}="","借方科目缺失",IF($P{r}="","贷方科目缺失","√"))))))))))))))', font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
dv_list(ws, f'E{R0}:E{R1}', f'={PAR_CAT_N}')
dv_list(ws, f'F{R0}:F{R1}', f'={B_ACT_N}')
dv_list(ws, f'H{R0}:H{R1}', f'={B_DEP_N}')
dv_list(ws, f'I{R0}:I{R1}', f'={B_EMP_N}')
dv_num(ws, f'K{R0}:K{R1}'); dv_num(ws, f'L{R0}:L{R1}'); dv_date(ws, f'B{R0}:B{R1}', YEAR_CELL)
ws.conditional_formatting.add(f'M{R0}:M{R1}',
    FormulaRule(formula=[f'AND($M{R0}<>"",$M{R0}<>"√")'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.freeze_panes = 'D5'
page(ws, titles='4:4')
print('  ✓ 资金流水')

# ============================================================ 销售录入
ws = wb.create_sheet(SH_SAL)
R0, R1 = SAL_R0, SAL_R1
title(ws, '销售录入', 'AD', '所有对外销售（含赊销、现销）都在这里登记一次；客户什么时候把钱打过来，到【资金流水】登记「客户回款」。')
widths(ws, {'A':15,'B':11,'C':9,'D':16,'E':16,'F':10,'G':12,'H':8,'I':10,'J':9,'K':20,'L':14,
            'M':11,'N':7,'O':13,'P':12,'Q':13,'R':12,'S':13,'T':12,'U':11,
            'V':8,'W':8,'X':12,'Y':8,'Z':8,'AA':12,'AB':8,'AC':8,'AD':12})
inband(ws, 'A', 'L', 3); sysband(ws, 'M', 'AD', 3)
headers(ws, 4, 1, ['单据编号','日期','月份','客户','商品 / 服务','数量','单价\n(不含税)','税率','部门','业务员','摘要','校验'])
headers(ws, 4, 13, ['规格','单位','不含税金额','销项税额','价税合计','单位成本\n(移动加权)','销售成本','毛利','到期日',
                    '借1','贷1','金额1','借2','贷2','金额2','借3','贷3','金额3'], fill=FILL_AUTO, font=F_HDR2)
OPN, BUYQ = Q(SH_OPEN), Q(SH_BUY)
for r in range(R0, R1 + 1):
    for c, fmt, tx in [('B', DATE, 0), ('D', None, 1), ('E', None, 1), ('F', QTY, 0), ('G', MONEY, 0),
                       ('H', PCT, 0), ('I', None, 0), ('J', None, 0), ('K', None, 1)]:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if tx else C, fmt=fmt)
    put(ws, f'A{r}', f'=IF($B{r}="","","XS"&TEXT($B{r},"yyyymm")&"-"&TEXT(ROW()-{R0-1},"000"))', font=F_LINK)
    put(ws, f'C{r}', f'=IF($B{r}="","",TEXT($B{r},"yyyy-mm"))', font=F_LINK)
    put(ws, f'M{r}', f'=IF($E{r}="","",{lookup(f"$E{r}", B_GDS_N, f"{BASE}!$R${GDS_R0}:$R${GDS_R1}")})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'N{r}', f'=IF($E{r}="","",{lookup(f"$E{r}", B_GDS_N, f"{BASE}!$S${GDS_R0}:$S${GDS_R1}")})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'O{r}', f'=IF(OR($F{r}="",$G{r}=""),"",ROUND(N($F{r})*N($G{r}),2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'P{r}', f'=IF($O{r}="","",ROUND(N($O{r})*N($H{r}),2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'Q{r}', f'=IF($O{r}="","",N($O{r})+N($P{r}))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'R{r}', f'=IF(OR($E{r}="",$B{r}=""),"",IF({lookup(f"$E{r}", B_GDS_N, f"{BASE}!$V${GDS_R0}:$V${GDS_R1}")}<>"是",0,'
                     f'IFERROR((SUMIFS({OPN}!$T${OG_R0}:$T${OG_R1},{OPN}!$O${OG_R0}:$O${OG_R1},$E{r})'
                     f'+SUMIFS({BUYQ}!$Q${BUY_R0}:$Q${BUY_R1},{BUYQ}!$E${BUY_R0}:$E${BUY_R1},$E{r},'
                     f'{BUYQ}!$I${BUY_R0}:$I${BUY_R1},"是",{BUYQ}!$B${BUY_R0}:$B${BUY_R1},"<="&$B{r}))'
                     f'/(SUMIFS({OPN}!$R${OG_R0}:$R${OG_R1},{OPN}!$O${OG_R0}:$O${OG_R1},$E{r})'
                     f'+SUMIFS({BUYQ}!$F${BUY_R0}:$F${BUY_R1},{BUYQ}!$E${BUY_R0}:$E${BUY_R1},$E{r},'
                     f'{BUYQ}!$I${BUY_R0}:$I${BUY_R1},"是",{BUYQ}!$B${BUY_R0}:$B${BUY_R1},"<="&$B{r})),'
                     f'{lookup(f"$E{r}", B_GDS_N, f"{BASE}!$U${GDS_R0}:$U${GDS_R1}", "0")})))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'S{r}', f'=IF($O{r}="","",ROUND(N($F{r})*N($R{r}),2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'T{r}', f'=IF($O{r}="","",N($O{r})-N($S{r}))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'U{r}', f'=IF(OR($B{r}="",$D{r}=""),"",$B{r}+N({lookup(f"$D{r}", B_CUS_N, f"{BASE}!$AB${CUS_R0}:$AB${CUS_R1}", "0")}))',
        font=F_LINK, fill=FILL_AUTO, fmt=DATE)
    put(ws, f'V{r}', f'=IF($O{r}="","","1122")', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'W{r}', f'=IF($O{r}="","","6001")', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'X{r}', f'=IF($O{r}="","",$O{r})', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'Y{r}', f'=IF(N($P{r})=0,"","1122")', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'Z{r}', f'=IF(N($P{r})=0,"","2221")', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'AA{r}', f'=IF(N($P{r})=0,"",$P{r})', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'AB{r}', f'=IF(N($S{r})=0,"","6401")', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'AC{r}', f'=IF(N($S{r})=0,"","1405")', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'AD{r}', f'=IF(N($S{r})=0,"",$S{r})', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'L{r}',
        f'=IF($B{r}="","",'
        f'IF(NOT(ISNUMBER($B{r})),"日期不是日期格式",'
        f'IF(YEAR($B{r})<>{YEAR_CELL},"不在会计年度",'
        f'IF($D{r}="","未选客户",'
        f'IF(ISNA(MATCH($D{r},{B_CUS_N},0)),"客户不在档案",'
        f'IF($E{r}="","未选商品",'
        f'IF(ISNA(MATCH($E{r},{B_GDS_N},0)),"商品不在档案",'
        f'IF(NOT(ISNUMBER($F{r})),"数量须为数字",'
        f'IF(NOT(ISNUMBER($G{r})),"单价须为数字",'
        f'IF(N($F{r})<=0,"数量须大于0",'
        f'IF(AND({lookup(f"$E{r}", B_GDS_N, f"{BASE}!$V${GDS_R0}:$V${GDS_R1}")}="是",N($R{r})=0),"无成本价·请先录采购或填参考成本",'
        f'IF(N($T{r})<0,"毛利为负·请核对","√")))))))))))', font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
dv_list(ws, f'D{R0}:D{R1}', f'={B_CUS_N}')
dv_list(ws, f'E{R0}:E{R1}', f'={B_GDS_N}')
dv_list(ws, f'I{R0}:I{R1}', f'={B_DEP_N}')
dv_list(ws, f'J{R0}:J{R1}', f'={B_EMP_N}')
dv_num(ws, f'F{R0}:F{R1}'); dv_num(ws, f'G{R0}:G{R1}'); dv_date(ws, f'B{R0}:B{R1}', YEAR_CELL)
dv_list(ws, f'H{R0}:H{R1}', '"0,0.01,0.03,0.05,0.06,0.09,0.13"')
ws.conditional_formatting.add(f'L{R0}:L{R1}',
    FormulaRule(formula=[f'AND($L{R0}<>"",$L{R0}<>"√")'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.freeze_panes = 'D5'
page(ws, titles='4:4')
print('  ✓ 销售录入')

# ============================================================ 采购录入
ws = wb.create_sheet(SH_BUY)
R0, R1 = BUY_R0, BUY_R1
title(ws, '采购录入', 'Z', '所有采购、进货、接受劳务在这里登记一次；实际付款到【资金流水】登记「供应商付款」。不入库的直接费用请把「是否入库」选否并指定费用科目。')
widths(ws, {'A':15,'B':11,'C':9,'D':16,'E':16,'F':10,'G':12,'H':8,'I':8,'J':13,'K':10,'L':9,'M':20,'N':14,
            'O':11,'P':7,'Q':13,'R':12,'S':13,'T':11,'U':8,'V':8,'W':12,'X':8,'Y':8,'Z':12})
inband(ws, 'A', 'N', 3); sysband(ws, 'O', 'Z', 3)
headers(ws, 4, 1, ['单据编号','日期','月份','供应商','商品 / 内容','数量','单价\n(不含税)','税率','是否\n入库',
                   '不入库时\n费用科目','部门','采购员','摘要','校验'])
headers(ws, 4, 15, ['规格','单位','不含税金额','进项税额','价税合计','到期日',
                    '借1','贷1','金额1','借2','贷2','金额2'], fill=FILL_AUTO, font=F_HDR2)
for r in range(R0, R1 + 1):
    for c, fmt, tx in [('B', DATE, 0), ('D', None, 1), ('E', None, 1), ('F', QTY, 0), ('G', MONEY, 0),
                       ('H', PCT, 0), ('I', None, 0), ('J', None, 0), ('K', None, 0), ('L', None, 0), ('M', None, 1)]:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if tx else C, fmt=fmt)
    put(ws, f'A{r}', f'=IF($B{r}="","","CG"&TEXT($B{r},"yyyymm")&"-"&TEXT(ROW()-{R0-1},"000"))', font=F_LINK)
    put(ws, f'C{r}', f'=IF($B{r}="","",TEXT($B{r},"yyyy-mm"))', font=F_LINK)
    put(ws, f'O{r}', f'=IF($E{r}="","",{lookup(f"$E{r}", B_GDS_N, f"{BASE}!$R${GDS_R0}:$R${GDS_R1}")})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'P{r}', f'=IF($E{r}="","",{lookup(f"$E{r}", B_GDS_N, f"{BASE}!$S${GDS_R0}:$S${GDS_R1}")})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'Q{r}', f'=IF(OR($F{r}="",$G{r}=""),"",ROUND(N($F{r})*N($G{r}),2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'R{r}', f'=IF($Q{r}="","",ROUND(N($Q{r})*N($H{r}),2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'S{r}', f'=IF($Q{r}="","",N($Q{r})+N($R{r}))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'T{r}', f'=IF(OR($B{r}="",$D{r}=""),"",$B{r}+N({lookup(f"$D{r}", B_SUP_N, f"{BASE}!$AH${SUP_R0}:$AH${SUP_R1}", "0")}))',
        font=F_LINK, fill=FILL_AUTO, fmt=DATE)
    put(ws, f'U{r}', f'=IF($Q{r}="","",IF($I{r}="是","1405",IF($J{r}="","6602",$J{r})))', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'V{r}', f'=IF($Q{r}="","","2202")', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'W{r}', f'=IF($Q{r}="","",$Q{r})', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'X{r}', f'=IF(N($R{r})=0,"","2221")', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'Y{r}', f'=IF(N($R{r})=0,"","2202")', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'Z{r}', f'=IF(N($R{r})=0,"",$R{r})', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'N{r}',
        f'=IF($B{r}="","",'
        f'IF(NOT(ISNUMBER($B{r})),"日期不是日期格式",'
        f'IF(YEAR($B{r})<>{YEAR_CELL},"不在会计年度",'
        f'IF($D{r}="","未选供应商",'
        f'IF(ISNA(MATCH($D{r},{B_SUP_N},0)),"供应商不在档案",'
        f'IF($E{r}="","未填商品或内容",'
        f'IF(NOT(ISNUMBER($F{r})),"数量须为数字",'
        f'IF(NOT(ISNUMBER($G{r})),"单价须为数字",'
        f'IF(N($F{r})<=0,"数量须大于0",'
        f'IF($I{r}="","请选是否入库",'
        f'IF(AND($I{r}="是",ISNA(MATCH($E{r},{B_GDS_N},0))),"入库商品必须在商品档案中",'
        f'IF(AND($I{r}="否",$J{r}<>"",ISNA(MATCH($J{r},{PAR_ACC_C},0))),"费用科目不存在","√"))))))))))))', font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
dv_list(ws, f'D{R0}:D{R1}', f'={B_SUP_N}')
dv_list(ws, f'E{R0}:E{R1}', f'={B_GDS_N}')
dv_list(ws, f'I{R0}:I{R1}', '"是,否"')
dv_list(ws, f'J{R0}:J{R1}', f'={PAR_ACC_C}')
dv_list(ws, f'K{R0}:K{R1}', f'={B_DEP_N}')
dv_list(ws, f'L{R0}:L{R1}', f'={B_EMP_N}')
dv_num(ws, f'F{R0}:F{R1}'); dv_num(ws, f'G{R0}:G{R1}'); dv_date(ws, f'B{R0}:B{R1}', YEAR_CELL)
dv_list(ws, f'H{R0}:H{R1}', '"0,0.01,0.03,0.05,0.06,0.09,0.13"')
ws.conditional_formatting.add(f'N{R0}:N{R1}',
    FormulaRule(formula=[f'AND($N{R0}<>"",$N{R0}<>"√")'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.freeze_panes = 'D5'
page(ws, titles='4:4')

# ============================================================ 其他凭证
ws = wb.create_sheet(SH_OTH)
R0, R1 = OTH_R0, OTH_R1
title(ws, '其他凭证', 'N',
      '不涉及收付款的账都在这里做：工资社保计提、折旧摊销、税金计提、坏账、年末结转损益、以及任何需要手工调整的分录。一行一借一贷；多借多贷请拆成多行、凭证号填同一个。')
widths(ws, {'A':15,'B':11,'C':9,'D':14,'E':26,'F':11,'G':11,'H':14,'I':10,'J':9,'K':16,'L':16,'M':16,'N':16})
inband(ws, 'A', 'L', 3); sysband(ws, 'M', 'N', 3)
headers(ws, 4, 1, ['凭证号','日期','月份','业务类型','摘要','借方科目\n编码','贷方科目\n编码','金额','部门','员工','备注','校验'])
headers(ws, 4, 13, ['借方科目名称','贷方科目名称'], fill=FILL_AUTO, font=F_HDR2)
for r in range(R0, R1 + 1):
    for c, fmt, tx in [('B', DATE, 0), ('D', None, 0), ('E', None, 1), ('F', None, 0), ('G', None, 0),
                       ('H', MONEY, 0), ('I', None, 0), ('J', None, 0), ('K', None, 1)]:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if tx else C, fmt=fmt)
    put(ws, f'A{r}', f'=IF($B{r}="","","QT"&TEXT($B{r},"yyyymm")&"-"&TEXT(ROW()-{R0-1},"000"))', font=F_LINK)
    put(ws, f'C{r}', f'=IF($B{r}="","",TEXT($B{r},"yyyy-mm"))', font=F_LINK)
    put(ws, f'M{r}', f'=IF($F{r}="","",{lookup(f"$F{r}", PAR_ACC_C, PAR_ACC_N, chr(34)+"⚠科目不存在"+chr(34))})',
        font=F_LINK, fill=FILL_AUTO, align=CL)
    put(ws, f'N{r}', f'=IF($G{r}="","",{lookup(f"$G{r}", PAR_ACC_C, PAR_ACC_N, chr(34)+"⚠科目不存在"+chr(34))})',
        font=F_LINK, fill=FILL_AUTO, align=CL)
    put(ws, f'L{r}',
        f'=IF($B{r}="","",'
        f'IF(NOT(ISNUMBER($B{r})),"日期不是日期格式",'
        f'IF(YEAR($B{r})<>{YEAR_CELL},"不在会计年度",'
        f'IF($F{r}="","未填借方科目",'
        f'IF($G{r}="","未填贷方科目",'
        f'IF(LEFT($M{r},1)="⚠","借方科目不存在",'
        f'IF(LEFT($N{r},1)="⚠","贷方科目不存在",'
        f'IF($F{r}=$G{r},"借贷科目不能相同",'
        f'IF(NOT(ISNUMBER($H{r})),"金额须为数字",'
        f'IF(N($H{r})<=0,"金额须大于0","√"))))))))))', font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
dv_list(ws, f'D{R0}:D{R1}', '"工资计提,社保公积金计提,折旧计提,无形资产摊销,长期待摊费用摊销,税金计提,坏账计提,存货跌价,结转损益,年初建账,往来重分类,其他调整"')
dv_list(ws, f'F{R0}:F{R1}', f'={PAR_ACC_C}')
dv_list(ws, f'G{R0}:G{R1}', f'={PAR_ACC_C}')
dv_list(ws, f'I{R0}:I{R1}', f'={B_DEP_N}')
dv_list(ws, f'J{R0}:J{R1}', f'={B_EMP_N}')
dv_num(ws, f'H{R0}:H{R1}', 'greaterThan', '0'); dv_date(ws, f'B{R0}:B{R1}', YEAR_CELL)
ws.conditional_formatting.add(f'L{R0}:L{R1}',
    FormulaRule(formula=[f'AND($L{R0}<>"",$L{R0}<>"√")'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.freeze_panes = 'D5'
page(ws, titles='4:4')
print('  ✓ 采购录入 / 其他凭证')

# ============================================================ 总账引擎
CASHQ, SALQ, OTHQ = Q(SH_CASH), Q(SH_SAL), Q(SH_OTH)
# (表, 金额列, 借方科目列, 贷方科目列, 月份列, 日期列, 起行, 止行)
SLOTS = [
    (CASHQ, 'N', 'O', 'P', 'C', 'B', CASH_R0, CASH_R1),
    (SALQ,  'X', 'V', 'W', 'C', 'B', SAL_R0,  SAL_R1),
    (SALQ,  'AA','Y', 'Z', 'C', 'B', SAL_R0,  SAL_R1),
    (SALQ,  'AD','AB','AC','C', 'B', SAL_R0,  SAL_R1),
    (BUYQ,  'W', 'U', 'V', 'C', 'B', BUY_R0,  BUY_R1),
    (BUYQ,  'Z', 'X', 'Y', 'C', 'B', BUY_R0,  BUY_R1),
    (OTHQ,  'H', 'F', 'G', 'C', 'B', OTH_R0,  OTH_R1),
]

def occ(side, acct_ref, period):
    """side: 'D' 借 / 'C' 贷；period: 'month' 本月 / 'year' 本年累计"""
    terms = []
    for sh, amt, dr, cr, mc, dc, r0, r1 in SLOTS:
        key = dr if side == 'D' else cr
        cond = (f'{sh}!${mc}${r0}:${mc}${r1},{YM_CELL}' if period == 'month'
                else f'{sh}!${dc}${r0}:${dc}${r1},"<="&{EOM_CELL}')
        terms.append(f'SUMIFS({sh}!${amt}${r0}:${amt}${r1},{sh}!${key}${r0}:${key}${r1},{acct_ref},{cond})')
    return '+'.join(terms)

ws = wb.create_sheet(SH_TB)
R0, R1 = ACC_R0, ACC_R1
title(ws, '科目余额表', 'P',
      '全部数据由四张录入表自动汇总，没有任何写死的科目清单和年份——新增科目、跨年度都自动生效。每一行业务同时产生一借一贷，所以借贷永远相等。')
widths(ws, {'A':11,'B':18,'C':8,'D':7,'E':14,'F':14,'G':14,'H':14,'I':14,'J':14,'K':15,'L':14,'M':14,'N':17,'O':15,'P':11})
put(ws, 'A3', '查询月份', font=F_H2, align=CR, border=None)
put(ws, 'B3', f'={YM_CELL}', font=F_TOT, fill=FILL_TOT)
put(ws, 'D3', '本月借方合计', font=F_NOTE, align=CR, border=None)
put(ws, 'E3', f'=SUM($G${R0}:$G${R1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, 'F3', '本月贷方合计', font=F_NOTE, align=CR, border=None)
put(ws, 'G3', f'=SUM($H${R0}:$H${R1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, 'H3', '平衡', font=F_NOTE, align=CR, border=None)
put(ws, 'I3', f'=IF(ABS($E$3-$G$3)<0.01,"✓ 借贷平衡","✗ 差 "&TEXT($E$3-$G$3,"#,##0.00"))',
    font=F_TOT, fill=FILL_CHK, align=CL)
ws.merge_cells('I3:K3')
headers(ws, 5, 1, ['科目编码','科目名称','类别','余额\n方向','期初借方','期初贷方','本月借方','本月贷方',
                   '本年累计\n借方','本年累计\n贷方','期末余额\n(借正贷负)','期末借方','期末贷方',
                   '资产负债表项目','利润表项目','现金流性质'])
OPNQ = Q(SH_OPEN)
for r in range(R0, R1 + 1):
    a = f'$A{r}'
    put(ws, f'A{r}', f'=IF({PARAM}!$A{r}="","",{PARAM}!$A{r})', font=F_LINK)
    for col, src in [('B', 'B'), ('C', 'C'), ('D', 'D'), ('N', 'F'), ('O', 'G'), ('P', 'H')]:
        put(ws, f'{col}{r}', f'=IF({a}="","",{PARAM}!${src}{r})', font=F_LINK,
            align=CL if col in ('B', 'N', 'O') else C)
    put(ws, f'E{r}', f'=IF({a}="","",SUMIFS({OPNQ}!$Z${OA_R0}:$Z${OA_R1},{OPNQ}!$W${OA_R0}:$W${OA_R1},{a}))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF({a}="","",SUMIFS({OPNQ}!$AA${OA_R0}:$AA${OA_R1},{OPNQ}!$W${OA_R0}:$W${OA_R1},{a}))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'G{r}', f'=IF({a}="","",{occ("D", a, "month")})', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF({a}="","",{occ("C", a, "month")})', font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF({a}="","",{occ("D", a, "year")})', font=F_LINK, fmt=MONEY)
    put(ws, f'J{r}', f'=IF({a}="","",{occ("C", a, "year")})', font=F_LINK, fmt=MONEY)
    put(ws, f'K{r}', f'=IF({a}="","",($E{r}+$I{r})-($F{r}+$J{r}))', font=F_TXT, fmt=MONEY)
    put(ws, f'L{r}', f'=IF({a}="","",IF($K{r}>=0,$K{r},0))', font=F_TXT, fmt=MONEY)
    put(ws, f'M{r}', f'=IF({a}="","",IF($K{r}<0,-$K{r},0))', font=F_TXT, fmt=MONEY)
    ws.row_dimensions[r].height = 16
TR = R1 + 1
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT); put(ws, f'B{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in ['C','D','N','O','P']: put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in ['E','F','G','H','I','J','K','L','M']:
    put(ws, f'{c}{TR}', f'=SUM({c}{R0}:{c}{R1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
ws.freeze_panes = 'C6'
page(ws, titles='5:5')

def tb_sum(col, item_col, item, sign=1):
    s = f'SUMIFS({Q(SH_TB)}!${col}${ACC_R0}:${col}${ACC_R1},{Q(SH_TB)}!${item_col}${ACC_R0}:${item_col}${ACC_R1},{item})'
    return s if sign > 0 else f'-{s}'

# ============================================================ 利润表
ws = wb.create_sheet(SH_PL)
title(ws, '利润表', 'D', '按【科目参数】里的「利润表项目」自动归集。收入类取贷方减借方，成本费用类取借方减贷方。')
widths(ws, {'A':34,'B':7,'C':17,'D':17})
put(ws, 'A3', '查询月份', font=F_H2, align=CR, border=None)
put(ws, 'B3', f'={YM_CELL}', font=F_TOT, fill=FILL_TOT); ws.merge_cells('B3:C3')
headers(ws, 5, 1, ['项    目','行次','本月金额','本年累计'])
PL = [  # (显示名, 行次, 类型, 利润表项目 或 计算式)
 ('一、营业收入',            1, 'sum',  ['主营业务收入', '其他业务收入']),
 ('        其中：主营业务收入', 2, 'rev', '主营业务收入'),
 ('               其他业务收入', 3, 'rev', '其他业务收入'),
 ('减：营业成本',             4, 'sum',  ['主营业务成本', '其他业务成本']),
 ('        其中：主营业务成本', 5, 'exp', '主营业务成本'),
 ('               其他业务成本', 6, 'exp', '其他业务成本'),
 ('        税金及附加',        7, 'exp', '税金及附加'),
 ('        销售费用',          8, 'exp', '销售费用'),
 ('        管理费用',          9, 'exp', '管理费用'),
 ('        研发费用',         10, 'exp', '研发费用'),
 ('        财务费用',         11, 'exp', '财务费用'),
 ('        资产减值损失',     12, 'exp', '资产减值损失'),
 ('        信用减值损失',     13, 'exp', '信用减值损失'),
 ('加：其他收益',             14, 'rev', '其他收益'),
 ('        投资收益',         15, 'rev', '投资收益'),
 ('        资产处置收益',     16, 'rev', '资产处置收益'),
 ('二、营业利润',             17, 'calc', 'R1+R14+R15+R16-R4-R7-R8-R9-R10-R11-R12-R13'),
 ('加：营业外收入',           18, 'rev', '营业外收入'),
 ('减：营业外支出',           19, 'exp', '营业外支出'),
 ('三、利润总额',             20, 'calc', 'R17+R18-R19'),
 ('减：所得税费用',           21, 'exp', '所得税费用'),
 ('四、净利润',               22, 'calc', 'R20-R21'),
]
ROWMAP = {}
r = 6
for name, no, kind, spec in PL:
    ROWMAP[no] = r
    bold = name.startswith(('一', '二', '三', '四'))
    put(ws, f'A{r}', name, font=F_TOT if bold else F_TXT, align=CL,
        fill=FILL_TOT if bold else None)
    put(ws, f'B{r}', no, font=F_TXT, fill=FILL_TOT if bold else None)
    for col, mc, yc in [('C', 'H', 'G'), ('D', 'J', 'I')]:
        cr_col, dr_col = (mc, yc) if col == 'C' else ('J', 'I')
        if kind == 'rev':
            f = f'={tb_sum(cr_col,"O",chr(34)+spec+chr(34))}-{tb_sum(dr_col,"O",chr(34)+spec+chr(34))}'
        elif kind == 'exp':
            f = f'={tb_sum(dr_col,"O",chr(34)+spec+chr(34))}-{tb_sum(cr_col,"O",chr(34)+spec+chr(34))}'
        elif kind == 'sum':
            parts = []
            for it in spec:
                if it.endswith('收入'):
                    parts.append(f'({tb_sum(cr_col,"O",chr(34)+it+chr(34))}-{tb_sum(dr_col,"O",chr(34)+it+chr(34))})')
                else:
                    parts.append(f'({tb_sum(dr_col,"O",chr(34)+it+chr(34))}-{tb_sum(cr_col,"O",chr(34)+it+chr(34))})')
            f = '=' + '+'.join(parts)
        else:
            f = '=' + spec
        put(ws, f'{col}{r}', f, font=F_TXT if kind != 'calc' else F_TOT,
            fill=FILL_TOT if bold else None, fmt=MONEY)
    ws.row_dimensions[r].height = 18
    r += 1
for rr in range(6, r):
    for col in ['C', 'D']:
        v = ws[f'{col}{rr}'].value
        if isinstance(v, str) and 'R' in v and v.startswith('=R'):
            import re as _re
            ws[f'{col}{rr}'] = '=' + _re.sub(r'R(\d+)', lambda m: f'{col}{ROWMAP[int(m.group(1))]}', v[1:])
ws.freeze_panes = 'A6'
page(ws, titles='5:5', landscape=False)
print('  ✓ 科目余额表 / 利润表')

# ============================================================ 资产负债表
ws = wb.create_sheet(SH_BS)
title(ws, '资产负债表', 'H',
      '按【科目参数】里的「资产负债表项目」自动归集。未分配利润＝利润分配类科目余额＋利润表本年净利润，年末做了结转损益也不会重复计算。')
widths(ws, {'A':30,'B':7,'C':16,'D':16,'E':30,'F':7,'G':16,'H':16})
put(ws, 'A3', '查询月份', font=F_H2, align=CR, border=None)
put(ws, 'B3', f'={YM_CELL}', font=F_TOT, fill=FILL_TOT); ws.merge_cells('B3:C3')
headers(ws, 5, 1, ['资      产','行次','期末余额','年初余额','负债和所有者权益','行次','期末余额','年初余额'])
ASSETS = ['货币资金','交易性金融资产','应收票据','应收账款','预付款项','其他应收款','存货','其他流动资产',
          '@流动资产合计', '债权投资','长期股权投资','投资性房地产','固定资产','在建工程','无形资产',
          '长期待摊费用','递延所得税资产','其他非流动资产','@非流动资产合计','@@资产总计']
LIABS = ['短期借款','应付票据','应付账款','预收款项','应付职工薪酬','应交税费','其他应付款',
         '一年内到期的非流动负债','其他流动负债','@流动负债合计','长期借款','长期应付款','递延收益',
         '递延所得税负债','其他非流动负债','@非流动负债合计','@负债合计','',
         '实收资本','资本公积','盈余公积','未分配利润','@所有者权益合计','@@负债和所有者权益总计']
def bs_item(item, col_end, col_beg, sign):
    # 期末取【科目余额表】K 列（期末余额·借正贷负）；年初取 E-F
    e = f'SUMIFS({Q(SH_TB)}!$K${ACC_R0}:$K${ACC_R1},{Q(SH_TB)}!$N${ACC_R0}:$N${ACC_R1},"{item}")'
    b = (f'SUMIFS({Q(SH_TB)}!$E${ACC_R0}:$E${ACC_R1},{Q(SH_TB)}!$N${ACC_R0}:$N${ACC_R1},"{item}")'
         f'-SUMIFS({Q(SH_TB)}!$F${ACC_R0}:$F${ACC_R1},{Q(SH_TB)}!$N${ACC_R0}:$N${ACC_R1},"{item}")')
    return (f'={e}' if sign > 0 else f'=-({e})'), (f'={b}' if sign > 0 else f'=-({b})')

def render(items, col_name, col_no, col_end, col_beg, sign, r_start):
    rows, sub = {}, []
    r = r_start
    for it in items:
        if it == '':
            for c in [col_name, col_no, col_end, col_beg]:
                put(ws, f'{c}{r}', None, border=None)
            r += 1; continue
        tot = it.startswith('@')
        big = it.startswith('@@')
        nm = it.lstrip('@')
        put(ws, f'{col_name}{r}', ('　' if not tot else '') + nm,
            font=F_TOT if tot else F_TXT, align=CL, fill=FILL_TOT if tot else None)
        put(ws, f'{col_no}{r}', r - r_start + 1, font=F_TXT, fill=FILL_TOT if tot else None)
        if not tot:
            fe, fb = bs_item(nm, col_end, col_beg, sign)
            if nm == '未分配利润':
                fe = fe[:-1] + f')+{Q(SH_PL)}!$D${ROWMAP[22]}' if fe.startswith('=-(') else fe
                fb = fb
            put(ws, f'{col_end}{r}', fe, font=F_TXT, fmt=MONEY)
            put(ws, f'{col_beg}{r}', fb, font=F_TXT, fmt=MONEY)
            sub.append(r)
        else:
            rows[nm] = (r, list(sub)); sub = []
            for c in [col_end, col_beg]:
                put(ws, f'{c}{r}', None, font=F_TOT, fill=FILL_TOT, fmt=MONEY)
        ws.row_dimensions[r].height = 18
        r += 1
    return rows, r

AR, aend = render(ASSETS, 'A', 'B', 'C', 'D', 1, 6)
LR, lend = render(LIABS, 'E', 'F', 'G', 'H', -1, 6)
def fill_tot(rows, key, cols, expr):
    r = rows[key][0]
    for c in cols:
        put(ws, f'{c}{r}', expr(c, r), font=F_TOT, fill=FILL_TOT, fmt=MONEY)
fill_tot(AR, '流动资产合计', ['C','D'], lambda c, r: f'=SUM({c}{6}:{c}{r-1})')
fill_tot(AR, '非流动资产合计', ['C','D'], lambda c, r: f'=SUM({c}{AR["流动资产合计"][0]+1}:{c}{r-1})')
fill_tot(AR, '资产总计', ['C','D'], lambda c, r: f'={c}{AR["流动资产合计"][0]}+{c}{AR["非流动资产合计"][0]}')
fill_tot(LR, '流动负债合计', ['G','H'], lambda c, r: f'=SUM({c}{6}:{c}{r-1})')
fill_tot(LR, '非流动负债合计', ['G','H'], lambda c, r: f'=SUM({c}{LR["流动负债合计"][0]+1}:{c}{r-1})')
fill_tot(LR, '负债合计', ['G','H'], lambda c, r: f'={c}{LR["流动负债合计"][0]}+{c}{LR["非流动负债合计"][0]}')
fill_tot(LR, '所有者权益合计', ['G','H'], lambda c, r: f'=SUM({c}{LR["负债合计"][0]+2}:{c}{r-1})')
fill_tot(LR, '负债和所有者权益总计', ['G','H'], lambda c, r: f'={c}{LR["负债合计"][0]}+{c}{LR["所有者权益合计"][0]}')
BAL_R = max(aend, lend) + 1
put(ws, f'A{BAL_R}', '平 衡 校 验', font=F_TOT, fill=FILL_CHK)
ws.merge_cells(f'A{BAL_R}:B{BAL_R}'); put(ws, f'B{BAL_R}', None, font=F_TOT, fill=FILL_CHK)
ta, tl = AR['资产总计'][0], LR['负债和所有者权益总计'][0]
put(ws, f'C{BAL_R}', f'=IF(ABS($C${ta}-$G${tl})<0.01,"✓ 平衡","✗ 差 "&TEXT($C${ta}-$G${tl},"#,##0.00"))',
    font=F_TOT, fill=FILL_CHK, align=CL)
ws.merge_cells(f'C{BAL_R}:E{BAL_R}')
for c in ['D','E']: put(ws, f'{c}{BAL_R}', None, font=F_TOT, fill=FILL_CHK)
put(ws, f'F{BAL_R}', '年初平衡', font=F_TOT, fill=FILL_CHK)
put(ws, f'G{BAL_R}', f'=IF(ABS($D${ta}-$H${tl})<0.01,"✓ 平衡","✗ 差 "&TEXT($D${ta}-$H${tl},"#,##0.00"))',
    font=F_TOT, fill=FILL_CHK, align=CL)
ws.merge_cells(f'G{BAL_R}:H{BAL_R}'); put(ws, f'H{BAL_R}', None, font=F_TOT, fill=FILL_CHK)
ws.freeze_panes = 'A6'
page(ws, titles='5:5')
print('  ✓ 资产负债表')

# ============================================================ 现金流量表
ws = wb.create_sheet(SH_CF)
title(ws, '现金流量表', 'D',
      '直接法，取自【资金流水】。每一笔收付款按收支科目自动归入现金流量表项目；账户之间的内部转账不计入。')
widths(ws, {'A':38,'B':7,'C':17,'D':17})
put(ws, 'A3', '查询月份', font=F_H2, align=CR, border=None)
put(ws, 'B3', f'={YM_CELL}', font=F_TOT, fill=FILL_TOT); ws.merge_cells('B3:C3')
headers(ws, 5, 1, ['项    目','行次','本月金额','本年累计'])
def cf(item, io_type, period):
    p = (f'{CASHQ}!$C${CASH_R0}:$C${CASH_R1},{YM_CELL}' if period == 'm'
         else f'{CASHQ}!$B${CASH_R0}:$B${CASH_R1},"<="&{EOM_CELL}')
    return (f'SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$Q${CASH_R0}:$Q${CASH_R1},"{item}",'
            f'{CASHQ}!$D${CASH_R0}:$D${CASH_R1},"{io_type}",{p})')
CF = [
 ('一、经营活动产生的现金流量', None, 'hd', None),
 ('　　销售商品、提供劳务收到的现金', 1, '收入', '销售商品、提供劳务收到的现金'),
 ('　　收到的其他与经营活动有关的现金', 2, '收入', '收到的其他与经营活动有关的现金'),
 ('　　　　经营活动现金流入小计', 3, 'sub+', None),
 ('　　购买商品、接受劳务支付的现金', 4, '支出', '购买商品、接受劳务支付的现金'),
 ('　　支付给职工以及为职工支付的现金', 5, '支出', '支付给职工以及为职工支付的现金'),
 ('　　支付的各项税费', 6, '支出', '支付的各项税费'),
 ('　　支付的其他与经营活动有关的现金', 7, '支出', '支付的其他与经营活动有关的现金'),
 ('　　　　经营活动现金流出小计', 8, 'sub-', None),
 ('　　经营活动产生的现金流量净额', 9, 'net', (3, 8)),
 ('二、投资活动产生的现金流量', None, 'hd', None),
 ('　　处置固定资产收回的现金净额', 10, '收入', '处置固定资产收回的现金净额'),
 ('　　取得投资收益收到的现金', 11, '收入', '取得投资收益收到的现金'),
 ('　　　　投资活动现金流入小计', 12, 'sub+', None),
 ('　　购建固定资产支付的现金', 13, '支出', '购建固定资产支付的现金'),
 ('　　购建无形资产支付的现金', 14, '支出', '购建无形资产支付的现金'),
 ('　　投资支付的现金', 15, '支出', '投资支付的现金'),
 ('　　　　投资活动现金流出小计', 16, 'sub-', None),
 ('　　投资活动产生的现金流量净额', 17, 'net', (12, 16)),
 ('三、筹资活动产生的现金流量', None, 'hd', None),
 ('　　吸收投资收到的现金', 18, '收入', '吸收投资收到的现金'),
 ('　　取得借款收到的现金', 19, '收入', '取得借款收到的现金'),
 ('　　　　筹资活动现金流入小计', 20, 'sub+', None),
 ('　　偿还债务支付的现金', 21, '支出', '偿还债务支付的现金'),
 ('　　分配股利、利润或偿付利息支付的现金', 22, '支出', '分配股利、利润或偿付利息支付的现金'),
 ('　　　　筹资活动现金流出小计', 23, 'sub-', None),
 ('　　筹资活动产生的现金流量净额', 24, 'net', (20, 23)),
 ('四、现金及现金等价物净增加额', 25, 'tot', (9, 17, 24)),
 ('　　加：期初现金及现金等价物余额', 26, 'beg', None),
 ('五、期末现金及现金等价物余额', 27, 'end', (25, 26)),
]
CFR, r, pend = {}, 6, []
for name, no, kind, spec in CF:
    if no: CFR[no] = r
    big = kind in ('hd', 'net', 'tot', 'end')
    put(ws, f'A{r}', name, font=F_TOT if big else F_TXT, align=CL, fill=FILL_TOT if big else None)
    put(ws, f'B{r}', no, font=F_TXT, fill=FILL_TOT if big else None)
    for col, per in [('C', 'm'), ('D', 'y')]:
        if kind == 'hd':
            put(ws, f'{col}{r}', None, font=F_TOT, fill=FILL_TOT, fmt=MONEY); continue
        if kind in ('收入', '支出'):
            f = '=' + cf(spec, kind, per); pend.append((col, r, kind))
        elif kind in ('sub+', 'sub-'):
            src = [rr for cc, rr, kk in pend if cc == col and kk == ('收入' if kind == 'sub+' else '支出')]
            f = '=' + '+'.join(f'{col}{x}' for x in src) if src else '=0'
            pend = [(cc, rr, kk) for cc, rr, kk in pend if not (cc == col and kk == ('收入' if kind == 'sub+' else '支出'))]
        elif kind == 'net':
            f = f'={col}{CFR[spec[0]]}-{col}{CFR[spec[1]]}'
        elif kind == 'tot':
            f = '=' + '+'.join(f'{col}{CFR[x]}' for x in spec)
        elif kind == 'beg':
            f = (f'=SUM({Q(SH_OPEN)}!$B${OC_R0}:$B${OC_R1})'
                 if col == 'D' else f'=$D{r}+SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$D${CASH_R0}:$D${CASH_R1},"收入",'
                                    f'{CASHQ}!$B${CASH_R0}:$B${CASH_R1},"<"&{BOM_CELL})'
                                    f'-SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$D${CASH_R0}:$D${CASH_R1},"支出",'
                                    f'{CASHQ}!$B${CASH_R0}:$B${CASH_R1},"<"&{BOM_CELL})')
        else:
            f = f'={col}{CFR[spec[0]]}+{col}{CFR[spec[1]]}'
        put(ws, f'{col}{r}', f, font=F_TOT if big else F_TXT, fill=FILL_TOT if big else None, fmt=MONEY)
    ws.row_dimensions[r].height = 18
    r += 1
put(ws, f'A{r+1}', '核对：本表「期末现金及现金等价物余额」应等于【账户余额】的期末合计、'
                   '也等于【资产负债表】的货币资金。三个数不一致时请到【财务勾稽核查】查原因。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{r+1}:D{r+1}')
ws.freeze_panes = 'A6'
page(ws, titles='5:5', landscape=False)

# ============================================================ 账户余额
ws = wb.create_sheet(SH_BAL)
R0, R1 = ACT_R0, ACT_R1
title(ws, '账户余额', 'I', '每个资金账户的期初、本月收支、本年累计与期末余额。期末余额合计应与现金流量表、资产负债表货币资金一致。')
widths(ws, {'A':18,'B':11,'C':14,'D':15,'E':15,'F':15,'G':15,'H':16,'I':13})
put(ws, 'A3', '查询月份', font=F_H2, align=CR, border=None)
put(ws, 'B3', f'={YM_CELL}', font=F_TOT, fill=FILL_TOT)
headers(ws, 5, 1, ['账户','类型','期初余额','本月收入\n(含转入)','本月支出\n(含转出)','本年收入\n(含转入)','本年支出\n(含转出)','期末余额','占比'])
def cash_flow_sum(kind, acct_ref, period, direction):
    """direction: in/out；kind: 本账户列 F 或转入列 G"""
    p = (f'{CASHQ}!$C${CASH_R0}:$C${CASH_R1},{YM_CELL}' if period == 'm'
         else f'{CASHQ}!$B${CASH_R0}:$B${CASH_R1},"<="&{EOM_CELL}')
    col = 'F' if kind == 'self' else 'G'
    return (f'SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!${col}${CASH_R0}:${col}${CASH_R1},{acct_ref},'
            f'{CASHQ}!$D${CASH_R0}:$D${CASH_R1},"{direction}",{p})')
for r in range(R0, R1 + 1):
    a = f'$A{r}'
    put(ws, f'A{r}', f'=IF({BASE}!$A{r}="","",{BASE}!$A{r})', font=F_LINK)
    put(ws, f'B{r}', f'=IF({a}="","",{BASE}!$B{r})', font=F_LINK)
    put(ws, f'C{r}', f'=IF({a}="","",SUMIFS({Q(SH_OPEN)}!$B${OC_R0}:$B${OC_R1},{Q(SH_OPEN)}!$A${OC_R0}:$A${OC_R1},{a}))',
        font=F_LINK, fmt=MONEY)
    for col, per in [('D', 'm'), ('F', 'y')]:
        put(ws, f'{col}{r}', f'=IF({a}="","",{cash_flow_sum("self", a, per, "收入")}+{cash_flow_sum("to", a, per, "转账")})',
            font=F_LINK, fmt=MONEY)
    for col, per in [('E', 'm'), ('G', 'y')]:
        put(ws, f'{col}{r}', f'=IF({a}="","",{cash_flow_sum("self", a, per, "支出")}+{cash_flow_sum("self", a, per, "转账")})',
            font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF({a}="","",$C{r}+$F{r}-$G{r})', font=F_TOT, fmt=MONEY)
    put(ws, f'I{r}', f'=IF(OR({a}="",$H${R1+1}=0),"",$H{r}/$H${R1+1})', font=F_TXT, fmt=PCT)
    ws.row_dimensions[r].height = 17
TR = R1 + 1
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT); put(ws, f'B{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in ['C','D','E','F','G','H']:
    put(ws, f'{c}{TR}', f'=SUM({c}{R0}:{c}{R1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'I{TR}', None, font=F_TOT, fill=FILL_TOT)
ws.freeze_panes = 'C6'
page(ws, titles='5:5')
print('  ✓ 现金流量表 / 账户余额')

# ============================================================ 库存收发存
ws = wb.create_sheet(SH_INV)
R0, R1 = GDS_R0, GDS_R1
title(ws, '库存收发存', 'P',
      '数量按「期初＋入库－出库」滚动，金额按「期初金额＋采购金额－结转成本」滚动，所以月末金额与科目余额表的「库存商品」永远一致。')
widths(ws, {'A':16,'B':11,'C':7,'D':12,'E':14,'F':12,'G':14,'H':12,'I':14,'J':12,'K':14,'L':13,'M':14,'N':14,'O':9,'P':15})
put(ws, 'A3', '查询月份', font=F_H2, align=CR, border=None)
put(ws, 'B3', f'={YM_CELL}', font=F_TOT, fill=FILL_TOT)
headers(ws, 5, 1, ['商品','规格','单位','月初数量','月初金额','本月入库\n数量','本月入库\n金额','本月出库\n数量',
                   '本月出库\n成本','月末数量','月末金额','月末\n单位成本','本月销售额','本月毛利','毛利率','预警'])
def buy_s(col, gref, cond):
    return (f'SUMIFS({BUYQ}!${col}${BUY_R0}:${col}${BUY_R1},{BUYQ}!$E${BUY_R0}:$E${BUY_R1},{gref},'
            f'{BUYQ}!$I${BUY_R0}:$I${BUY_R1},"是",{cond})')
def sal_s(col, gref, cond):
    return (f'SUMIFS({SALQ}!${col}${SAL_R0}:${col}${SAL_R1},{SALQ}!$E${SAL_R0}:$E${SAL_R1},{gref},{cond})')
BEFORE_B = f'{BUYQ}!$B${BUY_R0}:$B${BUY_R1},"<"&{BOM_CELL}'
BEFORE_S = f'{SALQ}!$B${SAL_R0}:$B${SAL_R1},"<"&{BOM_CELL}'
MON_B    = f'{BUYQ}!$C${BUY_R0}:$C${BUY_R1},{YM_CELL}'
MON_S    = f'{SALQ}!$C${SAL_R0}:$C${SAL_R1},{YM_CELL}'
for r in range(R0, R1 + 1):
    a = f'$A{r}'
    put(ws, f'A{r}', f'=IF({BASE}!$Q{r}="","",{BASE}!$Q{r})', font=F_LINK, align=CL)
    put(ws, f'B{r}', f'=IF({a}="","",{BASE}!$R{r})', font=F_LINK)
    put(ws, f'C{r}', f'=IF({a}="","",{BASE}!$S{r})', font=F_LINK)
    put(ws, f'D{r}', f'=IF({a}="","",SUMIFS({Q(SH_OPEN)}!$R${OG_R0}:$R${OG_R1},{Q(SH_OPEN)}!$O${OG_R0}:$O${OG_R1},{a})'
                     f'+{buy_s("F", a, BEFORE_B)}-{sal_s("F", a, BEFORE_S)})', font=F_LINK, fmt=QTY)
    put(ws, f'E{r}', f'=IF({a}="","",SUMIFS({Q(SH_OPEN)}!$T${OG_R0}:$T${OG_R1},{Q(SH_OPEN)}!$O${OG_R0}:$O${OG_R1},{a})'
                     f'+{buy_s("Q", a, BEFORE_B)}-{sal_s("S", a, BEFORE_S)})', font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF({a}="","",{buy_s("F", a, MON_B)})', font=F_LINK, fmt=QTY)
    put(ws, f'G{r}', f'=IF({a}="","",{buy_s("Q", a, MON_B)})', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF({a}="","",{sal_s("F", a, MON_S)})', font=F_LINK, fmt=QTY)
    put(ws, f'I{r}', f'=IF({a}="","",{sal_s("S", a, MON_S)})', font=F_LINK, fmt=MONEY)
    put(ws, f'J{r}', f'=IF({a}="","",$D{r}+$F{r}-$H{r})', font=F_TXT, fmt=QTY)
    put(ws, f'K{r}', f'=IF({a}="","",$E{r}+$G{r}-$I{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'L{r}', f'=IF(OR({a}="",$J{r}=0),"",$K{r}/$J{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'M{r}', f'=IF({a}="","",{sal_s("O", a, MON_S)})', font=F_LINK, fmt=MONEY)
    put(ws, f'N{r}', f'=IF({a}="","",$M{r}-$I{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'O{r}', f'=IF(OR({a}="",$M{r}=0),"",$N{r}/$M{r})', font=F_TXT, fmt=PCT)
    put(ws, f'P{r}', f'=IF({a}="","",IF($J{r}<0,"⚠ 负库存",IF(AND($J{r}=0,ABS($K{r})>0.01),"⚠ 有金额无数量",'
                     f'IF(AND($J{r}>0,$K{r}<0),"⚠ 金额为负",IF($N{r}<0,"⚠ 毛利为负","正常")))))', font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
TR = R1 + 1
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in ['B','C','L','P']: put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in ['D','E','F','G','H','I','J','K','M','N']:
    put(ws, f'{c}{TR}', f'=SUM({c}{R0}:{c}{R1})', font=F_TOT, fill=FILL_TOT,
        fmt=QTY if c in ('D','F','H','J') else MONEY)
put(ws, f'O{TR}', f'=IF($M${TR}=0,"",$N${TR}/$M${TR})', font=F_TOT, fill=FILL_TOT, fmt=PCT)
ws.conditional_formatting.add(f'P{R0}:P{R1}',
    FormulaRule(formula=[f'LEFT($P{R0},1)="⚠"'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.freeze_panes = 'D6'
page(ws, titles='5:5')

# ============================================================ 应收 / 应付跟进
def build_due(sheet, is_ar):
    ws = wb.create_sheet(sheet)
    R0, R1 = (CUS_R0, CUS_R1) if is_ar else (SUP_R0, SUP_R1)
    who, base_c, base_n, base_term, base_own = (
        ('客户', 'Y', 'Z', 'AB', 'AA') if is_ar else ('供应商', 'AE', 'AF', 'AH', 'AG'))
    title(ws, f'应{"收" if is_ar else "付"}跟进', 'N',
          f'月末应{"收" if is_ar else "付"}＝月初＋本月{"销售" if is_ar else "采购"}－本月{"回款" if is_ar else "付款"}；'
          f'账龄按到期日分桶，{"回款" if is_ar else "付款"}按最早的单据先冲抵。')
    widths(ws, {'A':18,'B':8,'C':14,'D':14,'E':14,'F':15,'G':13,'H':13,'I':13,'J':13,'K':14,'L':14,'M':11,'N':11})
    put(ws, 'A3', '查询月份', font=F_H2, align=CR, border=None)
    put(ws, 'B3', f'={YM_CELL}', font=F_TOT, fill=FILL_TOT)
    put(ws, 'D3', '截止日', font=F_NOTE, align=CR, border=None)
    put(ws, 'E3', f'={EOM_CELL}', font=F_TOT, fill=FILL_TOT, fmt=DATE)
    headers(ws, 5, 1, [who, '账期\n(天)', f'月初应{"收" if is_ar else "付"}', f'本月{"销售" if is_ar else "采购"}',
                       f'本月{"回款" if is_ar else "付款"}', f'月末应{"收" if is_ar else "付"}',
                       '未到期', '逾期\n1-30天', '逾期\n31-60天', '逾期\n61-90天', '逾期\n90天以上', '逾期合计',
                       '风险\n等级', '业务员' if is_ar else '采购员'])
    biz, amt_c, dat_c, due_c, party_c = ((SALQ, 'Q', 'B', 'U', 'D') if is_ar else (BUYQ, 'S', 'B', 'T', 'D'))
    pay_kind = '客户回款' if is_ar else '供应商付款'
    pay_dir  = '收入' if is_ar else '支出'
    op_party, op_date, op_amt = (('E', 'F', 'G') if is_ar else ('J', 'K', 'L'))
    OP = Q(SH_OPEN)
    o0, o1 = (OR_R0, OR_R1) if is_ar else (OP_R0, OP_R1)
    def opn(pref, cond=''):
        return f'SUMIFS({OP}!${op_amt}${o0}:${op_amt}${o1},{OP}!${op_party}${o0}:${op_party}${o1},{pref}{cond})'
    def bizsum(pref, cond):
        return f'SUMIFS({biz}!${amt_c}${o0 if False else (SAL_R0 if is_ar else BUY_R0)}:${amt_c}${(SAL_R1 if is_ar else BUY_R1)},' \
               f'{biz}!${party_c}${(SAL_R0 if is_ar else BUY_R0)}:${party_c}${(SAL_R1 if is_ar else BUY_R1)},{pref}{cond})'
    def paysum(pref, cond):
        return (f'SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$G${CASH_R0}:$G${CASH_R1},{pref},'
                f'{CASHQ}!$E${CASH_R0}:$E${CASH_R1},"{pay_kind}"{cond})')
    for r in range(R0, R1 + 1):
        a = f'$A{r}'
        put(ws, f'A{r}', f'=IF({BASE}!${base_n}{r}="","",{BASE}!${base_n}{r})', font=F_LINK, align=CL)
        put(ws, f'B{r}', f'=IF({a}="","",N({BASE}!${base_term}{r}))', font=F_LINK)
        bef_b = f',{biz}!${dat_c}${(SAL_R0 if is_ar else BUY_R0)}:${dat_c}${(SAL_R1 if is_ar else BUY_R1)},"<"&{BOM_CELL}'
        mon_b = f',{biz}!$C${(SAL_R0 if is_ar else BUY_R0)}:$C${(SAL_R1 if is_ar else BUY_R1)},{YM_CELL}'
        put(ws, f'C{r}', f'=IF({a}="","",{opn(a, f",{OP}!${op_date}${o0}:${op_date}${o1},chr34<chr34&{BOM_CELL}".replace("chr34",chr(34)))}'
                         f'+{bizsum(a, bef_b)}-{paysum(a, f",{CASHQ}!$B${CASH_R0}:$B${CASH_R1},chr34<chr34&{BOM_CELL}".replace("chr34",chr(34)))})',
            font=F_LINK, fmt=MONEY)
        put(ws, f'D{r}', f'=IF({a}="","",{bizsum(a, mon_b)})', font=F_LINK, fmt=MONEY)
        put(ws, f'E{r}', f'=IF({a}="","",{paysum(a, f",{CASHQ}!$C${CASH_R0}:$C${CASH_R1},{YM_CELL}")})', font=F_LINK, fmt=MONEY)
        put(ws, f'F{r}', f'=IF({a}="","",$C{r}+$D{r}-$E{r})', font=F_TOT, fmt=MONEY)
        pay_all = paysum(a, f',{CASHQ}!$B${CASH_R0}:$B${CASH_R1},"<="&{EOM_CELL}')
        def due_le(off):
            d = f'{EOM_CELL}' if off == 0 else f'{EOM_CELL}-{off}'
            return (f'MAX(0,{opn(a, f",{OP}!${op_date}${o0}:${op_date}${o1},chr34<=chr34&{d}".replace("chr34",chr(34)))}'
                    f'+SUMIFS({biz}!${amt_c}${(SAL_R0 if is_ar else BUY_R0)}:${amt_c}${(SAL_R1 if is_ar else BUY_R1)},'
                    f'{biz}!${party_c}${(SAL_R0 if is_ar else BUY_R0)}:${party_c}${(SAL_R1 if is_ar else BUY_R1)},{a},'
                    f'{biz}!${due_c}${(SAL_R0 if is_ar else BUY_R0)}:${due_c}${(SAL_R1 if is_ar else BUY_R1)},"<="&{d})'
                    f'-{pay_all})')
        put(ws, f'K{r}', f'=IF({a}="","",MIN($F{r},{due_le(90)}))', font=F_TXT, fmt=MONEY)
        put(ws, f'J{r}', f'=IF({a}="","",MIN($F{r},{due_le(60)})-$K{r})', font=F_TXT, fmt=MONEY)
        put(ws, f'I{r}', f'=IF({a}="","",MIN($F{r},{due_le(30)})-$K{r}-$J{r})', font=F_TXT, fmt=MONEY)
        put(ws, f'H{r}', f'=IF({a}="","",MIN($F{r},{due_le(0)})-$K{r}-$J{r}-$I{r})', font=F_TXT, fmt=MONEY)
        put(ws, f'G{r}', f'=IF({a}="","",$F{r}-$H{r}-$I{r}-$J{r}-$K{r})', font=F_TXT, fmt=MONEY)
        put(ws, f'L{r}', f'=IF({a}="","",$H{r}+$I{r}+$J{r}+$K{r})', font=F_TOT, fmt=MONEY)
        put(ws, f'M{r}', f'=IF({a}="","",IF($K{r}>0,"高危",IF($J{r}>0,"关注",IF($L{r}>0,"提醒","正常"))))',
            font=F_TXT, fill=FILL_CHK)
        put(ws, f'N{r}', f'=IF({a}="","",{BASE}!${base_own}{r})', font=F_LINK)
        ws.row_dimensions[r].height = 16
    TR = R1 + 1
    put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT)
    for c in ['B','M','N']: put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
    for c in ['C','D','E','F','G','H','I','J','K','L']:
        put(ws, f'{c}{TR}', f'=SUM({c}{R0}:{c}{R1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
    ws.conditional_formatting.add(f'M{R0}:M{R1}',
        FormulaRule(formula=[f'$M{R0}="高危"'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
    ws.freeze_panes = 'C6'
    page(ws, titles='5:5')
build_due(SH_AR, True)
build_due(SH_AP, False)
print('  ✓ 库存收发存 / 应收跟进 / 应付跟进')

# ============================================================ 经营分析
ws = wb.create_sheet(SH_ANA)
title(ws, '经营分析', 'K', '全年 12 个月趋势。数据全部取自科目余额表口径以外的业务表，可与利润表交叉验证。')
widths(ws, {'A':10,'B':15,'C':15,'D':14,'E':10,'F':14,'G':14,'H':14,'I':14,'J':15,'K':15})
headers(ws, 5, 1, ['月份','销售收入\n(不含税)','销售成本','毛利','毛利率','采购金额\n(不含税)','期间费用',
                   '资金流入','资金流出','资金净额','期末资金余额'])
for i in range(12):
    r = 6 + i
    m = f'TEXT(DATE({YEAR_CELL},{i+1},1),"yyyy-mm")'
    eom = f'EOMONTH(DATE({YEAR_CELL},{i+1},1),0)'
    put(ws, f'A{r}', f'={i+1}&"月"', font=F_TXT, fill=FILL_HDR2)
    put(ws, f'B{r}', f'=SUMIFS({SALQ}!$O${SAL_R0}:$O${SAL_R1},{SALQ}!$C${SAL_R0}:$C${SAL_R1},{m})', font=F_LINK, fmt=MONEY)
    put(ws, f'C{r}', f'=SUMIFS({SALQ}!$S${SAL_R0}:$S${SAL_R1},{SALQ}!$C${SAL_R0}:$C${SAL_R1},{m})', font=F_LINK, fmt=MONEY)
    put(ws, f'D{r}', f'=$B{r}-$C{r}', font=F_TXT, fmt=MONEY)
    put(ws, f'E{r}', f'=IFERROR($D{r}/$B{r},"")', font=F_TXT, fmt=PCT)
    put(ws, f'F{r}', f'=SUMIFS({BUYQ}!$Q${BUY_R0}:$Q${BUY_R1},{BUYQ}!$C${BUY_R0}:$C${BUY_R1},{m})', font=F_LINK, fmt=MONEY)
    ep = []
    for pat in ['"66*"', '"6403"', '"6711"']:
        ep.append(f'SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$O${CASH_R0}:$O${CASH_R1},{pat},'
                  f'{CASHQ}!$C${CASH_R0}:$C${CASH_R1},{m})')
        ep.append(f'SUMIFS({OTHQ}!$H${OTH_R0}:$H${OTH_R1},{OTHQ}!$F${OTH_R0}:$F${OTH_R1},{pat},'
                  f'{OTHQ}!$C${OTH_R0}:$C${OTH_R1},{m})')
        ep.append(f'SUMIFS({BUYQ}!$W${BUY_R0}:$W${BUY_R1},{BUYQ}!$U${BUY_R0}:$U${BUY_R1},{pat},'
                  f'{BUYQ}!$C${BUY_R0}:$C${BUY_R1},{m})')
    put(ws, f'G{r}', '=' + '+'.join(ep), font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$C${CASH_R0}:$C${CASH_R1},{m},'
                     f'{CASHQ}!$D${CASH_R0}:$D${CASH_R1},"收入")', font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$C${CASH_R0}:$C${CASH_R1},{m},'
                     f'{CASHQ}!$D${CASH_R0}:$D${CASH_R1},"支出")', font=F_TXT, fmt=MONEY)
    put(ws, f'J{r}', f'=$H{r}-$I{r}', font=F_TXT, fmt=MONEY)
    put(ws, f'K{r}', f'=SUM({Q(SH_OPEN)}!$B${OC_R0}:$B${OC_R1})+SUM($J$6:$J{r})', font=F_TXT, fmt=MONEY)
    ws.row_dimensions[r].height = 18
TR = 18
put(ws, f'A{TR}', '全年合计', font=F_TOT, fill=FILL_TOT)
for c in ['B','C','D','F','G','H','I','J']:
    put(ws, f'{c}{TR}', f'=SUM({c}6:{c}17)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'E{TR}', f'=IFERROR($D${TR}/$B${TR},"")', font=F_TOT, fill=FILL_TOT, fmt=PCT)
put(ws, f'K{TR}', f'=$K$17', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
ws.freeze_panes = 'B6'
page(ws, titles='5:5')

# ============================================================ 财务勾稽核查
ws = wb.create_sheet(SH_CHK)
title(ws, '财务勾稽核查', 'F', '月结前请确认全部为「✓ 正常」。这些都是跨表比对，不是自己和自己比，真实错账会被抓出来。')
widths(ws, {'A':4,'B':30,'C':17,'D':17,'E':15,'F':46})
headers(ws, 5, 1, ['#', '核查项目', '业务 / 明细数', '总账 / 报表数', '状态', '不一致时怎么办'])
TBQ, PLQ, BSQ, CFQ, BALQ, INVQ, ARQ, APQ = (Q(SH_TB), Q(SH_PL), Q(SH_BS), Q(SH_CF), Q(SH_BAL),
                                            Q(SH_INV), Q(SH_AR), Q(SH_AP))
def tb_k(code):
    return f'SUMIFS({TBQ}!$K${ACC_R0}:$K${ACC_R1},{TBQ}!$A${ACC_R0}:$A${ACC_R1},"{code}")'
def tb_open(code):
    return (f'SUMIFS({TBQ}!$E${ACC_R0}:$E${ACC_R1},{TBQ}!$A${ACC_R0}:$A${ACC_R1},"{code}")'
            f'-SUMIFS({TBQ}!$F${ACC_R0}:$F${ACC_R1},{TBQ}!$A${ACC_R0}:$A${ACC_R1},"{code}")')
CHECKS = [
 ('本月凭证借贷平衡', f'={TBQ}!$E$3', f'={TBQ}!$G$3',
  '每行业务自动生成一借一贷，不平说明有公式被改动或删除，请检查科目余额表的公式列。'),
 ('期初借贷平衡', f'=SUM({Q(SH_OPEN)}!$Z${OA_R0}:$Z${OA_R1})', f'=SUM({Q(SH_OPEN)}!$AA${OA_R0}:$AA${OA_R1})',
  '【期初数据】⑤会计科目期初的借方合计必须等于贷方合计，先把这里配平。'),
 ('资产负债表平衡', f'={BSQ}!$C${AR["资产总计"][0]}', f'={BSQ}!$G${LR["负债和所有者权益总计"][0]}',
  '通常是期初不平、或有科目没在【科目参数】里映射资产负债表项目。'),
 ('期初资金明细 ↔ 科目期初', f'=SUM({Q(SH_OPEN)}!$B${OC_R0}:$B${OC_R1})',
  f'={tb_open("1001")}+{tb_open("1002")}+{tb_open("1012")}',
  '【期初数据】①各账户期初余额合计，应等于⑤里 1001/1002/1012 三个科目的期初借方合计。'),
 ('期初应收明细 ↔ 科目期初', f'=SUM({Q(SH_OPEN)}!$G${OR_R0}:$G${OR_R1})', f'={tb_open("1122")}',
  '【期初数据】②应收期初合计，应等于⑤里 1122 应收账款的期初借方。'),
 ('期初应付明细 ↔ 科目期初', f'=SUM({Q(SH_OPEN)}!$L${OP_R0}:$L${OP_R1})', f'=-({tb_open("2202")})',
  '【期初数据】③应付期初合计，应等于⑤里 2202 应付账款的期初贷方。'),
 ('期初存货明细 ↔ 科目期初', f'=SUM({Q(SH_OPEN)}!$T${OG_R0}:$T${OG_R1})', f'={tb_open("1405")}',
  '【期初数据】④存货期初金额合计，应等于⑤里 1405 库存商品的期初借方。'),
 ('库存台账 ↔ 库存商品科目', f'={INVQ}!$K${GDS_R1+1}', f'={tb_k("1405")}',
  '两边都是「期初＋采购－结转成本」，不一致说明有人直接改了库存表公式，或用【其他凭证】动过 1405。'),
 ('应收跟进 ↔ 应收账款科目', f'={ARQ}!$F${CUS_R1+1}', f'={tb_k("1122")}',
  '多半是资金流水里「客户回款」的【往来单位】写了档案里没有的客户名，或客户名有多余空格。'),
 ('应付跟进 ↔ 应付账款科目', f'={APQ}!$F${SUP_R1+1}', f'=-({tb_k("2202")})',
  '多半是资金流水里「供应商付款」的【往来单位】没选档案里的供应商。'),
 ('账户余额 ↔ 货币资金科目', f'={BALQ}!$H${ACT_R1+1}',
  f'={tb_k("1001")}+{tb_k("1002")}+{tb_k("1012")}',
  '检查【基础资料】①里每个账户的「资金科目编码」是否都填了，且填的是 1001/1002/1012。'),
 ('现金流量表 ↔ 账户余额', f'={CFQ}!$D${CFR[27]}', f'={BALQ}!$H${ACT_R1+1}',
  '说明有收支科目没在【科目参数】②里配「现金流量表项目」，那笔钱进了总账但没进现金流量表。'),
 ('利润表净利润 ↔ 业务毛利勾稽', f'={Q(SH_ANA)}!$D$18', f'={PLQ}!$D${ROWMAP[1]}-{PLQ}!$D${ROWMAP[4]}',
  '经营分析的全年毛利应等于利润表的营业收入减营业成本。不等说明有销售或成本是用【其他凭证】直接做的。'),
]
r = 6
for i, (name, bf, cf_, note) in enumerate(CHECKS, 1):
    put(ws, f'A{r}', i, font=F_TXT)
    put(ws, f'B{r}', name, font=F_TXT, align=CL)
    put(ws, f'C{r}', bf, font=F_LINK, fmt=MONEY)
    put(ws, f'D{r}', cf_, font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', f'=IF(ABS($C{r}-$D{r})<0.01,"✓ 正常","✗ 差 "&TEXT($C{r}-$D{r},"#,##0.00"))',
        font=F_TOT, fill=FILL_CHK)
    put(ws, f'F{r}', note, font=F_NOTE, align=CL)
    ws.row_dimensions[r].height = 30
    r += 1
EXTRA = [
 ('录入表校验全部通过', f'=COUNTIF({CASHQ}!$M${CASH_R0}:$M${CASH_R1},"√")+COUNTIF({SALQ}!$L${SAL_R0}:$L${SAL_R1},"√")'
                       f'+COUNTIF({BUYQ}!$N${BUY_R0}:$N${BUY_R1},"√")+COUNTIF({OTHQ}!$L${OTH_R0}:$L${OTH_R1},"√")',
  f'=COUNT({CASHQ}!$B${CASH_R0}:$B${CASH_R1})+COUNT({SALQ}!$B${SAL_R0}:$B${SAL_R1})'
  f'+COUNT({BUYQ}!$B${BUY_R0}:$B${BUY_R1})+COUNT({OTHQ}!$B${OTH_R0}:$B${OTH_R1})',
  '左边是校验通过的行数，右边是已录入的业务行数，两者必须相等。不等说明有行报错，按各表【校验】列的红色提示逐行修正。'),
 ('查询月份有业务数据', f'=COUNTIF({CASHQ}!$C${CASH_R0}:$C${CASH_R1},{YM_CELL})'
                       f'+COUNTIF({SALQ}!$C${SAL_R0}:$C${SAL_R1},{YM_CELL})'
                       f'+COUNTIF({BUYQ}!$C${BUY_R0}:$C${BUY_R1},{YM_CELL})'
                       f'+COUNTIF({OTHQ}!$C${OTH_R0}:$C${OTH_R1},{YM_CELL})', '=0',
  '本项「相等」反而是异常：说明查询月份在四张录入表里一笔业务都没有，此时所有报表都是 0，不要误以为账是平的。'),
 ('损益科目未录期初', f'=SUMIFS({Q(SH_OPEN)}!$Z${OA_R0}:$Z${OA_R1},{Q(SH_OPEN)}!$W${OA_R0}:$W${OA_R1},">=5000")'
                     f'+SUMIFS({Q(SH_OPEN)}!$AA${OA_R0}:$AA${OA_R1},{Q(SH_OPEN)}!$W${OA_R0}:$W${OA_R1},">=5000")', '=0',
  '5xxx / 6xxx 损益类科目不能录期初余额。年中启用时，本年已发生的收入费用请到【其他凭证】按月补录。'),
]
for i, (name, bf, cf_, note) in enumerate(EXTRA, len(CHECKS) + 1):
    put(ws, f'A{r}', i, font=F_TXT)
    put(ws, f'B{r}', name, font=F_TXT, align=CL)
    put(ws, f'C{r}', bf, font=F_LINK, fmt='#,##0')
    put(ws, f'D{r}', cf_, font=F_LINK, fmt='#,##0')
    if i == len(CHECKS) + 2:
        put(ws, f'E{r}', f'=IF($C{r}>0,"✓ 正常","✗ 本月无数据")', font=F_TOT, fill=FILL_CHK)
    else:
        put(ws, f'E{r}', f'=IF(ABS($C{r}-$D{r})<0.01,"✓ 正常","✗ 有 "&TEXT(ABS($C{r}-$D{r}),"0")&" 处")',
            font=F_TOT, fill=FILL_CHK)
    put(ws, f'F{r}', note, font=F_NOTE, align=CL)
    ws.row_dimensions[r].height = 30
    r += 1
CHK_R0, CHK_R1 = 6, r - 1
SUM_R = r + 1
put(ws, f'B{SUM_R}', '总  体  结  论', font=F_TOT, fill=FILL_TOT)
put(ws, f'C{SUM_R}', f'=COUNTIF($E${CHK_R0}:$E${CHK_R1},"✓*")&" / {CHK_R1-CHK_R0+1} 项通过"',
    font=F_TOT, fill=FILL_TOT); ws.merge_cells(f'C{SUM_R}:D{SUM_R}')
put(ws, f'D{SUM_R}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'E{SUM_R}', f'=IF(COUNTIF($E${CHK_R0}:$E${CHK_R1},"✗*")=0,"✓ 可以月结","✗ 不可月结")',
    font=F_TOT, fill=FILL_CHK)
put(ws, f'F{SUM_R}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'A{SUM_R}', None, font=F_TOT, fill=FILL_TOT)
ws.conditional_formatting.add(f'E{CHK_R0}:E{SUM_R}',
    FormulaRule(formula=[f'LEFT($E{CHK_R0},1)="✗"'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.freeze_panes = 'A6'
page(ws, titles='5:5')
print('  ✓ 经营分析 / 财务勾稽核查')

# ============================================================ 首页
ws = wb.create_sheet(SH_HOME)
ws.sheet_view.showGridLines = False
widths(ws, {'A':13,'B':13,'C':3,'D':13,'E':13,'F':3,'G':13,'H':13,'I':3,'J':13,'K':13,'L':3})
ws.merge_cells('A1:L1')
put(ws, 'A1', '内账财务管理系统', font=Font(name='微软雅黑', size=20, bold=True, color='FFFFFF'),
    fill=PatternFill('solid', fgColor='1F3864'), align=CL, border=None)
ws.row_dimensions[1].height = 42
ws.merge_cells('A2:L2')
put(ws, 'A2', '　日常只录四张表：资金流水 · 销售录入 · 采购录入 · 其他凭证　→　总账与全部报表自动生成',
    font=Font(name='微软雅黑', size=9, color='FFFFFF'),
    fill=PatternFill('solid', fgColor='2F5597'), align=CL, border=None)
ws.row_dimensions[2].height = 20

put(ws, 'A3', '会计年度', font=F_H2, fill=FILL_HDR2)
put(ws, 'B3', 2026, font=Font(name='微软雅黑', size=12, bold=True, color='0000C0'), fill=FILL_IN, fmt='0"年"')
put(ws, 'D3', '查询月份', font=F_H2, fill=FILL_HDR2)
import datetime
put(ws, 'E3', datetime.datetime(2026, 1, 1), font=Font(name='微软雅黑', size=12, bold=True, color='0000C0'),
    fill=FILL_IN, fmt='yyyy年m月')
put(ws, 'G3', '← 全表只有这两格需要设置。跨年度把会计年度改成新的一年，所有报表自动切换，不用改任何公式。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells('G3:L3')
ws.row_dimensions[3].height = 26
put(ws, 'A5', '本月月份', font=F_NOTE, align=CR, border=None)
put(ws, 'B5', '=TEXT($E$3,"yyyy-mm")', font=F_NOTE, border=None)
put(ws, 'D5', '本月月初', font=F_NOTE, align=CR, border=None)
put(ws, 'E5', '=DATE(YEAR($E$3),MONTH($E$3),1)', font=F_NOTE, fmt=DATE, border=None)
put(ws, 'G5', '本月月末', font=F_NOTE, align=CR, border=None)
put(ws, 'H5', '=EOMONTH($E$3,0)', font=F_NOTE, fmt=DATE, border=None)
put(ws, 'B6', '=$E$5', font=F_NOTE, border=None); put(ws, 'B7', '=$H$5', font=F_NOTE, border=None)
ws.column_dimensions['B'].hidden = False
for rr in (5, 6, 7):
    ws.row_dimensions[rr].height = 13
ws.row_dimensions[6].hidden = True
ws.row_dimensions[7].hidden = True

CARDS = [
 ('本月营业收入', f'={PLQ}!$C${ROWMAP[1]}', MONEY, 'D6E4F0'),
 ('本月毛利',     f'={PLQ}!$C${ROWMAP[1]}-{PLQ}!$C${ROWMAP[4]}', MONEY, 'D6E4F0'),
 ('本月净利润',   f'={PLQ}!$C${ROWMAP[22]}', MONEY, 'D6E4F0'),
 ('本年累计净利润', f'={PLQ}!$D${ROWMAP[22]}', MONEY, 'D6E4F0'),
 ('货币资金余额', f'={BALQ}!$H${ACT_R1+1}', MONEY, 'E2EFDA'),
 ('应收账款余额', f'={ARQ}!$F${CUS_R1+1}', MONEY, 'E2EFDA'),
 ('应付账款余额', f'={APQ}!$F${SUP_R1+1}', MONEY, 'E2EFDA'),
 ('存货金额',     f'={INVQ}!$K${GDS_R1+1}', MONEY, 'E2EFDA'),
 ('逾期应收',     f'={ARQ}!$L${CUS_R1+1}', MONEY, 'FCE4E4'),
 ('本月经营现金净额', f'={CFQ}!$C${CFR[9]}', MONEY, 'FCE4E4'),
 ('资产总计',     f'={BSQ}!$C${AR["资产总计"][0]}', MONEY, 'FFF2CC'),
 ('所有者权益',   f'={BSQ}!$G${LR["所有者权益合计"][0]}', MONEY, 'FFF2CC'),
]
row = 9
for i, (lab, f, fmt, color) in enumerate(CARDS):
    c0 = 1 + (i % 4) * 3
    r = row + (i // 4) * 3
    a, b = L(c0), L(c0 + 1)
    ws.merge_cells(f'{a}{r}:{b}{r}')
    put(ws, f'{a}{r}', lab, font=Font(name='微软雅黑', size=9, bold=True, color='404040'),
        fill=PatternFill('solid', fgColor=color), align=CL, border=None)
    put(ws, f'{b}{r}', None, fill=PatternFill('solid', fgColor=color), border=None)
    ws.merge_cells(f'{a}{r+1}:{b}{r+1}')
    put(ws, f'{a}{r+1}', f, font=F_BIG, fill=PatternFill('solid', fgColor=color), align=CR, fmt=fmt, border=None)
    put(ws, f'{b}{r+1}', None, fill=PatternFill('solid', fgColor=color), border=None)
    ws.row_dimensions[r].height = 18
    ws.row_dimensions[r + 1].height = 30
    ws.row_dimensions[r + 2].height = 6

STAT_R = row + 9
ws.merge_cells(f'A{STAT_R}:B{STAT_R}')
put(ws, f'A{STAT_R}', '月结状态', font=F_H2, fill=FILL_HDR2, align=CL)
put(ws, f'B{STAT_R}', None, fill=FILL_HDR2)
ws.merge_cells(f'D{STAT_R}:L{STAT_R}')
put(ws, f'D{STAT_R}', f'={Q(SH_CHK)}!$E${SUM_R}&"　（"&{Q(SH_CHK)}!$C${SUM_R}&"）　"'
                      f'&IF({Q(SH_CHK)}!$E${SUM_R}="✓ 可以月结","全部勾稽通过，可以结账。","请到【财务勾稽核查】查看未通过的项目。")',
    font=F_TOT, fill=FILL_CHK, align=CL)
ws.row_dimensions[STAT_R].height = 26

NAV_R = STAT_R + 2
ws.merge_cells(f'A{NAV_R}:L{NAV_R}')
put(ws, f'A{NAV_R}', '　每月怎么用', font=F_HDR, fill=FILL_HDR, align=CL, border=None)
ws.row_dimensions[NAV_R].height = 22
NAV = [
 ('第 1 步', '启用时先做一次', '【科目参数】核对科目和收支科目 →【基础资料】建档案 →【期初数据】录期初余额'),
 ('第 2 步', '平时随时录', '收付款记【资金流水】；卖出记【销售录入】；进货记【采购录入】'),
 ('第 3 步', '月末做一次', '工资社保计提、折旧摊销、税金计提记【其他凭证】'),
 ('第 4 步', '月末检查', '把【首页】查询月份切到本月 →【财务勾稽核查】全部✓ → 打印报表'),
 ('第 5 步', '年末', '【其他凭证】做结转损益，次年在新文件的【期初数据】录新年度期初'),
]
r = NAV_R + 1
for a, b, c in NAV:
    put(ws, f'A{r}', a, font=F_TOT, fill=FILL_HDR2)
    ws.merge_cells(f'B{r}:C{r}')
    put(ws, f'B{r}', b, font=F_TXT, align=CL)
    put(ws, f'C{r}', None)
    ws.merge_cells(f'D{r}:L{r}')
    put(ws, f'D{r}', c, font=F_TXT, align=CL)
    for cc in 'EFGHIJKL': put(ws, f'{cc}{r}', None)
    ws.row_dimensions[r].height = 20
    r += 1
dv_num(ws, 'B3', 'between', '2000')
ws.page_setup.orientation = 'portrait'
ws.page_setup.paperSize = 9
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0

# ============================================================ 操作说明
ws = wb.create_sheet(SH_HELP)
ws.sheet_view.showGridLines = False
widths(ws, {'A':3,'B':22,'C':100})
ws.merge_cells('B2:C2')
put(ws, 'B2', '操作说明', font=F_TITLE, align=CL, border=None)
ws.row_dimensions[2].height = 28
HELP = [
 ('一、这套账的设计', None),
 ('只有四个录入口', '资金流水（实际收付的钱）、销售录入（卖出去的东西）、采购录入（买进来的东西）、其他凭证（不涉及收付款的账，如计提折旧）。除此之外所有表都是自动算出来的，不要手工改。'),
 ('每行自动生成分录', '每一行业务都会自动算出借方科目和贷方科目，金额相同。所以借贷永远相等，资产负债表永远平——除非期初就没配平。'),
 ('没有写死的科目和年份', '科目表可以随便加，年份改【首页】的会计年度就行，所有报表自动跟着走，不需要改任何公式。'),
 ('二、启用步骤（只做一次）', None),
 ('1. 核对科目参数', '【科目参数】里已经内置了小企业会计准则的全套科目。行业特殊的科目直接在下面加行；用不到的把「启用」改成停用（不要删行）。'),
 ('2. 配置收支科目', '【科目参数】②是资金流水下拉的来源。每一条要配「对方科目」——就是这笔钱的另一半。比如「房租物业」对方科目是 6602 管理费用，「客户回款」对方科目是 1122 应收账款。行业不同，主要就改这张表。'),
 ('3. 建立基础资料', '【基础资料】里录账户、部门、员工、商品、客户、供应商。停用请改「状态」，不要删整行。'),
 ('4. 录入期初', '【期初数据】录启用那一天的余额。注意两件事：⑤会计科目期初必须借贷相等；5xxx/6xxx 损益类科目不要录期初。'),
 ('5. 核对期初', '到【财务勾稽核查】看第 2～7 项是不是全部✓。期初没配平，后面所有报表都不会平。'),
 ('三、年中启用怎么办', None),
 ('期初只录资产负债', '比如 8 月才开始用，【期初数据】录的是 7 月 31 日的资产、负债、所有者权益余额，不要把 1～7 月的收入费用录进去。'),
 ('本年已发生损益', '1～7 月的收入和费用，到【其他凭证】按月各做一笔汇总分录补录（业务类型选「年初建账」）。这样利润表的本年累计才是完整的，资产负债表也不会不平。'),
 ('四、每月流程', None),
 ('平时', '收到钱、付出钱 → 资金流水；卖货开单 → 销售录入；进货收料 → 采购录入。每张表最右边的【校验】列必须是「√」，红色提示必须当天改掉。'),
 ('月末', '工资计提、社保计提、折旧、摊销、税金计提 → 其他凭证。然后把【首页】的查询月份切到本月。'),
 ('结账前', '【财务勾稽核查】必须全部✓。第 14 项「查询月份有业务数据」如果报错，说明这个月一笔账都没有，此时报表全是 0，不要误以为账平了。'),
 ('五、常见问题', None),
 ('资产负债表不平', '按顺序查勾稽核查的第 2 项（期初借贷）、第 3 项（表间平衡）。九成是期初没配平，或某个科目在【科目参数】里没填「资产负债表项目」。'),
 ('应收和总账对不上', '资金流水里「客户回款」的【往来单位】必须从下拉里选，手打的名字或多一个空格，SUMIFS 就匹配不上，钱进了总账却没冲掉这个客户的应收。'),
 ('现金流量表和银行对不上', '某个收支科目在【科目参数】②里没配「现金流量表项目」。查勾稽核查第 12 项。'),
 ('销售成本不对', '成本单价是按「期初 + 该笔销售日期之前的入库」算移动加权平均。如果这个商品还没有采购记录，会退回用【基础资料】里的参考成本；参考成本也没填就会是 0，校验列会提示。'),
 ('要加行怎么办', '各表都预留了行数（资金流水 500 行、销售采购各 400 行）。不够时请在数据区中间插入整行再复制公式，不要在最后一行下面直接粘贴，那会盖掉合计行。'),
 ('六、注意', None),
 ('不要删整行', '基础资料、期初数据都是横向分区共用行号，删整行会同时删掉旁边几个区的数据。要清空请选中单元格按 Delete。'),
 ('不要动白色和灰色区域', '黄色是录入区，白色/灰色是公式区。公式被覆盖后不会有任何报错，但报表会静默算错。'),
 ('年内不要做结转损益', '损益科目的余额本身就是本年利润，资产负债表已经自动接住了。结转损益只在 12 月做一次。'),
]
r = 4
for a, b in HELP:
    if b is None:
        ws.merge_cells(f'B{r}:C{r}')
        put(ws, f'B{r}', a, font=F_HDR, fill=FILL_HDR, align=CL)
        put(ws, f'C{r}', None, fill=FILL_HDR)
        ws.row_dimensions[r].height = 22
    else:
        put(ws, f'B{r}', a, font=F_TOT, fill=FILL_HDR2, align=CL)
        put(ws, f'C{r}', b, font=F_TXT, align=CL)
        ws.row_dimensions[r].height = 34
    r += 1
page(ws, titles=None, landscape=False)


# ============================================================ 固定资产台账
ws = wb.create_sheet(SH_FA)
R0, R1 = FA_R0, FA_R1
title(ws, '固定资产台账', 'N',
      '本表自动算出每月应提折旧。月末请把最下面的「本月折旧合计」按使用部门在【其他凭证】做一笔计提分录（借 管理费用/销售费用 贷 累计折旧）。')
widths(ws, {'A':13,'B':20,'C':12,'D':11,'E':12,'F':14,'G':9,'H':8,'I':13,'J':9,'K':14,'L':14,'M':9,'N':16})
put(ws, 'A3', '查询月份', font=F_H2, align=CR, border=None)
put(ws, 'B3', f'={YM_CELL}', font=F_TOT, fill=FILL_TOT)
inband(ws, 'A', 'H', 4); sysband(ws, 'I', 'N', 4)
headers(ws, 5, 1, ['资产编号','资产名称','类别','使用部门','购入日期','原值','使用\n年限','残值率'])
headers(ws, 5, 9, ['月折旧额','已提\n月数','累计折旧','账面净值','状态','备注'], fill=FILL_AUTO, font=F_HDR2)
for r in range(R0, R1 + 1):
    for c, fmt, tx in [('A', None, 0), ('B', None, 1), ('C', None, 0), ('D', None, 0),
                       ('E', DATE, 0), ('F', MONEY, 0), ('G', '0', 0), ('H', PCT, 0)]:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if tx else C, fmt=fmt)
    put(ws, f'I{r}', f'=IF(OR($B{r}="",N($G{r})=0),"",ROUND(N($F{r})*(1-N($H{r}))/N($G{r})/12,2))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'J{r}', f'=IF(OR($B{r}="",$E{r}=""),"",MEDIAN(0,(YEAR({EOM_CELL})-YEAR($E{r}))*12'
                     f'+MONTH({EOM_CELL})-MONTH($E{r}),N($G{r})*12))', font=F_LINK, fill=FILL_AUTO, fmt='0')
    put(ws, f'K{r}', f'=IF($B{r}="","",ROUND(N($I{r})*N($J{r}),2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'L{r}', f'=IF($B{r}="","",N($F{r})-N($K{r}))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'M{r}', f'=IF($B{r}="","",IF(N($J{r})>=N($G{r})*12,"已提完","计提中"))', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'N{r}', None, font=F_IN, fill=FILL_IN, align=CL)
    ws.row_dimensions[r].height = 16
TR = R1 + 1
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in ['B','C','D','E','G','H','J','M','N']: put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in ['F','K','L']:
    put(ws, f'{c}{TR}', f'=SUM({c}{R0}:{c}{R1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'I{TR}', f'=SUMIF($M{R0}:$M{R1},"计提中",$I{R0}:$I{R1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'A{TR+2}', '本月应计提折旧合计 →', font=F_H2, align=CR, border=None)
ws.merge_cells(f'A{TR+2}:C{TR+2}')
put(ws, f'D{TR+2}', f'=$I${TR}', font=F_BIG, fill=FILL_CHK, fmt=MONEY)
ws.merge_cells(f'D{TR+2}:E{TR+2}'); put(ws, f'E{TR+2}', None, fill=FILL_CHK)
put(ws, f'F{TR+2}', '请到【其他凭证】做：业务类型「折旧计提」，借方 6602 管理费用（或按部门用 6601 销售费用），贷方 1602 累计折旧。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'F{TR+2}:N{TR+2}')
dv_list(ws, f'D{R0}:D{R1}', f'={B_DEP_N}')
dv_list(ws, f'C{R0}:C{R1}', '"房屋建筑物,机器设备,运输工具,电子设备,办公家具,其他"')
dv_num(ws, f'F{R0}:F{R1}'); dv_num(ws, f'G{R0}:G{R1}', 'greaterThan', '0')
ws.freeze_panes = 'C6'; page(ws, titles='5:5')

# ============================================================ 工资表
ws = wb.create_sheet(SH_PAYROLL)
R0, R1 = PAY_R0, PAY_R1
title(ws, '工资表', 'N',
      '计提用最右边的「单位负担合计」，发放用「实发工资」。两笔分录分别在【其他凭证】（计提）和【资金流水】（发放）登记。')
widths(ws, {'A':12,'B':11,'C':13,'D':12,'E':12,'F':14,'G':12,'H':12,'I':11,'J':14,'K':12,'L':12,'M':15,'N':14})
put(ws, 'A3', '工资月份', font=F_H2, align=CR, border=None)
put(ws, 'B3', f'={YM_CELL}', font=F_TOT, fill=FILL_TOT)
inband(ws, 'A', 'I', 4); sysband(ws, 'J', 'N', 4)
headers(ws, 5, 1, ['姓名','部门','基本工资','绩效奖金','补贴津贴','应发合计','社保\n(个人)','公积金\n(个人)','个人所得税'])
headers(ws, 5, 10, ['实发工资','社保\n(单位)','公积金\n(单位)','单位负担合计','备注'], fill=FILL_AUTO, font=F_HDR2)
for r in range(R0, R1 + 1):
    put(ws, f'A{r}', None, font=F_IN, fill=FILL_IN)
    put(ws, f'B{r}', f'=IF($A{r}="","",{lookup(f"$A{r}", B_EMP_N, f"{BASE}!$L${EMP_R0}:$L${EMP_R1}")})', font=F_LINK)
    for c in ['C','D','E','G','H','I','K','L']:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",N($C{r})+N($D{r})+N($E{r}))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",N($F{r})-N($G{r})-N($H{r})-N($I{r}))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'M{r}', f'=IF($A{r}="","",N($F{r})+N($K{r})+N($L{r}))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'N{r}', None, font=F_IN, fill=FILL_IN, align=CL)
    ws.row_dimensions[r].height = 16
TR = R1 + 1
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT); put(ws, f'B{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in ['C','D','E','F','G','H','I','J','K','L','M']:
    put(ws, f'{c}{TR}', f'=SUM({c}{R0}:{c}{R1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'N{TR}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'A{TR+2}', '① 月末计提：借 6602 管理费用（按部门也可用 6601）', font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{TR+2}:E{TR+2}')
put(ws, f'F{TR+2}', f'=$M${TR}', font=F_TOT, fill=FILL_CHK, fmt=MONEY)
put(ws, f'G{TR+2}', '贷 2211 应付职工薪酬（同额）→ 记在【其他凭证】，业务类型选「工资计提」',
    font=F_NOTE, align=CL, border=None); ws.merge_cells(f'G{TR+2}:N{TR+2}')
put(ws, f'A{TR+3}', '② 实际发放：借 2211 应付职工薪酬', font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{TR+3}:E{TR+3}')
put(ws, f'F{TR+3}', f'=$J${TR}', font=F_TOT, fill=FILL_CHK, fmt=MONEY)
put(ws, f'G{TR+3}', '贷 银行存款 → 记在【资金流水】，收支科目选「工资发放」。代扣的社保个税单独在缴纳时登记。',
    font=F_NOTE, align=CL, border=None); ws.merge_cells(f'G{TR+3}:N{TR+3}')
dv_list(ws, f'A{R0}:A{R1}', f'={B_EMP_N}')
for c in ['C','D','E','G','H','I','K','L']: dv_num(ws, f'{c}{R0}:{c}{R1}')
ws.freeze_panes = 'C6'; page(ws, titles='5:5')
print('  ✓ 固定资产台账 / 工资表')

# ============================================================ 辅助核算
SH_AUX2 = '辅助核算'
ws = wb.create_sheet(SH_AUX2)
title(ws, '辅助核算', 'AA', '客户、供应商、部门、员工四个维度的本月汇总。数据全部来自录入表，与总账同源。')
widths(ws, {'A':16,'B':13,'C':13,'D':13,'E':13,'F':10,'G':3,
            'H':16,'I':13,'J':13,'K':13,'L':13,'M':10,'N':3,
            'O':12,'P':13,'Q':13,'R':13,'S':13,'T':13,'U':3,
            'V':11,'W':11,'X':13,'Y':13,'Z':13,'AA':13})
put(ws, 'A3', '查询月份', font=F_H2, align=CR, border=None)
put(ws, 'B3', f'={YM_CELL}', font=F_TOT, fill=FILL_TOT)
block(ws, 1, 6, 5, '① 客户', ['客户','本月销售','本月回款','期末应收','逾期应收','业务员'], CUS_R0, CUS_R1)
block(ws, 8, 6, 5, '② 供应商', ['供应商','本月采购','本月付款','期末应付','逾期应付','采购员'], SUP_R0, SUP_R1)
block(ws, 15, 6, 5, '③ 部门', ['部门','本月销售额','本月销售成本','本月毛利','本月期间费用','部门利润'], DEP_R0, DEP_R1)
block(ws, 22, 6, 5, '④ 员工', ['姓名','部门','经办销售额','经办采购额','经办收款','经办付款'], EMP_R0, EMP_R1)
ARQ2, APQ2 = Q(SH_AR), Q(SH_AP)
for r in range(CUS_R0, CUS_R1 + 1):
    a = f'$A{r}'
    put(ws, f'A{r}', f'=IF({ARQ2}!$A{r}="","",{ARQ2}!$A{r})', font=F_LINK, align=CL)
    for col, src in [('B', 'D'), ('C', 'E'), ('D', 'F'), ('E', 'L'), ('F', 'N')]:
        put(ws, f'{col}{r}', f'=IF({a}="","",{ARQ2}!${src}{r})', font=F_LINK,
            fmt=None if col == 'F' else MONEY)
    ws.row_dimensions[r].height = 16
for r in range(SUP_R0, SUP_R1 + 1):
    a = f'$H{r}'
    put(ws, f'H{r}', f'=IF({APQ2}!$A{r}="","",{APQ2}!$A{r})', font=F_LINK, align=CL)
    for col, src in [('I', 'D'), ('J', 'E'), ('K', 'F'), ('L', 'L'), ('M', 'N')]:
        put(ws, f'{col}{r}', f'=IF({a}="","",{APQ2}!${src}{r})', font=F_LINK,
            fmt=None if col == 'M' else MONEY)
    ws.row_dimensions[r].height = 16
# 科目编码是文本，用 ">=6400" 会被当数字比而永不匹配，必须用通配符
EXP_PATS = ['"66*"', '"6403"', '"6711"']
for r in range(DEP_R0, DEP_R1 + 1):
    a = f'$O{r}'
    put(ws, f'O{r}', f'=IF({BASE}!$G{r}="","",{BASE}!$G{r})', font=F_LINK)
    put(ws, f'P{r}', f'=IF({a}="","",SUMIFS({SALQ}!$O${SAL_R0}:$O${SAL_R1},{SALQ}!$I${SAL_R0}:$I${SAL_R1},{a},'
                     f'{SALQ}!$C${SAL_R0}:$C${SAL_R1},{YM_CELL}))', font=F_LINK, fmt=MONEY)
    put(ws, f'Q{r}', f'=IF({a}="","",SUMIFS({SALQ}!$S${SAL_R0}:$S${SAL_R1},{SALQ}!$I${SAL_R0}:$I${SAL_R1},{a},'
                     f'{SALQ}!$C${SAL_R0}:$C${SAL_R1},{YM_CELL}))', font=F_LINK, fmt=MONEY)
    put(ws, f'R{r}', f'=IF({a}="","",$P{r}-$Q{r})', font=F_TXT, fmt=MONEY)
    exp_terms = []
    for pat in EXP_PATS:
        exp_terms.append(f'SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$H${CASH_R0}:$H${CASH_R1},{a},'
                         f'{CASHQ}!$O${CASH_R0}:$O${CASH_R1},{pat},{CASHQ}!$C${CASH_R0}:$C${CASH_R1},{YM_CELL})')
        exp_terms.append(f'SUMIFS({OTHQ}!$H${OTH_R0}:$H${OTH_R1},{OTHQ}!$I${OTH_R0}:$I${OTH_R1},{a},'
                         f'{OTHQ}!$F${OTH_R0}:$F${OTH_R1},{pat},{OTHQ}!$C${OTH_R0}:$C${OTH_R1},{YM_CELL})')
        exp_terms.append(f'SUMIFS({BUYQ}!$W${BUY_R0}:$W${BUY_R1},{BUYQ}!$K${BUY_R0}:$K${BUY_R1},{a},'
                         f'{BUYQ}!$U${BUY_R0}:$U${BUY_R1},{pat},{BUYQ}!$C${BUY_R0}:$C${BUY_R1},{YM_CELL})')
    put(ws, f'S{r}', f'=IF({a}="","",' + '+'.join(exp_terms) + ')', font=F_LINK, fmt=MONEY)
    put(ws, f'T{r}', f'=IF({a}="","",$R{r}-$S{r})', font=F_TOT, fmt=MONEY)
    ws.row_dimensions[r].height = 16
for r in range(EMP_R0, EMP_R1 + 1):
    a = f'$V{r}'
    put(ws, f'V{r}', f'=IF({BASE}!$K{r}="","",{BASE}!$K{r})', font=F_LINK)
    put(ws, f'W{r}', f'=IF({a}="","",{BASE}!$L{r})', font=F_LINK)
    put(ws, f'X{r}', f'=IF({a}="","",SUMIFS({SALQ}!$O${SAL_R0}:$O${SAL_R1},{SALQ}!$J${SAL_R0}:$J${SAL_R1},{a},'
                     f'{SALQ}!$C${SAL_R0}:$C${SAL_R1},{YM_CELL}))', font=F_LINK, fmt=MONEY)
    put(ws, f'Y{r}', f'=IF({a}="","",SUMIFS({BUYQ}!$Q${BUY_R0}:$Q${BUY_R1},{BUYQ}!$L${BUY_R0}:$L${BUY_R1},{a},'
                     f'{BUYQ}!$C${BUY_R0}:$C${BUY_R1},{YM_CELL}))', font=F_LINK, fmt=MONEY)
    put(ws, f'Z{r}', f'=IF({a}="","",SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$I${CASH_R0}:$I${CASH_R1},{a},'
                     f'{CASHQ}!$D${CASH_R0}:$D${CASH_R1},"收入",{CASHQ}!$C${CASH_R0}:$C${CASH_R1},{YM_CELL}))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'AA{r}', f'=IF({a}="","",SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$I${CASH_R0}:$I${CASH_R1},{a},'
                      f'{CASHQ}!$D${CASH_R0}:$D${CASH_R1},"支出",{CASHQ}!$C${CASH_R0}:$C${CASH_R1},{YM_CELL}))',
        font=F_LINK, fmt=MONEY)
    ws.row_dimensions[r].height = 16
for c0, r1, cols in [('A', CUS_R1, ['B','C','D','E']), ('H', SUP_R1, ['I','J','K','L']),
                     ('O', DEP_R1, ['P','Q','R','S','T']), ('V', EMP_R1, ['X','Y','Z','AA'])]:
    TR = r1 + 1
    put(ws, f'{c0}{TR}', '合计', font=F_TOT, fill=FILL_TOT)
    for c in cols:
        put(ws, f'{c}{TR}', f'=SUM({c}{6}:{c}{r1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
ws.freeze_panes = 'A6'; page(ws, titles='5:5')

# ============================================================ 科目分析
SH_ACA = '科目分析'
ws = wb.create_sheet(SH_ACA)
title(ws, '科目分析（对方科目）', 'F',
      '选一个科目，看它本月的钱是从哪些科目来、又到哪些科目去。查账定位最快的一张表。')
widths(ws, {'A':13,'B':22,'C':17,'D':17,'E':14,'F':40})
put(ws, 'A3', '选择科目', font=F_H2, fill=FILL_HDR2)
put(ws, 'B3', '1122', font=Font(name='微软雅黑', size=11, bold=True, color='0000C0'), fill=FILL_IN)
put(ws, 'C3', f'=IFERROR(INDEX({PAR_ACC_N},MATCH($B$3,{PAR_ACC_C},0)),"⚠ 科目不存在")', font=F_TOT, fill=FILL_TOT, align=CL)
put(ws, 'D3', '查询月份', font=F_NOTE, align=CR, border=None)
put(ws, 'E3', f'={YM_CELL}', font=F_TOT, fill=FILL_TOT)
dv_list(ws, 'B3', f'={PAR_ACC_C}')
headers(ws, 5, 1, ['对方科目','对方科目名称','本月借方\n(本科目增加)','本月贷方\n(本科目减少)','净额','说明'])
def pair(side, other_ref):
    """side='D'：本科目在借方、对方在贷方"""
    ts = []
    for sh, amt, dr, cr, mc, dc, r0, r1 in SLOTS:
        me, ot = (dr, cr) if side == 'D' else (cr, dr)
        ts.append(f'SUMIFS({sh}!${amt}${r0}:${amt}${r1},{sh}!${me}${r0}:${me}${r1},$B$3,'
                  f'{sh}!${ot}${r0}:${ot}${r1},{other_ref},{sh}!${mc}${r0}:${mc}${r1},{YM_CELL})')
    return '+'.join(ts)
for r in range(ACC_R0, ACC_R1 + 1):
    a = f'$A{r}'
    put(ws, f'A{r}', f'=IF({PARAM}!$A{r}="","",{PARAM}!$A{r})', font=F_LINK)
    put(ws, f'B{r}', f'=IF({a}="","",{PARAM}!$B{r})', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF({a}="","",{pair("D", a)})', font=F_LINK, fmt=MONEY)
    put(ws, f'D{r}', f'=IF({a}="","",{pair("C", a)})', font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', f'=IF({a}="","",$C{r}-$D{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'F{r}', f'=IF(OR({a}="",AND($C{r}=0,$D{r}=0)),"",'
                     f'IF($C{r}>0,"借 "&$C$3&" / 贷 "&$B{r},"")&IF(AND($C{r}>0,$D{r}>0)," ；","")'
                     f'&IF($D{r}>0,"借 "&$B{r}&" / 贷 "&$C$3,""))', font=F_TXT, align=CL)
    ws.row_dimensions[r].height = 16
TR = ACC_R1 + 1
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT); put(ws, f'B{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in ['C','D','E']:
    put(ws, f'{c}{TR}', f'=SUM({c}{ACC_R0}:{c}{ACC_R1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'F{TR}', None, font=F_TOT, fill=FILL_TOT)
ws.freeze_panes = 'A6'; page(ws, titles='5:5')
print('  ✓ 辅助核算 / 科目分析')

# ============================================================ 往来对账单
SH_SOA1, SH_SOA2 = '客户对账单', '供应商对账单'
SOA_N = 100   # 明细行数

def add_seq(sheet, col, cond):
    ws2 = wb[sheet]
    r0, r1 = {SH_SAL: (SAL_R0, SAL_R1), SH_BUY: (BUY_R0, BUY_R1), SH_CASH: (CASH_R0, CASH_R1)}[sheet]
    put(ws2, f'{col}4', '对账\n序号', font=F_HDR2, fill=FILL_AUTO)
    ws2.column_dimensions[col].width = 8
    for r in range(r0, r1 + 1):
        put(ws2, f'{col}{r}', cond(r), font=F_LINK, fill=FILL_AUTO)

S1, S2 = Q(SH_SOA1), Q(SH_SOA2)
add_seq(SH_SAL, 'AE', lambda r: (
    f'=IF(AND($D{r}={S1}!$B$3,$B{r}>={S1}!$E$3,$B{r}<={S1}!$H$3),'
    f'COUNTIFS($D${SAL_R0}:$D{r},{S1}!$B$3,$B${SAL_R0}:$B{r},">="&{S1}!$E$3,$B${SAL_R0}:$B{r},"<="&{S1}!$H$3),"")'))
add_seq(SH_CASH, 'V', lambda r: (
    f'=IF(AND($G{r}={S1}!$B$3,$E{r}="客户回款",$B{r}>={S1}!$E$3,$B{r}<={S1}!$H$3),'
    f'COUNTIFS($G${CASH_R0}:$G{r},{S1}!$B$3,$E${CASH_R0}:$E{r},"客户回款",'
    f'$B${CASH_R0}:$B{r},">="&{S1}!$E$3,$B${CASH_R0}:$B{r},"<="&{S1}!$H$3),"")'))
add_seq(SH_BUY, 'AA', lambda r: (
    f'=IF(AND($D{r}={S2}!$B$3,$B{r}>={S2}!$E$3,$B{r}<={S2}!$H$3),'
    f'COUNTIFS($D${BUY_R0}:$D{r},{S2}!$B$3,$B${BUY_R0}:$B{r},">="&{S2}!$E$3,$B${BUY_R0}:$B{r},"<="&{S2}!$H$3),"")'))
add_seq(SH_CASH, 'W', lambda r: (
    f'=IF(AND($G{r}={S2}!$B$3,$E{r}="供应商付款",$B{r}>={S2}!$E$3,$B{r}<={S2}!$H$3),'
    f'COUNTIFS($G${CASH_R0}:$G{r},{S2}!$B$3,$E${CASH_R0}:$E{r},"供应商付款",'
    f'$B${CASH_R0}:$B{r},">="&{S2}!$E$3,$B${CASH_R0}:$B{r},"<="&{S2}!$H$3),"")'))

def build_soa(sheet, is_ar):
    ws = wb.create_sheet(sheet)
    who = '客户' if is_ar else '供应商'
    biz, seq_b, party_src = ((SALQ, 'AE', B_CUS_N) if is_ar else (BUYQ, 'AA', B_SUP_N))
    br0, br1 = (SAL_R0, SAL_R1) if is_ar else (BUY_R0, BUY_R1)
    seq_c = 'V' if is_ar else 'W'
    op_party, op_date, op_amt, o0, o1 = (('E', 'F', 'G', OR_R0, OR_R1) if is_ar
                                         else ('J', 'K', 'L', OP_R0, OP_R1))
    amt_c = 'Q' if is_ar else 'S'
    title(ws, f'{who}对账单', 'M',
          f'选好{who}和起止日期即可打印。左边是本期业务明细，右边是本期{"收款" if is_ar else "付款"}明细，'
          f'上方四个数是对账结论。')
    widths(ws, {'A':12,'B':18,'C':9,'D':11,'E':13,'F':12,'G':3,'H':12,'I':13,'J':13,'K':20,'L':11,'M':11})
    put(ws, 'A3', who, font=F_H2, fill=FILL_HDR2)
    put(ws, 'B3', None, font=Font(name='微软雅黑', size=11, bold=True, color='0000C0'), fill=FILL_IN)
    put(ws, 'D3', '起始日期', font=F_H2, fill=FILL_HDR2)
    put(ws, 'E3', f'={BOM_CELL}', font=Font(name='微软雅黑', size=10, bold=True, color='0000C0'),
        fill=FILL_IN, fmt=DATE)
    put(ws, 'G3', '结束日期', font=F_H2, fill=FILL_HDR2); ws.merge_cells('G3:G3')
    put(ws, 'H3', f'={EOM_CELL}', font=Font(name='微软雅黑', size=10, bold=True, color='0000C0'),
        fill=FILL_IN, fmt=DATE)
    put(ws, 'J3', '对账状态', font=F_H2, fill=FILL_HDR2)
    put(ws, 'K3', '待对方确认', font=F_IN, fill=FILL_IN, align=CL); ws.merge_cells('K3:M3')
    put(ws, 'L3', None, fill=FILL_IN); put(ws, 'M3', None, fill=FILL_IN)
    dv_list(ws, 'B3', f'={party_src}')
    SUMS = [(f'期初应{"收" if is_ar else "付"}', 'A', 'B'), (f'本期{"销售" if is_ar else "采购"}', 'D', 'E'),
            (f'本期{"收款" if is_ar else "付款"}', 'G', 'H'), (f'期末应{"收" if is_ar else "付"}', 'J', 'K')]
    OP = Q(SH_OPEN)
    f_open = (f'SUMIFS({OP}!${op_amt}${o0}:${op_amt}${o1},{OP}!${op_party}${o0}:${op_party}${o1},$B$3,'
              f'{OP}!${op_date}${o0}:${op_date}${o1},"<"&$E$3)'
              f'+SUMIFS({biz}!${amt_c}${br0}:${amt_c}${br1},{biz}!$D${br0}:$D${br1},$B$3,'
              f'{biz}!$B${br0}:$B${br1},"<"&$E$3)'
              f'-SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$G${CASH_R0}:$G${CASH_R1},$B$3,'
              f'{CASHQ}!$E${CASH_R0}:$E${CASH_R1},"{"客户回款" if is_ar else "供应商付款"}",'
              f'{CASHQ}!$B${CASH_R0}:$B${CASH_R1},"<"&$E$3)')
    f_biz = (f'SUMIFS({biz}!${amt_c}${br0}:${amt_c}${br1},{biz}!$D${br0}:$D${br1},$B$3,'
             f'{biz}!$B${br0}:$B${br1},">="&$E$3,{biz}!$B${br0}:$B${br1},"<="&$H$3)')
    f_pay = (f'SUMIFS({CASHQ}!$N${CASH_R0}:$N${CASH_R1},{CASHQ}!$G${CASH_R0}:$G${CASH_R1},$B$3,'
             f'{CASHQ}!$E${CASH_R0}:$E${CASH_R1},"{"客户回款" if is_ar else "供应商付款"}",'
             f'{CASHQ}!$B${CASH_R0}:$B${CASH_R1},">="&$E$3,{CASHQ}!$B${CASH_R0}:$B${CASH_R1},"<="&$H$3)')
    for (lab, lc, vc), f in zip(SUMS, [f_open, f_biz, f_pay, '=$B$5+$E$5-$H$5']):
        put(ws, f'{lc}5', lab, font=F_TOT, fill=FILL_TOT)
        put(ws, f'{vc}5', f if f.startswith('=') else '=' + f, font=F_BIG, fill=FILL_CHK, fmt=MONEY)
    for c in ['C','F','I','L','M']: put(ws, f'{c}5', None, font=F_TOT, fill=FILL_TOT)
    ws.row_dimensions[5].height = 28
    hdr(ws, 'A7', f'本期{"销售" if is_ar else "采购"}明细', 'A7:F7', font=F_HDR, fill=FILL_HDR)
    hdr(ws, 'H7', f'本期{"收款" if is_ar else "付款"}明细', 'H7:M7', font=F_HDR, fill=FILL_HDR)
    headers(ws, 8, 1, ['日期', '商品 / 内容', '数量', '单价', '价税合计', '到期日'], fill=FILL_HDR2, font=F_HDR2)
    headers(ws, 8, 8, ['日期', '收款账户' if is_ar else '付款账户', '金额', '摘要', '经办人', '部门'],
            fill=FILL_HDR2, font=F_HDR2)
    L_SRC = [('B', 'A'), ('E', 'B'), ('F', 'C'), ('G', 'D'), (amt_c, 'E'), ('U' if is_ar else 'T', 'F')]
    R_SRC = [('B', 'H'), ('F', 'I'), ('N', 'J'), ('J', 'K'), ('I', 'L'), ('H', 'M')]
    for i in range(SOA_N):
        r = 9 + i
        for src, dst in L_SRC:
            fmt = DATE if src in ('B', 'U', 'T') else (MONEY if src in ('G', amt_c) else (QTY if src == 'F' else None))
            put(ws, f'{dst}{r}', f'=IFERROR(INDEX({biz}!${src}${br0}:${src}${br1},'
                                 f'MATCH(ROW()-8,{biz}!${seq_b}${br0}:${seq_b}${br1},0)),"")',
                font=F_LINK, fmt=fmt, align=CL if src == 'E' else C)
        for src, dst in R_SRC:
            fmt = DATE if src == 'B' else (MONEY if src == 'N' else None)
            put(ws, f'{dst}{r}', f'=IFERROR(INDEX({CASHQ}!${src}${CASH_R0}:${src}${CASH_R1},'
                                 f'MATCH(ROW()-8,{CASHQ}!${seq_c}${CASH_R0}:${seq_c}${CASH_R1},0)),"")',
                font=F_LINK, fmt=fmt, align=CL if src == 'J' else C)
        ws.row_dimensions[r].height = 16
    TR = 9 + SOA_N
    put(ws, f'A{TR}', '小计', font=F_TOT, fill=FILL_TOT)
    for c in ['B','C','D','F']: put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
    put(ws, f'E{TR}', f'=SUM(E9:E{TR-1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
    put(ws, f'H{TR}', '小计', font=F_TOT, fill=FILL_TOT)
    for c in ['I','K','L','M']: put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
    put(ws, f'J{TR}', f'=SUM(J9:J{TR-1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
    put(ws, f'A{TR+2}', f'本对账单根据双方往来记录生成，如有差异请在 30 日内书面提出。'
                        f'　　制表：＿＿＿＿＿　　日期：＿＿＿＿＿　　对方确认（盖章）：＿＿＿＿＿＿＿＿',
        font=F_NOTE, align=CL, border=None)
    ws.merge_cells(f'A{TR+2}:M{TR+2}')
    put(ws, f'A{TR+3}', f'注：明细最多显示 {SOA_N} 行。超过时请缩短对账期间分次打印。',
        font=F_NOTE, align=CL, border=None)
    ws.merge_cells(f'A{TR+3}:M{TR+3}')
    ws.freeze_panes = 'A9'; page(ws, titles='7:8')
build_soa(SH_SOA1, True)
build_soa(SH_SOA2, False)
print('  ✓ 客户对账单 / 供应商对账单')

# ============================================================ 统一修饰
TABCOLOR = {
    SH_HOME: '1F3864', SH_HELP: '1F3864',
    SH_PARAM: '7F7F7F', SH_BASE: '7F7F7F', SH_OPEN: '7F7F7F',
    SH_CASH: 'ED7D31', SH_SAL: 'ED7D31', SH_BUY: 'ED7D31', SH_OTH: 'ED7D31',
    SH_TB: '548235', SH_ACA: '548235', SH_PL: '548235', SH_BS: '548235', SH_CF: '548235',
    SH_BAL: '2F5597', SH_INV: '2F5597', SH_AR: '2F5597', SH_AP: '2F5597',
    SH_FA: '2F5597', SH_PAYROLL: '2F5597', SH_AUX2: '2F5597',
    SH_SOA1: '2F5597', SH_SOA2: '2F5597', SH_ANA: '2F5597',
    SH_CHK: 'C00000',
}
for name, color in TABCOLOR.items():
    wb[name].sheet_properties.tabColor = color

for name, last, r0, r1 in [(SH_CASH, 'W', CASH_R0, CASH_R1), (SH_SAL, 'AE', SAL_R0, SAL_R1),
                           (SH_BUY, 'AA', BUY_R0, BUY_R1), (SH_OTH, 'N', OTH_R0, OTH_R1),
                           (SH_TB, 'P', ACC_R0, ACC_R1), (SH_INV, 'P', GDS_R0, GDS_R1),
                           (SH_AR, 'N', CUS_R0, CUS_R1), (SH_AP, 'N', SUP_R0, SUP_R1),
                           (SH_FA, 'N', FA_R0, FA_R1)]:
    hr = 4 if name in (SH_CASH, SH_SAL, SH_BUY, SH_OTH) else 5
    wb[name].auto_filter.ref = f'A{hr}:{last}{r1}'

for name in [SH_PARAM, SH_BASE, SH_OPEN, SH_CASH, SH_SAL, SH_BUY, SH_OTH, SH_FA, SH_PAYROLL]:
    ws2 = wb[name]
    r = (ws2.max_row + 2)
    put(ws2, f'A{r}', '颜色约定：淡黄色＝手工录入　｜　白色 / 灰色＝公式自动计算，请勿覆盖　｜　'
                      '绿色字＝引用其他工作表，改动会连锁出错。整行删除会破坏结构，清空请选中后按 Delete。',
        font=F_NOTE, align=CL, border=None)
print('  ✓ 标签配色 / 自动筛选 / 颜色图例')
ORDER = [SH_HOME, SH_HELP, SH_PARAM, SH_BASE, SH_OPEN,
         SH_CASH, SH_SAL, SH_BUY, SH_OTH,
         SH_TB, SH_ACA, SH_PL, SH_BS, SH_CF,
         SH_BAL, SH_INV, SH_AR, SH_AP, SH_FA, SH_PAYROLL,
         SH_AUX2, SH_SOA1, SH_SOA2, SH_ANA, SH_CHK]
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
OUT = '/home/user/temp/内账财务管理系统/内账财务管理系统_通用版.xlsx'
wb.save(OUT)
nf = sum(1 for s in wb.worksheets for row in s.iter_rows() for c in row
         if isinstance(c.value, str) and c.value.startswith('='))
print(f'  ✓ 首页 / 操作说明')
print(f'\n已保存：{OUT}')
print(f'工作表 {len(wb.worksheets)} 张，公式 {nf} 个，自动配平括号 {BAL_FIXED[0]} 处')
