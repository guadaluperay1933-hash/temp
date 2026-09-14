# -*- coding: utf-8 -*-
"""A060 · 农业基地 / 木业 / 科技 三公司财务账套 —— 生成脚本"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from gl_common import *            # noqa
import gl_common as G
from a060_data import *            # noqa
import a060_data as D
import model

M = model.build()
D0 = dt.date(YEAR, 1, 1)

wb = Workbook()
wb.remove(wb.active)

def col_idx(letters):
    n = 0
    for ch in letters: n = n * 26 + (ord(ch) - 64)
    return n

NO_Q = '"否"'

def lk(key, keyrng, valrng, dflt='""'):
    return f'IFERROR(INDEX({valrng},MATCH({key},{keyrng},0)),{dflt})'

def nlk(key, keyrng, valrng):
    return f'N(IFERROR(INDEX({valrng},MATCH({key},{keyrng},0)),0))'

SPARE_N = 12
def hide_tail(ws, r0, r1, keep, filter_col=None, ref=None, spare=SPARE_N, vals=('有',)):
    keep = set(keep)
    if keep:
        nxt = max(keep) + 1
        keep |= set(range(nxt, min(nxt + spare, r1 + 1)))
    for r in range(r0, r1 + 1):
        if r not in keep: ws.row_dimensions[r].hidden = True
    if ref:
        ws.auto_filter.ref = ref
        if filter_col is not None:
            ws.auto_filter.add_filter_column(filter_col, list(vals), blank=False)

# 科目余额表的坐标在 a060_data 里，前面的表也要引用它
TBA, TBC = f'{SH_TB}!$A${TB_0}:$A${TB_1}', f'{SH_TB}!$C${TB_0}:$C${TB_1}'
TBD, TBF = f'{SH_TB}!$D${TB_0}:$D${TB_1}', f'{SH_TB}!$F${TB_0}:$F${TB_1}'
TBG, TBL = f'{SH_TB}!$G${TB_0}:$G${TB_1}', f'{SH_TB}!$L${TB_0}:$L${TB_1}'
TBK, TBO = f'{SH_TB}!$K${TB_0}:$K${TB_1}', f'{SH_TB}!$O${TB_0}:$O${TB_1}'

def filter_band(ws, lastcol, co_default=None,
                note='年度/起止都留空＝全部期间；只填年度＝按整年；填了起止就以起止为准'):
    """第 3 行统一的筛选带，第 4 行（隐藏）把它解析成取数区间。返回 (起, 止, 公司)"""
    if co_default is not None:
        cells = [('选择公司', 'A', 'B', co_default, None), ('年度', 'C', 'D', YEAR, '0'),
                 ('起始日期', 'E', 'F', None, DATE), ('截止日期', 'G', 'H', None, DATE)]
        co, y, s_, e_ = '$B$3', '$D$3', '$F$3', '$H$3'
        lab_c, disp_c = 'I', 'J'
    else:
        cells = [('年度', 'A', 'B', YEAR, '0'), ('起始日期', 'C', 'D', None, DATE),
                 ('截止日期', 'E', 'F', None, DATE)]
        co, y, s_, e_ = None, '$B$3', '$D$3', '$F$3'
        lab_c, disp_c = 'G', 'H'
    for lab, lc, c2, dflt, fmt in cells:
        put(ws, f'{lc}3', lab, font=F_H2, fill=FILL_HDR2)
        put(ws, f'{c2}3', dflt, font=Font(name='微软雅黑', size=10, bold=True, color='0000C0'),
            fill=FILL_IN, fmt=fmt)
    put(ws, f'{lab_c}3', '当前取数', font=F_H2, fill=FILL_HDR2)
    end_c = L(max(col_idx(lastcol), col_idx(disp_c) + 1))     # 窄表允许筛选带往右探出去
    for ci in range(col_idx(disp_c), col_idx(end_c) + 1):
        put(ws, f'{L(ci)}3', None, font=F_TOT, fill=FILL_CHK, align=CL)
    ws.merge_cells(f'{disp_c}3:{end_c}3')
    put(ws, f'{disp_c}3',
        f'=IF(AND({y}="",{s_}="",{e_}=""),"全部期间",TEXT($A$4,"yyyy-mm-dd")&"  至  "&TEXT($B$4,"yyyy-mm-dd"))'
        + (f'&"　｜　公司："&IF({co}="","（未选）",{co})' if co else '')
        + f'&"　｜　"&"{note}"',
        font=F_TOT, fill=FILL_CHK, align=CL)
    put(ws, 'A4', f'=IF({s_}<>"",{s_},IF({y}<>"",DATE({y},1,1),DATE(1900,1,1)))', font=F_NOTE, fmt=DATE)
    put(ws, 'B4', f'=IF({e_}<>"",{e_},IF({y}<>"",DATE({y},12,31),DATE(2199,12,31)))', font=F_NOTE, fmt=DATE)
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].hidden = True
    if co_default is not None:
        dv_list(ws, 'B3', '"' + ','.join(CO_NAMES) + '"')
    return '$A$4', '$B$4', co

# ============================================================ 参数设置
ws = wb.create_sheet(SH_PARAM)
title(ws, '参数设置', 'J', '一年一套表：先把下面的会计年度改成本年度，再另存为新文件，就是新一年的账。'
                          '这套表是内账口径——全部含税核算、税按实缴进费用，右边有说明。')
widths(ws, {'A':14,'B':22,'C':14,'D':30,'E':14,'F':14,'G':14,'H':14,'I':14,'J':30})
put(ws, 'A5', '会计年度', font=F_H2, fill=FILL_HDR2)
put(ws, 'B5', YEAR, font=Font(name='微软雅黑', size=12, bold=True, color='0000C0'), fill=FILL_IN, fmt='0')
put(ws, 'C5', '期初日期', font=F_H2, fill=FILL_HDR2)
put(ws, 'D5', '=DATE($B$5,1,1)', font=F_TOT, fmt=DATE)
put(ws, 'E5', '期末日期', font=F_H2, fill=FILL_HDR2)
put(ws, 'F5', '=DATE($B$5,12,31)', font=F_TOT, fmt=DATE)
put(ws, 'G5', '附加税率（参考）', font=F_H2, fill=FILL_HDR2)
put(ws, 'H5', SURTAX_RATE, font=F_IN, fill=FILL_IN, fmt=PCT)
put(ws, 'I5', '所得税率（参考）', font=F_H2, fill=FILL_HDR2)
put(ws, 'J5', CIT_RATE, font=F_IN, fill=FILL_IN, fmt=PCT)
put(ws, 'A6', '说明', font=F_NOTE, align=CL, border=None)
ws.merge_cells('B6:J6')
put(ws, 'B6', '本账套是内账：收入成本费用一律含税核算，税按实际缴纳当期计入费用，不计提、不拆进项销项。'
              '这两个税率只是算税时给你参考用（附加税＝城建 7%＋教育费附加 3%＋地方教育附加 2%；'
              '所得税按小微实际税负 5%），表里的任何计算都不用它们。'
              '农林牧渔中的林木种植所得免征企业所得税、自产农产品销售免征增值税，农业基地公司示例里就没有税。',
    font=F_NOTE, align=CL, border=None)

hdr(ws, 'A8', '一、三家公司', span='A8:J8', font=F_H2, fill=FILL_HDR2)
hm = headers(ws, 9, 1, ['简称（各表都用这个）', '公司全称', '纳税人身份', '主营业务', '本年凭证数'])
ws.merge_cells('D9:I9'); ws.merge_cells('D10:I10'); ws.merge_cells('D11:I11'); ws.merge_cells('D12:I12')
for i, (s, full, kind, biz) in enumerate(COS):
    r = 10 + i
    put(ws, f'A{r}', s, font=F_TOT, fill=FILL_TOT)
    put(ws, f'B{r}', full, font=F_IN, fill=FILL_IN, align=CL)
    put(ws, f'C{r}', kind, font=F_IN, fill=FILL_IN)
    put(ws, f'D{r}', biz, font=F_TXT, align=CL)
    for c in 'EFGHI': put(ws, f'{c}{r}', None)
    put(ws, f'J{r}', f'=COUNTIFS({SH_VOU}!$F${V_0}:$F${V_1},$A{r})', font=F_LINK, fmt='#,##0')

hdr(ws, 'A14', '二、费用项目（资金流水 / 生产加工 / 其他分录 的「费用项目」下拉就是这一列）',
    span='A14:J14', font=F_H2, fill=FILL_HDR2)
headers(ws, 15, 1, ['费用项目', '默认归集科目', '本年合计（三家）'])
ws.merge_cells('C15:D15')
for i, (name, acc) in enumerate(EXPENSES):
    r = 16 + i
    put(ws, f'A{r}', name, font=F_IN, fill=FILL_IN)
    put(ws, f'B{r}', acc, font=F_IN, fill=FILL_IN)
    ws.merge_cells(f'C{r}:D{r}')
    put(ws, f'C{r}', f'=IF($A{r}="","",{SH_EXP}!$E${{}})'.format(7 + i),
        font=F_LINK, fmt=MONEY)
EXP_LAST = 15 + len(EXPENSES)
for r in range(EXP_LAST + 1, 16 + 60):
    put(ws, f'A{r}', None, font=F_IN, fill=FILL_IN)
    put(ws, f'B{r}', None, font=F_IN, fill=FILL_IN)
    ws.merge_cells(f'C{r}:D{r}'); put(ws, f'C{r}', None, fmt=MONEY)
EXP_END = 15 + 60
put(ws, 'F16', '内账口径（这套表的核心约定）：', font=F_H2, align=CL, border=None)
for i, (k, v) in enumerate([
        ('含税核算', '收入、成本、费用一律按实际成交的含税金额入账，不拆进项销项，也不设进项/销项科目'),
        ('实交税费', '增值税、附加税、所得税、个税都在实际缴纳那个月，按缴的钱直接进费用'),
        ('缴税怎么录', '【资金流水】选「缴纳税费」→ 对方科目选税金及附加或所得税费用 → 费用项目选税种'),
        ('不计提税金', '不做「计提应交税费」的分录，资产负债表上也就没有应交税费余额（期初有欠税除外）'),
        ('开票税率', '商品档案上的税率只用来提醒开票时选哪一档，不参与任何计算')]):
    put(ws, f'F{17+i}', k, font=F_TOT, fill=FILL_TOT)
    ws.merge_cells(f'G{17+i}:J{17+i}')
    put(ws, f'G{17+i}', v, font=F_TXT, align=CL)
page(ws)
PARAM_Y = f'{SH_PARAM}!$B$5'
EXP_RNG = f'{SH_PARAM}!$A$16:$A${EXP_END}'

# ============================================================ 会计科目表
ws = wb.create_sheet(SH_ACC)
title(ws, '会计科目表', 'J',
      '报表就是按「报表项目」这一列汇总出来的。加科目照着填一行即可；'
      '「余额方向」决定期末余额怎么算，「报表项目」必须用资产负债表/利润表上已有的项目名。')
widths(ws, {'A':12,'B':26,'C':13,'D':10,'E':13,'F':18,'G':11,'H':14,'I':14,'J':30})
headers(ws, HR, 1, ['科目编码', '科目名称', '科目类别', '余额方向', '所属报表', '报表项目',
                    '往来核算', '本年借方发生', '本年贷方发生', '备注'])
for i, (code, name, cls, dr, rpt, item, aux) in enumerate(ACCS):
    r = ACC_0 + i
    put(ws, f'A{r}', code, font=F_IN, fill=FILL_IN)
    put(ws, f'B{r}', name, font=F_IN, fill=FILL_IN, align=CL)
    put(ws, f'C{r}', cls, font=F_IN, fill=FILL_IN)
    put(ws, f'D{r}', dr, font=F_IN, fill=FILL_IN)
    put(ws, f'E{r}', rpt, font=F_IN, fill=FILL_IN)
    put(ws, f'F{r}', item, font=F_IN, fill=FILL_IN, align=CL)
    put(ws, f'G{r}', aux, font=F_IN, fill=FILL_IN)
for r in range(ACC_0, ACC_1 + 1):
    for c in 'ABCDEFG':
        if ws[f'{c}{r}'].value is None: put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN)
    put(ws, f'H{r}', f'=IF($B{r}="","",SUMIFS({SH_VOU}!$I${V_0}:$I${V_1},{SH_VOU}!$H${V_0}:$H${V_1},$B{r}))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($B{r}="","",SUMIFS({SH_VOU}!$J${V_0}:$J${V_1},{SH_VOU}!$H${V_0}:$H${V_1},$B{r}))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'J{r}', None, font=F_IN, fill=FILL_IN, align=CL)
dv_list(ws, f'C{ACC_0}:C{ACC_1}', '"资产,负债,权益,成本,损益-收入,损益-费用"')
dv_list(ws, f'D{ACC_0}:D{ACC_1}', '"借,贷"')
dv_list(ws, f'E{ACC_0}:E{ACC_1}', '"资产负债表,利润表"')
dv_list(ws, f'G{ACC_0}:G{ACC_1}', '"是,否"')
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'A{ACC_0}'

A_CODE = f'{SH_ACC}!$A${ACC_0}:$A${ACC_1}'
A_NAME = f'{SH_ACC}!$B${ACC_0}:$B${ACC_1}'
A_CLS  = f'{SH_ACC}!$C${ACC_0}:$C${ACC_1}'
A_DIR  = f'{SH_ACC}!$D${ACC_0}:$D${ACC_1}'
A_RPT  = f'{SH_ACC}!$E${ACC_0}:$E${ACC_1}'
A_ITEM = f'{SH_ACC}!$F${ACC_0}:$F${ACC_1}'

# ============================================================ 记账规则
ws = wb.create_sheet(SH_RULE)
title(ws, '记账规则（业务类型 → 借贷分录）', 'N',
      '录入表只填业务，借贷两条腿由这张表决定。科目里的 @ 是占位符：'
      '@结算＝按结算方式取（挂账→应收/应付、现金→库存现金、银行→银行存款）；'
      '@存货/@收入/@成本＝按商品档案该商品设的科目取；@账户＝资金流水那行选的账户；@对方＝那行填的对方科目。'
      '看不惯哪条改哪条，改完全表立刻跟着变。')
widths(ws, {'A':16,'B':30})
for c in 'CDEFGHIJKLMN': widths(ws, {c: 15})
RB_0 = 6
def rule_block(r0, rules, nlegs, label):
    hdr(ws, f'A{r0-2}', label, span=f'A{r0-2}:{L(2+nlegs*3)}{r0-2}', font=F_H2, fill=FILL_HDR2)
    names = ['业务类型', '说明']
    for k in range(1, nlegs + 1): names += [f'腿{k} 科目', f'腿{k} 借贷', f'腿{k} 取值']
    headers(ws, r0 - 1, 1, names)
    for i, (t, desc, legs) in enumerate(rules):
        r = r0 + i
        put(ws, f'A{r}', t, font=F_TOT, fill=FILL_TOT)
        put(ws, f'B{r}', desc, font=F_TXT, align=CL)
        for k, (acc, dr, code) in enumerate(legs):
            put(ws, f'{L(3+k*3)}{r}', acc or None, font=F_IN, fill=FILL_IN)
            put(ws, f'{L(4+k*3)}{r}', dr or None, font=F_IN, fill=FILL_IN)
            put(ws, f'{L(5+k*3)}{r}', code or None, font=F_IN, fill=FILL_IN)
    return r0 + len(rules) - 1

RB_1 = rule_block(RB_0, RULE_BUY, LEG_BY, '一、购销流水（4 条腿：结算 / 收入或存货 / 结转成本借 / 结转成本贷）')
RC_0 = RB_1 + 4
RC_1 = rule_block(RC_0, RULE_CASH, LEG_CS, '二、资金流水（2 条腿：账户 / 对方）')
RP_0 = RC_1 + 4
RP_1 = rule_block(RP_0, RULE_PROD, LEG_PD, '三、生产加工（2 条腿）')
put(ws, f'A{RP_1+2}', '四、其他分录', font=F_H2, fill=FILL_HDR2)
ws.merge_cells(f'B{RP_1+2}:N{RP_1+2}')
put(ws, f'B{RP_1+2}', '没有规则表——直接在【其他分录】里填借方科目、贷方科目和金额，一借一贷。'
                      '一借多贷就拆成几行写。', font=F_TXT, align=CL)
page(ws)
R_BUY  = f'{SH_RULE}!$A${RB_0}:$N${RB_1}'
R_BUYK = f'{SH_RULE}!$A${RB_0}:$A${RB_1}'
R_CASH = f'{SH_RULE}!$A${RC_0}:$H${RC_1}'
R_CASHK= f'{SH_RULE}!$A${RC_0}:$A${RC_1}'
R_PROD = f'{SH_RULE}!$A${RP_0}:$H${RP_1}'
R_PRODK= f'{SH_RULE}!$A${RP_0}:$A${RP_1}'
BUY_KINDS  = '"' + ','.join(r[0] for r in RULE_BUY) + '"'
CASH_KINDS = '"' + ','.join(r[0] for r in RULE_CASH) + '"'
PROD_KINDS = '"' + ','.join(r[0] for r in RULE_PROD) + '"'

# ============================================================ 商品档案
ws = wb.create_sheet(SH_GD)
title(ws, '商品档案', 'M',
      '三家公司的料和成品都登在这里，用「所属公司」分开。存货/收入/成本科目决定这个商品的分录进哪个科目。'
      '内账是含税核算，「开票参考税率」只是登记开票时用哪档，不参与任何计算。')
widths(ws, {'A':11,'B':24,'C':18,'D':8,'E':10,'F':11,'G':13,'H':15,'I':15,'J':13,'K':26,'L':13,'M':10})
headers(ws, HR, 1, ['商品编码', '商品名称', '规格型号', '单位', '所属公司', '存货类别',
                    '存货科目', '收入科目', '成本科目', '开票参考税率', '备注',
                    '本年结存金额', '校验'])
for i, g in enumerate(GOODS):
    r = GD_0 + i
    for j, v in enumerate(g[:9]):
        put(ws, f'{L(1+j)}{r}', v if v != '' else None, font=F_IN, fill=FILL_IN,
            align=CL if j in (1, 2) else C)
    put(ws, f'J{r}', g[9], font=F_IN, fill=FILL_IN, fmt=PCT)
    put(ws, f'K{r}', g[10] or None, font=F_IN, fill=FILL_IN, align=CL)
for r in range(GD_0, GD_1 + 1):
    for c in 'ABCDEFGHIK':
        if ws[f'{c}{r}'].value is None: put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN)
    if ws[f'J{r}'].value is None: put(ws, f'J{r}', None, font=F_IN, fill=FILL_IN, fmt=PCT)
    put(ws, f'L{r}', f'=IF($A{r}="","",SUMIFS({SH_INV}!$U${GD_0}:$U${GD_1},{SH_INV}!$B${GD_0}:$B${GD_1},$A{r},'
                     f'{SH_INV}!$A${GD_0}:$A${GD_1},$E{r}))', font=F_LINK, fmt=MONEY)
    put(ws, f'M{r}', f'=IF($A{r}="","",IF(COUNTIF($A${GD_0}:$A${GD_1},$A{r})>1,"编码重复",'
                     f'IF($B{r}="","缺名称",IF(COUNTIF({A_NAME},$G{r})=0,"存货科目不存在",'
                     f'IF(AND($F{r}<>"原材料",COUNTIF({A_NAME},$H{r})=0),"收入科目不存在","OK")))))',
        font=F_TXT)
dv_list(ws, f'E{GD_0}:E{GD_1}', '"' + ','.join(CO_NAMES) + '"')
dv_list(ws, f'F{GD_0}:F{GD_1}', '"原材料,半成品,产成品,周转材料"')
dv_list(ws, f'G{GD_0}:G{GD_1}', f'={A_NAME}')
dv_list(ws, f'H{GD_0}:H{GD_1}', f'={A_NAME}')
dv_list(ws, f'I{GD_0}:I{GD_1}', f'={A_NAME}')
ws.conditional_formatting.add(f'M{GD_0}:M{GD_1}',
    FormulaRule(formula=[f'AND($M{GD_0}<>"",$M{GD_0}<>"OK")'], fill=FILL_WARN))
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'C{GD_0}'

G_CODE = f'{SH_GD}!$A${GD_0}:$A${GD_1}'
G_NAME = f'{SH_GD}!$B${GD_0}:$B${GD_1}'
G_SPEC = f'{SH_GD}!$C${GD_0}:$C${GD_1}'
G_UNIT = f'{SH_GD}!$D${GD_0}:$D${GD_1}'
G_CO   = f'{SH_GD}!$E${GD_0}:$E${GD_1}'
G_CLS  = f'{SH_GD}!$F${GD_0}:$F${GD_1}'
G_INV  = f'{SH_GD}!$G${GD_0}:$G${GD_1}'
G_REV  = f'{SH_GD}!$H${GD_0}:$H${GD_1}'
G_CST  = f'{SH_GD}!$I${GD_0}:$I${GD_1}'
G_RATE = f'{SH_GD}!$J${GD_0}:$J${GD_1}'

# ============================================================ 往来单位
ws = wb.create_sheet(SH_PT)
title(ws, '往来单位', 'I', '客户、供应商、以及三家公司互相之间。'
                          '「是否内部」标「是」的，购销流水会自动认成内部交易，合并报表按它抵销。')
widths(ws, {'A':24,'B':12,'C':24,'D':10,'E':14,'F':14,'G':14,'H':14,'I':10})
headers(ws, HR, 1, ['单位名称', '类型', '说明', '是否内部',
                    '本年应收发生', '本年应付发生', '期末应收余额', '期末应付余额', '校验'])
for i, (name, kind, desc, inner) in enumerate(PARTNERS):
    r = PT_0 + i
    put(ws, f'A{r}', name, font=F_IN, fill=FILL_IN, align=CL)
    put(ws, f'B{r}', kind, font=F_IN, fill=FILL_IN)
    put(ws, f'C{r}', desc, font=F_IN, fill=FILL_IN, align=CL)
    put(ws, f'D{r}', inner, font=F_IN, fill=FILL_IN)
for r in range(PT_0, PT_1 + 1):
    for c in 'ABCD':
        if ws[f'{c}{r}'].value is None: put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN)
    VK, VP, VI, VJ = (f'{SH_VOU}!$H${V_0}:$H${V_1}', f'{SH_VOU}!$K${V_0}:$K${V_1}',
                      f'{SH_VOU}!$I${V_0}:$I${V_1}', f'{SH_VOU}!$J${V_0}:$J${V_1}')
    put(ws, f'E{r}', f'=IF($A{r}="","",SUMIFS({VI},{VK},"应收账款",{VP},$A{r}))', font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",SUMIFS({VJ},{VK},"应付账款",{VP},$A{r}))', font=F_LINK, fmt=MONEY)
    put(ws, f'G{r}', f'=IF($A{r}="","",$E{r}-SUMIFS({VJ},{VK},"应收账款",{VP},$A{r}))', font=F_TOT, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($A{r}="","",$F{r}-SUMIFS({VI},{VK},"应付账款",{VP},$A{r}))', font=F_TOT, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($A{r}="","",IF(COUNTIF($A${PT_0}:$A${PT_1},$A{r})>1,"名称重复","OK"))', font=F_TXT)
dv_list(ws, f'B{PT_0}:B{PT_1}', '"客户,供应商,客户兼供应商,内部关联,其他"')
dv_list(ws, f'D{PT_0}:D{PT_1}', '"是,否"')
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'B{PT_0}'
T_NAME  = f'{SH_PT}!$A${PT_0}:$A${PT_1}'
T_INNER = f'{SH_PT}!$D${PT_0}:$D${PT_1}'

# ============================================================ 期初余额
ws = wb.create_sheet(SH_OPEN)
title(ws, f'期初余额（{YEAR}-01-01）', 'N',
      '左边填各公司的科目年初余额（借贷各一列），右边填存货的年初数量和金额。'
      '两边要对得上：右边某公司「库存商品」的金额合计，必须等于左边该公司库存商品的借方余额，下面有校验。')
widths(ws, {'A':10,'B':22,'C':15,'D':15,'E':22,'F':11,'G':3,
            'H':10,'I':11,'J':22,'K':12,'L':15,'M':12,'N':11})
headers(ws, HR, 1, ['公司', '科目名称', '借方余额', '贷方余额', '备注', '校验'])
headers(ws, HR, 8, ['公司', '商品编码', '商品名称', '期初数量', '期初金额', '期初单价', '校验'])
for i, (co, acc, dr, cr) in enumerate(OPEN_ACC):
    r = OB_0 + i
    put(ws, f'A{r}', co, font=F_IN, fill=FILL_IN)
    put(ws, f'B{r}', acc, font=F_IN, fill=FILL_IN, align=CL)
    put(ws, f'C{r}', dr or None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'D{r}', cr or None, font=F_IN, fill=FILL_IN, fmt=MONEY)
for r in range(OB_0, OB_1 + 1):
    for c in 'ABE':
        if ws[f'{c}{r}'].value is None: put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL)
    for c in 'CD':
        if ws[f'{c}{r}'].value is None: put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'F{r}', f'=IF(AND($A{r}="",$B{r}=""),"",IF(SUMPRODUCT(($B${OB_0}:$B${OB_1}=$B{r})*($A${OB_0}:$A${OB_1}=$A{r}))>1,"同公司同科目重复",IF(ISNA(MATCH($B{r},{A_NAME},0)),"科目不存在",'
                     f'IF(AND(N($C{r})<>0,N($D{r})<>0),"借贷不能同时有数","OK"))))', font=F_TXT)
for i, (co, gc, q, a) in enumerate(OPEN_INV):
    r = OI_0 + i
    put(ws, f'H{r}', co, font=F_IN, fill=FILL_IN)
    put(ws, f'I{r}', gc, font=F_IN, fill=FILL_IN)
    put(ws, f'K{r}', q, font=F_IN, fill=FILL_IN, fmt=QTY)
    put(ws, f'L{r}', a, font=F_IN, fill=FILL_IN, fmt=MONEY)
for r in range(OI_0, OI_1 + 1):
    for c in 'HI':
        if ws[f'{c}{r}'].value is None: put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN)
    put(ws, f'J{r}', f'=IF($I{r}="","",{lk(f"$I{r}", G_CODE, G_NAME)})', font=F_LINK, align=CL)
    if ws[f'K{r}'].value is None: put(ws, f'K{r}', None, font=F_IN, fill=FILL_IN, fmt=QTY)
    if ws[f'L{r}'].value is None: put(ws, f'L{r}', None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'M{r}', f'=IF(N($K{r})=0,"",ROUND(N($L{r})/N($K{r}),4))', font=F_LINK, fmt='#,##0.0000')
    put(ws, f'N{r}', f'=IF($I{r}="","",IF(ISNA(MATCH($I{r},{G_CODE},0)),"商品不存在",'
                     f'IF({lk(f"$I{r}", G_CODE, G_CO)}<>$H{r},"商品不属于该公司","OK")))', font=F_TXT)
CHK0 = OB_1 + 2        # 小结块必须排在科目区之外，否则 SUMIFS 会把自己圈进去变循环引用
put(ws, f'H{CHK0}', '存货两边对账', font=F_H2, fill=FILL_HDR2)
ws.merge_cells(f'I{CHK0}:N{CHK0}')
put(ws, f'I{CHK0}', '右边存货金额合计 vs 左边「原材料 + 库存商品」借方余额', font=F_NOTE, align=CL)
for i, co in enumerate(CO_NAMES):
    r = CHK0 + 1 + i
    put(ws, f'H{r}', co, font=F_TOT, fill=FILL_TOT)
    put(ws, f'I{r}', f'=SUMIFS($L${OI_0}:$L${OI_1},$H${OI_0}:$H${OI_1},$H{r})', font=F_LINK, fmt=MONEY)
    ws.merge_cells(f'J{r}:K{r}')
    put(ws, f'J{r}', f'=SUMIFS($C${OB_0}:$C${OB_1},$A${OB_0}:$A${OB_1},$H{r},$B${OB_0}:$B${OB_1},"原材料")'
                     f'+SUMIFS($C${OB_0}:$C${OB_1},$A${OB_0}:$A${OB_1},$H{r},$B${OB_0}:$B${OB_1},"库存商品")',
        font=F_LINK, fmt=MONEY)
    ws.merge_cells(f'L{r}:N{r}')
    put(ws, f'L{r}', f'=IF(ROUND($I{r}-$J{r},2)=0,"√ 一致","✗ 差 "&TEXT($I{r}-$J{r},"#,##0.00"))',
        font=F_TOT, align=CL)
    ws.conditional_formatting.add(f'L{r}', FormulaRule(formula=[f'LEFT($L{r},1)="✗"'], fill=FILL_WARN))
BALR = CHK0 + 5
put(ws, f'A{BALR}', '借贷平衡', font=F_H2, fill=FILL_HDR2)
for i, co in enumerate(CO_NAMES):
    r = BALR + 1 + i
    put(ws, f'A{r}', co, font=F_TOT, fill=FILL_TOT)
    put(ws, f'B{r}', '借方合计 / 贷方合计', font=F_NOTE, align=CL)
    put(ws, f'C{r}', f'=SUMIFS($C${OB_0}:$C${OB_1},$A${OB_0}:$A${OB_1},$A{r})', font=F_TOT, fmt=MONEY)
    put(ws, f'D{r}', f'=SUMIFS($D${OB_0}:$D${OB_1},$A${OB_0}:$A${OB_1},$A{r})', font=F_TOT, fmt=MONEY)
    put(ws, f'E{r}', f'=IF(ROUND($C{r}-$D{r},2)=0,"√ 借贷平衡","✗ 差 "&TEXT($C{r}-$D{r},"#,##0.00"))',
        font=F_TOT, align=CL)
    ws.conditional_formatting.add(f'E{r}', FormulaRule(formula=[f'LEFT($E{r},1)="✗"'], fill=FILL_WARN))
dv_list(ws, f'A{OB_0}:A{OB_1}', '"' + ','.join(CO_NAMES) + '"')
dv_list(ws, f'B{OB_0}:B{OB_1}', f'={A_NAME}')
dv_list(ws, f'H{OI_0}:H{OI_1}', '"' + ','.join(CO_NAMES) + '"')
dv_list(ws, f'I{OI_0}:I{OI_1}', f'={G_CODE}')
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'A{OB_0}'
O_CO, O_ACC = f'{SH_OPEN}!$A${OB_0}:$A${OB_1}', f'{SH_OPEN}!$B${OB_0}:$B${OB_1}'
O_DR, O_CR = f'{SH_OPEN}!$C${OB_0}:$C${OB_1}', f'{SH_OPEN}!$D${OB_0}:$D${OB_1}'
OI_CO, OI_GC = f'{SH_OPEN}!$H${OI_0}:$H${OI_1}', f'{SH_OPEN}!$I${OI_0}:$I${OI_1}'
OI_Q, OI_A = f'{SH_OPEN}!$K${OI_0}:$K${OI_1}', f'{SH_OPEN}!$L${OI_0}:$L${OI_1}'

# ============================================================ 购销流水
ws = wb.create_sheet(SH_BUY)
title(ws, '购销流水（录入表 · 全部行都显示）', 'Z',
      '进货、卖货都在这一张表上，用「公司」列分开。**内账含税核算**：单价填实际成交的含税价，'
      '金额＝数量×单价，不拆进项销项，收入和成本都是含税数。'
      '这里只管货和票，一律挂应收/应付；收钱付钱到【资金流水】记一笔，账户余额和往来才对得上。'
      '退货用「销售退回 / 采购退回」，数量照正数填。'
      '向农户收购原木：往来单位选那个农户，发票类型选「收购发票」，单价就填收购价。')
BW = {'A':6,'B':11,'C':7,'D':11,'E':17,'F':10,'G':22,'H':15,'I':7,'J':10,'K':12,'L':10,'M':13,
      'N':24,'O':14,'P':15,'Q':14,'R':12,'S':11,'T':13,'U':13,'V':12,'W':13,'X':8,'Y':7,'Z':10}
widths(ws, BW)
BH = ['序号','日期','公司','业务类型','往来单位','商品编码','商品名称','规格型号','单位','数量',
      '含税单价','发票类型','发票号码','摘要','备注','校验',
      '金额（含税）','结算科目','存货科目','收入科目','成本科目','成本单价','成本金额',
      '内部','年度','所属月份']
headers(ws, HR, 1, BH[:16])
headers(ws, HR, 17, BH[16:], fill=FILL_AUTO, font=F_HDR2)
EXB = {i: e for i, e in enumerate(EX_BUY)}
for i in range(BY_1 - BY_0 + 1):
    r = BY_0 + i
    put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-{BY_0-1})', font=F_NOTE, fmt='0')
    vals = {}
    if i in EXB:
        d, co, kind, pt, gc, qty, price, inv, invno, memo = EXB[i]
        vals = {'B': dt.date.fromisoformat(d), 'C': co, 'D': kind, 'E': pt, 'F': gc,
                'J': qty, 'K': price, 'L': inv, 'M': invno or None, 'N': memo}
    for c, fmt, al in (('B', DATE, C), ('C', None, C), ('D', None, C), ('E', None, CL), ('F', None, C),
                       ('J', QTY, C), ('K', '#,##0.0000', C), ('L', None, C), ('M', None, C),
                       ('N', None, CL), ('O', None, CL)):
        put(ws, f'{c}{r}', vals.get(c), font=F_IN, fill=FILL_IN, fmt=fmt, align=al)
    put(ws, f'G{r}', f'=IF($F{r}="","",{lk(f"$F{r}", G_CODE, G_NAME)})', font=F_LINK, align=CL)
    put(ws, f'H{r}', f'=IF($F{r}="","",{lk(f"$F{r}", G_CODE, G_SPEC)})', font=F_LINK, align=CL)
    put(ws, f'I{r}', f'=IF($F{r}="","",{lk(f"$F{r}", G_CODE, G_UNIT)})', font=F_LINK)
    put(ws, f'Q{r}', f'=IF($F{r}="","",ROUND(N($J{r})*N($K{r}),2))', font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'R{r}', f'=IF($D{r}="","",IF(LEFT($D{r},2)="销售","应收账款","应付账款"))',
        font=F_LINK, fill=FILL_AUTO)
    put(ws, f'S{r}', f'=IF($F{r}="","",{lk(f"$F{r}", G_CODE, G_INV)})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'T{r}', f'=IF($F{r}="","",{lk(f"$F{r}", G_CODE, G_REV)})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'U{r}', f'=IF($F{r}="","",{lk(f"$F{r}", G_CODE, G_CST)})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'V{r}', f'=IF($F{r}="","",SUMIFS({SH_INV}!$N${GD_0}:$N${GD_1},{SH_INV}!$A${GD_0}:$A${GD_1},$C{r},'
                     f'{SH_INV}!$B${GD_0}:$B${GD_1},$F{r}))', font=F_LINK, fill=FILL_AUTO, fmt='#,##0.0000')
    put(ws, f'W{r}', f'=IF(OR($D{r}="销售出库",$D{r}="销售退回"),ROUND(N($J{r})*N($V{r}),2),0)',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    _inner = lk(f'$E{r}', T_NAME, T_INNER, dflt=NO_Q)
    put(ws, f'X{r}', f'=IF($E{r}="","",{_inner})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'Y{r}', f'=IF($B{r}="","",YEAR($B{r}))', font=F_LINK, fill=FILL_AUTO, fmt='0')
    put(ws, f'Z{r}', f'=IF($B{r}="","",TEXT($B{r},"yyyy-mm"))', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'P{r}', f'=IF(AND($B{r}="",$D{r}="",$F{r}=""),"",'
                     f'IF(NOT(ISNUMBER($B{r})),"日期无效",'
                     f'IF(ISNA(MATCH($C{r},{SH_PARAM}!$A$10:$A$12,0)),"公司无效",'
                     f'IF(ISNA(MATCH($D{r},{R_BUYK},0)),"业务类型无效",'
                     f'IF(ISNA(MATCH($F{r},{G_CODE},0)),"商品编码无效",'
                     f'IF({lk(f"$F{r}", G_CODE, G_CO)}<>$C{r},"商品不属于该公司",'
                     f'IF(AND($E{r}<>"",ISNA(MATCH($E{r},{T_NAME},0))),"往来单位未建档",'
                     f'IF(N($J{r})=0,"数量为 0",'
                     f'IF(N($K{r})=0,"单价为 0",'
                     f'IF(AND(LEFT($D{r},2)="销售",$U{r}=""),"该商品没设成本科目",'
                     f'IF(YEAR($B{r})<>{PARAM_Y},"不在本会计年度","OK")))))))))))', font=F_TXT)
dv_list(ws, f'C{BY_0}:C{BY_1}', '"' + ','.join(CO_NAMES) + '"')
dv_list(ws, f'D{BY_0}:D{BY_1}', BUY_KINDS)
dv_list(ws, f'E{BY_0}:E{BY_1}', f'={T_NAME}')
dv_list(ws, f'F{BY_0}:F{BY_1}', f'={G_CODE}')
dv_list(ws, f'L{BY_0}:L{BY_1}', '"' + ','.join(INVTYPE) + '"')
dv_num(ws, f'K{BY_0}:K{BY_1}', 'greaterThanOrEqual', '-1000000000')
ws.conditional_formatting.add(f'P{BY_0}:P{BY_1}',
    FormulaRule(formula=[f'AND($P{BY_0}<>"",$P{BY_0}<>"OK")'], fill=FILL_WARN))
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'G{BY_0}'
KB = lambda c: f'{SH_BUY}!${c}${BY_0}:${c}${BY_1}'

# ============================================================ 资金账户
ws = wb.create_sheet(SH_ACCT)
title(ws, '资金账户', 'L',
      '几家公司共用几个账号，就按实际开的户登在这里。'
      '「对应科目」决定这个账户在账上走库存现金还是银行存款；'
      '共用账户照样在【资金流水】里按公司分别记账，账户余额是账户的，科目余额是各公司自己的，两边在下面对账。')
widths(ws, {'A':10,'B':18,'C':28,'D':20,'E':10,'F':13,'G':16,'H':15,'I':15,'J':15,'K':15,'L':11})
headers(ws, HR, 1, ['账户编码','账户名称','开户行 / 说明','账号','账户类型','对应科目','使用公司',
                    '期初余额','本期收入','本期支出','期末余额','校验'])
for i, a in enumerate(ACCOUNTS):
    r = AC_0 + i
    for j, v in enumerate(a[:7]):
        put(ws, f'{L(1+j)}{r}', v, font=F_IN, fill=FILL_IN, align=CL if j in (2, 3, 6) else C)
    put(ws, f'H{r}', a[7], font=F_IN, fill=FILL_IN, fmt=MONEY)
for r in range(AC_0, AC_1 + 1):
    for c in 'ABCDEFG':
        if ws[f'{c}{r}'].value is None: put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL)
    if ws[f'H{r}'].value is None: put(ws, f'H{r}', None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($B{r}="","",SUMIFS({SH_CASH}!$I${CS_0}:$I${CS_1},{SH_CASH}!$D${CS_0}:$D${CS_1},$B{r}))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($B{r}="","",SUMIFS({SH_CASH}!$J${CS_0}:$J${CS_1},{SH_CASH}!$D${CS_0}:$D${CS_1},$B{r}))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'K{r}', f'=IF($B{r}="","",ROUND(N($H{r})+$I{r}-$J{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'L{r}', f'=IF($B{r}="","",IF(COUNTIF($B${AC_0}:$B${AC_1},$B{r})>1,"账户名称重复",'
                     f'IF(ISNA(MATCH($F{r},{A_NAME},0)),"对应科目不存在",'
                     f'IF($K{r}<0,"✗ 期末余额为负","OK"))))', font=F_TXT)
    ws.conditional_formatting.add(f'L{r}', FormulaRule(formula=[f'AND($L{r}<>"",$L{r}<>"OK")'], fill=FILL_WARN))
dv_list(ws, f'E{AC_0}:E{AC_1}', '"现金,银行,其他货币资金"')
dv_list(ws, f'F{AC_0}:F{AC_1}', f'={A_NAME}')
TOTR = AC_1 + 2
put(ws, f'A{TOTR}', '合  计', font=F_TOT, fill=FILL_TOT)
ws.merge_cells(f'B{TOTR}:G{TOTR}')
put(ws, f'B{TOTR}', '全部账户', font=F_TOT, fill=FILL_TOT, align=CL)
for c in 'HIJK':
    put(ws, f'{c}{TOTR}', f'=SUM({c}${AC_0}:{c}${AC_1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'L{TOTR}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'A{TOTR+2}', '账实对账', font=F_H2, fill=FILL_HDR2)
ws.merge_cells(f'B{TOTR+2}:L{TOTR+2}')
put(ws, f'B{TOTR+2}', '全部账户余额合计，应当等于三家公司「库存现金＋银行存款」的科目余额合计。'
                      '共用账户会让某一家的账面银行存款对不上单个账号，但三家加总一定对得上。',
    font=F_NOTE, align=CL)
CO_CASH = ('+'.join(f'SUMIFS({TBL},{TBA},{SH_PARAM}!$A${10+i},{TBC},"库存现金")'
                    f'+SUMIFS({TBL},{TBA},{SH_PARAM}!$A${10+i},{TBC},"银行存款")' for i in range(len(COS))))
put(ws, f'A{TOTR+3}', '账户余额合计', font=F_TOT, fill=FILL_TOT)
ws.merge_cells(f'B{TOTR+3}:C{TOTR+3}')
put(ws, f'B{TOTR+3}', f'=$K${TOTR}', font=F_TOT, fmt=MONEY)
put(ws, f'D{TOTR+3}', '三家科目余额合计', font=F_TOT, fill=FILL_TOT)
ws.merge_cells(f'E{TOTR+3}:F{TOTR+3}')
put(ws, f'E{TOTR+3}', f'={CO_CASH}', font=F_TOT, fmt=MONEY)
ws.merge_cells(f'G{TOTR+3}:L{TOTR+3}')
put(ws, f'G{TOTR+3}', f'=IF(ROUND($B${TOTR+3}-$E${TOTR+3},2)=0,"√ 账实一致",'
                      f'"✗ 差 "&TEXT($B${TOTR+3}-$E${TOTR+3},"#,##0.00")&"　（多半是购销或其他分录直接动了现金/银行，没走资金流水）")',
    font=F_TOT, align=CL)
ws.conditional_formatting.add(f'G{TOTR+3}', FormulaRule(formula=[f'LEFT($G${TOTR+3},1)="✗"'], fill=FILL_WARN))
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'C{AC_0}'
AK_CODE = f'{SH_ACCT}!$A${AC_0}:$A${AC_1}'
AK_NAME = f'{SH_ACCT}!$B${AC_0}:$B${AC_1}'
AK_ACC  = f'{SH_ACCT}!$F${AC_0}:$F${AC_1}'
AK_OPEN = f'{SH_ACCT}!$H${AC_0}:$H${AC_1}'

# ============================================================ 资金流水（混合录入 · 实时结余额）
ws = wb.create_sheet(SH_CASH)
title(ws, '资金流水（录入表 · 收支混排、逐行结出账户余额）', 'T',
      '所有账户的收支都按发生顺序记在这一张表上：收进来的填「收入金额」，付出去的填「支出金额」，'
      '右边「账户余额」按账户逐行滚出来——几家公司共用一个账号也没关系，余额是按账户算的，'
      '账是按「公司」列各记各的。金额一律填实际收付的含税金额，不拆税。'
      '按日期先后往下录，余额才有意义（录反了校验列会提示）。'
      '收付货款选好「往来单位」，就自动冲【往来台账】里那家的应收/应付；'
      '交税选「缴纳税费」，对方科目选税金及附加或所得税费用、费用项目选税种，当期直接进费用。')
widths(ws, {'A':6,'B':11,'C':7,'D':17,'E':14,'F':17,'G':21,'H':14,'I':14,'J':14,'K':26,
            'L':12,'M':14,'N':18,'O':10,'P':13,'Q':12,'R':7,'S':15,'T':7})
headers(ws, HR, 1, ['序号','日期','公司','账户','业务类型','往来单位','对方科目','费用项目',
                    '收入金额','支出金额','摘要','关联单号','备注','校验'])
headers(ws, HR, 15, ['所属月份','金额','账户科目','收支','账户余额','年度'],
        fill=FILL_AUTO, font=F_HDR2)
EXC = {i: e for i, e in enumerate(EX_CASH)}
for i in range(CS_1 - CS_0 + 1):
    r = CS_0 + i
    put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-{CS_0-1})', font=F_NOTE, fmt='0')
    vals = {}
    if i in EXC:
        d, co, acct, kind, pt, ctr, ei, cin, cout, ref, memo = EXC[i]
        vals = {'B': dt.date.fromisoformat(d), 'C': co, 'D': acct, 'E': kind, 'F': pt or None,
                'G': ctr or None, 'H': ei or None, 'I': cin or None, 'J': cout or None,
                'K': memo, 'L': ref or None}
    for c, fmt, al in (('B', DATE, C), ('C', None, C), ('D', None, C), ('E', None, C), ('F', None, CL),
                       ('G', None, CL), ('H', None, C), ('I', MONEY, C), ('J', MONEY, C),
                       ('K', None, CL), ('L', None, C), ('M', None, CL)):
        put(ws, f'{c}{r}', vals.get(c), font=F_IN, fill=FILL_IN, fmt=fmt, align=al)
    put(ws, f'O{r}', f'=IF($B{r}="","",TEXT($B{r},"yyyy-mm"))', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'P{r}', f'=IF($E{r}="","",ROUND(N($I{r})+N($J{r}),2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'Q{r}', f'=IF($D{r}="","",{lk(f"$D{r}", AK_NAME, AK_ACC)})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'R{r}', f'=IF($E{r}="","",IFERROR(IF(INDEX({R_CASH},MATCH($E{r},{R_CASHK},0),4)="借","收","支"),""))',
        font=F_LINK, fill=FILL_AUTO)
    put(ws, f'S{r}', f'=IF($D{r}="","",ROUND(SUMIFS({AK_OPEN},{AK_NAME},$D{r})'
                     f'+SUMIFS($I${CS_0}:$I{r},$D${CS_0}:$D{r},$D{r})'
                     f'-SUMIFS($J${CS_0}:$J{r},$D${CS_0}:$D{r},$D{r}),2))',
        font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'T{r}', f'=IF($B{r}="","",YEAR($B{r}))', font=F_LINK, fill=FILL_AUTO, fmt='0')
    need_ctr = f'SUMPRODUCT(--(INDEX({R_CASH},MATCH($E{r},{R_CASHK},0),0)="@对方"))>0'
    prev = f'$B{r-1}' if r > CS_0 else '""'
    put(ws, f'N{r}', f'=IF(AND($B{r}="",$E{r}="",$I{r}="",$J{r}=""),"",'
                     f'IF(NOT(ISNUMBER($B{r})),"日期无效",'
                     f'IF(ISNA(MATCH($C{r},{SH_PARAM}!$A$10:$A$12,0)),"公司无效",'
                     f'IF(ISNA(MATCH($D{r},{AK_NAME},0)),"账户未建档",'
                     f'IF(ISNA(MATCH($E{r},{R_CASHK},0)),"业务类型无效",'
                     f'IF(AND(N($I{r})<>0,N($J{r})<>0),"收入和支出只能填一列",'
                     f'IF($P{r}=0,"金额为 0",'
                     f'IF(AND($R{r}="收",N($I{r})=0),"这类业务是收款，金额填到收入列",'
                     f'IF(AND($R{r}="支",N($J{r})=0),"这类业务是付款，金额填到支出列",'
                     f'IF(AND({need_ctr},$G{r}=""),"这个类型要填对方科目",'
                     f'IF(AND($G{r}<>"",ISNA(MATCH($G{r},{A_NAME},0))),"对方科目不存在",'
                     f'IF(AND($E{r}="缴纳税费",$H{r}=""),"缴税请在费用项目里选税种",'
                     f'IF(AND($F{r}<>"",ISNA(MATCH($F{r},{T_NAME},0))),"往来单位未建档",'
                     f'IF(YEAR($B{r})<>{PARAM_Y},"不在本会计年度",'
                     f'IF(AND({prev}<>"",ISNUMBER({prev}),$B{r}<{prev}),"日期比上一行早，余额按行序滚，建议按日期录",'
                     f'"OK")))))))))))))))', font=F_TXT)
dv_list(ws, f'C{CS_0}:C{CS_1}', '"' + ','.join(CO_NAMES) + '"')
dv_list(ws, f'D{CS_0}:D{CS_1}', f'={AK_NAME}')
dv_list(ws, f'E{CS_0}:E{CS_1}', CASH_KINDS)
dv_list(ws, f'F{CS_0}:F{CS_1}', f'={T_NAME}')
dv_list(ws, f'G{CS_0}:G{CS_1}', f'={A_NAME}')
dv_list(ws, f'H{CS_0}:H{CS_1}', f'={EXP_RNG}')
ws.conditional_formatting.add(f'N{CS_0}:N{CS_1}',
    FormulaRule(formula=[f'AND($N{CS_0}<>"",$N{CS_0}<>"OK")'], fill=FILL_WARN))
ws.conditional_formatting.add(f'S{CS_0}:S{CS_1}',
    CellIsRule(operator='lessThan', formula=['0'], fill=FILL_WARN))
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'F{CS_0}'
KC = lambda c: f'{SH_CASH}!${c}${CS_0}:${c}${CS_1}'

# ============================================================ 生产加工
ws = wb.create_sheet(SH_PROD)
title(ws, '生产加工（录入表 · 全部行都显示）', 'W',
      '一个批次号 = 一次加工。同一批次里：「领用投入」按加权平均成本自动出库，'
      '「加工费用」填金额和对方科目（人工先挂应付职工薪酬、折旧挂累计折旧），'
      '「产成品入库」不填金额——本批全部投入按「权重」自动分摊到产出。'
      '一批只有一个产出，权重留空即可；一批两个产出，就照价值比例填 3 和 1 这样。')
widths(ws, {'A':6,'B':11,'C':7,'D':11,'E':12,'F':10,'G':22,'H':7,'I':11,'J':13,'K':18,'L':13,
            'M':8,'N':24,'O':13,'P':14,'Q':11,'R':12,'S':13,'T':13,'U':13,'V':7,'W':9,'X':9})
headers(ws, HR, 1, ['序号','日期','公司','批次号','业务类型','商品编码','商品名称','单位','数量',
                    '加工费金额','对方科目','费用项目','权重','摘要','备注','校验'])
headers(ws, HR, 17, ['存货科目','成本单价','投入金额','产出成本','记账金额','年度','权重实际','累计权重'],
        fill=FILL_AUTO, font=F_HDR2)
EXP_ = {i: e for i, e in enumerate(EX_PROD)}
for i in range(PD_1 - PD_0 + 1):
    r = PD_0 + i
    put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-{PD_0-1})', font=F_NOTE, fmt='0')
    vals = {}
    if i in EXP_:
        d, co, bt, kind, gc, qty, fee, ctr, ei, w, memo = EXP_[i]
        vals = {'B': dt.date.fromisoformat(d), 'C': co, 'D': bt, 'E': kind, 'F': gc or None,
                'I': qty or None, 'J': fee or None, 'K': ctr or None, 'L': ei or None,
                'M': w or None, 'N': memo}
    for c, fmt, al in (('B', DATE, C), ('C', None, C), ('D', None, C), ('E', None, C), ('F', None, C),
                       ('I', QTY, C), ('J', MONEY, C), ('K', None, CL), ('L', None, C),
                       ('M', '#,##0.##', C), ('N', None, CL), ('O', None, CL)):
        put(ws, f'{c}{r}', vals.get(c), font=F_IN, fill=FILL_IN, fmt=fmt, align=al)
    put(ws, f'G{r}', f'=IF($F{r}="","",{lk(f"$F{r}", G_CODE, G_NAME)})', font=F_LINK, align=CL)
    put(ws, f'H{r}', f'=IF($F{r}="","",{lk(f"$F{r}", G_CODE, G_UNIT)})', font=F_LINK)
    put(ws, f'Q{r}', f'=IF($F{r}="","",{lk(f"$F{r}", G_CODE, G_INV)})', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'R{r}', f'=IF($F{r}="","",SUMIFS({SH_INV}!$N${GD_0}:$N${GD_1},{SH_INV}!$A${GD_0}:$A${GD_1},$C{r},'
                     f'{SH_INV}!$B${GD_0}:$B${GD_1},$F{r}))', font=F_LINK, fill=FILL_AUTO, fmt='#,##0.0000')
    put(ws, f'W{r}', f'=IF($E{r}<>"产成品入库","",IF(N($M{r})=0,1,N($M{r})))',
        font=F_NOTE, fill=FILL_AUTO, fmt='#,##0.##')
    put(ws, f'X{r}', f'=IF($E{r}<>"产成品入库","",SUMIFS($W${PD_0}:$W{r},$C${PD_0}:$C{r},$C{r},'
                     f'$D${PD_0}:$D{r},$D{r},$E${PD_0}:$E{r},"产成品入库"))',
        font=F_NOTE, fill=FILL_AUTO, fmt='#,##0.##')
    put(ws, f'S{r}', f'=IF($E{r}="领用投入",ROUND(N($I{r})*N($R{r}),2),'
                     f'IF($E{r}="加工费用",ROUND(N($J{r}),2),0))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    wsum = (f'SUMIFS($W${PD_0}:$W${PD_1},$C${PD_0}:$C${PD_1},$C{r},$D${PD_0}:$D${PD_1},$D{r},'
            f'$E${PD_0}:$E${PD_1},"产成品入库")')
    insum = f'SUMIFS($S${PD_0}:$S${PD_1},$C${PD_0}:$C${PD_1},$C{r},$D${PD_0}:$D${PD_1},$D{r})'
    put(ws, f'T{r}', f'=IF($E{r}<>"产成品入库",0,IF({wsum}=0,0,'
                     f'ROUND({insum}*N($X{r})/{wsum},2)-ROUND({insum}*(N($X{r})-N($W{r}))/{wsum},2)))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'U{r}', f'=IF($E{r}="","",ROUND($S{r}+$T{r},2))', font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'V{r}', f'=IF($B{r}="","",YEAR($B{r}))', font=F_LINK, fill=FILL_AUTO, fmt='0')
    put(ws, f'P{r}', f'=IF(AND($B{r}="",$E{r}=""),"",'
                     f'IF(NOT(ISNUMBER($B{r})),"日期无效",'
                     f'IF(ISNA(MATCH($C{r},{SH_PARAM}!$A$10:$A$12,0)),"公司无效",'
                     f'IF($D{r}="","缺批次号",'
                     f'IF(ISNA(MATCH($E{r},{R_PRODK},0)),"业务类型无效",'
                     f'IF(AND($E{r}<>"加工费用",ISNA(MATCH($F{r},{G_CODE},0))),"商品编码无效",'
                     f'IF(AND($E{r}<>"加工费用",{lk(f"$F{r}", G_CODE, G_CO)}<>$C{r}),"商品不属于该公司",'
                     f'IF(AND($E{r}<>"加工费用",N($I{r})=0),"数量为 0",'
                     f'IF(AND($E{r}="加工费用",N($J{r})=0),"加工费金额为 0",'
                     f'IF(AND($E{r}="加工费用",ISNA(MATCH($K{r},{A_NAME},0))),"对方科目不存在",'
                     f'IF(YEAR($B{r})<>{PARAM_Y},"不在本会计年度","OK")))))))))))', font=F_TXT)
dv_list(ws, f'C{PD_0}:C{PD_1}', '"' + ','.join(CO_NAMES) + '"')
dv_list(ws, f'E{PD_0}:E{PD_1}', PROD_KINDS)
dv_list(ws, f'F{PD_0}:F{PD_1}', f'={G_CODE}')
dv_list(ws, f'K{PD_0}:K{PD_1}', f'={A_NAME}')
dv_list(ws, f'L{PD_0}:L{PD_1}', f'={EXP_RNG}')
ws.conditional_formatting.add(f'P{PD_0}:P{PD_1}',
    FormulaRule(formula=[f'AND($P{PD_0}<>"",$P{PD_0}<>"OK")'], fill=FILL_WARN))
for c in 'WX': ws.column_dimensions[c].hidden = True
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'F{PD_0}'
KD = lambda c: f'{SH_PROD}!${c}${PD_0}:${c}${PD_1}'

# ============================================================ 其他分录
ws = wb.create_sheet(SH_OTH)
title(ws, '其他分录（录入表 · 全部行都显示）', 'L',
      '计提折旧摊销、计提税金、计提工资、结转——凡是不动钱也不动货的分录都写这里，一借一贷。'
      '一借多贷就拆成几行。')
widths(ws, {'A':6,'B':11,'C':7,'D':32,'E':21,'F':21,'G':14,'H':16,'I':13,'J':16,'K':14,'L':7})
headers(ws, HR, 1, ['序号','日期','公司','摘要','借方科目','贷方科目','金额','往来单位','费用项目','备注','校验'])
headers(ws, HR, 12, ['年度'], fill=FILL_AUTO, font=F_HDR2)
OTH_AMT = {i: (o['amt'] if o['amt'] is not None else None) for i, o in enumerate(M['oth'])}
for i in range(OT_1 - OT_0 + 1):
    r = OT_0 + i
    put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-{OT_0-1})', font=F_NOTE, fmt='0')
    vals = {}
    if i < len(EX_OTH):
        d, co, memo, dr, cr, _amt, pt, ei = EX_OTH[i]
        vals = {'B': dt.date.fromisoformat(d), 'C': co, 'D': memo, 'E': dr, 'F': cr,
                'G': M['oth'][i]['amt'], 'H': pt or None, 'I': ei or None}
    for c, fmt, al in (('B', DATE, C), ('C', None, C), ('D', None, CL), ('E', None, CL), ('F', None, CL),
                       ('G', MONEY, C), ('H', None, CL), ('I', None, C), ('J', None, CL)):
        put(ws, f'{c}{r}', vals.get(c), font=F_IN, fill=FILL_IN, fmt=fmt, align=al)
    put(ws, f'L{r}', f'=IF($B{r}="","",YEAR($B{r}))', font=F_LINK, fill=FILL_AUTO, fmt='0')
    put(ws, f'K{r}', f'=IF(AND($B{r}="",$E{r}="",$G{r}=""),"",'
                     f'IF(NOT(ISNUMBER($B{r})),"日期无效",'
                     f'IF(ISNA(MATCH($C{r},{SH_PARAM}!$A$10:$A$12,0)),"公司无效",'
                     f'IF(ISNA(MATCH($E{r},{A_NAME},0)),"借方科目不存在",'
                     f'IF(ISNA(MATCH($F{r},{A_NAME},0)),"贷方科目不存在",'
                     f'IF($E{r}=$F{r},"借贷科目相同",'
                     f'IF(N($G{r})=0,"金额为 0",'
                     f'IF(YEAR($B{r})<>{PARAM_Y},"不在本会计年度","OK"))))))))', font=F_TXT)
dv_list(ws, f'C{OT_0}:C{OT_1}', '"' + ','.join(CO_NAMES) + '"')
dv_list(ws, f'E{OT_0}:E{OT_1}', f'={A_NAME}')
dv_list(ws, f'F{OT_0}:F{OT_1}', f'={A_NAME}')
dv_list(ws, f'H{OT_0}:H{OT_1}', f'={T_NAME}')
dv_list(ws, f'I{OT_0}:I{OT_1}', f'={EXP_RNG}')
ws.conditional_formatting.add(f'K{OT_0}:K{OT_1}',
    FormulaRule(formula=[f'AND($K{OT_0}<>"",$K{OT_0}<>"OK")'], fill=FILL_WARN))
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'D{OT_0}'
KO = lambda c: f'{SH_OTH}!${c}${OT_0}:${c}${OT_1}'

# ============================================================ 记账分录（自动展开）
ws = wb.create_sheet(SH_VOU)
title(ws, '记账分录（自动生成 · 不要手工改）', 'P',
      '四张录入表按【记账规则】自动展开成的借贷明细，所有报表都从这张表取数。'
      '空白行已隐藏，新增业务后点 数据 → 筛选 → 重新应用。')
widths(ws, {'A':7,'B':11,'C':8,'D':6,'E':11,'F':7,'G':30,'H':21,'I':14,'J':14,'K':16,
            'L':13,'M':7,'N':12,'O':13,'P':6})
headers(ws, HR, 1, ['序号','来源表','源行','腿','日期','公司','摘要','科目名称','借方金额','贷方金额',
                    '往来单位','规则科目','借贷','取值','金额','有'])
QB, QC, QD, QE = SH_BUY, SH_CASH, SH_PROD, SH_OTH
keep_v = []
seq = 0

def vou_row(v, src, s, k, srcsheet, rule_rng, rule_key, kind_ref, tokens, amt_map,
            date_c, co_c, memo_c, pt_c, guard):
    """写一条分录腿。tokens: {占位符: 源列}; amt_map: {取值码: 源列}"""
    put(ws, f'A{v}', f'=IF($H{v}="","",ROW()-{V_0-1})', font=F_NOTE, fmt='0')
    put(ws, f'B{v}', src, font=F_NOTE)
    put(ws, f'C{v}', s, font=F_NOTE, fmt='0')
    put(ws, f'D{v}', k, font=F_NOTE, fmt='0')
    if rule_rng:
        base = f'MATCH({kind_ref},{rule_key},0)'
        put(ws, f'L{v}', f'=IF({guard},"",IFERROR(INDEX({rule_rng},{base},{3+(k-1)*3}),""))',
            font=F_NOTE, fill=FILL_AUTO)
        put(ws, f'M{v}', f'=IF($L{v}="","",IFERROR(INDEX({rule_rng},{base},{4+(k-1)*3}),""))',
            font=F_NOTE, fill=FILL_AUTO)
        put(ws, f'N{v}', f'=IF($L{v}="","",IFERROR(INDEX({rule_rng},{base},{5+(k-1)*3}),""))',
            font=F_NOTE, fill=FILL_AUTO)
        tok = f'$L{v}'
        expr = tok
        for t, col in tokens.items():
            expr = f'IF({tok}="{t}",{srcsheet}!${col}{s},' + expr + ')'
        put(ws, f'H{v}', f'=IF({tok}="","",{expr})', font=F_TXT, align=CL)
        amt = '0'
        for code, col in amt_map.items():
            amt = f'IF($N{v}="{code}",N({srcsheet}!${col}{s}),' + amt + ')'
        put(ws, f'O{v}', f'=IF($H{v}="",0,ROUND({amt},2))', font=F_NOTE, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'E{v}', f'=IF($H{v}="","",{srcsheet}!${date_c}{s})', font=F_LINK, fmt=DATE)
    put(ws, f'F{v}', f'=IF($H{v}="","",{srcsheet}!${co_c}{s})', font=F_LINK)
    put(ws, f'G{v}', f'=IF($H{v}="","",{srcsheet}!${memo_c}{s})', font=F_LINK, align=CL)
    put(ws, f'K{v}', (f'=IF($H{v}="","",{srcsheet}!${pt_c}{s})' if pt_c else None), font=F_LINK, align=CL)
    put(ws, f'I{v}', f'=IF($M{v}="借",$O{v},0)', font=F_TOT, fmt=MONEY)
    put(ws, f'J{v}', f'=IF($M{v}="贷",$O{v},0)', font=F_TOT, fmt=MONEY)
    put(ws, f'P{v}', f'=IF(AND($H{v}<>"",ROUND($O{v},2)<>0),"有","")', font=F_NOTE)

# --- 购销流水 5 腿 ---
RB_LEGS = {t: legs for t, _, legs in RULE_BUY}
for i in range(BY_1 - BY_0 + 1):
    s = BY_0 + i
    for k in range(1, LEG_BY + 1):
        v = V_BY0 + i * LEG_BY + (k - 1)
        vou_row(v, SH_BUY, s, k, QB, R_BUY, R_BUYK, f'{QB}!$D{s}',
                {'@结算': 'R', '@存货': 'S', '@收入': 'T', '@成本': 'U'},
                {'金额': 'Q', '成本金额': 'W'},
                'B', 'C', 'N', 'E', f'{QB}!$D{s}=""')
        if i < len(EX_BUY) and RB_LEGS[EX_BUY[i][2]][k - 1][0]:
            keep_v.append(v)
# --- 资金流水 3 腿 ---
RC_LEGS = {t: legs for t, _, legs in RULE_CASH}
for i in range(CS_1 - CS_0 + 1):
    s = CS_0 + i
    for k in range(1, LEG_CS + 1):
        v = V_CS0 + i * LEG_CS + (k - 1)
        vou_row(v, SH_CASH, s, k, QC, R_CASH, R_CASHK, f'{QC}!$E{s}',
                {'@账户': 'Q', '@对方': 'G'},
                {'金额': 'P'},
                'B', 'C', 'K', 'F', f'{QC}!$E{s}=""')
        if i < len(EX_CASH) and RC_LEGS[EX_CASH[i][3]][k - 1][0]:
            keep_v.append(v)
# --- 生产加工 2 腿 ---
RP_LEGS = {t: legs for t, _, legs in RULE_PROD}
for i in range(PD_1 - PD_0 + 1):
    s = PD_0 + i
    for k in range(1, LEG_PD + 1):
        v = V_PD0 + i * LEG_PD + (k - 1)
        vou_row(v, SH_PROD, s, k, QD, R_PROD, R_PRODK, f'{QD}!$E{s}',
                {'@存货': 'Q', '@对方': 'K'}, {'记账金额': 'U'},
                'B', 'C', 'N', None, f'{QD}!$E{s}=""')
        if i < len(EX_PROD) and RP_LEGS[EX_PROD[i][3]][k - 1][0]:
            keep_v.append(v)
# --- 其他分录 2 腿（不走规则表，一借一贷） ---
for i in range(OT_1 - OT_0 + 1):
    s = OT_0 + i
    for k in (1, 2):
        v = V_OT0 + i * LEG_OT + (k - 1)
        acc_c, side = ('E', '借') if k == 1 else ('F', '贷')
        put(ws, f'A{v}', f'=IF($H{v}="","",ROW()-{V_0-1})', font=F_NOTE, fmt='0')
        put(ws, f'B{v}', SH_OTH, font=F_NOTE)
        put(ws, f'C{v}', s, font=F_NOTE, fmt='0')
        put(ws, f'D{v}', k, font=F_NOTE, fmt='0')
        put(ws, f'L{v}', None, font=F_NOTE, fill=FILL_AUTO)
        put(ws, f'M{v}', f'=IF($H{v}="","","{side}")', font=F_NOTE, fill=FILL_AUTO)
        put(ws, f'N{v}', None, font=F_NOTE, fill=FILL_AUTO)
        put(ws, f'H{v}', f'=IF(OR({QE}!$B{s}="",N({QE}!$G{s})=0),"",{QE}!${acc_c}{s})', font=F_TXT, align=CL)
        put(ws, f'O{v}', f'=IF($H{v}="",0,ROUND(N({QE}!$G{s}),2))', font=F_NOTE, fill=FILL_AUTO, fmt=MONEY)
        put(ws, f'E{v}', f'=IF($H{v}="","",{QE}!$B{s})', font=F_LINK, fmt=DATE)
        put(ws, f'F{v}', f'=IF($H{v}="","",{QE}!$C{s})', font=F_LINK)
        put(ws, f'G{v}', f'=IF($H{v}="","",{QE}!$D{s})', font=F_LINK, align=CL)
        put(ws, f'K{v}', f'=IF($H{v}="","",{QE}!$H{s})', font=F_LINK, align=CL)
        put(ws, f'I{v}', f'=IF($M{v}="借",$O{v},0)', font=F_TOT, fmt=MONEY)
        put(ws, f'J{v}', f'=IF($M{v}="贷",$O{v},0)', font=F_TOT, fmt=MONEY)
        put(ws, f'P{v}', f'=IF(AND($H{v}<>"",ROUND($O{v},2)<>0),"有","")', font=F_NOTE)
        if i < len(EX_OTH): keep_v.append(v)
for c in 'LMNO': ws.column_dimensions[c].hidden = True
hide_tail(ws, V_0, V_1, keep_v, filter_col=col_idx('P') - 1, ref=f'A{HR}:P{V_1}')
put(ws, 'A2', HIDE_NOTE, font=F_NOTE, align=CL, border=None)
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'E{V_0}'
VE, VF, VG = f'{SH_VOU}!$E${V_0}:$E${V_1}', f'{SH_VOU}!$F${V_0}:$F${V_1}', f'{SH_VOU}!$G${V_0}:$G${V_1}'
VH, VI, VJ = f'{SH_VOU}!$H${V_0}:$H${V_1}', f'{SH_VOU}!$I${V_0}:$I${V_1}', f'{SH_VOU}!$J${V_0}:$J${V_1}'
VK = f'{SH_VOU}!$K${V_0}:$K${V_1}'

# ============================================================ 科目余额表
ws = wb.create_sheet(SH_TB)
title(ws, '科目余额表', 'P',
      '三家公司上下排开。「期初余额」＝年初余额＋起始日之前的发生额，所以把起止日期改成某个月，'
      '这张表就是那个月的月初、本月发生、月末。资产负债表取「期末余额」，利润表取「本期发生净额」。')
widths(ws, {'A':8,'B':11,'C':22,'D':12,'E':7,'F':18,'G':14,'H':14,'I':14,'J':14,'K':14,'L':14,
            'M':14,'N':14,'O':14,'P':6})
DR_S, DR_E, _ = filter_band(ws, 'P')
headers(ws, HR, 1, ['公司','科目编码','科目名称','类别','方向','报表项目',
                    '年初余额','期初余额','本期借方','本期贷方','本期发生净额','期末余额',
                    '本年借方','本年贷方','本年发生净额','有'])
TR_TB = 6
put(ws, f'A{TR_TB}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BCDEF': put(ws, f'{c}{TR_TB}', None, font=F_TOT, fill=FILL_TOT)
for c in 'GHIJKLMNO':
    put(ws, f'{c}{TR_TB}', f'=SUM({c}${TB_0}:{c}${TB_1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'P{TR_TB}', None, font=F_TOT, fill=FILL_TOT)
keep_tb = []
for ci in range(len(COS)):
    for ai in range(ACC_CAP):
        r = TB_0 + ci * ACC_CAP + ai
        ar = ACC_0 + ai
        put(ws, f'A{r}', f'=IF({SH_ACC}!$B{ar}="","",{SH_PARAM}!$A${10+ci})', font=F_TXT)
        put(ws, f'B{r}', f'=IF({SH_ACC}!$B{ar}="","",{SH_ACC}!$A{ar})', font=F_TXT)
        put(ws, f'C{r}', f'=IF({SH_ACC}!$B{ar}="","",{SH_ACC}!$B{ar})', font=F_TXT, align=CL)
        put(ws, f'D{r}', f'=IF($C{r}="","",{SH_ACC}!$C{ar})', font=F_TXT)
        put(ws, f'E{r}', f'=IF($C{r}="","",{SH_ACC}!$D{ar})', font=F_TXT)
        put(ws, f'F{r}', f'=IF($C{r}="","",{SH_ACC}!$F{ar})', font=F_TXT, align=CL)
        od = f'SUMIFS({O_DR},{O_CO},$A{r},{O_ACC},$C{r})'
        oc = f'SUMIFS({O_CR},{O_CO},$A{r},{O_ACC},$C{r})'
        put(ws, f'G{r}', f'=IF($C{r}="","",IF($E{r}="借",{od}-{oc},{oc}-{od}))', font=F_LINK, fmt=MONEY)
        pd_ = f'SUMIFS({VI},{VF},$A{r},{VH},$C{r},{VE},"<"&{DR_S})'
        pc_ = f'SUMIFS({VJ},{VF},$A{r},{VH},$C{r},{VE},"<"&{DR_S})'
        put(ws, f'H{r}', f'=IF($C{r}="","",$G{r}+IF($E{r}="借",{pd_}-{pc_},{pc_}-{pd_}))',
            font=F_LINK, fmt=MONEY)
        rng = f',{VE},">="&{DR_S},{VE},"<="&{DR_E}'
        put(ws, f'I{r}', f'=IF($C{r}="","",SUMIFS({VI},{VF},$A{r},{VH},$C{r}{rng}))', font=F_LINK, fmt=MONEY)
        put(ws, f'J{r}', f'=IF($C{r}="","",SUMIFS({VJ},{VF},$A{r},{VH},$C{r}{rng}))', font=F_LINK, fmt=MONEY)
        put(ws, f'K{r}', f'=IF($C{r}="","",IF($E{r}="借",$I{r}-$J{r},$J{r}-$I{r}))', font=F_TOT, fmt=MONEY)
        put(ws, f'L{r}', f'=IF($C{r}="","",ROUND($H{r}+$K{r},2))', font=F_TOT, fmt=MONEY)
        yr = f',{VE},"<="&{DR_E}'
        put(ws, f'M{r}', f'=IF($C{r}="","",SUMIFS({VI},{VF},$A{r},{VH},$C{r}{yr}))', font=F_LINK, fmt=MONEY)
        put(ws, f'N{r}', f'=IF($C{r}="","",SUMIFS({VJ},{VF},$A{r},{VH},$C{r}{yr}))', font=F_LINK, fmt=MONEY)
        put(ws, f'O{r}', f'=IF($C{r}="","",IF($E{r}="借",$M{r}-$N{r},$N{r}-$M{r}))', font=F_TOT, fmt=MONEY)
        put(ws, f'P{r}', f'=IF($C{r}="","",IF(OR(ROUND($G{r},2)<>0,ROUND($M{r},2)<>0,'
                         f'ROUND($N{r},2)<>0),"有",""))', font=F_NOTE)
        if ai < len(ACCS): keep_tb.append(r)
hide_tail(ws, TB_0, TB_1, keep_tb, filter_col=col_idx('P') - 1, ref=f'A{HR}:P{TB_1}', spare=3)
put(ws, 'A2', HIDE_NOTE, font=F_NOTE, align=CL, border=None)
page(ws, titles=f'{HR}:{TR_TB}')
ws.freeze_panes = f'C{TB_0}'

# ============================================================ 进销存台账（全年，成本基准）
ws = wb.create_sheet(SH_INV)
title(ws, '进销存台账（全年 · 成本基准表）', 'V',
      '这张表不随日期筛选变——「加权平均单价」是全年口径，购销流水结转销售成本、生产加工领用出库，'
      '都按这一列取价。要看某个区间的收发存，用【收发存查询】。'
      '加权平均单价 ＝（期初金额＋采购入库金额＋生产入库金额）÷ 对应数量，不含出库，不会循环引用。')
widths(ws, {'A':8,'B':10,'C':22,'D':15,'E':7,'F':10,'G':11,'H':12,'I':13,'J':12,'K':13,'L':12,
            'M':13,'N':13,'O':12,'P':13,'Q':12,'R':13,'S':12,'T':12,'U':13,'V':13})
headers(ws, HR, 1, ['公司','商品编码','商品名称','规格型号','单位','类别','存货科目',
                    '期初数量','期初金额','采购入库数量','采购入库金额','生产入库数量','生产入库金额',
                    '加权平均单价','销售出库数量','销售出库金额','生产领用数量','生产领用金额',
                    '结存数量','结存单价','结存金额','校验'])
for r in range(GD_0, GD_1 + 1):
    g = f'{SH_GD}!$A{r}'
    put(ws, f'A{r}', f'=IF({g}="","",{SH_GD}!$E{r})', font=F_TXT)
    put(ws, f'B{r}', f'=IF({g}="","",{g})', font=F_TXT)
    put(ws, f'C{r}', f'=IF({g}="","",{SH_GD}!$B{r})', font=F_TXT, align=CL)
    put(ws, f'D{r}', f'=IF({g}="","",{SH_GD}!$C{r})', font=F_TXT, align=CL)
    put(ws, f'E{r}', f'=IF({g}="","",{SH_GD}!$D{r})', font=F_TXT)
    put(ws, f'F{r}', f'=IF({g}="","",{SH_GD}!$F{r})', font=F_TXT)
    put(ws, f'G{r}', f'=IF({g}="","",{SH_GD}!$G{r})', font=F_TXT)
    ck = f'{KB("C")},$A{r},{KB("F")},$B{r}'
    dk = f'{KD("C")},$A{r},{KD("F")},$B{r}'
    put(ws, f'H{r}', f'=IF($B{r}="","",SUMIFS({OI_Q},{OI_CO},$A{r},{OI_GC},$B{r}))', font=F_LINK, fmt=QTY)
    put(ws, f'I{r}', f'=IF($B{r}="","",SUMIFS({OI_A},{OI_CO},$A{r},{OI_GC},$B{r}))', font=F_LINK, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($B{r}="","",SUMIFS({KB("J")},{ck},{KB("D")},"采购入库")'
                     f'-SUMIFS({KB("J")},{ck},{KB("D")},"采购退回"))', font=F_LINK, fmt=QTY)
    put(ws, f'K{r}', f'=IF($B{r}="","",SUMIFS({KB("Q")},{ck},{KB("D")},"采购入库")'
                     f'-SUMIFS({KB("Q")},{ck},{KB("D")},"采购退回"))', font=F_LINK, fmt=MONEY)
    put(ws, f'L{r}', f'=IF($B{r}="","",SUMIFS({KD("I")},{dk},{KD("E")},"产成品入库"))', font=F_LINK, fmt=QTY)
    put(ws, f'M{r}', f'=IF($B{r}="","",SUMIFS({KD("T")},{dk},{KD("E")},"产成品入库"))', font=F_LINK, fmt=MONEY)
    put(ws, f'N{r}', f'=IF($B{r}="","",IF(($H{r}+$J{r}+$L{r})=0,0,'
                     f'($I{r}+$K{r}+$M{r})/($H{r}+$J{r}+$L{r})))', font=F_TOT, fmt='#,##0.0000')
    put(ws, f'O{r}', f'=IF($B{r}="","",SUMIFS({KB("J")},{ck},{KB("D")},"销售出库")'
                     f'-SUMIFS({KB("J")},{ck},{KB("D")},"销售退回"))', font=F_LINK, fmt=QTY)
    put(ws, f'P{r}', f'=IF($B{r}="","",SUMIFS({KB("W")},{ck},{KB("D")},"销售出库")'
                     f'-SUMIFS({KB("W")},{ck},{KB("D")},"销售退回"))', font=F_LINK, fmt=MONEY)
    put(ws, f'Q{r}', f'=IF($B{r}="","",SUMIFS({KD("I")},{dk},{KD("E")},"领用投入"))', font=F_LINK, fmt=QTY)
    put(ws, f'R{r}', f'=IF($B{r}="","",SUMIFS({KD("S")},{dk},{KD("E")},"领用投入"))', font=F_LINK, fmt=MONEY)
    put(ws, f'S{r}', f'=IF($B{r}="","",$H{r}+$J{r}+$L{r}-$O{r}-$Q{r})', font=F_TOT, fmt=QTY)
    put(ws, f'U{r}', f'=IF($B{r}="","",ROUND($I{r}+$K{r}+$M{r}-$P{r}-$R{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'T{r}', f'=IF($B{r}="","",IF($S{r}=0,0,ROUND($U{r}/$S{r},4)))', font=F_TXT, fmt='#,##0.0000')
    put(ws, f'V{r}', f'=IF($B{r}="","",IF($S{r}<-0.0001,"✗ 结存数量为负",'
                     f'IF(AND($S{r}=0,ROUND($U{r},2)<>0),"✗ 数量 0 金额不为 0",'
                     f'IF(ROUND($U{r},2)<-0.005,"✗ 结存金额为负","OK"))))', font=F_TXT)
    ws.conditional_formatting.add(f'V{r}', FormulaRule(formula=[f'LEFT($V{r},1)="✗"'], fill=FILL_WARN))
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'C{GD_0}'

# ============================================================ 收发存查询（按日期）
ws = wb.create_sheet(SH_INVQ)
IQ_0, IQ_1 = 7, 7 + (GD_1 - GD_0)
title(ws, '收发存查询（按起止日期）', 'N',
      '期初＝年初数＋起始日之前的收发；本期只算区间内的。金额按【进销存台账】的全年加权平均单价折算，'
      '所以这里是「按量还原」的口径，跟总账存货余额在期末（全年）时一致。')
widths(ws, {'A':8,'B':10,'C':22,'D':7,'E':13,'F':14,'G':13,'H':14,'I':13,'J':14,'K':13,'L':14,'M':13,'N':6})
IQ_S, IQ_E, _ = filter_band(ws, 'N')
headers(ws, HR, 1, ['公司','商品编码','商品名称','单位','期初数量','期初金额','本期入库数量','本期入库金额',
                    '本期出库数量','本期出库金额','期末数量','期末金额','加权平均单价','有'])
put(ws, 'A6', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BCD': put(ws, f'{c}6', None, font=F_TOT, fill=FILL_TOT)
for c in 'EFGHIJKLM':
    put(ws, f'{c}6', f'=SUM({c}${IQ_0}:{c}${IQ_1})' if c in 'FHJL' else
        (f'=SUM({c}${IQ_0}:{c}${IQ_1})'), font=F_TOT, fill=FILL_TOT, fmt=MONEY if c in 'FHJL' else QTY)
put(ws, 'N6', None, font=F_TOT, fill=FILL_TOT)
for i in range(GD_1 - GD_0 + 1):
    r, gr = IQ_0 + i, GD_0 + i
    put(ws, f'A{r}', f'=IF({SH_INV}!$B{gr}="","",{SH_INV}!$A{gr})', font=F_TXT)
    put(ws, f'B{r}', f'=IF({SH_INV}!$B{gr}="","",{SH_INV}!$B{gr})', font=F_TXT)
    put(ws, f'C{r}', f'=IF($B{r}="","",{SH_INV}!$C{gr})', font=F_TXT, align=CL)
    put(ws, f'D{r}', f'=IF($B{r}="","",{SH_INV}!$E{gr})', font=F_TXT)
    put(ws, f'M{r}', f'=IF($B{r}="","",{SH_INV}!$N{gr})', font=F_TXT, fmt='#,##0.0000')
    ck = f'{KB("C")},$A{r},{KB("F")},$B{r}'
    dk = f'{KD("C")},$A{r},{KD("F")},$B{r}'
    def qsum(before):
        d = (f',{KB("B")},"<"&{IQ_S}' if before else f',{KB("B")},">="&{IQ_S},{KB("B")},"<="&{IQ_E}')
        p = (f',{KD("B")},"<"&{IQ_S}' if before else f',{KD("B")},">="&{IQ_S},{KD("B")},"<="&{IQ_E}')
        inq = (f'SUMIFS({KB("J")},{ck},{KB("D")},"采购入库"{d})-SUMIFS({KB("J")},{ck},{KB("D")},"采购退回"{d})'
               f'+SUMIFS({KD("I")},{dk},{KD("E")},"产成品入库"{p})')
        outq = (f'SUMIFS({KB("J")},{ck},{KB("D")},"销售出库"{d})-SUMIFS({KB("J")},{ck},{KB("D")},"销售退回"{d})'
                f'+SUMIFS({KD("I")},{dk},{KD("E")},"领用投入"{p})')
        return inq, outq
    b_in, b_out = qsum(True)
    c_in, c_out = qsum(False)
    put(ws, f'E{r}', f'=IF($B{r}="","",{SH_INV}!$H{gr}+{b_in}-({b_out}))', font=F_LINK, fmt=QTY)
    put(ws, f'G{r}', f'=IF($B{r}="","",{c_in})', font=F_LINK, fmt=QTY)
    put(ws, f'I{r}', f'=IF($B{r}="","",{c_out})', font=F_LINK, fmt=QTY)
    put(ws, f'K{r}', f'=IF($B{r}="","",$E{r}+$G{r}-$I{r})', font=F_TOT, fmt=QTY)
    for qc, ac in (('E', 'F'), ('G', 'H'), ('I', 'J'), ('K', 'L')):
        put(ws, f'{ac}{r}', f'=IF($B{r}="","",ROUND(${qc}{r}*$M{r},2))',
            font=F_LINK if ac != 'L' else F_TOT, fmt=MONEY)
    put(ws, f'N{r}', f'=IF($B{r}="","",IF(OR($E{r}<>0,$G{r}<>0,$I{r}<>0),"有",""))', font=F_NOTE)
hide_tail(ws, IQ_0, IQ_1, [IQ_0 + i for i in range(len(GOODS))],
          filter_col=col_idx('N') - 1, ref=f'A{HR}:N{IQ_1}')
put(ws, 'A2', HIDE_NOTE, font=F_NOTE, align=CL, border=None)
page(ws, titles=f'{HR}:6')
ws.freeze_panes = f'C{IQ_0}'

# ============================================================ 期初余额 · 往来明细（补在同一张表右侧）
ws = wb[SH_OPEN]
widths(ws, {'O':3,'P':10,'Q':22,'R':15,'S':15,'T':11})
headers(ws, HR, 16, ['公司', '往来单位', '年初应收', '年初应付', '校验'])
for i, (co, pt, ar, ap) in enumerate(OPEN_PT):
    r = OI_0 + i
    put(ws, f'P{r}', co, font=F_IN, fill=FILL_IN)
    put(ws, f'Q{r}', pt, font=F_IN, fill=FILL_IN, align=CL)
    put(ws, f'R{r}', ar or None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'S{r}', ap or None, font=F_IN, fill=FILL_IN, fmt=MONEY)
for r in range(OI_0, OI_1 + 1):
    for c in 'PQ':
        if ws[f'{c}{r}'].value is None: put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL)
    for c in 'RS':
        if ws[f'{c}{r}'].value is None: put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'T{r}', f'=IF($Q{r}="","",IF(ISNA(MATCH($Q{r},{T_NAME},0)),"单位未建档","OK"))', font=F_TXT)
put(ws, f'P{CHK0}', '往来两边对账', font=F_H2, fill=FILL_HDR2)
ws.merge_cells(f'Q{CHK0}:T{CHK0}')
put(ws, f'Q{CHK0}', '按单位拆的年初应收/应付合计，要等于左边该公司应收账款/应付账款的余额', font=F_NOTE, align=CL)
for i, co in enumerate(CO_NAMES):
    r = CHK0 + 1 + i
    put(ws, f'P{r}', co, font=F_TOT, fill=FILL_TOT)
    for j, (src, acc, side) in enumerate([('R', '应收账款', 'C'), ('S', '应付账款', 'D')]):
        c = 'Q' if j == 0 else 'S'
        put(ws, f'{c}{r}',
            f'=SUMIFS(${src}${OI_0}:${src}${OI_1},$P${OI_0}:$P${OI_1},$P{r})'
            f'-SUMIFS(${side}${OB_0}:${side}${OB_1},$A${OB_0}:$A${OB_1},$P{r},$B${OB_0}:$B${OB_1},"{acc}")',
            font=F_LINK, fmt=MONEY)
        put(ws, f'{"R" if j==0 else "T"}{r}',
            f'=IF(ROUND(${c}{r},2)=0,"√ 应收一致" ,"✗ 差 "&TEXT(${c}{r},"#,##0.00"))' if j == 0 else
            f'=IF(ROUND(${c}{r},2)=0,"√ 应付一致","✗ 差 "&TEXT(${c}{r},"#,##0.00"))',
            font=F_TOT, align=CL)
dv_list(ws, f'P{OI_0}:P{OI_1}', '"' + ','.join(CO_NAMES) + '"')
dv_list(ws, f'Q{OI_0}:Q{OI_1}', f'={T_NAME}')
OP_CO, OP_PT = f'{SH_OPEN}!$P${OI_0}:$P${OI_1}', f'{SH_OPEN}!$Q${OI_0}:$Q${OI_1}'
OP_AR, OP_AP = f'{SH_OPEN}!$R${OI_0}:$R${OI_1}', f'{SH_OPEN}!$S${OI_0}:$S${OI_1}'

# ============================================================ 报表定义
def tb_end(co, item):   return f'SUMIFS({TBL},{TBA},{co},{TBF},"{item}")'
def tb_beg(co, item):   return f'SUMIFS({TBG},{TBA},{co},{TBF},"{item}")'
def pl_cur(co, item):   return f'SUMIFS({TBK},{TBA},{co},{TBF},"{item}")'
def pl_ytd(co, item):   return f'SUMIFS({TBO},{TBA},{co},{TBF},"{item}")'
def net_profit(co, col):
    rng = TBK if col == 'cur' else TBO
    return (f'(SUMIFS({rng},{TBA},{co},{TBD},"损益-收入")-SUMIFS({rng},{TBA},{co},{TBD},"损益-费用"))')

# (缩进标签, 取数项目 或 None, 类型)  类型: h 小标题 / i 明细 / s 合计（公式另给）
BS_L = [
    ('流动资产：', None, 'h'),
    ('　　货币资金', '货币资金', 'i'),
    ('　　应收账款', '应收账款', 'i'),
    ('　　预付款项', '预付款项', 'i'),
    ('　　其他应收款', '其他应收款', 'i'),
    ('　　存货', '存货', 'i'),
    ('流动资产合计', None, 's'),
    ('非流动资产：', None, 'h'),
    ('　　固定资产原价', '固定资产原价', 'i'),
    ('　　减：累计折旧', '累计折旧', 'i'),
    ('　　固定资产净值', None, 's'),
    ('　　生产性生物资产原价', '生产性生物资产原价', 'i'),
    ('　　减：生物资产累计折旧', '生物资产累计折旧', 'i'),
    ('　　生产性生物资产净值', None, 's'),
    ('　　无形资产原价', '无形资产原价', 'i'),
    ('　　减：累计摊销', '累计摊销', 'i'),
    ('　　无形资产净值', None, 's'),
    ('　　长期待摊费用', '长期待摊费用', 'i'),
    ('非流动资产合计', None, 's'),
    ('资产总计', None, 's'),
]
BS_R = [
    ('流动负债：', None, 'h'),
    ('　　短期借款', '短期借款', 'i'),
    ('　　应付账款', '应付账款', 'i'),
    ('　　预收款项', '预收款项', 'i'),
    ('　　应付职工薪酬', '应付职工薪酬', 'i'),
    ('　　应交税费', '应交税费', 'i'),
    ('　　其他应付款', '其他应付款', 'i'),
    ('流动负债合计', None, 's'),
    ('非流动负债：', None, 'h'),
    ('　　长期借款', '长期借款', 'i'),
    ('负债合计', None, 's'),
    ('所有者权益：', None, 'h'),
    ('　　实收资本', '实收资本', 'i'),
    ('　　资本公积', '资本公积', 'i'),
    ('　　盈余公积', '盈余公积', 'i'),
    ('　　未分配利润', None, 's'),
    ('所有者权益合计', None, 's'),
    ('负债和所有者权益总计', None, 's'),
    ('　', None, 'h'),
    ('平衡校验（资产总计 − 负债和所有者权益）', None, 's'),
]
PL_L = [
    ('一、营业收入', '营业收入', 'i'),
    ('　　减：营业成本', '营业成本', 'i'),
    ('　　　　税金及附加', '税金及附加', 'i'),
    ('　　　　销售费用', '销售费用', 'i'),
    ('　　　　管理费用', '管理费用', 'i'),
    ('　　　　研发费用', '研发费用', 'i'),
    ('　　　　财务费用', '财务费用', 'i'),
    ('　　加：投资收益', '投资收益', 'i'),
    ('二、营业利润', None, 's'),
    ('　　加：营业外收入', '营业外收入', 'i'),
    ('　　减：营业外支出', '营业外支出', 'i'),
    ('三、利润总额', None, 's'),
    ('　　减：所得税费用', '所得税费用', 'i'),
    ('四、净利润', None, 's'),
]

# ============================================================ 分公司报表：资产负债表 / 利润表
BS_SHEETS, PL_SHEETS = {}, {}
for ci, (short, full, _k, _b) in enumerate(COS):
    co_ref = f'{SH_PARAM}!$A${10+ci}'
    # ---------------- 资产负债表 ----------------
    name = f'{short}-资产负债表'
    BS_SHEETS[short] = name
    ws = wb.create_sheet(name)
    title(ws, f'{full} · 资产负债表', 'G',
          '「年初数」取期初余额表，「期末数」按下面的截止日期算到那一天。'
          '未分配利润＝年初未分配利润＋本年净利润（报表自动结转，不用自己做结转损益分录）。')
    BD_S, BD_E, _ = filter_band(ws, 'G')
    widths(ws, {'A':32,'B':17,'C':17,'D':2,'E':32,'F':17,'G':17})
    headers(ws, HR, 1, ['资产', '年初数', '期末数'])
    put(ws, f'D{HR}', None, font=F_HDR, fill=FILL_HDR)
    headers(ws, HR, 5, ['负债和所有者权益', '年初数', '期末数'])
    rowof = {}
    for side, items, (lc, bc, ec) in (('L', BS_L, ('A', 'B', 'C')), ('R', BS_R, ('E', 'F', 'G'))):
        for i, (lab, item, kind) in enumerate(items):
            r = 6 + i
            rowof[(side, lab.strip('　'))] = r
            bold = kind != 'i'
            put(ws, f'{lc}{r}', lab, font=F_TOT if bold else F_TXT, align=CL,
                fill=FILL_TOT if kind == 's' else (FILL_HDR2 if kind == 'h' else None))
            for cc in (bc, ec):
                put(ws, f'{cc}{r}', None, font=F_TOT if bold else F_TXT, fmt=MONEY,
                    fill=FILL_TOT if kind == 's' else None)
            if kind == 'i':
                put(ws, f'{bc}{r}', f'={tb_beg(co_ref, item)}', font=F_TXT, fmt=MONEY)
                put(ws, f'{ec}{r}', f'={tb_end(co_ref, item)}', font=F_TXT, fmt=MONEY)
    RL, RR = (lambda k: rowof[('L', k)]), (lambda k: rowof[('R', k)])
    for cc, side in (('B', 'beg'), ('C', 'end')):
        put(ws, f'{cc}{RL("流动资产合计")}', f'=SUM({cc}{RL("货币资金")}:{cc}{RL("存货")})', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
        put(ws, f'{cc}{RL("固定资产净值")}', f'={cc}{RL("固定资产原价")}-{cc}{RL("减：累计折旧")}', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
        put(ws, f'{cc}{RL("生产性生物资产净值")}', f'={cc}{RL("生产性生物资产原价")}-{cc}{RL("减：生物资产累计折旧")}', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
        put(ws, f'{cc}{RL("无形资产净值")}', f'={cc}{RL("无形资产原价")}-{cc}{RL("减：累计摊销")}', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
        put(ws, f'{cc}{RL("非流动资产合计")}', f'={cc}{RL("固定资产净值")}+{cc}{RL("生产性生物资产净值")}'
                                            f'+{cc}{RL("无形资产净值")}+{cc}{RL("长期待摊费用")}', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
        put(ws, f'{cc}{RL("资产总计")}', f'={cc}{RL("流动资产合计")}+{cc}{RL("非流动资产合计")}', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
    for cc, side in (('F', 'beg'), ('G', 'end')):
        put(ws, f'{cc}{RR("流动负债合计")}', f'=SUM({cc}{RR("短期借款")}:{cc}{RR("其他应付款")})', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
        put(ws, f'{cc}{RR("负债合计")}', f'={cc}{RR("流动负债合计")}+{cc}{RR("长期借款")}', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
        prof = net_profit(co_ref, 'ytd') if side == 'end' else '0'
        put(ws, f'{cc}{RR("未分配利润")}',
            f'={tb_beg(co_ref, "未分配利润") if side == "beg" else tb_end(co_ref, "未分配利润")}+{prof}',
            font=F_TOT, fmt=MONEY, fill=FILL_TOT)
        put(ws, f'{cc}{RR("所有者权益合计")}', f'=SUM({cc}{RR("实收资本")}:{cc}{RR("未分配利润")})', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
        put(ws, f'{cc}{RR("负债和所有者权益总计")}', f'={cc}{RR("负债合计")}+{cc}{RR("所有者权益合计")}', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
    bchk = RR('平衡校验（资产总计 − 负债和所有者权益）')
    for cc, ac in (('F', 'B'), ('G', 'C')):
        tot = RL('资产总计'); lia = RR('负债和所有者权益总计')
        put(ws, f'{cc}{bchk}', f'=IF(ROUND({ac}{tot}-{cc}{lia},2)=0,"√ 平衡","✗ 差 "&TEXT({ac}{tot}-{cc}{lia},"#,##0.00"))',
            font=F_TOT, fill=FILL_CHK, align=C, fmt=None)
        ws.conditional_formatting.add(f'{cc}{bchk}', FormulaRule(formula=[f'LEFT(${cc}${bchk},1)="✗"'], fill=FILL_WARN))
    page(ws, landscape=True)
    ws.print_area = f'A1:G{RR("平衡校验（资产总计 − 负债和所有者权益）")}'
    ws.freeze_panes = 'A6'

    # ---------------- 利润表 ----------------
    name = f'{short}-利润表'
    PL_SHEETS[short] = name
    ws = wb.create_sheet(name)
    title(ws, f'{full} · 利润表', 'C',
          '「本期金额」＝起止日期区间内的发生额；「本年累计」＝年初到截止日期的累计。'
          '只填年度不填起止，两列一样。')
    BD_S, BD_E, _ = filter_band(ws, 'C')
    widths(ws, {'A':42,'B':20,'C':20})
    headers(ws, HR, 1, ['项　目', '本期金额', '本年累计'])
    prow = {}
    for i, (lab, item, kind) in enumerate(PL_L):
        r = 6 + i
        prow[lab.strip('　')] = r
        bold = kind == 's'
        put(ws, f'A{r}', lab, font=F_TOT if bold else F_TXT, align=CL, fill=FILL_TOT if bold else None)
        for cc, fn in (('B', pl_cur), ('C', pl_ytd)):
            if kind == 'i':
                put(ws, f'{cc}{r}', f'={fn(co_ref, item)}', font=F_TXT, fmt=MONEY)
            else:
                put(ws, f'{cc}{r}', None, font=F_TOT, fmt=MONEY, fill=FILL_TOT)
    P = lambda k: prow[k]
    for cc in ('B', 'C'):
        put(ws, f'{cc}{P("二、营业利润")}',
            f'={cc}{P("一、营业收入")}-{cc}{P("减：营业成本")}-{cc}{P("税金及附加")}-{cc}{P("销售费用")}'
            f'-{cc}{P("管理费用")}-{cc}{P("研发费用")}-{cc}{P("财务费用")}+{cc}{P("加：投资收益")}',
            font=F_TOT, fmt=MONEY, fill=FILL_TOT)
        put(ws, f'{cc}{P("三、利润总额")}',
            f'={cc}{P("二、营业利润")}+{cc}{P("加：营业外收入")}-{cc}{P("减：营业外支出")}',
            font=F_TOT, fmt=MONEY, fill=FILL_TOT)
        put(ws, f'{cc}{P("四、净利润")}', f'={cc}{P("三、利润总额")}-{cc}{P("减：所得税费用")}',
            font=F_TOT, fmt=MONEY, fill=FILL_TOT)
    NR = P('四、净利润') + 2
    put(ws, f'A{NR}', '毛利率（营业收入−营业成本）÷ 营业收入', font=F_H2, fill=FILL_HDR2, align=CL)
    for cc in ('B', 'C'):
        put(ws, f'{cc}{NR}', f'=IF(ROUND({cc}{P("一、营业收入")},2)=0,"—",'
                             f'({cc}{P("一、营业收入")}-{cc}{P("减：营业成本")})/{cc}{P("一、营业收入")})',
            font=F_TOT, fmt=PCT, fill=FILL_CHK)
    put(ws, f'A{NR+1}', '净利率　净利润 ÷ 营业收入', font=F_H2, fill=FILL_HDR2, align=CL)
    for cc in ('B', 'C'):
        put(ws, f'{cc}{NR+1}', f'=IF(ROUND({cc}{P("一、营业收入")},2)=0,"—",'
                               f'{cc}{P("四、净利润")}/{cc}{P("一、营业收入")})',
            font=F_TOT, fmt=PCT, fill=FILL_CHK)
    page(ws, landscape=False)
    ws.print_area = f'A1:C{NR+1}'
    ws.freeze_panes = 'A6'

# ============================================================ 三公司合并报表
ws = wb.create_sheet(SH_MERGE)
title(ws, '三公司合并报表（含内部交易抵销）', 'G',
      '三家单体数＋小计＋内部抵销＝合并数。内部抵销里，内部销售收入和对应成本、内部应收应付是自动算的；'
      '「未实现内部存货利润」需要你自己估一个数填进去（下面有算法提示），不填就按 0 算。')
MG_S, MG_E, _ = filter_band(ws, 'G')
widths(ws, {'A':30,'B':16,'C':16,'D':16,'E':17,'F':16,'G':17})
CO_REFS = [f'{SH_PARAM}!$A${10+i}' for i in range(len(COS))]
FULL_REFS = [f'{SH_PARAM}!$B${10+i}' for i in range(len(COS))]
IS_ = (f'(SUMIFS({KB("Q")},{KB("X")},"是",{KB("D")},"销售出库")'
       f'-SUMIFS({KB("Q")},{KB("X")},"是",{KB("D")},"销售退回"))')
def inner_bal(acc, plus, minus, openrng):
    parts = []
    for n in FULL_REFS:
        parts.append(f'(SUMIFS({openrng},{OP_PT},{n})'
                     f'+SUMIFS({plus},{VH},"{acc}",{VK},{n},{VE},"<="&{MG_E})'
                     f'-SUMIFS({minus},{VH},"{acc}",{VK},{n},{VE},"<="&{MG_E}))')
    return '(' + '+'.join(parts) + ')'
IR_AR = inner_bal('应收账款', VI, VJ, OP_AR)
IR_AP = inner_bal('应付账款', VJ, VI, OP_AP)
UPR = 5
put(ws, 'A5', '未实现内部存货利润', font=F_H2, fill=FILL_HDR2)
put(ws, 'B5', 0, font=Font(name='微软雅黑', size=10, bold=True, color='0000C0'), fill=FILL_IN, fmt=MONEY)
ws.merge_cells('C5:G5')
put(ws, 'C5', '＝期末还压在买方仓库里的、从内部买来的存货 × 卖方在这批货上的毛利率。'
              '例：木业年末剩的原木是从农业买的，农业卖原木毛利率 50%，剩 125 方×938 元＝11.7 万，'
              '就填 11.7 万×50%≈5.9 万。不做合并报送的话，留 0 也行。', font=F_NOTE, align=CL)
UP = '$B$5'

MG_PL_R0 = 8
headers(ws, 7, 1, ['利润表项目'] + [c[1] for c in COS] + ['小计', '内部抵销', '合并数'])
prow = {}
for i, (lab, item, kind) in enumerate(PL_L):
    r = MG_PL_R0 + i
    prow[lab.strip('　')] = r
    bold = kind == 's'
    put(ws, f'A{r}', lab, font=F_TOT if bold else F_TXT, align=CL, fill=FILL_TOT if bold else None)
    for j, co in enumerate(CO_REFS):
        cc = L(2 + j)
        put(ws, f'{cc}{r}', (f'={pl_ytd(co, item)}' if kind == 'i' else None),
            font=F_TOT if bold else F_TXT, fmt=MONEY, fill=FILL_TOT if bold else None)
    put(ws, f'E{r}', f'=SUM(B{r}:D{r})', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
    elim = '0'
    if item == '营业收入': elim = f'=-{IS_}'
    elif item == '营业成本': elim = f'=-({IS_}-N({UP}))'
    put(ws, f'F{r}', elim if isinstance(elim, str) and elim.startswith('=') else 0,
        font=F_TXT, fmt=MONEY, fill=FILL_IN if not str(elim).startswith('=') else None)
    put(ws, f'G{r}', f'=E{r}+F{r}', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
P = lambda k: prow[k]
for cc in 'BCDEFG':
    put(ws, f'{cc}{P("二、营业利润")}',
        f'={cc}{P("一、营业收入")}-{cc}{P("减：营业成本")}-{cc}{P("税金及附加")}-{cc}{P("销售费用")}'
        f'-{cc}{P("管理费用")}-{cc}{P("研发费用")}-{cc}{P("财务费用")}+{cc}{P("加：投资收益")}',
        font=F_TOT, fmt=MONEY, fill=FILL_TOT)
    put(ws, f'{cc}{P("三、利润总额")}',
        f'={cc}{P("二、营业利润")}+{cc}{P("加：营业外收入")}-{cc}{P("减：营业外支出")}',
        font=F_TOT, fmt=MONEY, fill=FILL_TOT)
    put(ws, f'{cc}{P("四、净利润")}', f'={cc}{P("三、利润总额")}-{cc}{P("减：所得税费用")}',
        font=F_TOT, fmt=MONEY, fill=FILL_TOT)

MERGE_BS = [
    ('资产', None, 'h'),
    ('　货币资金', [('货币资金', 1)], 'i'),
    ('　应收账款', [('应收账款', 1)], 'i'),
    ('　预付款项', [('预付款项', 1)], 'i'),
    ('　其他应收款', [('其他应收款', 1)], 'i'),
    ('　存货', [('存货', 1)], 'i'),
    ('　固定资产净值', [('固定资产原价', 1), ('累计折旧', -1)], 'i'),
    ('　生产性生物资产净值', [('生产性生物资产原价', 1), ('生物资产累计折旧', -1)], 'i'),
    ('　无形资产净值', [('无形资产原价', 1), ('累计摊销', -1)], 'i'),
    ('　长期待摊费用', [('长期待摊费用', 1)], 'i'),
    ('资产总计', None, 'sum'),
    ('负债', None, 'h'),
    ('　短期借款', [('短期借款', 1)], 'i'),
    ('　应付账款', [('应付账款', 1)], 'i'),
    ('　预收款项', [('预收款项', 1)], 'i'),
    ('　应付职工薪酬', [('应付职工薪酬', 1)], 'i'),
    ('　应交税费', [('应交税费', 1)], 'i'),
    ('　其他应付款', [('其他应付款', 1)], 'i'),
    ('　长期借款', [('长期借款', 1)], 'i'),
    ('负债合计', None, 'sum'),
    ('所有者权益', None, 'h'),
    ('　实收资本', [('实收资本', 1)], 'i'),
    ('　资本公积', [('资本公积', 1)], 'i'),
    ('　盈余公积', [('盈余公积', 1)], 'i'),
    ('　未分配利润', [('未分配利润', 1)], 'p'),
    ('所有者权益合计', None, 'sum'),
    ('负债和所有者权益总计', None, 'tot'),
    ('平衡校验（资产总计 − 负债和所有者权益）', None, 'chk'),
]
MG_BS_R0 = P('四、净利润') + 3
headers(ws, MG_BS_R0 - 1, 1, ['资产负债表项目'] + [c[1] for c in COS] + ['小计', '内部抵销', '合并数'])
brow, block = {}, []
blocks = {}
for i, (lab, items, kind) in enumerate(MERGE_BS):
    r = MG_BS_R0 + i
    key = lab.strip('　')
    brow[key] = r
    if kind == 'h': block = []
    elif kind in ('i', 'p'): block.append(r)
    elif kind == 'sum': blocks[key] = list(block)
    bold = kind not in ('i', 'p')
    put(ws, f'A{r}', lab, font=F_TOT if bold else F_TXT, align=CL,
        fill=FILL_HDR2 if kind == 'h' else (FILL_TOT if bold else None))
    for j, co in enumerate(CO_REFS):
        cc = L(2 + j)
        if kind in ('i', 'p'):
            f = '+'.join(f'{"" if sg > 0 else "-"}{tb_end(co, it)}' for it, sg in items).replace('+-', '-')
            if kind == 'p': f += f'+{net_profit(co, "ytd")}'
            put(ws, f'{cc}{r}', f'={f}', font=F_TXT, fmt=MONEY)
        else:
            put(ws, f'{cc}{r}', None, font=F_TOT, fmt=MONEY, fill=FILL_TOT)
    put(ws, f'E{r}', (f'=SUM(B{r}:D{r})' if kind != 'h' else None), font=F_TOT, fmt=MONEY,
        fill=FILL_TOT if kind != 'h' else FILL_HDR2)
    el = None
    if key == '应收账款': el = f'=-{IR_AR}'
    elif key == '应付账款': el = f'=-{IR_AP}'
    elif key == '存货': el = f'=-N({UP})'
    elif key == '未分配利润': el = f'=-N({UP})'
    put(ws, f'F{r}', el if el else (0 if kind in ('i', 'p') else None), font=F_TXT, fmt=MONEY,
        fill=FILL_TOT if bold and kind != 'h' else (FILL_HDR2 if kind == 'h' else None))
    put(ws, f'G{r}', (f'=E{r}+F{r}' if kind != 'h' else None), font=F_TOT, fmt=MONEY,
        fill=FILL_TOT if kind != 'h' else FILL_HDR2)
B = lambda k: brow[k]
for cc in 'BCDEFG':
    for key in ('资产总计', '负债合计', '所有者权益合计'):
        put(ws, f'{cc}{B(key)}', '=' + '+'.join(f'{cc}{x}' for x in blocks[key]),
            font=F_TOT, fmt=MONEY, fill=FILL_TOT)
    put(ws, f'{cc}{B("负债和所有者权益总计")}', f'={cc}{B("负债合计")}+{cc}{B("所有者权益合计")}',
        font=F_TOT, fmt=MONEY, fill=FILL_TOT)
    put(ws, f'{cc}{B("平衡校验（资产总计 − 负债和所有者权益）")}',
        f'=IF(ROUND({cc}{B("资产总计")}-{cc}{B("负债和所有者权益总计")},2)=0,"√ 平衡",'
        f'"✗ 差 "&TEXT({cc}{B("资产总计")}-{cc}{B("负债和所有者权益总计")},"#,##0.00"))',
        font=F_TOT, fill=FILL_CHK, fmt=None)
MG_LAST = B('平衡校验（资产总计 − 负债和所有者权益）')
put(ws, f'A{MG_LAST+2}', '内部交易抵销明细', font=F_H2, fill=FILL_HDR2)
for i, (lab, f) in enumerate([('内部销售收入（含税）', f'={IS_}'),
                              ('内部应收余额', f'={IR_AR}'),
                              ('内部应付余额', f'={IR_AP}'),
                              ('未实现内部存货利润（手工填，上面 B5）', f'=N({UP})')]):
    r = MG_LAST + 3 + i
    ws.merge_cells(f'A{r}:D{r}')
    put(ws, f'A{r}', lab, font=F_TXT, align=CL)
    ws.merge_cells(f'E{r}:G{r}')
    put(ws, f'E{r}', f, font=F_TOT, fmt=MONEY, fill=FILL_CHK)
page(ws)
ws.freeze_panes = 'A8'

# ============================================================ 往来台账
ws = wb.create_sheet(SH_AR)
AR_0, AR_1 = 7, 7 + (PT_1 - PT_0)
title(ws, '往来台账（应收 / 应付）', 'L',
      '按上面选的公司出这家的往来。应收＝购销开票挂上去的，减掉【资金流水】里「收客户货款」冲的；'
      '应付同理。年初数取【期初余额】右侧那张按单位拆的表。')
AR_S, AR_E, AR_CO = filter_band(ws, 'L', co_default=CO_NAMES[1])
widths(ws, {'A':22,'B':12,'C':15,'D':15,'E':15,'F':15,'G':15,'H':15,'I':15,'J':15,'K':15,'L':6})
headers(ws, HR, 1, ['往来单位','类型','年初应收','本期应收发生','本期收款','期末应收',
                    '年初应付','本期应付发生','本期付款','期末应付','净往来（收−付）','有'])
put(ws, 'A6', '合  计', font=F_TOT, fill=FILL_TOT)
put(ws, 'B6', None, font=F_TOT, fill=FILL_TOT)
for c in 'CDEFGHIJK':
    put(ws, f'{c}6', f'=SUM({c}${AR_0}:{c}${AR_1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, 'L6', None, font=F_TOT, fill=FILL_TOT)
for i in range(PT_1 - PT_0 + 1):
    r, pr = AR_0 + i, PT_0 + i
    put(ws, f'A{r}', f'=IF({SH_PT}!$A{pr}="","",{SH_PT}!$A{pr})', font=F_TXT, align=CL)
    put(ws, f'B{r}', f'=IF($A{r}="","",{SH_PT}!$B{pr})', font=F_TXT)
    for (cc, acc, plus, minus, openrng, dc, ec) in (
            ('C', '应收账款', VI, VJ, OP_AR, 'D', 'E'), ('G', '应付账款', VJ, VI, OP_AP, 'H', 'I')):
        pre = (f'SUMIFS({plus},{VF},{AR_CO},{VH},"{acc}",{VK},$A{r},{VE},"<"&{AR_S})'
               f'-SUMIFS({minus},{VF},{AR_CO},{VH},"{acc}",{VK},$A{r},{VE},"<"&{AR_S})')
        put(ws, f'{cc}{r}', f'=IF($A{r}="","",SUMIFS({openrng},{OP_CO},{AR_CO},{OP_PT},$A{r})+{pre})',
            font=F_LINK, fmt=MONEY)
        rng = f',{VE},">="&{AR_S},{VE},"<="&{AR_E}'
        put(ws, f'{dc}{r}', f'=IF($A{r}="","",SUMIFS({plus},{VF},{AR_CO},{VH},"{acc}",{VK},$A{r}{rng}))',
            font=F_LINK, fmt=MONEY)
        put(ws, f'{ec}{r}', f'=IF($A{r}="","",SUMIFS({minus},{VF},{AR_CO},{VH},"{acc}",{VK},$A{r}{rng}))',
            font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",ROUND($C{r}+$D{r}-$E{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",ROUND($G{r}+$H{r}-$I{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'K{r}', f'=IF($A{r}="","",$F{r}-$J{r})', font=F_TOT, fmt=MONEY)
    put(ws, f'L{r}', f'=IF($A{r}="","",IF(OR(ROUND($C{r},2)<>0,ROUND($D{r},2)<>0,ROUND($E{r},2)<>0,'
                     f'ROUND($G{r},2)<>0,ROUND($H{r},2)<>0,ROUND($I{r},2)<>0),"有",""))', font=F_NOTE)
hide_tail(ws, AR_0, AR_1, [AR_0 + i for i in range(len(PARTNERS))],
          filter_col=col_idx('L') - 1, ref=f'A{HR}:L{AR_1}')
put(ws, 'A2', HIDE_NOTE, font=F_NOTE, align=CL, border=None)
page(ws, titles=f'{HR}:6')
ws.freeze_panes = f'B{AR_0}'

# ============================================================ 税费台账
ws = wb.create_sheet(SH_TAX)
title(ws, '税费台账（分月实缴 · 全年）', 'K',
      '内账按实交税费计入费用：这张表就是【资金流水】里「缴纳税费」那些行，按公司、按月、按税种摊开。'
      '税种取的是那一行选的「费用项目」。右边放了当月含税收入，合计行给出全年税负率。'
      '这张表不随日期筛选变，看的就是全年十二个月。')
widths(ws, {'A':10,'B':11,'C':15,'D':16,'E':13,'F':15,'G':15,'H':13,'I':15,'J':16,'K':24})
headers(ws, HR, 1, ['公司','月份'] + TAX_ITEMS + ['其他税费','本月合计','本月含税收入','说明'])
TX_0 = 6
r = TX_0
TAXC = [L(3 + i) for i in range(len(TAX_ITEMS))]          # C D E F G
OTHC, SUMC, REVC = L(3 + len(TAX_ITEMS)), L(4 + len(TAX_ITEMS)), L(5 + len(TAX_ITEMS))
for ci, (short, full, _k, _b) in enumerate(COS):
    co = f'{SH_PARAM}!$A${10+ci}'
    r0 = r
    for m in range(1, 13):
        m1, m2 = f'DATE({PARAM_Y},{m},1)', f'DATE({PARAM_Y},{m+1},0)'
        cd = f',{KC("B")},">="&{m1},{KC("B")},"<="&{m2}'
        bd = f',{KB("B")},">="&{m1},{KB("B")},"<="&{m2}'
        put(ws, f'A{r}', f'={co}', font=F_TXT)
        put(ws, f'B{r}', f'=TEXT({m1},"yyyy-mm")', font=F_TXT)
        for j, t in enumerate(TAX_ITEMS):
            put(ws, f'{TAXC[j]}{r}', f'=SUMIFS({KC("P")},{KC("C")},{co},{KC("E")},"缴纳税费",'
                                     f'{KC("H")},"{t}"{cd})', font=F_LINK, fmt=MONEY)
        put(ws, f'{OTHC}{r}', f'=ROUND(SUMIFS({KC("P")},{KC("C")},{co},{KC("E")},"缴纳税费"{cd})'
                              + ''.join(f'-{c}{r}' for c in TAXC) + ',2)', font=F_LINK, fmt=MONEY)
        put(ws, f'{SUMC}{r}', f'=ROUND(SUM({TAXC[0]}{r}:{OTHC}{r}),2)', font=F_TOT, fmt=MONEY)
        put(ws, f'{REVC}{r}', f'=SUMIFS({KB("Q")},{KB("C")},{co},{KB("D")},"销售出库"{bd})'
                              f'-SUMIFS({KB("Q")},{KB("C")},{co},{KB("D")},"销售退回"{bd})'
                              f'+SUMIFS({KC("P")},{KC("C")},{co},{KC("E")},"现销收款"{cd})'
                              f'+SUMIFS({KC("P")},{KC("C")},{co},{KC("E")},"其他收入"{cd})',
            font=F_LINK, fmt=MONEY)
        put(ws, f'K{r}', None, font=F_NOTE, align=CL)
        r += 1
    put(ws, f'A{r}', f'={co}', font=F_TOT, fill=FILL_TOT)
    put(ws, f'B{r}', '全年合计', font=F_TOT, fill=FILL_TOT)
    for c in TAXC + [OTHC, SUMC, REVC]:
        put(ws, f'{c}{r}', f'=SUM({c}{r0}:{c}{r-1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
    put(ws, f'K{r}', f'=IF(ROUND(${REVC}{r},2)=0,"本年无含税收入",'
                     f'"全年税负率 "&TEXT(${SUMC}{r}/${REVC}{r},"0.00%")&"（实缴税费 ÷ 含税收入）")',
        font=F_TOT, fill=FILL_CHK, align=CL)
    r += 2
TAX_END = r
put(ws, f'A{TAX_END}', '口径', font=F_H2, fill=FILL_HDR2)
ws.merge_cells(f'B{TAX_END}:K{TAX_END}')
put(ws, f'B{TAX_END}', '内账不计提应交税费，也不拆进项销项——交多少就是多少，缴的当月直接进「税金及附加」'
                       '或「所得税费用」。想看应交未交，那是外账口径，本表不做。', font=F_NOTE, align=CL)
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'C{TX_0}'

# ============================================================ 费用统计
ws = wb.create_sheet(SH_EXP)
title(ws, '费用统计（按起止日期）', 'F',
      '上半张按「费用项目」统计：【资金流水】里「费用支出/其他支出/支付利息」三类的不含税金额'
      '＋【生产加工】的加工费＋【其他分录】的计提数。'
      '发放工资、缴纳税费、还款这些是付钱不是费用（计提时已经算过一次），所以不重复计入；'
      '下半张按会计科目统计，直接取科目余额表的本期发生额。两张的口径不同：'
      '上面是按你自己分的项目看钱花在哪，下面是按报表科目看。')
EXQ_S, EXQ_E, _ = filter_band(ws, 'F')
widths(ws, {'A':20,'B':17,'C':17,'D':17,'E':18,'F':11})
EXQ_0 = 7
EXQ_1 = EXQ_0 + (EXP_END - 16)
headers(ws, HR, 1, ['费用项目'] + [c[1] for c in COS] + ['三家合计', '有'])
put(ws, 'A6', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BCDE':
    put(ws, f'{c}6', f'=SUM({c}${EXQ_0}:{c}${EXQ_1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, 'F6', None, font=F_TOT, fill=FILL_TOT)
for i in range(EXQ_1 - EXQ_0 + 1):
    r, pr = EXQ_0 + i, 16 + i
    put(ws, f'A{r}', f'=IF({SH_PARAM}!$A{pr}="","",{SH_PARAM}!$A{pr})', font=F_TXT, align=CL)
    for j, co in enumerate(CO_REFS):
        cc = L(2 + j)
        # 资金流水只取真正形成费用的三类，「发放工资 / 缴纳税费 / 还款」是付钱不是费用，
        # 计提已经在生产加工和其他分录里算过一次了，再算一次就重了
        f = ('+'.join(
                f'SUMIFS({KC("P")},{KC("C")},{co},{KC("H")},$A{r},{KC("E")},"{t}",'
                f'{KC("B")},">="&{EXQ_S},{KC("B")},"<="&{EXQ_E})'
                for t in ('费用支出', '其他支出', '支付利息'))
             + f'+SUMIFS({KD("J")},{KD("C")},{co},{KD("L")},$A{r},'
               f'{KD("B")},">="&{EXQ_S},{KD("B")},"<="&{EXQ_E})'
             + f'+SUMIFS({KO("G")},{KO("C")},{co},{KO("I")},$A{r},'
               f'{KO("B")},">="&{EXQ_S},{KO("B")},"<="&{EXQ_E})')
        put(ws, f'{cc}{r}', f'=IF($A{r}="","",{f})', font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', f'=IF($A{r}="","",SUM(B{r}:D{r}))', font=F_TOT, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",IF(ROUND($E{r},2)<>0,"有",""))', font=F_NOTE)
hide_tail(ws, EXQ_0, EXQ_1, [EXQ_0 + i for i in range(len(EXPENSES))],
          filter_col=col_idx('F') - 1, ref=f'A{HR}:F{EXQ_1}')
EXA_0 = EXQ_1 + 2
EXP_ACCS = ['主营业务成本', '税金及附加', '销售费用', '管理费用', '财务费用', '研发费用', '营业外支出', '所得税费用']
hdr(ws, f'A{EXA_0}', '按会计科目（本期发生额）', span=f'A{EXA_0}:F{EXA_0}', font=F_H2, fill=FILL_HDR2)
headers(ws, EXA_0 + 1, 1, ['会计科目'] + [c[1] for c in COS] + ['三家合计', ''])
for i, acc in enumerate(EXP_ACCS):
    r = EXA_0 + 2 + i
    put(ws, f'A{r}', acc, font=F_TXT, align=CL)
    for j, co in enumerate(CO_REFS):
        put(ws, f'{L(2+j)}{r}', f'=SUMIFS({TBK},{TBA},{co},{TBC},"{acc}")', font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', f'=SUM(B{r}:D{r})', font=F_TOT, fmt=MONEY)
    put(ws, f'F{r}', None)
r = EXA_0 + 2 + len(EXP_ACCS)
put(ws, f'A{r}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BCDE':
    put(ws, f'{c}{r}', f'=SUM({c}{EXA_0+2}:{c}{r-1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'F{r}', None, font=F_TOT, fill=FILL_TOT)
put(ws, 'A2', HIDE_NOTE, font=F_NOTE, align=CL, border=None)
page(ws, titles=f'{HR}:6')
ws.freeze_panes = f'B{EXQ_0}'

# ============================================================ 内部交易核对
ws = wb.create_sheet(SH_IC)
title(ws, '内部交易核对', 'I',
      '三家之间买卖的，卖方开多少、买方入多少必须一样——内账含税核算，两边都是同一个含税金额，'
      '差一分都要查。右边再对一次期末内部往来（卖方应收 vs 买方应付）。')
IC_S, IC_E, _ = filter_band(ws, 'I')
widths(ws, {'A':16,'B':16,'C':18,'D':18,'E':14,'F':18,'G':18,'H':14,'I':24})
headers(ws, HR, 1, ['卖方公司','买方公司','卖方开票金额（含税）','买方入账金额（含税）','购销差额',
                    '期末卖方应收','期末买方应付','往来差额','结论'])
pairs = [(i, j) for i in range(len(COS)) for j in range(len(COS)) if i != j]
for n, (si, bi) in enumerate(pairs):
    r = IC_0 + n
    sc, bc = CO_REFS[si], CO_REFS[bi]
    sf, bf = FULL_REFS[si], FULL_REFS[bi]
    dr = f',{KB("B")},">="&{IC_S},{KB("B")},"<="&{IC_E}'
    put(ws, f'A{r}', f'={sc}', font=F_TXT)
    put(ws, f'B{r}', f'={bc}', font=F_TXT)
    put(ws, f'C{r}', f'=SUMIFS({KB("Q")},{KB("C")},{sc},{KB("E")},{bf},{KB("D")},"销售出库"{dr})'
                     f'-SUMIFS({KB("Q")},{KB("C")},{sc},{KB("E")},{bf},{KB("D")},"销售退回"{dr})',
        font=F_LINK, fmt=MONEY)
    put(ws, f'D{r}', f'=SUMIFS({KB("Q")},{KB("C")},{bc},{KB("E")},{sf},{KB("D")},"采购入库"{dr})'
                     f'-SUMIFS({KB("Q")},{KB("C")},{bc},{KB("E")},{sf},{KB("D")},"采购退回"{dr})',
        font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', f'=ROUND($C{r}-$D{r},2)', font=F_TOT, fmt=MONEY)
    ar = (f'(SUMIFS({OP_AR},{OP_CO},{sc},{OP_PT},{bf})'
          f'+SUMIFS({VI},{VF},{sc},{VH},"应收账款",{VK},{bf},{VE},"<="&{IC_E})'
          f'-SUMIFS({VJ},{VF},{sc},{VH},"应收账款",{VK},{bf},{VE},"<="&{IC_E}))')
    ap = (f'(SUMIFS({OP_AP},{OP_CO},{bc},{OP_PT},{sf})'
          f'+SUMIFS({VJ},{VF},{bc},{VH},"应付账款",{VK},{sf},{VE},"<="&{IC_E})'
          f'-SUMIFS({VI},{VF},{bc},{VH},"应付账款",{VK},{sf},{VE},"<="&{IC_E}))')
    put(ws, f'F{r}', f'={ar}', font=F_LINK, fmt=MONEY)
    put(ws, f'G{r}', f'={ap}', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=ROUND($F{r}-$G{r},2)', font=F_TOT, fmt=MONEY)
    put(ws, f'I{r}', f'=IF(AND(ROUND($C{r},2)=0,ROUND($D{r},2)=0),"本期无内部交易",'
                     f'IF(AND(ROUND($E{r},2)=0,ROUND($H{r},2)=0),"√ 两边一致",'
                     f'IF(ROUND($E{r},2)<>0,"✗ 开票与入账差 "&TEXT($E{r},"#,##0.00"),'
                     f'"✗ 往来余额差 "&TEXT($H{r},"#,##0.00"))))', font=F_TOT, align=CL)
    ws.conditional_formatting.add(f'I{r}', FormulaRule(formula=[f'LEFT($I{r},1)="✗"'], fill=FILL_WARN))
IC_END = IC_0 + len(pairs) - 1
page(ws)
ws.freeze_panes = f'C{IC_0}'

# ============================================================ 校验中心
PL_ROW  = {lab.strip('　'): 6 + i for i, (lab, _a, _b) in enumerate(PL_L)}
BSL_ROW = {lab.strip('　'): 6 + i for i, (lab, _a, _b) in enumerate(BS_L)}
BSR_ROW = {lab.strip('　'): 6 + i for i, (lab, _a, _b) in enumerate(BS_R)}

ws = wb.create_sheet(SH_CHK)
title(ws, '校验中心', 'G',
      '这张表自己告诉你数对不对。全绿就可以出报表；有红的先把红的改掉。'
      '涉及报表的几项，取的是各报表页当前选的期间，默认是整年。')
widths(ws, {'A':6,'B':34,'C':12,'D':18,'E':18,'F':16,'G':46})
headers(ws, HR, 1, ['序号', '检查项', '公司', '左值', '右值', '结果', '说明'])
CK = []
for ci, (short, full, _k, _b) in enumerate(COS):
    co = f'{SH_PARAM}!$A${10+ci}'
    CK.append(('期初借贷是否平衡', short,
               f'=SUMIFS({O_DR},{O_CO},{co})', f'=SUMIFS({O_CR},{O_CO},{co})', 'eq',
               '期初余额表借方合计应等于贷方合计'))
    CK.append(('期初存货金额是否对上科目', short,
               f'=SUMIFS({OI_A},{OI_CO},{co})',
               f'=SUMIFS({O_DR},{O_CO},{co},{O_ACC},"原材料")+SUMIFS({O_DR},{O_CO},{co},{O_ACC},"库存商品")',
               'eq', '右边存货明细金额合计 vs 左边原材料＋库存商品借方余额'))
    CK.append(('期初应收是否对上科目', short,
               f'=SUMIFS({OP_AR},{OP_CO},{co})',
               f'=SUMIFS({O_DR},{O_CO},{co},{O_ACC},"应收账款")', 'eq', '按单位拆的年初应收 vs 应收账款科目'))
    CK.append(('期初应付是否对上科目', short,
               f'=SUMIFS({OP_AP},{OP_CO},{co})',
               f'=SUMIFS({O_CR},{O_CO},{co},{O_ACC},"应付账款")', 'eq', '按单位拆的年初应付 vs 应付账款科目'))
    CK.append(('本年试算平衡（借＝贷）', short,
               f'=SUMIFS({VI},{VF},{co})', f'=SUMIFS({VJ},{VF},{co})', 'eq',
               '记账分录里这家公司的借方合计应等于贷方合计'))
    CK.append(('资产 ＝ 负债＋所有者权益', short,
               f"='{BS_SHEETS[short]}'!$C${BSL_ROW['资产总计']}",
               f"='{BS_SHEETS[short]}'!$G${BSR_ROW['负债和所有者权益总计']}", 'eq',
               '取该公司资产负债表当前期末数'))
    CK.append(('利润表净利 ＝ 收入−成本−费用', short,
               f"='{PL_SHEETS[short]}'!$C${PL_ROW['四、净利润']}",
               f'={net_profit(co, "ytd")}', 'eq', '利润表本年累计净利 vs 按损益类科目直接算'))
CK.append(('账户余额合计 ＝ 三家货币资金', '全部',
           f'={SH_ACCT}!$K${AC_1+2}', f'={CO_CASH}', 'eq',
           '所有账户余额加总，应等于三家库存现金＋银行存款'))
CK.append(('三家试算平衡（借＝贷）', '全部',
           f'=SUM({VI})', f'=SUM({VJ})', 'eq', '整本账的借贷合计'))
CNT = [
    ('购销流水有没有报错行', f'=SUMPRODUCT(--({SH_BUY}!$P${BY_0}:$P${BY_1}<>"OK"),--({SH_BUY}!$P${BY_0}:$P${BY_1}<>""))',
     '校验列不是 OK 的行数，应为 0'),
    ('资金流水有没有报错行', f'=SUMPRODUCT(--({SH_CASH}!$N${CS_0}:$N${CS_1}<>"OK"),--({SH_CASH}!$N${CS_0}:$N${CS_1}<>""))',
     '校验列不是 OK 的行数，应为 0'),
    ('生产加工有没有报错行', f'=SUMPRODUCT(--({SH_PROD}!$P${PD_0}:$P${PD_1}<>"OK"),--({SH_PROD}!$P${PD_0}:$P${PD_1}<>""))',
     '校验列不是 OK 的行数，应为 0'),
    ('其他分录有没有报错行', f'=SUMPRODUCT(--({SH_OTH}!$K${OT_0}:$K${OT_1}<>"OK"),--({SH_OTH}!$K${OT_0}:$K${OT_1}<>""))',
     '校验列不是 OK 的行数，应为 0'),
    ('期初余额有没有报错行', f'=SUMPRODUCT(--({SH_OPEN}!$F${OB_0}:$F${OB_1}<>"OK"),--({SH_OPEN}!$F${OB_0}:$F${OB_1}<>""))'
     f'+SUMPRODUCT(--({SH_OPEN}!$N${OI_0}:$N${OI_1}<>"OK"),--({SH_OPEN}!$N${OI_0}:$N${OI_1}<>""))', '应为 0'),
    ('基础档案有没有报错行', f'=SUMPRODUCT(--({SH_GD}!$N${GD_0}:$N${GD_1}<>"OK"),--({SH_GD}!$N${GD_0}:$N${GD_1}<>""))'
     f'+SUMPRODUCT(--({SH_PT}!$I${PT_0}:$I${PT_1}<>"OK"),--({SH_PT}!$I${PT_0}:$I${PT_1}<>""))'
     f'+SUMPRODUCT(--({SH_ACCT}!$L${AC_0}:$L${AC_1}<>"OK"),--({SH_ACCT}!$L${AC_0}:$L${AC_1}<>""))',
     '商品档案＋往来单位＋资金账户，应为 0'),
    ('存货有没有负结存', f'=COUNTIF({SH_INV}!$V${GD_0}:$V${GD_1},"✗*")',
     '结存数量或金额为负的品种数，应为 0；有就是卖多了或领多了'),
    ('内部交易两边对不对得上', f'=COUNTIF({SH_IC}!$I${IC_0}:$I${IC_END},"✗*")',
     '开票额与入账额、往来余额对不上的对数，应为 0'),
    ('销售有没有漏结转成本', f'=COUNTIFS({SH_BUY}!$D${BY_0}:$D${BY_1},"销售出库",'
     f'{SH_BUY}!$W${BY_0}:$W${BY_1},0)', '有销售但成本金额算出来是 0 的行数，一般是存货没入库'),
    ('会计科目有没有漏配报表项目', f'=COUNTIFS({SH_ACC}!$B${ACC_0}:$B${ACC_1},"<>",{SH_ACC}!$F${ACC_0}:$F${ACC_1},"")',
     '报表项目留空的科目，报表会取不到它，应为 0'),
    ('有没有账户余额透支', f'=COUNTIF({SH_ACCT}!$L${AC_0}:$L${AC_1},"✗*")', '期末余额为负的账户数，应为 0'),
    ('生产成本有没有挂着没结转', f'=ROUND(SUMIFS({TBL},{TBC},"生产成本"),2)',
     '三家生产成本科目的期末余额；批次全部结转完应为 0，还有在产品就不为 0'),
]
CHK_0 = 6
r = CHK_0
for i, (lab, co, lf, rf, kind, note) in enumerate(CK):
    put(ws, f'A{r}', i + 1, font=F_NOTE, fmt='0')
    put(ws, f'B{r}', lab, font=F_TXT, align=CL)
    put(ws, f'C{r}', co, font=F_TXT)
    put(ws, f'D{r}', lf, font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', rf, font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF(ROUND($D{r}-$E{r},2)=0,"√ 一致","✗ 差 "&TEXT($D{r}-$E{r},"#,##0.00"))',
        font=F_TOT, fill=FILL_CHK, align=C)
    put(ws, f'G{r}', note, font=F_NOTE, align=CL)
    ws.conditional_formatting.add(f'F{r}', FormulaRule(formula=[f'LEFT($F{r},1)="✗"'], fill=FILL_WARN))
    r += 1
for j, (lab, f, note) in enumerate(CNT):
    put(ws, f'A{r}', len(CK) + j + 1, font=F_NOTE, fmt='0')
    put(ws, f'B{r}', lab, font=F_TXT, align=CL)
    put(ws, f'C{r}', '全部', font=F_TXT)
    put(ws, f'D{r}', f, font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', 0, font=F_NOTE, fmt=MONEY)
    put(ws, f'F{r}', f'=IF(ROUND($D{r},2)=0,"√ 没问题","✗ "&TEXT($D{r},"#,##0.##")&" 处")',
        font=F_TOT, fill=FILL_CHK, align=C)
    put(ws, f'G{r}', note, font=F_NOTE, align=CL)
    ws.conditional_formatting.add(f'F{r}', FormulaRule(formula=[f'LEFT($F{r},1)="✗"'], fill=FILL_WARN))
    r += 1
CHK_1 = r - 1
put(ws, f'B{r+1}', '汇总', font=F_H2, fill=FILL_HDR2)
put(ws, f'C{r+1}', None, font=F_TOT, fill=FILL_CHK)
put(ws, f'D{r+1}', f'=COUNTIF($F${CHK_0}:$F${CHK_1},"√*")', font=F_TOT, fill=FILL_CHK, fmt='0')
put(ws, f'E{r+1}', f'=COUNTIF($F${CHK_0}:$F${CHK_1},"✗*")', font=F_TOT, fill=FILL_CHK, fmt='0')
put(ws, f'F{r+1}', f'=IF($E${r+1}=0,"√ 全部通过","✗ 有 "&$E${r+1}&" 项要处理")',
    font=F_TOT, fill=FILL_CHK, align=C)
put(ws, f'G{r+1}', '左边是通过数，中间是未通过数', font=F_NOTE, align=CL)
ws.conditional_formatting.add(f'F{r+1}', FormulaRule(formula=[f'LEFT($F${r+1},1)="✗"'], fill=FILL_WARN))
CHK_SUM = r + 1
page(ws, titles=f'{HR}:{HR}')
ws.freeze_panes = f'C{CHK_0}'

# ============================================================ 首页
ws = wb.create_sheet(SH_HOME)
title(ws, f'三公司财务账套 · {YEAR} 年', 'H',
      '农业基地 / 木业 / 科技，一套表三家账。录【购销流水】【资金流水】【生产加工】【其他分录】四张表，'
      '报表自己出。一年一套：把【参数设置】的会计年度改成下一年，另存一个文件就是新账。')
widths(ws, {'A':22,'B':18,'C':18,'D':18,'E':18,'F':4,'G':24,'H':40})
put(ws, 'A4', '会计年度', font=F_H2, fill=FILL_HDR2)
put(ws, 'B4', f'={PARAM_Y}', font=F_BIG, fill=FILL_CARD, fmt='0')
put(ws, 'C4', '本年凭证数', font=F_H2, fill=FILL_HDR2)
put(ws, 'D4', f'=COUNTIF({SH_VOU}!$P${V_0}:$P${V_1},"有")', font=F_BIG, fill=FILL_CARD, fmt='#,##0')
ws.merge_cells('E4:H4')
put(ws, 'E4', f'=IF({SH_CHK}!$E${CHK_SUM}=0,"√ 全部校验通过，可以出报表",'
              f'"✗ 校验中心还有 "&{SH_CHK}!$E${CHK_SUM}&" 项没过，先去【校验中心】看红色那几行")',
    font=F_TOT, fill=FILL_CHK, align=CL)
ws.row_dimensions[4].height = 26

hdr(ws, 'A6', '一、三家公司关键数（口径＝各报表页当前选的期间，默认整年）', span='A6:H6',
    font=F_H2, fill=FILL_HDR2)
headers(ws, 7, 1, ['指　标'] + [c[1] for c in COS] + ['三家合计'])
HOME_ITEMS = [
    ('营业收入',   lambda s: f"'{PL_SHEETS[s]}'!$C${PL_ROW['一、营业收入']}"),
    ('营业成本',   lambda s: f"'{PL_SHEETS[s]}'!$C${PL_ROW['减：营业成本']}"),
    ('毛利',       None),
    ('营业利润',   lambda s: f"'{PL_SHEETS[s]}'!$C${PL_ROW['二、营业利润']}"),
    ('利润总额',   lambda s: f"'{PL_SHEETS[s]}'!$C${PL_ROW['三、利润总额']}"),
    ('净利润',     lambda s: f"'{PL_SHEETS[s]}'!$C${PL_ROW['四、净利润']}"),
    ('资产总计',   lambda s: f"'{BS_SHEETS[s]}'!$C${BSL_ROW['资产总计']}"),
    ('货币资金',   lambda s: f"'{BS_SHEETS[s]}'!$C${BSL_ROW['货币资金']}"),
    ('存货',       lambda s: f"'{BS_SHEETS[s]}'!$C${BSL_ROW['存货']}"),
    ('应收账款',   lambda s: f"'{BS_SHEETS[s]}'!$C${BSL_ROW['应收账款']}"),
    ('应付账款',   lambda s: f"'{BS_SHEETS[s]}'!$G${BSR_ROW['应付账款']}"),
    ('负债合计',   lambda s: f"'{BS_SHEETS[s]}'!$G${BSR_ROW['负债合计']}"),
    ('所有者权益', lambda s: f"'{BS_SHEETS[s]}'!$G${BSR_ROW['所有者权益合计']}"),
]
HR0 = 8
for i, (lab, fn) in enumerate(HOME_ITEMS):
    r = HR0 + i
    bold = lab in ('毛利', '净利润', '资产总计')
    put(ws, f'A{r}', lab, font=F_TOT if bold else F_TXT, align=CL, fill=FILL_TOT if bold else None)
    for j, (short, _f, _k, _b) in enumerate(COS):
        cc = L(2 + j)
        v = (f'={cc}{HR0}-{cc}{HR0+1}' if fn is None else f'={fn(short)}')
        put(ws, f'{cc}{r}', v, font=F_TOT if bold else F_TXT, fmt=MONEY,
            fill=FILL_TOT if bold else FILL_CARD)
    put(ws, f'E{r}', f'=SUM(B{r}:D{r})', font=F_TOT, fmt=MONEY, fill=FILL_TOT)
GM = HR0 + len(HOME_ITEMS)
put(ws, f'A{GM}', '毛利率', font=F_TOT, align=CL, fill=FILL_TOT)
for j in range(len(COS) + 1):
    cc = L(2 + j)
    put(ws, f'{cc}{GM}', f'=IF(ROUND({cc}{HR0},2)=0,"—",{cc}{HR0+2}/{cc}{HR0})',
        font=F_TOT, fmt=PCT, fill=FILL_TOT)

hdr(ws, f'A{GM+2}', '二、常用入口', span=f'A{GM+2}:E{GM+2}', font=F_H2, fill=FILL_HDR2)
NAV = [
    ('① 每天录这四张', [SH_BUY, SH_CASH, SH_PROD, SH_OTH]),
    ('② 先建好档案', [SH_PARAM, SH_GD, SH_PT, SH_ACCT, SH_OPEN, SH_ACC, SH_RULE]),
    ('③ 出报表', [BS_SHEETS['农业'], PL_SHEETS['农业'], BS_SHEETS['木业'], PL_SHEETS['木业'],
                  BS_SHEETS['科技'], PL_SHEETS['科技'], SH_MERGE]),
    ('④ 查台账', [SH_TB, SH_VOU, SH_INV, SH_INVQ, SH_AR, SH_TAX, SH_EXP, SH_IC]),
    ('⑤ 对数', [SH_CHK, SH_HELP]),
]
r = GM + 3
for grp, sheets in NAV:
    put(ws, f'A{r}', grp, font=F_TOT, fill=FILL_TOT, align=CL)
    c = 2
    for sn in sheets:
        if c > 5:
            r += 1; c = 2
            put(ws, f'A{r}', None, font=F_TOT, fill=FILL_TOT)
        cell = put(ws, f'{L(c)}{r}', sn, font=Font(name='微软雅黑', size=9, color='0563C1', underline='single'),
                   fill=FILL_CARD, align=C)
        cell.hyperlink = f"#'{sn}'!A1"
        c += 1
    while c <= 5:
        put(ws, f'{L(c)}{r}', None, fill=FILL_CARD); c += 1
    r += 1

hdr(ws, f'A{r+1}', '三、表格容量（快满了就去【生成脚本】把容量数字改大重跑，或直接新开一年）',
    span=f'A{r+1}:E{r+1}', font=F_H2, fill=FILL_HDR2)
CAPS = [('购销流水', f'COUNT({SH_BUY}!$B${BY_0}:$B${BY_1})', BY_1 - BY_0 + 1),
        ('资金流水', f'COUNT({SH_CASH}!$B${CS_0}:$B${CS_1})', CS_1 - CS_0 + 1),
        ('生产加工', f'COUNT({SH_PROD}!$B${PD_0}:$B${PD_1})', PD_1 - PD_0 + 1),
        ('其他分录', f'COUNT({SH_OTH}!$B${OT_0}:$B${OT_1})', OT_1 - OT_0 + 1),
        ('商品档案', f'COUNTA({SH_GD}!$A${GD_0}:$A${GD_1})', GD_1 - GD_0 + 1),
        ('往来单位', f'COUNTA({SH_PT}!$A${PT_0}:$A${PT_1})', PT_1 - PT_0 + 1)]
r += 2
for lab, cnt, cap in CAPS:
    put(ws, f'A{r}', lab, font=F_TXT, align=CL)
    put(ws, f'B{r}', f'={cnt}', font=F_TOT, fmt='#,##0')
    put(ws, f'C{r}', cap, font=F_TXT, fmt='#,##0')
    put(ws, f'D{r}', f'=IF($C{r}=0,"",$B{r}/$C{r})', font=F_TXT, fmt=PCT)
    put(ws, f'E{r}', f'=IF($B{r}/$C{r}>0.9,"✗ 快满了","√ 够用")', font=F_TOT, fill=FILL_CHK)
    ws.conditional_formatting.add(f'E{r}', FormulaRule(formula=[f'LEFT($E{r},1)="✗"'], fill=FILL_WARN))
    r += 1
put(ws, f'G7', '账户余额', font=F_H2, fill=FILL_HDR2)
put(ws, f'H7', f'={SH_ACCT}!$K${AC_1+2}', font=F_BIG, fill=FILL_CARD, fmt=MONEY)
put(ws, f'G8', '账实核对', font=F_H2, fill=FILL_HDR2)
put(ws, f'H8', f'={SH_ACCT}!$G${AC_1+5}', font=F_TOT, fill=FILL_CHK, align=CL)
put(ws, f'G9', '内部交易', font=F_H2, fill=FILL_HDR2)
put(ws, f'H9', f'=IF(COUNTIF({SH_IC}!$I${IC_0}:$I${IC_END},"✗*")=0,"√ 三家之间对得上",'
               f'"✗ 有 "&COUNTIF({SH_IC}!$I${IC_0}:$I${IC_END},"✗*")&" 对对不上")',
    font=F_TOT, fill=FILL_CHK, align=CL)
put(ws, f'G10', '合并后收入', font=F_H2, fill=FILL_HDR2)
put(ws, f'H10', f'={SH_MERGE}!$G${prow["一、营业收入"]}', font=F_TOT, fill=FILL_CARD, fmt=MONEY)
put(ws, f'G11', '合并后净利', font=F_H2, fill=FILL_HDR2)
put(ws, f'H11', f'={SH_MERGE}!$G${prow["四、净利润"]}', font=F_TOT, fill=FILL_CARD, fmt=MONEY)
page(ws)

# ============================================================ 使用说明
ws = wb.create_sheet(SH_HELP)
title(ws, '使用说明', 'C', '照着这个顺序走一遍，这套表就能自己跑起来。')
widths(ws, {'A':6, 'B':26, 'C':120})
headers(ws, HR, 1, ['', '主题', '说　明'])
HELP = [
 ('一、这套表怎么转', ''),
 ('', '四张录入表', '【购销流水】进货卖货（只管货和票，一律挂应收应付）｜【资金流水】所有账户的收支'
                  '｜【生产加工】领料、加工费、产成品入库｜【其他分录】折旧摊销计提结转。'),
 ('', '自动往下走', '四张表 →【记账分录】按【记账规则】展开成借贷 →【科目余额表】→ 三家的资产负债表、'
                  '利润表和【三公司合并报表】。中间不用你手工做凭证。'),
 ('', '一年一套', '【参数设置】把会计年度改成新的一年，把四张录入表的数据清掉，'
                 '再把上年各表的期末数抄到【期初余额】，就是新一年的账。建议每年另存一个文件。'),
 ('二、先把档案建好', ''),
 ('', '参数设置', '会计年度、附加税率、所得税率，以及「费用项目」清单（资金流水和其他分录的下拉就取这一列）。'),
 ('', '商品档案', '三家的料和成品。存货/收入/成本科目决定这个商品记到哪个科目。'
                  '税率一列只是开票时选哪一档的提醒，内账含税核算，不参与计算。'),
 ('', '往来单位', '客户、供应商、农户，以及三家公司互相。标了「是否内部＝是」的，购销流水会认成内部交易，合并报表按它抵销。'),
 ('', '资金账户', '实际开了几个户就登几个。几家公司共用一个账号也照登，账户余额按账户算，账务按公司各记各的。'),
 ('', '期初余额', '左边科目年初余额（借贷各一列），中间存货年初数量金额，右边往来按单位拆。三块都有对账提示。'),
 ('', '会计科目表 / 记账规则', '一般不用动。要加科目就照格式加一行，「报表项目」必须用报表上已有的项目名。'
                            '要改某类业务的借贷走向，就去【记账规则】改那一行。'),
 ('三、日常录入', ''),
 ('', '购销流水', '选公司、业务类型、往来单位、商品编码，填数量和**含税单价**即可，金额＝数量×单价，不拆税。'
                '退货用「销售退回/采购退回」，数量填正数。'),
 ('', '向农户收购原木', '往来单位选那个农户，发票类型选「收购发票」，单价就填实际收购价。'
                      '内账含税核算，收购价全额进原材料成本，不再拆 9% 进项——'
                      '能不能抵扣是外账报税的事，内账只看实际付了多少钱。'),
 ('', '资金流水', '收支混在一张表上按时间顺序录：收进来填「收入金额」，付出去填「支出金额」，金额都是含税实收实付数，'
                '右边「账户余额」逐行滚出来。业务类型决定对方科目怎么走，'
                '规则里写「@对方」的才要你自己选对方科目。收付货款记得选往来单位，'
                '这样【往来台账】才冲得掉；「关联单号」写对应的购销流水序号，日后好核销。'),
 ('', '几家共用一个账号', '共用账户在【资金账户】里「使用公司」写清楚。资金流水每一行照样选自己的公司，'
                        '所以账户余额是这个账号的真实余额，而各公司账上的银行存款是各自的——'
                        '【资金账户】底下有一行「账实对账」把两边加总对上。'),
 ('', '生产加工', '一个批次号一次加工：先记「领用投入」和「加工费用」，再记「产成品入库」。'
                '产出行不用填金额，本批投入合计会按「权重」自动分摊；一批只有一个产出，权重留空就行。'),
 ('', '其他分录', '折旧、摊销、计提工资、计提房租，一借一贷。一借多贷拆成几行。'
                '**税不在这里计提**——内账按实缴，交税那天在【资金流水】记一笔就行。'
                '也不用做「结转本年利润」，报表会自动把损益结进未分配利润，你自己再结一次会重复。'),
 ('四、看报表', ''),
 ('', '按起止日期筛选', '所有汇总和查询表第 3 行都有「年度 / 起始日期 / 截止日期」。'
                      '都留空＝全部期间；只填年度＝整年；填了起止就按起止。录入表和明细表不筛选，全部显示。'),
 ('', '资产负债表', '年初数取期初余额表；期末数算到截止日期。未分配利润＝年初未分配＋本年净利润。'
                  '最后一行「平衡校验」必须是 √。'),
 ('', '利润表', '「本期金额」是起止区间的发生额，「本年累计」是年初到截止日的累计。下面附毛利率和净利率。'),
 ('', '合并报表', '三家单体＋小计＋内部抵销＝合并数。内部销售收入/成本、内部应收应付自动抵；'
                '「未实现内部存货利润」要你自己估一个填进去（上面有算法），不填按 0。'),
 ('', '进销存台账 / 收发存查询', '【进销存台账】是全年口径、成本基准表，加权平均单价就在这张表上算，'
                              '购销结转成本和生产领用都按它取价，所以它不随日期筛选变。'
                              '要看某段时间的收发存，用【收发存查询】。'),
 ('', '往来台账 / 税费台账 / 费用统计', '往来台账按公司看每家客户供应商的应收应付；'
                                   '税费台账按公司按月按税种列实缴税费和税负率；'
                                   '费用统计上半张按费用项目、下半张按会计科目。'),
 ('五、几个口径说明', ''),
 ('', '存货计价', '全年加权平均（含税口径）：单价＝（期初金额＋采购入库金额＋生产入库金额）÷ 对应数量。'
                '不含出库，所以不会循环引用。年中看的成本是按全年均价倒推的，和逐笔移动平均会有小差异。'),
 ('', '含税核算 / 实交税费', '这是内账，全部含税：卖多少钱记多少收入，进货花多少钱记多少成本，'
                          '费用付多少记多少，一律不拆进项销项，科目表里也没有进项/销项税额。'
                          '税走【资金流水】的「缴纳税费」——交增值税、附加税、印花税、个税进「税金及附加」，'
                          '交企业所得税进「所得税费用」，费用项目里选具体税种，'
                          '【税费台账】按月按税种摊开、给出全年税负率。'
                          '内部交易两边都是同一个含税金额，所以【内部交易核对】要求分毫不差。'),
 ('', '税率参数只是参考', '【参数设置】上的附加税率和所得税率不参与任何计算，'
                        '只是你自己估税时的提醒。实际交了多少，就在【资金流水】记多少。'
                        '农林牧渔里的林木种植所得免征企业所得税、自产农产品销售免征增值税，'
                        '所以农业基地公司示例里全年税费是 0。'),
 ('', '示例数据', '四张录入表里都放了示例（购销 20 条、资金 34 条、生产 36 条、其他分录 10 条），'
                '专门用来现场验公式。正式用之前，把这些行整行删掉即可，'
                '【期初余额】和【商品档案】里的示例也一并换成你自己的。'),
]
r = HR + 1
for a, *rest in HELP:
    if len(rest) == 1:
        ws.merge_cells(f'A{r}:C{r}')
        put(ws, f'A{r}', a, font=F_H2, fill=FILL_HDR2, align=CL)
        ws.row_dimensions[r].height = 22
    else:
        b, c = rest
        put(ws, f'A{r}', None)
        put(ws, f'B{r}', b, font=F_TOT, align=CL)
        put(ws, f'C{r}', c, font=F_TXT, align=CL)
        ws.row_dimensions[r].height = 34
    r += 1
page(ws, landscape=False)

# ============================================================ 排序 / 标签色 / 保存
ORDER = [SH_HOME, SH_HELP, SH_PARAM, SH_ACC, SH_RULE, SH_GD, SH_PT, SH_ACCT, SH_OPEN,
         SH_BUY, SH_CASH, SH_PROD, SH_OTH, SH_VOU, SH_TB]
for s in CO_NAMES: ORDER += [BS_SHEETS[s], PL_SHEETS[s]]
ORDER += [SH_MERGE, SH_INV, SH_INVQ, SH_AR, SH_TAX, SH_EXP, SH_IC, SH_CHK]
assert sorted(ORDER) == sorted(wb.sheetnames), set(ORDER) ^ set(wb.sheetnames)
wb._sheets = [wb[n] for n in ORDER]
TAB = {SH_HOME: '1F3864', SH_HELP: '7F7F7F',
       SH_BUY: 'ED7D31', SH_CASH: 'ED7D31', SH_PROD: 'ED7D31', SH_OTH: 'ED7D31',
       SH_PARAM: 'A6A6A6', SH_ACC: 'A6A6A6', SH_RULE: 'A6A6A6', SH_GD: 'A6A6A6',
       SH_PT: 'A6A6A6', SH_ACCT: 'A6A6A6', SH_OPEN: 'A6A6A6',
       SH_VOU: 'BFBFBF', SH_TB: '2F5597', SH_MERGE: '1F3864',
       SH_INV: '548235', SH_INVQ: '548235', SH_AR: '548235', SH_TAX: '548235',
       SH_EXP: '548235', SH_IC: '548235', SH_CHK: '70AD47'}
for s in CO_NAMES:
    TAB[BS_SHEETS[s]] = '2F5597'; TAB[PL_SHEETS[s]] = '2F5597'
for n, c in TAB.items(): wb[n].sheet_properties.tabColor = c
wb.active = 0

OUT = '../三公司财务账套_A060.xlsx'
import sys
if len(sys.argv) > 1: OUT = sys.argv[1]
wb.save(OUT)
n_f = sum(1 for s in wb.worksheets for row in s.iter_rows()
          for c in row if isinstance(c.value, str) and c.value.startswith('='))
print(f'已生成 {OUT}')
print(f'工作表 {len(wb.sheetnames)} 张，公式 {n_f:,} 条，括号配平修正 {G.BAL_FIXED[0]} 处')
