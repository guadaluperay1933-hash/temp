# -*- coding: utf-8 -*-
"""③ 给《01 水果进销存台账》加四张表 + 改一处口径：
   · 【库存总结余】  按 货主＋品种 汇总，不分等级，一眼看出「聂新平苹果共入 200、出 150、结余 50」
   · 【果然鲜采购明细】公司买下客户果品的逐笔 + 底部按货主汇总（含采购代发运费）
   · 【果然鲜销售汇总】按购买方汇总卖了多少、运费多少、收了多少、还欠多少
   · 【库存等级接口】供《03》跨文件取「某货主某等级还剩多少」
   · 【果然鲜销售明细】F 列改成「购买方/客户」：填了购买方就显示购买方，没填才退回客户/领用方；
     命中规则从「销售类型以果然鲜开头」放宽到「果然鲜开头 或 填了销售单价/销售金额」
   跑法：python3 fix03_新表.py <入> <出>"""
import sys, os, openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from style01 import *
from openpyxl.utils import get_column_letter as L

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
IN_R, IN_1 = 4, 3003            # 原料入库明细 / 原料出库明细 / 成品出库明细 数据行
AUTO_1 = 3003                   # _自动清单 各块行数

RK = lambda c: f'原料入库明细!${c}${IN_R}:${c}${IN_1}'
CK = lambda c: f'原料出库明细!${c}${IN_R}:${c}${IN_1}'
FK = lambda c: f'成品出库明细!${c}${IN_R}:${c}${IN_1}'

# ============================================================ _自动清单 新增去重块
ws = wb['_自动清单']
NEW = [
    # (列, 表头, 行公式模板)
    ('CD', '购买方', '=IF(成品出库明细!$B{r}="","",IF(成品出库明细!$H{r}<>"",成品出库明细!$H{r},成品出库明细!$G{r}))'),
    ('CE', '首次', '=IF(CD{r}="",0,IF(MATCH(CD{r},$CD$4:$CD$3003,0)=ROW()-3,1,0))'),
    ('CF', '累计', '=N(CF{p})+CE{r}'),
    ('CH', '购买方(紧凑)', '=IFERROR(INDEX($CD$4:$CD$3003,MATCH(ROW()-3,$CF$4:$CF$3003,0)),"")'),
    ('CJ', '采购命中', '=IF(AND(成品出库明细!$B{r}<>"",OR(N(成品出库明细!$T{r})<>0,N(成品出库明细!$V{r})<>0)),1,0)'),
    ('CK', '累计', '=N(CK{p})+CJ{r}'),
    ('CM', '销售命中', '=IF(AND(成品出库明细!$B{r}<>"",OR(N(成品出库明细!$Y{r})<>0,N(成品出库明细!$AB{r})<>0)),1,0)'),
    ('CN', '累计', '=N(CN{p})+CM{r}'),
]
for col, head, fml in NEW:
    put(ws, f'{col}3', head, font=F_TOT, fill=FILL_HDR2)
    for r in range(4, AUTO_1 + 1):
        put(ws, f'{col}{r}', fml.format(r=r, p=r - 1), font=F_NOTE, border=None)
put(ws, 'CD1', '=MAX($CF$4:$CF$3003)', font=F_NOTE, border=None)
put(ws, 'CJ1', '=MAX($CK$4:$CK$3003)', font=F_NOTE, border=None)
put(ws, 'CM1', '=MAX($CN$4:$CN$3003)', font=F_NOTE, border=None)
print('  _自动清单：补 购买方去重 / 采购命中 / 销售命中 三块')

# ============================================================ ① 库存总结余
if '库存总结余' in wb.sheetnames: del wb['库存总结余']
ws = wb.create_sheet('库存总结余', wb.sheetnames.index('库存结余') + 1)
N_TOT = 300
R0, R1 = 4, 3 + N_TOT
title(ws, '库 存 总 结 余 表（货主 ＋ 品种 · 不分等级 · 全自动）',
      'N', sub_last='I', sub='★ 就是你要的那个直观数：一个货主一个品种一行 —— 入库多少、出库多少、还剩多少，等级全部合并。'
      '要看细到等级的，翻前面那张【库存结余】。填右边【开始日期／结束日期】可按区间统计（留空＝不限时间）。')
widths(ws, {'A': 6, 'B': 14, 'C': 11, 'D': 7, 'E': 12, 'F': 13, 'G': 12, 'H': 13, 'I': 12,
            'J': 13, 'K': 13, 'L': 14, 'M': 11, 'N': 16, 'O': 3, 'P': 12, 'Q': 12})
put(ws, 'J2', '开始日期', font=F_TOT, fill=FILL_HDR2)
put(ws, 'K2', None, font=F_IN, fill=FILL_IN, fmt=DATE)
put(ws, 'L2', '结束日期', font=F_TOT, fill=FILL_HDR2)
put(ws, 'M2', None, font=F_IN, fill=FILL_IN, fmt=DATE)
put(ws, 'P1', '=IF($K$2="",DATE(1900,1,1),$K$2)', font=F_NOTE, fmt=DATE, border=None)
put(ws, 'P2', '=IF($M$2="",DATE(2999,12,31),$M$2)', font=F_NOTE, fmt=DATE, border=None)
ws.column_dimensions['P'].hidden = True
headers(ws, 3, ['序号', '货主/存放人', '品种', '单位', '入库数量', '入库总重KG',
                '出库数量', '出库总重KG', '结余数量', '结余总重KG',
                '成品出库数量', '成品出库总重KG', '库存状态', '备注'])
DR = ',原料入库明细!$B$4:$B$3003,">="&$P$1,原料入库明细!$B$4:$B$3003,"<="&$P$2'
DRC = ',原料出库明细!$B$4:$B$3003,">="&$P$1,原料出库明细!$B$4:$B$3003,"<="&$P$2'
DRF = ',成品出库明细!$B$4:$B$3003,">="&$P$1,成品出库明细!$B$4:$B$3003,"<="&$P$2'
for r in range(R0, R1 + 1):
    put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-3)', font=F_LINK)
    put(ws, f'B{r}', f'=IFERROR(INDEX(_自动清单!$Z$4:$Z$3003,ROW()-3),"")', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IFERROR(INDEX(_自动清单!$AA$4:$AA$3003,ROW()-3),"")', font=F_LINK)
    put(ws, f'D{r}', f'=IF($B{r}="","",IFERROR(INDEX({RK("I")},MATCH($B{r}&"|"&$C{r},'
                     f'_自动清单!$S$4:$S$3003,0)),""))', font=F_LINK)
    put(ws, f'E{r}', f'=IF($B{r}="","",SUMIFS({RK("H")},{RK("F")},$B{r},{RK("D")},$C{r}{DR}))',
        font=F_LINK, fmt=NUM)
    put(ws, f'F{r}', f'=IF($B{r}="","",ROUND(SUMIFS({RK("K")},{RK("F")},$B{r},{RK("D")},$C{r}{DR}),2))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'G{r}', f'=IF($B{r}="","",SUMIFS({CK("H")},{CK("F")},$B{r},{CK("D")},$C{r}{DRC}))',
        font=F_LINK, fmt=NUM)
    put(ws, f'H{r}', f'=IF($B{r}="","",ROUND(SUMIFS({CK("K")},{CK("F")},$B{r},{CK("D")},$C{r}{DRC}),2))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($B{r}="","",N($E{r})-N($G{r}))', font=F_TOT, fmt=NUM)
    put(ws, f'J{r}', f'=IF($B{r}="","",ROUND(N($F{r})-N($H{r}),2))', font=F_TOT, fmt=MONEY)
    put(ws, f'K{r}', f'=IF($B{r}="","",SUMIFS({FK("J")},{FK("F")},$B{r},{FK("D")},$C{r}{DRF}))',
        font=F_LINK, fmt=NUM)
    put(ws, f'L{r}', f'=IF($B{r}="","",ROUND(SUMIFS({FK("M")},{FK("F")},$B{r},{FK("D")},$C{r}{DRF}),2))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'M{r}', f'=IF($B{r}="","",IF($I{r}<=0,"无库存",IF($I{r}>=基础资料!$U$4,"充足",'
                     f'IF($I{r}>=基础资料!$U$5,"正常","偏低"))))', font=F_TXT, fill=FILL_AUTO)
    put(ws, f'N{r}', f'=IF($B{r}="","",$B{r}&" 的"&$C{r}&"：入库 "&TEXT(N($E{r}),"#,##0")'
                     f'&"，出库 "&TEXT(N($G{r}),"#,##0")&"，结余 "&TEXT(N($I{r}),"#,##0"))',
        font=F_NOTE, align=CL)
    ws.row_dimensions[r].height = 16
TOT = R1 + 1
put(ws, f'A{TOT}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BCD': put(ws, f'{c}{TOT}', None, font=F_TOT, fill=FILL_TOT)
for c in 'EFGHIJKL':
    put(ws, f'{c}{TOT}', f'=ROUND(SUM({c}{R0}:{c}{R1}),2)', font=F_TOT, fill=FILL_TOT,
        fmt=NUM if c in 'EGIK' else MONEY)
for c in 'MN': put(ws, f'{c}{TOT}', None, font=F_TOT, fill=FILL_TOT)
ws.freeze_panes = 'C4'
ws.auto_filter.ref = f'A3:N{R1}'
page(ws, titles='3:3')
print('  ✓ 新表【库存总结余】（货主＋品种，不分等级）')

# ============================================================ ② 果然鲜采购明细
if '果然鲜采购明细' in wb.sheetnames: del wb['果然鲜采购明细']
ws = wb.create_sheet('果然鲜采购明细', wb.sheetnames.index('果然鲜销售明细'))
P_N = 400
PR0, PR1 = 4, 3 + P_N
title(ws, '果 然 鲜 采 购 明 细（公司买下客户果品 · 自动）', 'N', sub_last='H',
      sub='★ 全自动：把【成品出库明细】里填了「采购单价 / 采购金额（应付）」的行挑出来 —— '
      '也就是公司（果然鲜）从货主手里买断的那批货，金额是我们要付给货主的钱。'
      '采购代发运费单列一列，已经含在「采购金额（应付）」里。填右边起止日期可按区间查。')
widths(ws, {'A': 6, 'B': 12, 'C': 14, 'D': 13, 'E': 11, 'F': 12, 'G': 9, 'H': 7, 'I': 11,
            'J': 13, 'K': 15, 'L': 11, 'M': 16, 'N': 12, 'O': 3, 'P': 6, 'Q': 12})
put(ws, 'I2', '开始日期', font=F_TOT, fill=FILL_HDR2)
put(ws, 'J2', None, font=F_IN, fill=FILL_IN, fmt=DATE)
put(ws, 'K2', '结束日期', font=F_TOT, fill=FILL_HDR2)
put(ws, 'L2', None, font=F_IN, fill=FILL_IN, fmt=DATE)
put(ws, 'Q1', '=IF($J$2="",DATE(1900,1,1),$J$2)', font=F_NOTE, fmt=DATE, border=None)
put(ws, 'Q2', '=IF($L$2="",DATE(2999,12,31),$L$2)', font=F_NOTE, fmt=DATE, border=None)
ws.column_dimensions['P'].hidden = True
ws.column_dimensions['Q'].hidden = True
headers(ws, 3, ['序号', '日期', '票据单号', '货主（卖给公司的）', '品种', '等级', '数量', '单位',
                '采购单价', '采购代发运费', '采购金额（应付）', '付款状态', '备注', '购买方（再卖给谁）'])
def fx(col, ptr):
    return f'IF(INDEX({FK(col)},{ptr})=0,"",INDEX({FK(col)},{ptr}))'
for r in range(PR0, PR1 + 1):
    p = f'$P{r}'
    put(ws, f'P{r}', f'=IFERROR(MATCH(ROW()-3,_自动清单!$CK$4:$CK$3003,0),"")', font=F_NOTE, border=None)
    put(ws, f'A{r}', f'=IF({p}="","",ROW()-3)', font=F_LINK)
    put(ws, f'B{r}', f'=IF({p}="","",{fx("B", p)})', font=F_LINK, fmt=DATE)
    put(ws, f'C{r}', f'=IF({p}="","",{fx("C", p)})', font=F_LINK)
    put(ws, f'D{r}', f'=IF({p}="","",{fx("F", p)})', font=F_LINK, align=CL)
    put(ws, f'E{r}', f'=IF({p}="","",{fx("D", p)})', font=F_LINK)
    put(ws, f'F{r}', f'=IF({p}="","",{fx("E", p)})', font=F_LINK)
    put(ws, f'G{r}', f'=IF({p}="","",{fx("J", p)})', font=F_LINK, fmt=NUM)
    put(ws, f'H{r}', f'=IF({p}="","",{fx("K", p)})', font=F_LINK)
    put(ws, f'I{r}', f'=IF({p}="","",{fx("T", p)})', font=F_LINK, fmt=MONEY)
    put(ws, f'J{r}', f'=IF({p}="","",{fx("U", p)})', font=F_LINK, fmt=MONEY)
    put(ws, f'K{r}', f'=IF({p}="","",{fx("V", p)})', font=F_TOT, fmt=MONEY)
    put(ws, f'L{r}', f'=IF({p}="","",{fx("X", p)})', font=F_LINK)
    put(ws, f'M{r}', f'=IF({p}="","",{fx("AF", p)})', font=F_NOTE, align=CL)
    put(ws, f'N{r}', f'=IF({p}="","",IF(INDEX({FK("H")},{p})<>"",INDEX({FK("H")},{p}),{fx("G", p)}))',
        font=F_LINK, align=CL)
    ws.row_dimensions[r].height = 16
PT = PR1 + 1
put(ws, f'A{PT}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BCDEFH LMN'.replace(' ', ''): put(ws, f'{c}{PT}', None, font=F_TOT, fill=FILL_TOT)
for c in ['G', 'I', 'J', 'K']:
    put(ws, f'{c}{PT}', f'=ROUND(SUM({c}{PR0}:{c}{PR1}),2)', font=F_TOT, fill=FILL_TOT,
        fmt=NUM if c == 'G' else MONEY)
# 底部：按货主汇总
S0 = PT + 2
ws.merge_cells(f'A{S0}:N{S0}')
put(ws, f'A{S0}', '按 货 主 汇 总（公司一共买了谁多少果、付了多少、还欠多少）',
    font=F_TOT, fill=FILL_HDR2, align=CL)
headers(ws, S0 + 1, ['序号', '货主/客户', '笔数', '数量', '采购代发运费', '采购金额（应付）',
                     '已付金额', '未付金额', '', '', '', '', '', ''], fill=FILL_HDR2, font=F_TOT)
SS = S0 + 2
for i in range(120):
    r = SS + i
    put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-{SS - 1})', font=F_LINK)
    put(ws, f'B{r}', f'=IFERROR(INDEX(_自动清单!$BO$4:$BO$1203,ROW()-{SS - 1}),"")',
        font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($B{r}="","",COUNTIF($D${PR0}:$D${PR1},$B{r}))', font=F_LINK, fmt=NUM)
    put(ws, f'D{r}', f'=IF($B{r}="","",SUMIF($D${PR0}:$D${PR1},$B{r},$G${PR0}:$G${PR1}))',
        font=F_LINK, fmt=NUM)
    put(ws, f'E{r}', f'=IF($B{r}="","",ROUND(SUMIF($D${PR0}:$D${PR1},$B{r},$J${PR0}:$J${PR1}),2))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($B{r}="","",ROUND(SUMIF($D${PR0}:$D${PR1},$B{r},$K${PR0}:$K${PR1}),2))',
        font=F_TOT, fmt=MONEY)
    put(ws, f'G{r}', f'=IF($B{r}="","",ROUND(SUMIFS($K${PR0}:$K${PR1},$D${PR0}:$D${PR1},$B{r},'
                     f'$L${PR0}:$L${PR1},"已付清"),2))', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($B{r}="","",ROUND(N($F{r})-N($G{r}),2))', font=F_TOT, fmt=MONEY)
    for c in 'IJKLMN': put(ws, f'{c}{r}', None, border=None)
    ws.row_dimensions[r].height = 16
ST = SS + 120
put(ws, f'A{ST}', '合  计', font=F_TOT, fill=FILL_TOT)
put(ws, f'B{ST}', None, font=F_TOT, fill=FILL_TOT)
for c in 'CDEFGH':
    put(ws, f'{c}{ST}', f'=ROUND(SUM({c}{SS}:{c}{ST - 1}),2)', font=F_TOT, fill=FILL_TOT,
        fmt=NUM if c in 'CD' else MONEY)
ws.freeze_panes = 'C4'
page(ws, titles='3:3')
print('  ✓ 新表【果然鲜采购明细】（逐笔 + 按货主汇总）')

# ============================================================ ③ 果然鲜销售汇总
if '果然鲜销售汇总' in wb.sheetnames: del wb['果然鲜销售汇总']
ws = wb.create_sheet('果然鲜销售汇总', wb.sheetnames.index('果然鲜销售明细') + 1)
title(ws, '果 然 鲜 销 售 汇 总（按购买方 · 自动）', 'K',
      '★ 全自动：把【果然鲜销售明细】按「购买方」汇总 —— 卖给谁、卖了多少、运费多少、'
      '收了多少、还欠多少。购买方取【成品出库明细】新增的「购买方」列，没填就退回「客户/领用方」。'
      '想按日期区间看，去【果然鲜销售明细】右上角填起止日期，这里跟着变。')
widths(ws, {'A': 6, 'B': 16, 'C': 8, 'D': 10, 'E': 12, 'F': 15, 'G': 11, 'H': 11, 'I': 15,
            'J': 13, 'K': 13})
headers(ws, 3, ['序号', '购买方', '笔数', '数量', '总毛重KG', '销售金额（应收）',
                '自结运费', '代发运费', '应收合计', '已收金额', '未收金额'])
D0, D1 = 4, 403        # 果然鲜销售明细 的数据行
SD = '果然鲜销售明细'
for i in range(120):
    r = 4 + i
    put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-3)', font=F_LINK)
    put(ws, f'B{r}', f'=IFERROR(INDEX(_自动清单!$CH$4:$CH$3003,ROW()-3),"")', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($B{r}="","",COUNTIF({SD}!$F${D0}:$F${D1},$B{r}))', font=F_LINK, fmt=NUM)
    put(ws, f'D{r}', f'=IF($B{r}="","",SUMIF({SD}!$F${D0}:$F${D1},$B{r},{SD}!$I${D0}:$I${D1}))',
        font=F_LINK, fmt=NUM)
    put(ws, f'E{r}', f'=IF($B{r}="","",ROUND(SUMIF({SD}!$F${D0}:$F${D1},$B{r},{SD}!$K${D0}:$K${D1}),2))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($B{r}="","",ROUND(SUMIF({SD}!$F${D0}:$F${D1},$B{r},{SD}!$O${D0}:$O${D1}),2))',
        font=F_TOT, fmt=MONEY)
    put(ws, f'G{r}', f'=IF($B{r}="","",ROUND(SUMIF({SD}!$F${D0}:$F${D1},$B{r},{SD}!$P${D0}:$P${D1}),2))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($B{r}="","",ROUND(SUMIF({SD}!$F${D0}:$F${D1},$B{r},{SD}!$Q${D0}:$Q${D1}),2))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($B{r}="","",ROUND(N($F{r}),2))', font=F_TOT, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($B{r}="","",ROUND(SUMIFS({SD}!$O${D0}:$O${D1},{SD}!$F${D0}:$F${D1},$B{r},'
                     f'{SD}!$S${D0}:$S${D1},"已收清"),2))', font=F_LINK, fmt=MONEY)
    put(ws, f'K{r}', f'=IF($B{r}="","",ROUND(N($I{r})-N($J{r}),2))', font=F_TOT, fmt=MONEY)
    ws.row_dimensions[r].height = 16
T2 = 124
put(ws, f'A{T2}', '合  计', font=F_TOT, fill=FILL_TOT)
put(ws, f'B{T2}', None, font=F_TOT, fill=FILL_TOT)
for c in 'CDEFGHIJK':
    put(ws, f'{c}{T2}', f'=ROUND(SUM({c}4:{c}123),2)', font=F_TOT, fill=FILL_TOT,
        fmt=NUM if c in 'CD' else MONEY)
put(ws, f'A{T2 + 2}', '注：「应收合计」＝销售金额（应收），运费已经含在销售金额里（自结运费、代发运费另列一栏只是让你看清构成）。'
                      '「已收金额」按【果然鲜销售明细】收付状态＝已收清 的行统计。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{T2 + 2}:K{T2 + 2}')
ws.freeze_panes = 'C4'
page(ws, titles='3:3', landscape=False)
print('  ✓ 新表【果然鲜销售汇总】（按购买方）')

# ============================================================ ④ 库存等级接口（供《03》取数）
if '库存等级接口' in wb.sheetnames: del wb['库存等级接口']
ws = wb.create_sheet('库存等级接口')
title(ws, '库 存 等 级 接 口（供《03 财务账套》跨文件取数 · 请勿改动结构）', 'H',
      '★ 全自动，1:1 对应【库存结余】的行。《03》的「库存价值与欠款比对」按固定位置引用本表 A3:G403，'
      '请不要插入/删除行列，也不要改表名。')
widths(ws, {'A': 6, 'B': 14, 'C': 11, 'D': 13, 'E': 7, 'F': 12, 'G': 13, 'H': 14})
headers(ws, 3, ['源行', '货主/客户', '品种', '等级', '单位', '结余数量', '结余总重KG', '备注'])
for i in range(400):
    r, sr = 4 + i, 4 + i
    put(ws, f'A{r}', str(i + 1), font=F_NOTE)
    ws[f'A{r}'] = i + 1
    put(ws, f'B{r}', f'=IF(库存结余!$B{sr}="","",库存结余!$B{sr})', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF(库存结余!$B{sr}="","",库存结余!$C{sr})', font=F_LINK)
    put(ws, f'D{r}', f'=IF(库存结余!$B{sr}="","",库存结余!$D{sr})', font=F_LINK)
    put(ws, f'E{r}', f'=IF(库存结余!$B{sr}="","",库存结余!$F{sr})', font=F_LINK)
    put(ws, f'F{r}', f'=IF(库存结余!$B{sr}="","",N(库存结余!$K{sr}))', font=F_TOT, fmt=NUM)
    put(ws, f'G{r}', f'=IF(库存结余!$B{sr}="","",N(库存结余!$L{sr}))', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF(库存结余!$B{sr}="","",库存结余!$O{sr})', font=F_NOTE)
    ws.row_dimensions[r].height = 15
put(ws, 'J1', '有效行数', font=F_TOT, fill=FILL_HDR2)
put(ws, 'K1', '=COUNTIF($B$4:$B$403,"?*")', font=F_TOT, fill=FILL_AUTO, fmt=NUM)
ws.freeze_panes = 'C4'
page(ws, titles='3:3', landscape=False)
print('  ✓ 新表【库存等级接口】（供《03》取数）')

# ============================================================ ⑤ 果然鲜销售明细：口径与购买方
ws = wb['果然鲜销售明细']
ws['F3'].value = '购买方/客户'
for r in range(4, 404):
    ws[f'F{r}'] = (f'=IF($W{r}="","",IF(INDEX({FK("H")},$W{r})<>"",INDEX({FK("H")},$W{r}),'
                   f'IF(INDEX({FK("G")},$W{r})=0,"",INDEX({FK("G")},$W{r}))))')
a2 = ws['A2']
if isinstance(a2.value, str) and '购买方' not in a2.value:
    a2.value = a2.value + '　★ 「购买方/客户」列：填了【成品出库明细】的「购买方」就显示购买方，没填才退回客户/领用方。'
# 命中规则放宽：果然鲜开头 或 填了销售单价/销售金额
ws2 = wb['_自动清单']
n = 0
for r in range(4, AUTO_1 + 1):
    c = ws2[f'CA{r}']
    if isinstance(c.value, str) and c.value.startswith('='):
        c.value = (f'=IF(AND(成品出库明细!$B{r}<>"",'
                   f'OR(LEFT(成品出库明细!$I{r},3)="果然鲜",N(成品出库明细!$Y{r})<>0,N(成品出库明细!$AB{r})<>0),'
                   f'N(成品出库明细!$B{r})>=果然鲜销售明细!$X$1,N(成品出库明细!$B{r})<=果然鲜销售明细!$X$2),1,0)')
        n += 1
print(f'  ✓ 果然鲜销售明细：F 列改「购买方/客户」，命中规则放宽（{n} 行）')

wb.save(OUT)
print('已写', OUT)
