# -*- coding: utf-8 -*-
"""② 修《01》的两张接口表：
   (1) 【对账明细接口】C「往来单位」原来对成品出库行取的是「客户/领用方」，
       结果给货主出对账单时，成品出库那几十行一行都对不上 —— 改成取「存放人/货主」。
       这就是「04 表对账单成品出库数量有误」的根子。
   (2) 【对账明细接口】补 5 列：购买方 / 销售单价 / 自结运费 / 代发运费 / 采购代发运费，
       供《04》出「果然鲜销售对账单」。
   跑法：python3 fix02_接口.py <入> <出>"""
import sys, copy, openpyxl
from openpyxl.utils import get_column_letter as L

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)

# ---------------- (1)(2) 对账明细接口 ----------------
ws = wb['对账明细接口']
R0, R1 = 4, 1203
fixed = 0
for r in range(R0, R1 + 1):
    c = ws.cell(row=r, column=3)          # C 往来单位
    if isinstance(c.value, str) and '成品出库明细!$G$4:$G$3003' in c.value:
        c.value = c.value.replace('成品出库明细!$G$4:$G$3003', '成品出库明细!$F$4:$F$3003')
        fixed += 1
print(f'(1) 对账明细接口 C 列「往来单位」成品出库分支 G→F：改了 {fixed} 行')

NEWCOLS = [
    ('X', 12, '购买方',
     '=IF(OR($Z{r}="",$Z{r}<=6000,$Z{r}>9000),"",'
     'IF(INDEX(成品出库明细!$H$4:$H$3003,$Z{r}-6000)<>"",INDEX(成品出库明细!$H$4:$H$3003,$Z{r}-6000),'
     'IF(INDEX(成品出库明细!$G$4:$G$3003,$Z{r}-6000)=0,"",INDEX(成品出库明细!$G$4:$G$3003,$Z{r}-6000))))'),
    ('Y', 11, '销售单价',
     '=IF(OR($Z{r}="",$Z{r}<=6000,$Z{r}>9000),"",'
     'IF(N(INDEX(成品出库明细!$Y$4:$Y$3003,$Z{r}-6000))=0,"",N(INDEX(成品出库明细!$Y$4:$Y$3003,$Z{r}-6000))))'),
    ('AA', 11, '自结运费',
     '=IF(OR($Z{r}="",$Z{r}<=6000,$Z{r}>9000),"",'
     'IF(N(INDEX(成品出库明细!$Z$4:$Z$3003,$Z{r}-6000))=0,"",N(INDEX(成品出库明细!$Z$4:$Z$3003,$Z{r}-6000))))'),
    ('AB', 11, '代发运费',
     '=IF(OR($Z{r}="",$Z{r}<=6000,$Z{r}>9000),"",'
     'IF(N(INDEX(成品出库明细!$AA$4:$AA$3003,$Z{r}-6000))=0,"",N(INDEX(成品出库明细!$AA$4:$AA$3003,$Z{r}-6000))))'),
    ('AC', 13, '采购代发运费',
     '=IF(OR($Z{r}="",$Z{r}<=6000,$Z{r}>9000),"",'
     'IF(N(INDEX(成品出库明细!$U$4:$U$3003,$Z{r}-6000))=0,"",N(INDEX(成品出库明细!$U$4:$U$3003,$Z{r}-6000))))'),
    ('AD', 11, '销售类型',
     '=IF(OR($Z{r}="",$Z{r}<=6000,$Z{r}>9000),"",'
     'IF(INDEX(成品出库明细!$I$4:$I$3003,$Z{r}-6000)=0,"",INDEX(成品出库明细!$I$4:$I$3003,$Z{r}-6000)))'),
]
hdr_src = ws.cell(row=3, column=23)       # 照抄 W 列表头的样子
for col, w, name, fml in NEWCOLS:
    ws.column_dimensions[col].width = w
    h = ws[f'{col}3']; h.value = name
    if hdr_src.has_style: h._style = copy.copy(hdr_src._style)
    body_src = ws.cell(row=4, column=23)
    for r in range(R0, R1 + 1):
        c = ws[f'{col}{r}']
        c.value = fml.format(r=r)
        if body_src.has_style: c._style = copy.copy(ws.cell(row=r, column=23)._style)
print(f'(2) 对账明细接口 补列：{"、".join(n for _,_,n,_ in NEWCOLS)}')

# 注：原来打算把「按客户往来」那段的「项目」列改成 "客户应收"，
# 但那一列是 D/E 两列的判断条件（=IF($C<>"成品销售款",0,...)），改了整段就废了。
# 所以改在《03》那边认款项名，本文件不动。

# ---------------- (3) 成品出库明细：两个被写死的「手工录入」列 ----------------
# 「采购代发运费」和「销售单价」表头是手工录入色，格子里却留着旧版的公式：
#   采购代发运费 = 数量×采购单价  → 采购金额（应付）＝采购货款＋采购代发运费 就变成了 2 倍货款
#   销售单价     = 采购那四列相加 → 根本没法填单价，销售金额也跟着错
# 现在 42 行数据的采购单价全是空的，这两条链一路返回空串，所以看不出来；
# 只要有人第一次填采购单价，01/03/04 三本的应收应付会同时算错。这里把这两列清成真正的手工格。
ws = wb['成品出库明细']
FR, LR = 4, 3003
cleared = 0
for col, name in (('U', '采购代发运费'), ('Y', '销售单价')):
    for r in range(FR, LR + 1):
        c = ws[f'{col}{r}']
        if isinstance(c.value, str) and c.value.startswith('='):
            c.value = None; cleared += 1
print(f'(3) 成品出库明细：清掉「采购代发运费」「销售单价」两列里写死的旧公式 {cleared} 格，还原成手工录入')

# ---------------- (4) 成品出库明细：加一列自动的「实际购买方」 ----------------
# 填了「购买方」就是购买方，没填就沿用「客户/领用方」—— 给下游 SUMIFS 当条件用，
# 免得每处都写一遍 IF 兜底。放在最右边 AK，不动 AJ（那一列是旧版遗留，收入结算汇总还在引用）。
put_head = ws['G3']
ws['AK3'] = '实际购买方(自动)'
if put_head.has_style: ws['AK3']._style = copy.copy(put_head._style)
ws.column_dimensions['AK'].width = 14
for r in range(FR, LR + 1):
    ws[f'AK{r}'] = f'=IF($B{r}="","",IF($H{r}<>"",$H{r},$G{r}))'
print('(4) 成品出库明细：新增 AK「实际购买方(自动)」')

# ---------------- (5) 财务取数接口：应收对象改用「实际购买方」 ----------------
ws = wb['财务取数接口']
n = 0
for r in range(4, ws.max_row + 1):
    for col in (4, 8):
        c = ws.cell(row=r, column=col)
        if isinstance(c.value, str) and '成品出库明细!$G$4:$G$3003' in c.value:
            c.value = c.value.replace('成品出库明细!$G$4:$G$3003', '成品出库明细!$AK$4:$AK$3003')
            n += 1
print(f'(5) 财务取数接口：{n} 处「按客户应收」的匹配列从「客户/领用方」改成「实际购买方」')

wb.save(OUT)
print('已写', OUT)
