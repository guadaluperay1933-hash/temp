# -*- coding: utf-8 -*-
"""④ 修《03 财务账套与报表》：
   (1) 【对接源_01水果】跨文件链接指错了文件 —— 公式里写的是 [1]（＝02 物料册），
       应当是 [2]（＝01 水果册）。这就是「筛选客户显示的是水果销售/筐子销售」的根子：
       01 的客户往来数据根本没接进来，接进来的是 02 的项目名。
   (2) 【_自动清单】的「往来单位」列改成先取《03》本册【原料入库计价】按货主汇总那一段的
       真实客户名，再接上【进销存对接】里认得出来的往来款行；款项名从只认「客户应收」
       放宽到成品销售款/公司购买果品款/筐子销售款/物料销售款。
   (3) 【应收应付汇总表】同步放宽款项名。
   跑法：python3 fix04_03表.py <入> <出>"""
import sys, re, openpyxl

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)

# ---------------- (1) 对接源_01水果：[1] → [2] ----------------
ws = wb['对接源_01水果']
n = 0
for row in ws.iter_rows():
    for c in row:
        if isinstance(c.value, str) and '[1]' in c.value:
            c.value = c.value.replace('[1]', '[2]'); n += 1
print(f'(1) 对接源_01水果：{n} 个单元格的跨文件链接从 [1]（02物料）改成 [2]（01水果）')

# ---------------- (2) _自动清单 往来单位清单 ----------------
# A 列是往来单位的池子，原来分四段：
#   A4:A305    进销存对接（跨册来的客户往来）—— 过滤词写的是「客户应收」，可 01/02 接口
#              的项目名其实是 成品销售款/公司购买果品款/筐子销售款/物料销售款，一行都命不中
#   A306:A905  资金日记账 —— 取的是 $F「业务类型」，应该取 $E「往来单位」。
#              这就是下拉里满屏「水果销售/筐子销售」的真正原因：那是业务类型，不是客户。
#   A906:A1205 成本费用登记表!$C（应付对象，本来就对）
#   A1206:A1805 记账凭证!$J（往来单位，本来就对）
# 这里：修前两段，再在后面接一段 A1806:A1925，直接取 01 送过来、落在本册【原料入库计价】
# 按货主汇总那一段的真实客户名 —— 就算跨文件链接断了，下拉里也始终有人。
KINDS = ['客户应收', '成品销售款', '公司购买果品款', '筐子销售款', '物料销售款']
ws = wb['_自动清单']
A_LAST = 1925
n1 = n2 = n3 = 0
for r in range(4, 306):
    cond = '+'.join(f'(进销存对接!$D{r}="{k}")' for k in KINDS)
    ws[f'A{r}'] = f'=IF(({cond})>0,进销存对接!$C{r},"")'
    n1 += 1
for r in range(306, 906):
    ws[f'A{r}'] = f'=IF(资金日记账!$E{r - 302}="","",资金日记账!$E{r - 302})'
    n2 += 1
for r in range(1806, A_LAST + 1):
    ws[f'A{r}'] = f'=IFERROR(INDEX(原料入库计价!$B$909:$B$1028,ROW()-1805),"")'
    ws[f'B{r}'] = f'=IF(A{r}="",0,IF(MATCH(A{r},$A$4:$A${A_LAST},0)=ROW()-3,1,0))'
    ws[f'C{r}'] = f'=N(C{r - 1})+B{r}'
    n3 += 1
# 三个辅助列的范围一起放到新的末行
for r in range(4, A_LAST + 1):
    b, e = ws[f'B{r}'], ws[f'E{r}']
    if isinstance(b.value, str) and '$A$4:$A$1805' in b.value:
        b.value = b.value.replace('$A$4:$A$1805', f'$A$4:$A${A_LAST}')
    if isinstance(e.value, str):
        e.value = e.value.replace('$A$4:$A$1805', f'$A$4:$A${A_LAST}').replace(
            '$C$4:$C$1805', f'$C$4:$C${A_LAST}')
print(f'(2) _自动清单 往来单位池：进销存对接段放宽款项名 {n1} 行；'
      f'资金日记账段从「业务类型 F」改回「往来单位 E」{n2} 行；'
      f'新接一段【原料入库计价】的真实客户名 {n3} 行')

# ---------------- (3) 应收应付汇总表 ----------------
ws = wb['应收应付汇总表']
n = 0
OLD = '进销存对接!$D$4:$D$305,"客户应收"'
NEW = '进销存对接!$I$4:$I$305,"客户往来"'
for row in ws.iter_rows():
    for c in row:
        if isinstance(c.value, str) and OLD in c.value:
            c.value = c.value.replace(OLD, NEW); n += 1
print(f'(3) 应收应付汇总表：{n} 个单元格改走【进销存对接】新的「往来标记」列')

# 进销存对接 补一列「往来标记」：只要是这几种款项名就打「客户往来」
ws = wb['进销存对接']
ws['I3'].value = '记账类别 / 往来标记'
for r in range(4, 306):
    old = ws[f'I{r}'].value
    ws[f'I{r}'].value = (f'=IF(OR({",".join(chr(36)+"D"+str(r)+"="+chr(34)+k+chr(34) for k in KINDS)}),'
                         f'"客户往来",'
                         f'IFERROR(INDEX(收入成本映射!$C$4:$C$34,MATCH($D{r},收入成本映射!$B$4:$B$34,0)),"不记账"))')
print('(4) 进销存对接 I 列：既是记账类别，也顺带把往来款行标成「客户往来」')

wb.save(OUT)
print('已写', OUT)
