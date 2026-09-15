# -*- coding: utf-8 -*-
"""⑦ 在三本册子的【主页】底部加一段「本轮新增」，把新表列出来并做成可点击的跳转。
   用 HYPERLINK 公式而不是超链接对象：超链接对象经过一些工具处理会被转成指向外部文件的死链，
   公式形式在 Excel / WPS 里都稳。
   跑法：python3 fix07_主页.py <目录>"""
import sys, os, copy, openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from style01 import *

D = sys.argv[1]
ADD = {
 '01_水果进销存台账模板.xlsx': [
   ('库存总结余', '★按货主＋品种汇总，不分等级 —— 入库多少、出库多少、结余多少，一个直观数', '自动'),
   ('果然鲜采购明细', '★公司买下客户果品的逐笔明细 ＋ 按货主汇总（含采购代发运费、已付未付）', '自动'),
   ('果然鲜销售汇总', '★按购买方汇总：卖了多少、运费多少、收了多少、还欠多少', '自动'),
   ('库存等级接口', '供《03 财务账套》跨文件取「某货主某等级还剩多少」', '自动'),
 ],
 '03_财务账套与报表模板.xlsx': [
   ('对接源_01库存', '跨文件取《01》库存等级接口', '自动'),
   ('库存价值与欠款比对', '★下半部分按等级填单价算库存价值，上半部分按客户比对「库存价值 ↔ 应收款 ↔ 差额」', '录入+自动'),
 ],
 '04_综合查询对账单模板.xlsx': [
   ('果然鲜销售对账单', '★按购买方出：卖了什么、多少钱、运费多少、收了多少、还欠多少', '录入+自动'),
 ],
}
for fn, rows in ADD.items():
    p = os.path.join(D, fn)
    wb = openpyxl.load_workbook(p)
    ws = wb['主页']
    r0 = ws.max_row + 2
    ws.merge_cells(f'A{r0}:D{r0}')
    put(ws, f'A{r0}', '本 轮 新 增', font=F_TOT, fill=FILL_HDR2, align=CL)
    ws.row_dimensions[r0].height = 20
    hdr = r0 + 1
    for i, t in enumerate(['序号', '子表名称', '用途说明', '属性'], 1):
        put(ws, f'{ws.cell(row=hdr, column=i).column_letter}{hdr}', t, font=F_HDR, fill=FILL_HDR)
    ws.row_dimensions[hdr].height = 20
    for i, (nm, desc, kind) in enumerate(rows, 1):
        r = hdr + i
        put(ws, f'A{r}', i, font=F_TXT)
        put(ws, f'B{r}', f'=HYPERLINK("#\'{nm}\'!A1","{nm}")',
            font=Font(name='微软雅黑', size=10, bold=True, color='0563C1', underline='single'),
            align=C)
        put(ws, f'C{r}', desc, font=F_TXT, align=CL)
        put(ws, f'D{r}', kind, font=F_TXT)
        ws.row_dimensions[r].height = 18
    wb.save(p)
    print(f'  ✓ {fn} 主页补了 {len(rows)} 条新表入口')
