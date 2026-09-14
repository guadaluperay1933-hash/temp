# -*- coding: utf-8 -*-
from copy import copy
import openpyxl
SRC, DST = 'A052_final10.xlsx', 'A052_final11.xlsx'
LOG=[]
def log(t,m): LOG.append(f'[{t}] {m}'); print(f'[{t}] {m}')
wb = openpyxl.load_workbook(SRC); R, S = wb['月度汇报表'], wb['使用说明']
FR='资金流水'
BLANK = f'({FR}!$B$5:$B$3032="")'
HASAMT = f'((({FR}!$K$5:$K$3032<>"")+({FR}!$L$5:$L$3032<>""))>0)'

r = 187
for c in range(1,9): R.cell(r,c)._style = copy(R.cell(186,c)._style)
R.row_dimensions[r].height = R.row_dimensions[186].height
R[f'A{r}'] = '有金额却没填日期的行'
R[f'B{r}'] = f'=SUMPRODUCT({BLANK}*{HASAMT})'
R[f'C{r}'] = 0
R[f'D{r}'] = f'=N($B{r})-N($C{r})'
R[f'E{r}'] = (f'=IF(ABS($D{r})<0.01,"√ 一致","✗ 有 "&TEXT($D{r},"0")&" 行，合计收 "'
              f'&TEXT(SUMPRODUCT({BLANK}*N({FR}!$K$5:$K$3032)),"#,##0.00")&" 支 "'
              f'&TEXT(SUMPRODUCT({BLANK}*N({FR}!$L$5:$L$3032)),"#,##0.00")&" 元没进任何报表")')
R[f'F{r}'] = '没有日期，发生月和归属月都算不出来，这些钱在所有按月/按区间的报表里都不存在'
for c in (7,8): R.cell(r,c).value=None
R.merge_cells(f'F{r}:H{r}')
log('修复40', '【月度汇报表】勾稽校验加第 14 行：有金额却没填日期的流水行。'
              '本表就有 11 行这种情况（收 212,741.53、支 119,830.90）—— 「赎回理财」150,589.92、'
              '「收昌吉市财政局」61,675、8 笔工资报销借款共 11.98 万 —— 它们进了表头 SUBTOTAL 的合计，'
              '却进不了任何按月、按区间的报表，所以「资金流水」表头合计与「月度趋势数据」全年合计对不上，'
              '差的正是这 21.27 万和 11.98 万。日期补上就闭合了')

S['A83'] = ('⑭ 勾稽校验最后加了两行体检：日期被录成文本的行、有金额却没填日期的行。'
            '本表有 11 行有金额没日期（收 212,741.53、支 119,830.90），它们进了资金流水表头的合计，'
            '却进不了任何按月报表 —— 这就是表头合计与趋势表全年合计对不上的原因，补上日期即可。')
S['A83']._style = copy(S['A82']._style); S.row_dimensions[83].height = 31.5
wb.save(DST); open('fix17_log.txt','w',encoding='utf-8').write('\n'.join(LOG))
print('已保存', DST)
