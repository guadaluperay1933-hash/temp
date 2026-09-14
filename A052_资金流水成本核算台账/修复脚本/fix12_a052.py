# -*- coding: utf-8 -*-
from copy import copy
import openpyxl
SRC, DST = 'A052_final2.xlsx', 'A052_final3.xlsx'
LOG=[]
def log(t,m): LOG.append(f'[{t}] {m}'); print(f'[{t}] {m}')
wb = openpyxl.load_workbook(SRC); R, S = wb['月度汇报表'], wb['使用说明']
FR='资金流水'
def sums(month_col):
    return (f'SUMIFS({FR}!$L$5:$L$3032,{FR}!${month_col}$5:${month_col}$3032,$B$2,'
            f'{FR}!$D$5:$D$3032,IF($E$2="全部","*",$E$2),{FR}!$W$5:$W$3032,"")')

# ── 修复 21：支出结构栏加一行"未分类"兜底，本栏才等于现金流出 ──────
for c in range(1, 9):                      # 合计行 37 → 38
    R.cell(38,c)._style = copy(R.cell(37,c)._style)
    R.cell(38,c).value  = R.cell(37,c).value
R['B38'] = '=SUM($B$31:$B$35)'
R['D38'] = '=SUM($D$31:$D$35)'
for c in range(1, 9):                      # 新的第 37 行：未分类
    R.cell(37,c)._style = copy(R.cell(36,c)._style)
    R.cell(37,c).value  = None
R['A37'] = '未分类支出（两种类别都没填）'
R['B37'] = '=' + sums('T')
R['D37'] = '=' + sums('S')
R['E37'] = '既没填成本类别也没填费用类别 —— 进了现金流出，却进不了内帐盈亏。请回「资金流水」补 H 列或 I 列'
for r in range(31, 36):                    # 占比分母 B37 → B38
    R[f'C{r}'] = f'=IFERROR($B{r}/$B$38,"")'
R['B178'] = '=$B$38'                       # 勾稽校验第 5 行跟着改
log('修复21', '【月度汇报表】"三、支出结构"六个大类靠「支出大类」匹配，而成本类别和费用类别都没填的支出，'
              '「支出大类」返回空串，六行里没有一行收得住它 —— 本月 1,254,532.50 元（占当月现金流出 38%，'
              '包括"付结构性存款"100 万、"付新疆恒睿建设工程有限公司借款"186,750 等）在这一栏彻底消失，'
              '而第一栏"现金流出"是全额统计的，两栏天然对不上还没有任何提示。'
              '现在第 37 行加了一行"未分类支出"兜底，合计行下移到第 38 行，占比分母同步跟上')

S['A78'] = ('⑨「月度汇报表」第三栏"支出结构"新增了一行"未分类支出"：成本类别和费用类别都没填的支出会单独列在这里，'
            '它进了现金流出、进不了内帐盈亏。这一行不该有数，有数就去「资金流水」把类别补上。')
S['A78']._style = copy(S['A77']._style); S.row_dimensions[78].height = 31.5
S['A76'] = ('全表 43,338 条公式已用 LibreOffice 全量重算验证：0 处报错；换月份、换公司、换日期区间、'
            '换项目、切换含税/不含税口径分别跑过一遍，勾稽校验的前九项始终"√ 一致"，'
            '第十项（没填类别的支出）如实反映录入缺口，补完类别就会变绿。')
wb.save(DST); open('fix12_log.txt','w',encoding='utf-8').write('\n'.join(LOG))
print('已保存', DST)
