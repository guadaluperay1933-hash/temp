# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
SRC, DST = 'A052_fixed6.xlsx', 'A052_fixed7.xlsx'
LOG=[]
def log(t,m): LOG.append(f'[{t}] {m}'); print(f'[{t}] {m}')
wb = openpyxl.load_workbook(SRC)
F, Q = wb['资金流水'], wb['任意时段查询']

# ── 修复 11：V610 断链 ────────────────────────────────────────────
assert 'N(V608)' in F['V610'].value
F['V610'] = '=IF(AND($B610<>"",$F610<>"",$F610=项目明细账!$C$2),N(V609)+1,N(V609))'
log('修复11', '【资金流水】V610「项目行号」引用的是 V608（隔了一行），全表 3028 格里唯一的断链。'
             'V 列是「本项目第几笔」的滚动计数器，断在这里会让第 609 行的计数被丢掉、第 611 行拿到重复序号；'
             '「项目明细账」用 MATCH 精确定位序号，重复序号只命中第一个，该项目就永远少列一笔、末笔被挤掉。改为 N(V609)')

# ── 修复 12：条件格式锚点错位 ─────────────────────────────────────
exp_rule = cell_rule = None
for cf in F.conditional_formatting:
    for r in cf.rules:
        if r.type == 'expression' and r.formula == ['$U5="跨期"'] and exp_rule is None:
            exp_rule = r
        if r.type == 'cellIs' and str(cf.sqref).startswith('K5:L13') and cell_rule is None:
            cell_rule = r
assert exp_rule is not None and cell_rule is not None
n_old = sum(len(cf.rules) for cf in F.conditional_formatting)
F.conditional_formatting = ConditionalFormattingList()
exp_rule.priority, cell_rule.priority = 1, 2
F.conditional_formatting.add('A5:X3032', exp_rule)
F.conditional_formatting.add('K5:L3032', cell_rule)
log('修复12', f'【资金流水】{n_old} 条条件格式被粘贴数据打成了碎片，而且锚点全错位：'
              f'B172:L176 那条按 $U168 判、B177:L189 按 $U160 判、B190:L608 与 B627:L3032 按 $U173 判、'
              f'B610:L626 按 $U592 判 —— 第 168 行往下的 B~L 列是照着十几行以外那一行的跨期标志上色的，'
              f'真跨期的不黄、不跨期的反倒黄，同一行左右两半颜色还打架。'
              f'现在合并成两条干净的规则：A5:X3032 按 $U5="跨期" 标色、K5:L3032 负数标红')

# ── 修复 13：任意时段查询 收入小计口径 ─────────────────────────────
Q['A23'] = '小计（不含往来及保证金）'
Q['B23'] = '=SUM(B9:B22)-SUMIF(基础资料!$R$4:$R$17,"往来及保证金",$B$9:$B$22)'
for r in range(9, 23):
    Q[f'C{r}'] = f'=IF(基础资料!$R{r-5}="往来及保证金","",IFERROR(B{r}/$B$23,""))'
log('修复13', '【任意时段查询】B23「收入小计」把保证金退回/借款收回/股东投入/内部调拨这四类往来款也加了进去，'
              '而同页 E3「区间损益收入」是剔除了它们的，两个数对不上；C 列占比又拿含往来的分子去除以不含往来的 E3，'
              '有保证金进账的月份占比之和会超过 100%。现在小计剔除往来及保证金、占比改用小计当分母、'
              '往来行的占比留空 —— 与「月度汇报表」第二栏的处理完全一致')

wb.save(DST); open('fix7_log.txt','w',encoding='utf-8').write('\n'.join(LOG))
print('已保存', DST)
