# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
SRC, DST = 'A052_final7.xlsx', 'A052_final8.xlsx'
LOG=[]
def log(t,m): LOG.append(f'[{t}] {m}'); print(f'[{t}] {m}')
wb = openpyxl.load_workbook(SRC)
B, F, R = wb['基础资料'], wb['资金流水'], wb['月度汇报表']

# ── 修复 33：账户没配"所属公司"时，公司列返回数字 0，全集团合计会把这些钱吃掉 ──
for r in range(5, 3033):
    F[f'D{r}'] = (f'=IF($C{r}="","",IFERROR(IF(INDEX(基础资料!$F$4:$F$23,MATCH($C{r},基础资料!$E$4:$E$23,0))="",'
                  f'"未配所属公司",INDEX(基础资料!$F$4:$F$23,MATCH($C{r},基础资料!$E$4:$E$23,0))),"未匹配账户"))')
for r in range(115, 135):
    i = r - 111                                   # 基础资料 第 4 行起
    R[f'B{r}'] = f'=IF(基础资料!$E{i}="","",IF(基础资料!$F{i}="","未配所属公司",基础资料!$F{i}))'
log('修复33', '【资金流水】D 列「公司(自动)」用 INDEX 取账户档案的所属公司，账户档案里"备用7~备用10"'
              '这四个位子没填所属公司，INDEX 取到空格返回的是**数字 0**（不是空文本）。'
              '而全表大量公式写的是 SUMIFS(…,$D 列,IF($E$2="全部","*",$E$2))，通配符 "*" 只匹配文本、'
              '不匹配数字 —— 一旦启用这几个备用账户记流水，这些钱在"全部"集团合计里会整笔蒸发，'
              '现金流入流出、内帐盈亏全都少算，而且没有任何报错。现在改为显示"未配所属公司"（文本），'
              '钱进得了合计、也一眼看得出该去补档案。【月度汇报表】分账户栏同步')

# ── 修复 34：下拉全都没开"输入无效数据时警告"，等于没有约束 ────────
n=0
for sh in wb.sheetnames:
    for dv in wb[sh].data_validations.dataValidation:
        if dv.type in ('list','whole') and not dv.showErrorMessage:
            dv.showErrorMessage = True
            dv.errorStyle = 'stop'
            dv.errorTitle = '这个值不在清单里'
            dv.error = '请从下拉列表里选；确实是新增的，先到【基础资料】把它加进对应清单，下拉会自动跟着变。'
            n+=1
log('修复34', f'{n} 处下拉全部打开"输入无效数据时显示出错警告(停止)"。原来只提供下拉、不拦截手输，'
              f'而这套表的每一个汇总都靠字符串精确匹配，下拉不拦截等于没有约束 —— '
              f'脏数据已经进来了：应收应付台账的公司被打成"中"、开票登记混进 31 个不在客户清单里的名字、'
              f'资金流水混进 2 个连项目档案都没有的项目名，这几笔在按公司/按项目的报表里永远找不到')

# ── 修复 35：基础资料自己的几个关键字段反而没有下拉 ─────────────────
SPEC = [
    ('$C$4:$C$23',  '"一般纳税人,小规模纳税人"', '纳税人类型'),
    ('$F$4:$F$23',  '基础资料!$A$4:$A$23',       '账户所属公司'),
    ('$K$4:$K$153', '基础资料!$A$4:$A$23',       '项目所属公司'),
    ('$R$4:$R$17',  '"经营收入,非经营收入,往来及保证金"', '收入大类'),
    ('$W$4:$W$49',  '"项目成本,管理费用,人工社保,税费,财务费用,往来及保证金"', '费用大类'),
]
for sq, f1, name in SPEC:
    dv = DataValidation(type='list', formula1=f1, allow_blank=True, showErrorMessage=True,
                        errorStyle='stop', errorTitle=f'{name}只能从清单里选',
                        error='这一列的值下游报表是逐字精确匹配的，打错一个字，这一类的数就整块不进报表了。')
    B.add_data_validation(dv); dv.add(sq)
log('修复35', '【基础资料】给 纳税人类型 / 账户所属公司 / 项目所属公司 / 收入大类 / 费用大类 五列补上下拉。'
              '这几列的值下游全是硬匹配：费用大类多一个字，这类支出就不进内帐盈亏；'
              '纳税人类型打错，这家公司的发票在"小规模合计"和"一般纳税人合计"里会同时消失；'
              '收入大类决定一笔收入走开票口径还是流水口径、要不要剔除往来 —— 原来这几列一条校验都没有')

wb.save(DST); open('fix15_log.txt','w',encoding='utf-8').write('\n'.join(LOG))
print('已保存', DST)
