# -*- coding: utf-8 -*-
import openpyxl
SRC, DST = 'A052_fixed4.xlsx', 'A052_fixed5.xlsx'
LOG=[]
def log(t,m): LOG.append(f'[{t}] {m}'); print(f'[{t}] {m}')
wb = openpyxl.load_workbook(SRC)
R = wb['月度汇报表']
FR='资金流水'
for r in range(115, 135):
    guard = f'OR($A{r}="",OR($B{r}="",AND($E$2<>"全部",$B{r}<>$E$2)))'
    acct  = f',{FR}!$C$5:$C$3032,$A{r}'
    R[f'C{r}'] = (
        f'=IF({guard},"",'
        f'SUMIFS({FR}!$M$5:$M$3032,{FR}!$E$5:$E$3032,"期初余额"{acct},{FR}!$S$5:$S$3032,"<="&$B$2)'
        f'+SUMIFS({FR}!$K$5:$K$3032{acct},{FR}!$S$5:$S$3032,"<"&$B$2)'
        f'-SUMIFS({FR}!$L$5:$L$3032{acct},{FR}!$S$5:$S$3032,"<"&$B$2))')
log('修复5', '【月度汇报表】"六、分账户资金余额"C115:C134 的「期初余额」也是把 M 列滚动余额整段相加'
             '（双思威基本户一个账户就算出 1.64 亿），一并改为 期初余额行 + 上月末以前的收 − 支')
wb.save(DST)
open('fix5_log.txt','w',encoding='utf-8').write('\n'.join(LOG))
print('已保存', DST)
