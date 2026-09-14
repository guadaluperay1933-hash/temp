# -*- coding: utf-8 -*-
"""独立复算：不看表里的公式，直接从录入表的原始数据用 python 重算一遍，
再和 LibreOffice 全量重算后的表内值逐项核对。27 项全部一致才算修对了。
用法：先 python3 /mnt/skills/public/xlsx/scripts/recalc.py 台账.xlsx，再跑本脚本。"""
import sys, datetime, openpyxl
SRC = sys.argv[1] if len(sys.argv)>1 else '资金流水成本核算台账_A052.xlsx'
v=openpyxl.load_workbook(SRC,data_only=True)
F,I,L,R,P,Q,T,Y = (v['资金流水'],v['开票登记'],v['应收应付台账'],v['月度汇报表'],
                   v['项目核算表'],v['任意时段查询'],v['月度趋势数据'],v['资金日报表'])
num=lambda x: x if isinstance(x,(int,float)) else 0
flows=[dict(date=F.cell(r,2).value,acct=F.cell(r,3).value,comp=F.cell(r,4).value,typ=F.cell(r,5).value,
            proj=F.cell(r,6).value,K=num(F.cell(r,11).value),Lv=num(F.cell(r,12).value),M=F.cell(r,13).value,
            S=F.cell(r,19).value,T=F.cell(r,20).value,W=F.cell(r,23).value,X=F.cell(r,24).value)
       for r in range(5,3033) if F.cell(r,2).value is not None]
inv=[dict(date=I.cell(r,2).value,proj=I.cell(r,7).value,J=num(I.cell(r,10).value),Lv=num(I.cell(r,12).value),
          M=num(I.cell(r,13).value),Pv=num(I.cell(r,16).value),Rm=I.cell(r,18).value)
     for r in range(5,505) if I.cell(r,2).value is not None]
M=R['B2'].value; BIG=('项目成本','管理费用','人工社保','税费','财务费用')
op=sum(num(x['M']) for x in flows if x['typ']=='期初余额' and num(x['S'])<=M)
end=op+sum(x['K'] for x in flows if num(x['S'])<=M)-sum(x['Lv'] for x in flows if num(x['S'])<=M)
d0,d1=Y['B2'].value,Y['E2'].value
opd=sum(num(x['M']) for x in flows if x['typ']=='期初余额' and x['date']<=d0)
q0,q1=Q['B2'].value,Q['B3'].value
CASES=[
 ('本月开票收入 B6',   sum(x['M'] for x in inv if x['Rm']==M), R['B6']),
 ('本月开票税额 B8',   sum(x['Lv'] for x in inv if x['Rm']==M), R['B8']),
 ('现金流入 D6',      sum(x['K'] for x in flows if x['S']==M), R['D6']),
 ('现金流出 F6',      sum(x['Lv'] for x in flows if x['S']==M), R['F6']),
 ('损益收入 B7',      sum(x['M'] for x in inv if x['Rm']==M)+sum(x['K'] for x in flows if x['T']==M and x['X']=='非经营收入'), R['B7']),
 ('损益成本费用 D7',   sum(x['Lv'] for x in flows if x['T']==M and x['W'] in BIG), R['D7']),
 ('期末资金总额 H8',   end, R['H8']),
 ('已开票未收款 D8',   sum(x['Pv'] for x in inv if x['Pv']>0), R['D8']),
 ('台账应收 F8',      sum(num(L.cell(r,12).value) for r in range(5,305) if L.cell(r,2).value=='应收'), R['F8']),
 ('台账应付 B9',      sum(num(L.cell(r,12).value) for r in range(5,305) if L.cell(r,2).value=='应付'), R['B9']),
 ('未分类支出 B37',    sum(x['Lv'] for x in flows if x['T']==M and x['W'] in (None,'')), R['B37']),
 ('分公司现金流出 D111', sum(x['Lv'] for x in flows if x['S']==M), R['D111']),
 ('分账户期末 F135',   end, R['F135']),
 ('开票汇总价税 F161',  sum(x['M'] for x in inv if x['Rm']==M), R['F161']),
 ('项目累计开票 F155',  sum(x['M'] for x in inv if x['proj']), P['F155']),
 ('项目成本合计 X155',  sum(x['Lv'] for x in flows if x['proj'] and x['W']=='项目成本'), P['X155']),
 ('项目累计收款 G155',  sum(x['K'] for x in flows if x['proj']), P['G155']),
 ('项目未收款 I155',   sum(x['Pv'] for x in inv if x['proj'] and x['Pv']>0), P['I155']),
 ('趋势·全年开票 C16',  sum(x['M'] for x in inv if isinstance(x['Rm'],int) and 202601<=x['Rm']<=202612), T['C16']),
 ('趋势·全年现金收 D16', sum(x['K'] for x in flows if isinstance(x['S'],int) and 202601<=x['S']<=202612), T['D16']),
 ('趋势·全年现金支 E16', sum(x['Lv'] for x in flows if isinstance(x['S'],int) and 202601<=x['S']<=202612), T['E16']),
 ('日报·期初 A5',     opd+sum(x['K'] for x in flows if x['date']<d0)-sum(x['Lv'] for x in flows if x['date']<d0), Y['A5']),
 ('日报·期间收 C5',    sum(x['K'] for x in flows if d0<=x['date']<=d1)+sum(num(x['M']) for x in flows if x['typ']=='期初余额' and d0<x['date']<=d1), Y['C5']),
 ('日报·期间付 E5',    sum(x['Lv'] for x in flows if d0<=x['date']<=d1), Y['E5']),
 ('查询·区间开票 E2',   sum(x['M'] for x in inv if q0<=x['date']<=q1), Q['E2']),
 ('查询·区间现金支 H2',  sum(x['Lv'] for x in flows if q0<=x['date']<=q1), Q['H2']),
]
bad=0
for name, mine, cell in CASES:
    ok = abs(num(mine)-num(cell.value))<0.01
    bad += not ok
    print(('  ✓' if ok else '  ✗'), f'{name:22s} 独立复算 {num(mine):>18,.2f}   表内 {num(cell.value):>18,.2f}')
print(f'\n{len(CASES)-bad}/{len(CASES)} 项一致')
sys.exit(1 if bad else 0)
