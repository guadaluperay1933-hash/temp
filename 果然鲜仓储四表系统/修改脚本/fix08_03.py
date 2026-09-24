# -*- coding: utf-8 -*-
"""A050 第八轮 · 《03 财务账套与报表》

 ① 科目表里「资金账户」有两列（K 和 S），日记账的下拉却挂在只有 6 个的 S 列上 ——
    新加的账户永远选不到。现在只留 K 列一份，右边 T 列跟它一一对应写会计科目。
 ② 下拉区域太短、写死行数：改成「按表头名字找列、有多少条就多长」的动态区域。
 ③ 费用项目 71 个，但「费用对应科目」只填到第 48 个 ——
    后面 23 个选了以后对应科目取回来是 0，就是「有的不会自动跳转科目」。补齐 + 加兜底。
 ④ 自动规则只有 23 条，科目表里的业务类型有 43 个：缺的补上，区域从 40 行放宽到 143 行。
 ⑤ 期初余额、资金与往来余额表原来只认 6 个资金账户，一起放宽到 40 个 / 14 个。

跑法：python3 fix08_03.py <入> <出>
"""
import sys, os, copy
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from style01 import *
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)

K0, K1 = 4, 123          # 科目表下拉选项区的行范围（原来到 123，保持）
ACC_N = 40               # 资金账户最多放多少个
RULE0, RULE1 = 4, 143    # 自动规则行范围

# ══ ① 科目表：资金账户并成一列，T 列跟 K 列对齐 ══
ws = wb['科目表']
acc = [ws.cell(r, 11).value for r in range(4, 141) if ws.cell(r, 11).value not in (None, '')]
SUBJ = {   # 新增账户 → 会计科目
    '工行-姜诗林': '银行存款', '建行-张涛': '银行存款',
    '慢慢微信': '其他货币资金', '果然鲜微信': '其他货币资金',
}
BASE = {'库存现金': '库存现金', '基本户-银行存款': '银行存款', '一般户-银行存款': '银行存款',
        '微信': '其他货币资金', '支付宝': '其他货币资金', '承兑汇票': '应收票据'}
BASE.update(SUBJ)
st_T = copy.copy(ws['T4']._style)
for i, a in enumerate(acc):
    r = 4 + i
    c = ws.cell(r, 20)                      # T 列
    c._style = copy.copy(st_T)
    if c.value in (None, ''):
        c.value = BASE.get(a, '其他货币资金')
ws['T3'].value = '对应会计科目\n(对着左边 K 列的资金账户)'
for r in range(4, 141):                     # S 列整列清掉，只留一句说明
    ws.cell(r, 19).value = None
ws['S3'].value = '（已并入 K 列）'
ws['S3'].font = F_NOTE
print(f'  ✓ 科目表：资金账户只留 K 列（{len(acc)} 个），T 列对应会计科目已补齐')

# 费用项目 → 对应科目，把空着的 23 个补上
FEE = {
    '医药费': '管理费用', '赔付出库': '主营业务成本', '借款': '短期借款', '车充电': '管理费用',
    '水果销售': '主营业务成本', '还款': '应付账款', '代办费': '管理费用',
    '转入微信': '其他货币资金', '车辆使用费': '管理费用', '包装物料款': '主营业务成本',
    '还信用卡': '其他应付款', '滴滴打车': '管理费用', '个人消费': '其他应收款',
    '儿子消费': '其他应收款', '婆婆消费': '其他应收款', '杨洋消费': '其他应收款',
    '鱼消费': '其他应收款', '商店': '管理费用', '超市': '管理费用', '快递费': '销售费用',
    '蛋糕': '管理费用', '住宿费': '管理费用', '劳保': '管理费用',
}
st_U = copy.copy(ws['U4']._style)
n = 0
for r in range(4, 141):
    m = ws.cell(r, 13).value                # M 费用项目
    if m in (None, ''): continue
    u = ws.cell(r, 21)                      # U 费用对应科目
    if u.value in (None, ''):
        u._style = copy.copy(st_U)
        u.value = FEE.get(m, '管理费用'); n += 1
print(f'  ✓ 科目表：费用对应科目补了 {n} 个')

biz = [ws.cell(r, 10).value for r in range(4, 141) if ws.cell(r, 10).value not in (None, '')]
subj_names = [ws.cell(r, 2).value for r in range(4, 82) if ws.cell(r, 2).value not in (None, '')]
fee_items = [ws.cell(r, 13).value for r in range(4, 141) if ws.cell(r, 13).value not in (None, '')]

# ══ ② 动态下拉区（按表头名字找列） ══
def dyn(sheet, title, hdr_last, r0, r1, hrow=3):
    m = f'MATCH("{title}",{sheet}!$A${hrow}:${hdr_last}${hrow},0)'
    c = f'COUNTA(INDEX({sheet}!$A${r0}:${hdr_last}${r1},0,{m}))'
    return f'OFFSET({sheet}!$A${r0},0,{m}-1,MAX(1,{c}),1)'

DN = {
    '资金账户表': dyn('科目表', '资金账户', 'Q', K0, K1),
    '业务类型表': dyn('科目表', '业务类型', 'Q', K0, K1),
    '费用项目表': dyn('科目表', '费用项目', 'Q', K0, K1),
    '收入项目表': dyn('科目表', '收入项目', 'Q', K0, K1),
    '往来单位表': dyn('科目表', '往来单位', 'Q', K0, K1),
    '结算方式表': dyn('科目表', '结算方式', 'Q', K0, K1),
    '税率表':     dyn('科目表', '税率',     'Q', K0, K1),
    '是否表':     dyn('科目表', '是否',     'Q', K0, K1),
    '记账类别表': dyn('科目表', '记账类别', 'Q', K0, K1),
    '科目名称表': dyn('科目表', '科目名称', 'Q', 4, 81),
    '报表项目表': dyn('科目表', '报表项目', 'Q', 4, 81),
}
for nm, f in DN.items():
    if nm in wb.defined_names: del wb.defined_names[nm]
    wb.defined_names.add(DefinedName(nm, attr_text=f))
print(f'  ✓ 建了 {len(DN)} 个动态下拉区')

# ══ ③ 自动规则：补齐每一个业务类型 ══
ws = wb['自动规则']
old = {}
order = []
for r in range(4, 42):
    b = ws.cell(r, 2).value
    if b in (None, ''): continue
    old[b] = (ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value)
    order.append(b)
DEF = {
    '周转费': ('应收账款', '应收账款', '收客户周转打冷费 → 冲减应收'),
    '运费': ('应收账款', '应付账款', '收客户运费 → 冲减应收；付承运方运费 → 冲减应付'),
    '其他服务费': ('应收账款', '应收账款', '零星服务费 → 冲减应收'),
    '押金收取': ('其他应付款', '其他应付款', '收客户押金是欠客户的钱，不是收入'),
    '押金退还': ('其他应付款', '其他应付款', '退押金＝借记冲减其他应付款'),
    '装卸费支出': ('应付账款', '应付账款', '付给装卸队的钱 → 冲减应付'),
    '电费': ('应付账款', '应付账款', '电费 → 冲减应付'),
    '水费': ('应付账款', '应付账款', '水费 → 冲减应付'),
    '包装物料销售': ('应收账款', '应收账款', '包装物料卖断的货款 → 冲减应收'),
    '业务招待费': ('管理费用', '管理费用', '招待支出直接进管理费用'),
    '差旅费': ('管理费用', '管理费用', '差旅支出直接进管理费用'),
    '车辆使用费': ('管理费用', '管理费用', '油费、过路费、维修 → 管理费用'),
    '车费': ('管理费用', '管理费用', '打车、班车 → 管理费用'),
    '餐费': ('管理费用', '管理费用', '员工餐 → 管理费用'),
    '办公用品': ('管理费用', '管理费用', '办公消耗 → 管理费用'),
    '财务': ('财务费用', '财务费用', '手续费一类 → 财务费用'),
    '利息': ('财务费用', '财务费用', '利息收入/支出 → 财务费用'),
    '租库定金': ('预付账款', '预付账款', '付出去的定金先挂预付账款，正式起租再转租金'),
    '账户互转': ('其他货币资金', '其他货币资金', '两个自有账户之间调钱，一出一入各记一行；'
                 '对方科目请在日记账最右边「科目手工覆盖」填上另一个账户的科目'),
    '家用': ('其他应收款', '其他应收款', '老板家用，先挂其他应收款，年底再定性'),
    '个人消费': ('其他应收款', '其他应收款', '与经营无关的个人支出，挂其他应收款'),
    '个人收入': ('其他应收款', '其他应收款', '个人进账，挂其他应收款'),
    '国外': ('其他应收款', '其他应收款', '出国相关支出，先挂其他应收款'),
    '永城财务': ('其他应收款', '其他应收款', '关联方往来，挂其他应收款'),
}
for b in biz:
    if b not in old:
        c, d, e = DEF.get(b, ('其他应收款', '其他应付款', '（自动补的默认口径，可以改）'))
        old[b] = (c, d, e); order.append(b)
assert len(order) <= RULE1 - RULE0 + 1, len(order)

st = {c: copy.copy(ws[f'{c}4']._style) for c in 'ABCDE'}
h4 = ws.row_dimensions[4].height
note = ws['A42'].value
for m in [str(x) for x in ws.merged_cells.ranges]:
    if m.startswith('A42'): ws.unmerge_cells(m)
for r in range(4, 43):
    for c in range(1, 6): ws.cell(r, c).value = None
for i, b in enumerate(order):
    r = RULE0 + i
    c_, d_, e_ = old[b]
    for col in 'ABCDE': ws[f'{col}{r}']._style = copy.copy(st[col])
    ws[f'A{r}'].value = i + 1
    ws[f'B{r}'].value = b
    ws[f'C{r}'].value = c_
    ws[f'D{r}'].value = d_
    ws[f'E{r}'].value = e_
    if h4: ws.row_dimensions[r].height = h4
for r in range(RULE0 + len(order), RULE1 + 1):
    for col in 'ABCDE': ws[f'{col}{r}']._style = copy.copy(st[col])
    if h4: ws.row_dimensions[r].height = h4
NR = RULE1 + 1
ws.merge_cells(f'A{NR}:E{NR}')
put(ws, f'A{NR}', '提示：本表由【科目表】J 列「业务类型」自动对齐 —— 新增业务类型请先在科目表 J 列加一行，'
                  '再回到本表空白行填上「收到钱时 / 付出钱时」两个科目。没填的，收款默认走「其他应收款」、'
                  '付款默认走「其他应付款」；个别单据要例外，就在日记账最右边「科目手工覆盖」里填。',
    font=F_NOTE, align=CL, border=None)
ws.data_validations.dataValidation = []
for sq, nm in ((f'B{RULE0}:B{RULE1}', '业务类型表'), (f'C{RULE0}:D{RULE1}', '科目名称表')):
    dv = DataValidation(type='list', formula1=f'={nm}', allow_blank=True, showDropDown=False,
                            showInputMessage=True)
    ws.add_data_validation(dv); dv.add(sq)
print(f'  ✓ 自动规则：{len(order)} 条（原 {len(order) - len(biz) + len([b for b in biz if b in old])}…'
      f'补了 {len([b for b in biz if b in DEF])} 条），区域放宽到第 {RULE1} 行')

# ══ ④ 期初余额：资金账户跟科目表 K 列走，放宽到 40 个 ══
ws = wb['期初余额']
E0, E1 = 4, 4 + ACC_N - 1
keep = {ws.cell(r, 8).value: ws.cell(r, 9).value for r in range(4, 12)
        if ws.cell(r, 8).value not in (None, '')}
stH, stI = copy.copy(ws['H4']._style), copy.copy(ws['I4']._style)
for i in range(ACC_N):
    r, k = E0 + i, 4 + i
    ws[f'H{r}']._style = copy.copy(stH)
    ws[f'I{r}']._style = copy.copy(stI)
    ws[f'H{r}'].value = f'=IF(科目表!$K{k}="","",科目表!$K{k})'
    nm = acc[i] if i < len(acc) else None
    ws[f'I{r}'].value = keep.get(nm, 0 if nm else None)
ws['H3'].value = '资金账户\n(自动跟科目表)'
print(f'  ✓ 期初余额：资金账户自动跟科目表 K 列，放宽到 {ACC_N} 个')

# ══ ⑤ 资金日记账 ══
ws = wb['资金日记账']
ws.data_validations.dataValidation = []
for sq, nm in (('C4:C603', '资金账户表'), ('F4:F603', '业务类型表'),
               ('G4:G603', '费用项目表'), ('T4:T603', '科目名称表')):
    dv = DataValidation(type='list', formula1=f'={nm}', allow_blank=True, showDropDown=False,
                            showInputMessage=True)
    dv.errorTitle, dv.error = '不在科目表里', '请从下拉里选；没有的先去【科目表】对应列加一行'
    dv.promptTitle, dv.prompt = '从科目表取', '点右边小箭头选。科目表里新加的内容会自动出现在这里'
    ws.add_data_validation(dv); dv.add(sq)
dv = DataValidation(type='list', formula1='=_自动清单!$E$4:$E$203', allow_blank=True, showDropDown=False,
                            showInputMessage=True)
ws.add_data_validation(dv); dv.add('E4:E603')
dv = DataValidation(type='list', formula1='"是"', allow_blank=True, showDropDown=False,
                            showInputMessage=True)
ws.add_data_validation(dv); dv.add('U4:U603')

IU = f'INDEX(科目表!$U${K0}:$U${K1},MATCH($G{{r}},科目表!$M${K0}:$M${K1},0))'
IC = f'INDEX(自动规则!$C${RULE0}:$C${RULE1},MATCH($F{{r}},自动规则!$B${RULE0}:$B${RULE1},0))'
ID = f'INDEX(自动规则!$D${RULE0}:$D${RULE1},MATCH($F{{r}},自动规则!$B${RULE0}:$B${RULE1},0))'
IT = f'INDEX(科目表!$T${K0}:$T${K1},MATCH($C{{r}},科目表!$K${K0}:$K${K1},0))'

def pick(expr, default, r):
    """取回来是空格子（INDEX 会返回 0）或者压根没配，就用兜底科目"""
    e = expr.format(r=r)
    return f'IF(IFERROR({e},"")&""="","{default}",{e})'

for r in range(4, 604):
    put(ws, f'J{r}', f'=IF($C{r}="","",ROUND(IFERROR(INDEX(期初余额!$I${E0}:$I${E1},'
                     f'MATCH($C{r},期初余额!$H${E0}:$H${E1},0)),0)'
                     f'+SUMIFS($H$4:$H{r},$C$4:$C{r},$C{r})-SUMIFS($I$4:$I{r},$C$4:$C{r},$C{r}),2))',
        font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'K{r}', f'=IF($T{r}<>"",$T{r},IF($C{r}="","",'
                     f'IF(AND(N($I{r})>0,$G{r}<>""),{pick(IU, "管理费用", r)},'
                     f'IF(N($H{r})>0,{pick(IC, "其他应收款", r)},{pick(ID, "其他应付款", r)}))))',
        font=F_TXT, fill=FILL_AUTO)
    put(ws, f'M{r}', f'=IF($C{r}="","",{pick(IT, "其他货币资金", r)})',
        font=F_TXT, fill=FILL_AUTO)
ws['A2'].value = ('★ 资金账户 / 业务类型 / 费用项目三个下拉，全部直接取【科目表】K、J、M 三列 —— '
                  '在科目表里加一行，这里马上就能选到。「对应科目」按【自动规则】和科目表 U 列自动判断；'
                  '配了但没填科目的，收款走「其他应收款」、付款走「其他应付款」、费用走「管理费用」兜底，'
                  '个别单据要例外就在最右边「科目手工覆盖」里填。')
print('  ✓ 资金日记账：下拉改动态区；对应科目取不到时按兜底科目，不再显示 0')

# ══ ⑥ 资金与往来余额表：资金账户从 6 个放到 14 个 ══
ws = wb['资金与往来余额表']
N_ACC, N_CUS = 14, 110
snap = {k: [copy.copy(ws.cell(k, c)._style) for c in range(1, 10)]
        for k in (5, 11, 12, 14, 15, 16, 126)}
hh = {k: ws.row_dimensions[k].height for k in (5, 11, 12, 14, 15, 16)}
for m in [str(x) for x in ws.merged_cells.ranges]:
    if m.startswith('A14'): ws.unmerge_cells(m)
for r in range(5, 140):
    for c in range(1, 10):
        ws.cell(r, c).value = None
        ws.cell(r, c)._style = copy.copy(snap[16][c - 1])
    ws.row_dimensions[r].height = 18

def row(r, src):
    for c in range(1, 10): ws.cell(r, c)._style = copy.copy(snap[src][c - 1])
    if src in hh and hh[src]: ws.row_dimensions[r].height = hh[src]

A0 = 5
for i in range(N_ACC):
    r, k = A0 + i, 4 + i
    row(r, 5)
    ws[f'A{r}'].value = f'=IF(科目表!$K{k}="","",科目表!$K{k})'
    ws[f'B{r}'].value = (f'=IF($A{r}="","",ROUND(IFERROR(INDEX(期初余额!$I${E0}:$I${E1},'
                         f'MATCH($A{r},期初余额!$H${E0}:$H${E1},0)),0),2))')
    ws[f'C{r}'].value = f'=IF($A{r}="","",ROUND(SUMIFS(资金日记账!$H$4:$H$603,资金日记账!$C$4:$C$603,$A{r}),2))'
    ws[f'D{r}'].value = f'=IF($A{r}="","",ROUND(SUMIFS(资金日记账!$I$4:$I$603,资金日记账!$C$4:$C$603,$A{r}),2))'
    ws[f'E{r}'].value = f'=IF($A{r}="","",ROUND($B{r}+$C{r}-$D{r},2))'
A1 = A0 + N_ACC - 1
T1 = A1 + 1
row(T1, 11)
ws[f'A{T1}'].value = '合计'
for c in 'BCDE': ws[f'{c}{T1}'].value = f'=ROUND(SUM({c}{A0}:{c}{A1}),2)'
T2 = T1 + 1
row(T2, 12)
ws[f'A{T2}'].value = '账面货币资金（科目余额表）'
ws[f'E{T2}'].value = ('=ROUND(SUMIFS(科目余额表!$I$4:$I$81,科目余额表!$D$4:$D$81,"货币资金")'
                      '-SUMIFS(科目余额表!$J$4:$J$81,科目余额表!$D$4:$D$81,"货币资金"),2)')
ws[f'F{T2}'].value = (f'=IF(ROUND($E{T2}-$E{T1},2)=0,"✔ 与日记账一致",'
                      f'"✘ 差额 "&TEXT($E{T2}-$E{T1},"#,##0.00"))')
S2 = T2 + 2
row(S2, 14)
ws.merge_cells(f'A{S2}:I{S2}')
ws[f'A{S2}'].value = ('二、往来单位余额（往来单位自动生成 · 应收＝业务册已确认的应收 '
                      '− 资金日记账实收；应付＝账面凭证）')
H2 = S2 + 1
row(H2, 15)
for c, t in zip('ABCDEFGHI', ['往来单位', '期初应收', '本期应收发生', '资金实收', '应收余额',
                              '期初应付', '本期应付发生', '资金实付', '应付余额']):
    ws[f'{c}{H2}'].value = t
C0 = H2 + 1
C1 = C0 + N_CUS - 1
off = C0 - 1
for r in range(C0, C1 + 1):
    row(r, 16)
    ws[f'A{r}'].value = f'=IFERROR(INDEX(_自动清单!$E$4:$E$253,ROW()-{off}),"")'
    ws[f'C{r}'].value = (f'=IF($A{r}="","",ROUND(SUMIFS(进销存对接!$G$4:$G$305,'
                         f'进销存对接!$D$4:$D$305,"客户应收",进销存对接!$C$4:$C$305,$A{r}),2))')
    ws[f'D{r}'].value = (f'=IF($A{r}="","",ROUND(SUMIFS(资金日记账!$H$4:$H$603,'
                         f'资金日记账!$E$4:$E$603,$A{r},资金日记账!$K$4:$K$603,"应收账款"),2))')
    ws[f'E{r}'].value = f'=IF($A{r}="","",ROUND(N($B{r})+$C{r}-$D{r},2))'
    ws[f'G{r}'].value = (f'=IF($A{r}="","",ROUND(SUMIFS(进销存对接!$F$4:$F$305,'
                         f'进销存对接!$D$4:$D$305,"客户应收",进销存对接!$C$4:$C$305,$A{r})'
                         f'+SUMIFS(记账凭证!$I$4:$I$3609,记账凭证!$J$4:$J$3609,$A{r},'
                         f'记账凭证!$E$4:$E$3609,"应付账款"),2))')
    ws[f'H{r}'].value = (f'=IF($A{r}="","",ROUND(SUMIFS(记账凭证!$H$4:$H$3609,'
                         f'记账凭证!$J$4:$J$3609,$A{r},记账凭证!$E$4:$E$3609,"应付账款"),2))')
    ws[f'I{r}'].value = f'=IF($A{r}="","",ROUND(N($F{r})+$G{r}-$H{r},2))'
T3 = C1 + 1
row(T3, 126)
ws[f'A{T3}'].value = '合计'
for c in 'BCDEFGHI': ws[f'{c}{T3}'].value = f'=ROUND(SUM({c}{C0}:{c}{C1}),2)'
print(f'  ✓ 资金与往来余额表：资金账户 {N_ACC} 行（{A0}~{A1}），往来单位 {N_CUS} 行（{C0}~{C1}）')

wb.save(OUT)
print('已写', OUT)
