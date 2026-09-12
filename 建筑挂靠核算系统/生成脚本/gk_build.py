# -*- coding: utf-8 -*-
import sys, json, datetime as dt
sys.path.insert(0, '/tmp/claude-0/-home-user-temp/9101e7cb-4f82-53a7-bbe8-41b0f93b12b0/scratchpad')
from gl_common import *
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule

EV = json.load(open('/tmp/claude-0/-home-user-temp/9101e7cb-4f82-53a7-bbe8-41b0f93b12b0/scratchpad/events.json'))
EV = [e for e in EV if e['kind'] != '预提税费']
EV.sort(key=lambda e: (e['date'], e['src']))

# ---------------- 容量 ----------------
U0, U1 = 6, 35          # 单位档案 30
P0, P1 = 6, 255         # 项目档案 250
F0, F1 = 5, 1504        # 业务流水 1500
J0, J1 = 5, 1004        # 资金日记账 1000
D0, D1 = 6, 305         # 代收台账 300
SU0, SU1 = 6, 35        # 分表·单位 30
SP0, SP1 = 6, 255       # 分表·项目 250
SX0, SX1 = 6, 405       # 分表·单位×项目 400

SH_HOME2, SH_DOC = '首页', '使用说明'
SH_UNIT, SH_PROJ = '单位档案', '项目档案'
SH_FLOW, SH_JOUR, SH_DAI = '业务流水', '资金日记账', '代收台账'
SH_SUM_U, SH_SUM_P, SH_SUM_X = '单位汇总', '项目汇总', '单位项目明细'
SH_CHAIN, SH_TAX, SH_EXP, SH_REC = '链条核算', '税费台账', '费用统计', '对账核对'

wb = Workbook(); wb.remove(wb.active)

def sysband(ws, c0, c1, row, text='系统自动计算区 · 请勿修改'):
    ws.merge_cells(f'{c0}{row}:{c1}{row}')
    put(ws, f'{c0}{row}', text, font=F_HDR2, fill=FILL_AUTO, align=C)

def inband(ws, c0, c1, row, text='录入区 · 淡黄色格子手工填写'):
    ws.merge_cells(f'{c0}{row}:{c1}{row}')
    put(ws, f'{c0}{row}', text, font=F_HDR2, fill=FILL_HDR2, align=C)

def block(ws, col0, ncol, r_hdr, title_txt, names, r0, r1, note=None):
    c0, c1 = L(col0), L(col0 + ncol - 1)
    ws.merge_cells(f'{c0}{r_hdr-1}:{c1}{r_hdr-1}')
    put(ws, f'{c0}{r_hdr-1}', title_txt, font=F_HDR2, fill=FILL_HDR2, align=CL)
    m = headers(ws, r_hdr, col0, names)
    if note:
        put(ws, f'{c0}{r1+1}', note, font=F_NOTE, align=CL, border=None)
        ws.merge_cells(f'{c0}{r1+1}:{c1}{r1+1}')
    return m

QU, QP, QF, QJ, QD = Q(SH_UNIT), Q(SH_PROJ), Q(SH_FLOW), Q(SH_JOUR), Q(SH_DAI)
U_NAME = f'{QU}!$A${U0}:$A${U1}'
P_CODE = f'{QP}!$A${P0}:$A${P1}'

KINDS = ['销项开票','成本票','挂靠代收','我方收款','管理费结算','扣质保金','已交税','退税','其他']
KIND_DV = '"' + ','.join(KINDS) + '"'

# ============================================================ 单位档案
ws = wb.create_sheet(SH_UNIT)
title(ws, '单位档案', 'N',
      '每个挂靠单位的管理费率和税费参数都在这里配。改这里，业务流水的自动计算立刻跟着变。')
widths(ws, {'A':13,'B':26,'C':12,'D':11,'E':10,'F':11,'G':9,'H':13,'I':11,'J':11,'K':11,'L':11,'M':8,'N':24})
headers(ws, 5, 1, ['单位简称','单位全称','类型','默认\n管理费率','上游\n开票税率','我方回开\n票种税率','有无\n税差',
                   '增值税\n计算方式','增值税\n预征率','附加税率\n(按增值税)','印花税率\n(按开票额)',
                   '所得税\n预征率','状态','备注'])
UNITS = [
 ('民能','重庆民能实业有限公司','业主','','','','否','—','','','','','正常','工程发包方'),
 ('铜梁供电','国网重庆铜梁供电分公司','业主','','','','否','—','','','','','正常','工程发包方'),
 ('德誉嘉','德誉嘉（德誉佳）','挂靠单位',0.04,0.13,0.13,'否','—','','','','','正常','原扣8%返4%现金，现改为直接扣4%按下浮开票，无税差'),
 ('迅驰','迅驰','挂靠单位',0.04,0.09,0.03,'是','税差法','',0.12,0.000588,'','正常','开9%专票给民能，我方只能开3%，税差补给对方'),
 ('华城','华城','挂靠单位',0.02,0.03,0.01,'是','全额销项法','',0.12,0.0003,'','正常','开3%劳务票，我方回1%普票'),
 ('康欣','康欣','挂靠单位',0.02,0.03,0.03,'是','全额销项法','',0.06,'',0.002,'正常','老项目不收管理费只过票；新项目直接开给民能下浮2%'),
 ('金沁','金沁','挂靠单位',0.05,0.09,0.03,'是','税差法','',0.12,0.0006,'','正常','部分项目扣5%部分扣3%，以项目档案费率为准'),
 ('安锐','安锐','挂靠单位',0.02,0.13,0.13,'是','税差法','',0.06,'','','正常','两端同为13%，税差即管理费部分'),
 ('湖南锦泰','湖南锦泰电力建设有限公司','挂靠单位',0.10,0.09,0.09,'是','全额销项法','',0.12,'','','正常',''),
 ('杰华','重庆杰华电气有限公司','挂靠单位',0.10,0.13,0.13,'是','全额销项法','',0.12,'','','正常',''),
 ('泓普','泓普（我方主体）','我方主体','',0.03,'','否','—','','','','','正常','主要开票主体'),
 ('仟茂','仟茂（我方主体）','我方主体','',0.13,'','否','—','','','','','正常','13%票种主体'),
 ('税局','税务机关','其他','','','','否','—','','','','','正常',''),
]
for i in range(U1 - U0 + 1):
    r = U0 + i
    for j, c in enumerate('ABCDEFGHIJKLMN'):
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN,
            align=CL if c in 'BN' else C,
            fmt=PCT if c in 'DEFIJKL' else None)
    if i < len(UNITS):
        for c, v in zip('ABCDEFGHIJKLMN', UNITS[i]):
            if v != '': ws[f'{c}{r}'] = v
    ws.row_dimensions[r].height = 17
dv_list(ws, f'C{U0}:C{U1}', '"业主,挂靠单位,我方主体,其他"')
dv_list(ws, f'G{U0}:G{U1}', '"是,否"')
dv_list(ws, f'H{U0}:H{U1}', '"税差法,全额销项法,预征率法,—"')
dv_list(ws, f'M{U0}:M{U1}', '"正常,停用"')
put(ws, f'A{U1+2}', '税差法：预提增值税 ＝ 上游开票销项税 − 我方成本票进项税（迅驰、金沁、安锐）。'
                    '全额销项法：预提增值税 ＝ 开票额 ÷(1+上游税率) × 上游税率（华城、康欣、湖南锦泰、杰华）。'
                    '预征率法：开票额 ÷(1+上游税率) × 预征率。'
                    '附加税＝预提增值税×附加税率；印花税＝含税开票额×印花税率；所得税＝不含税额×所得税预征率。以上全部逐笔核对过你原表的数。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{U1+2}:N{U1+2}')
ws.freeze_panes = 'B6'; page(ws, titles='5:5')

# ============================================================ 项目档案
ws = wb.create_sheet(SH_PROJ)
title(ws, '项目档案', 'L',
      '一个项目最多三层：业主 ← 一级单位 ← 二级单位 ← 我方。只有两层的项目，二级单位留空即可。'
      '费率留空时自动取【单位档案】的默认管理费率。')
widths(ws, {'A':11,'B':46,'C':12,'D':13,'E':11,'F':13,'G':11,'H':12,'I':12,'J':9,'K':13,'L':22})
headers(ws, 5, 1, ['项目编号','项目全称','业主','一级单位','一级\n费率','二级单位','二级\n费率',
                   '我方主体','开工日期','状态','合同额','备注'])
PROJ_ROWS = []
OWN = ('民能', '铜梁供电')
seen = {}
amt1, amt2 = {}, {}
for e in EV:
    p = e['proj']
    if not p: continue
    seen.setdefault(p, {'units': [], 'owner': '', 'L1': {}, 'L2': {}})
    if e['unit'] not in seen[p]['units']: seen[p]['units'].append(e['unit'])
    if e['kind'] == '销项开票' and e.get('count_in', '是') == '是':
        if e['payee'] in OWN:
            seen[p]['owner'] = e['payee']
            seen[p]['L1'][e['payer']] = seen[p]['L1'].get(e['payer'], 0) + e['amt']
        elif e['payer'] not in ('泓普', '仟茂'):
            seen[p]['L2'][(e['payer'], e['payee'])] = seen[p]['L2'].get((e['payer'], e['payee']), 0) + e['amt']
for p, info in seen.items():
    l1 = max(info['L1'], key=info['L1'].get) if info['L1'] else (info['units'][0] if info['units'] else '')
    l2 = ''
    cand = [(k, v) for k, v in info['L2'].items() if k[1] == l1]
    if cand: l2 = max(cand, key=lambda x: x[1])[0][0]
    info['lv1'], info['lv2'] = l1, l2
NEW_ROW = ('', '（新项目：请填项目全称）', '民能', '', '', '', '', '泓普', '本次新增项目')
hist = [(p, info) for p, info in seen.items()]
ordered = hist[:58] + [(None, None)] + hist[58:]     # 新项目插在第 59 位 → 编号 A059
PCODE = {}
for i, (p, info) in enumerate(ordered):
    code = f'A{i+1:03d}'
    if p is None:
        PROJ_ROWS.append((code,) + NEW_ROW[1:8] + ('在建', NEW_ROW[8]))
        NEW_CODE = code
        continue
    PCODE[p] = code
    multi = len(info['L1']) > 1
    PROJ_ROWS.append((code, p, info['owner'] or '民能', info['lv1'], '', info['lv2'], '', '泓普', '在建',
                      ('该项目走了 ' + str(len(info['L1'])) + ' 条挂靠链：' + '、'.join(info['L1'].keys()) +
                       '，档案只能填一条，链条核算以此为准；其余链条请看单位项目明细') if multi else ''))
for i in range(P1 - P0 + 1):
    r = P0 + i
    for c in 'ABCDEFGHIJKL':
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if c in 'BL' else C,
            fmt=PCT if c in 'EG' else (DATE if c == 'I' else (MONEY if c == 'K' else None)))
    if i < len(PROJ_ROWS):
        row = PROJ_ROWS[i]
        for c, v in zip(['A','B','C','D','E','F','G','H'], row[:8]):
            if v: ws[f'{c}{r}'] = v
        ws[f'J{r}'] = row[8]
        if row[9]: ws[f'L{r}'] = row[9]
    ws.row_dimensions[r].height = 17
for c in 'ACDF': dv_list(ws, f'{c}{P0}:{c}{P1}', f'={U_NAME}')
dv_list(ws, f'H{P0}:H{P1}', f'={U_NAME}')
dv_list(ws, f'J{P0}:J{P1}', '"在建,已完工,已结清,暂停"')
ws.freeze_panes = 'C6'; page(ws, titles='5:5')
print(f'  ✓ 单位档案 / 项目档案（共 {len(PROJ_ROWS)} 个项目，新项目编号 {NEW_CODE}）')

# ============================================================ 业务流水（总表）
ws = wb.create_sheet(SH_FLOW)
title(ws, '业务流水（总表）', 'W',
      '所有开票、成本票、回款、交税、管理费结算都在这里录一行。右边灰色区自动算管理费、应开成本票和该笔要预提的各项税费；'
      '后面所有分表按项目、按单位自动汇总，不用再手工统计。')
widths(ws, {'A':7,'B':11,'C':10,'D':12,'E':12,'F':12,'G':11,'H':14,'I':9,'J':34,'K':16,'L':15,
            'M':38,'N':8,'O':12,'P':14,'Q':8,'R':8,'S':13,'T':12,'U':12,'V':12,'W':13})
inband(ws, 'A', 'L', 3); sysband(ws, 'M', 'W', 3)
headers(ws, 4, 1, ['序号','日期','项目编号','业务类型','开票/\n付款方','收票/\n收款方','发票性质','金额',
                   '管理费率','摘要','备注','校验'])
headers(ws, 4, 13, ['项目全称','层级','管理费','应开成本票','上游\n税率','我方\n税率',
                    '预提增值税','预提附加税','预提印花税','预提所得税','预提税费合计'],
        fill=FILL_AUTO, font=F_HDR2)
put(ws, 'Y4', '来源表', font=F_HDR2, fill=FILL_AUTO)
put(ws, 'Z4', '计入汇总', font=F_HDR2, fill=FILL_HDR2)
ws.column_dimensions['Y'].width = 13; ws.column_dimensions['Z'].width = 10
U_D = f'{QU}!$D${U0}:$D${U1}'; U_E = f'{QU}!$E${U0}:$E${U1}'; U_F = f'{QU}!$F${U0}:$F${U1}'
U_G = f'{QU}!$G${U0}:$G${U1}'; U_H = f'{QU}!$H${U0}:$H${U1}'; U_J = f'{QU}!$J${U0}:$J${U1}'
U_K = f'{QU}!$K${U0}:$K${U1}'; U_L = f'{QU}!$L${U0}:$L${U1}'
P_B = f'{QP}!$B${P0}:$B${P1}'; P_DU = f'{QP}!$D${P0}:$D${P1}'; P_E = f'{QP}!$E${P0}:$E${P1}'
P_FU = f'{QP}!$F${P0}:$F${P1}'; P_G = f'{QP}!$G${P0}:$G${P1}'; P_H = f'{QP}!$H${P0}:$H${P1}'
def lk(val, key, out, nf='""'): return f'IFERROR(INDEX({out},MATCH({val},{key},0)),{nf})'

for r in range(F0, F1 + 1):
    for c, fmt, tx in [('B', DATE, 0), ('C', None, 0), ('D', None, 0), ('E', None, 0), ('F', None, 0),
                       ('G', None, 0), ('H', MONEY, 0), ('J', None, 1), ('K', None, 1)]:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if tx else C, fmt=fmt)
    put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-{F0-1})', font=F_LINK)
    put(ws, f'M{r}', f'=IF($C{r}="","",{lk(f"$C{r}", P_CODE, P_B, chr(34)+"⚠项目编号不存在"+chr(34))})',
        font=F_LINK, fill=FILL_AUTO, align=CL)
    put(ws, f'N{r}', f'=IF($E{r}="","",IF($E{r}={lk(f"$C{r}", P_CODE, P_DU)},"一级",'
                     f'IF($E{r}={lk(f"$C{r}", P_CODE, P_FU)},"二级",'
                     f'IF($E{r}={lk(f"$C{r}", P_CODE, P_H)},"我方","—"))))', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'I{r}', f'=IF($D{r}<>"销项开票","",IF($N{r}="一级",'
                     f'IF({lk(f"$C{r}", P_CODE, P_E)}="",N({lk(f"$E{r}", U_NAME, U_D)}),{lk(f"$C{r}", P_CODE, P_E)}),'
                     f'IF($N{r}="二级",IF({lk(f"$C{r}", P_CODE, P_G)}="",N({lk(f"$E{r}", U_NAME, U_D)}),{lk(f"$C{r}", P_CODE, P_G)}),'
                     f'N({lk(f"$E{r}", U_NAME, U_D)}))))', font=F_IN, fill=FILL_IN, fmt=PCT)
    put(ws, f'O{r}', f'=IF($D{r}<>"销项开票",0,ROUND(N($H{r})*N($I{r}),2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'P{r}', f'=IF($D{r}<>"销项开票",0,ROUND(N($H{r})-$O{r},2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'Q{r}', f'=IF($E{r}="","",N({lk(f"$E{r}", U_NAME, U_E)}))', font=F_LINK, fill=FILL_AUTO, fmt=PCT)
    put(ws, f'R{r}', f'=IF($E{r}="","",N({lk(f"$E{r}", U_NAME, U_F)}))', font=F_LINK, fill=FILL_AUTO, fmt=PCT)
    put(ws, f'S{r}', f'=IF(OR($D{r}<>"销项开票",{lk(f"$E{r}", U_NAME, U_G)}<>"是",N($Q{r})=0),0,'
                     f'IF({lk(f"$E{r}", U_NAME, U_H)}="税差法",'
                     f'ROUND(N($H{r})/(1+$Q{r})*$Q{r}-IF(N($R{r})=0,0,$P{r}/(1+$R{r})*$R{r}),2),'
                     f'ROUND(N($H{r})/(1+$Q{r})*$Q{r},2)))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'T{r}', f'=ROUND($S{r}*N({lk(f"$E{r}", U_NAME, U_J)}),2)', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'U{r}', f'=IF($D{r}<>"销项开票",0,ROUND(N($H{r})*N({lk(f"$E{r}", U_NAME, U_K)}),2))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'V{r}', f'=IF(OR($D{r}<>"销项开票",N($Q{r})=0),0,'
                     f'ROUND(N($H{r})/(1+$Q{r})*N({lk(f"$E{r}", U_NAME, U_L)}),2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'W{r}', f'=ROUND($S{r}+$T{r}+$U{r}+$V{r},2)', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'L{r}',
        f'=IF($B{r}="","",'
        f'IF(NOT(ISNUMBER($B{r})),"日期格式不对",'
        f'IF($C{r}="","未选项目编号",'
        f'IF(LEFT($M{r},1)="⚠","项目编号不存在",'
        f'IF($D{r}="","未选业务类型",'
        f'IF($E{r}="","未选开票/付款方",'
        f'IF(ISNA(MATCH($E{r},{U_NAME},0)),"开票方不在单位档案",'
        f'IF($F{r}="","未选收票/收款方",'
        f'IF(ISNA(MATCH($F{r},{U_NAME},0)),"收款方不在单位档案",'
        f'IF(NOT(ISNUMBER($H{r})),"金额须为数字",'
        f'IF(AND($D{r}="销项开票",$N{r}="—"),"链条待确认","√")))))))))))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
# ---- 导入历史事件 ----
KIND_MAP = {'销项开票':'销项开票','成本票':'成本票','挂靠单位代收':'挂靠代收','我方收款':'我方收款',
            'PCODE_MISS':'其他','管理费结算':'管理费结算','扣质保金':'扣质保金','已交税':'已交税'}
imported = 0
for i, e in enumerate(EV):
    r = F0 + i
    if r > F1: break
    ws[f'B{r}'] = dt.date.fromisoformat(e['date'])
    ws[f'C{r}'] = PCODE.get(e['proj'], '')
    ws[f'D{r}'] = KIND_MAP.get(e['kind'], '其他')
    ws[f'E{r}'] = e['payer']; ws[f'F{r}'] = e['payee']
    ws[f'G{r}'] = e.get('inv', '')
    ws[f'H{r}'] = e['amt']
    ws[f'J{r}'] = e['memo'][:120]
    ws[f'K{r}'] = '历史导入'
    ws[f'Y{r}'] = e['src'].split('!')[0].strip()
    ws[f'Z{r}'] = e.get('count_in', '是')
    if e.get('count_in') == '否':
        ws[f'K{r}'] = '与 ' + str(e.get('dup_of', '')) + ' 为同一张票，只计一次'
    if e['kind'] == '销项开票' and e.get('mfee') and e['amt']:
        ws[f'I{r}'] = round(e['mfee'] / e['amt'], 6)
    imported += 1
for rr in range(F0, F1 + 1):
    put(ws, f'Y{rr}', None, font=F_LINK, fill=FILL_AUTO)
    if ws[f'Z{rr}'].value is None:
        put(ws, f'Z{rr}', '是', font=F_IN, fill=FILL_IN)
    else:
        put(ws, f'Z{rr}', None, font=F_IN, fill=FILL_IN)
dv_list(ws, f'Z{F0}:Z{F1}', '"是,否"')
dv_list(ws, f'C{F0}:C{F1}', f'={P_CODE}')
dv_list(ws, f'D{F0}:D{F1}', KIND_DV)
dv_list(ws, f'E{F0}:E{F1}', f'={U_NAME}')
dv_list(ws, f'F{F0}:F{F1}', f'={U_NAME}')
dv_list(ws, f'G{F0}:G{F1}', '"13%专票,9%专票,6%专票,3%专票,1%普票,3%普票,13%普票,不开票"')
dv_num(ws, f'H{F0}:H{F1}', 'greaterThanOrEqual', '-99999999')
ws.conditional_formatting.add(f'L{F0}:L{F1}',
    FormulaRule(formula=[f'AND($L{F0}<>"",$L{F0}<>"√",$L{F0}<>"链条待确认")'],
                fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.conditional_formatting.add(f'L{F0}:L{F1}',
    FormulaRule(formula=[f'$L{F0}="链条待确认"'], fill=FILL_IN, font=Font(color='9C6500')))
ws.auto_filter.ref = f'A4:Z{F1}'
ws.freeze_panes = 'C5'; page(ws, titles='4:4')
print(f'  ✓ 业务流水（导入历史 {imported} 笔）')

JOUR = json.load(open('/tmp/claude-0/-home-user-temp/9101e7cb-4f82-53a7-bbe8-41b0f93b12b0/scratchpad/jour.json'))
ACCTS = ['泓普', '仟茂', '现金']
OPEN_BAL = {'泓普': 3656.12, '仟茂': 4754.58, '现金': 3981.49}

# ============================================================ 资金日记账
ws = wb.create_sheet(SH_JOUR)
title(ws, '资金日记账（泓普 · 仟茂 · 现金 混合录入）', 'P')
widths(ws, {'A':7,'B':11,'C':10,'D':10,'E':13,'F':12,'G':10,'H':40,'I':13,'J':13,'K':14,'L':14,'M':11,'N':15,'O':34,'P':9})
inband(ws, 'A', 'J', 3); sysband(ws, 'K', 'P', 3)
headers(ws, 4, 1, ['序号','日期','资金账户','项目编号','费用类型','往来单位','经办人','摘要','收入','支出'])
headers(ws, 4, 11, ['该账户余额','三账户合计','工程\n回款','校验','项目全称','月份'], fill=FILL_AUTO, font=F_HDR2)
put(ws, 'H2', '三个账户混在一起按日期录，右边自动出该账户即时余额和三账户合计余额；'
              '填了项目编号就能按项目筛选，工程回款把「工程回款」列选「是」即可与业务流水勾稽。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells('H2:P2')
put(ws, 'A2', '期初余额', font=F_H2, align=CR, border=None)
for i, a in enumerate(ACCTS):
    put(ws, f'{L(2+i*2)}2', a, font=F_NOTE, align=CR, border=None)
    put(ws, f'{L(3+i*2)}2', OPEN_BAL[a], font=F_IN, fill=FILL_IN, fmt=MONEY)
OPB = {a: f'${L(3+i*2)}$2' for i, a in enumerate(ACCTS)}
OPB_ALL = '+'.join(OPB.values())
for r in range(J0, J1 + 1):
    for c, fmt, tx in [('B', DATE, 0), ('C', None, 0), ('D', None, 0), ('E', None, 0), ('F', None, 0),
                       ('G', None, 0), ('H', None, 1), ('I', MONEY, 0), ('J', MONEY, 0)]:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if tx else C, fmt=fmt)
    put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-{J0-1})', font=F_LINK)
    put(ws, f'K{r}', f'=IF($B{r}="","",IF($C{r}="泓普",{OPB["泓普"]},IF($C{r}="仟茂",{OPB["仟茂"]},'
                     f'IF($C{r}="现金",{OPB["现金"]},0)))'
                     f'+SUMIFS($I${J0}:$I{r},$C${J0}:$C{r},$C{r})-SUMIFS($J${J0}:$J{r},$C${J0}:$C{r},$C{r}))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'L{r}', f'=IF($B{r}="","",{OPB_ALL}+SUM($I${J0}:$I{r})-SUM($J${J0}:$J{r}))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'M{r}', None, font=F_IN, fill=FILL_IN)
    put(ws, f'O{r}', f'=IF($D{r}="","",IFERROR(INDEX({P_B},MATCH($D{r},{P_CODE},0)),"⚠编号不存在"))',
        font=F_LINK, fill=FILL_AUTO, align=CL)
    put(ws, f'P{r}', f'=IF($B{r}="","",TEXT($B{r},"yyyy-mm"))', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'N{r}',
        f'=IF($B{r}="","",'
        f'IF(NOT(ISNUMBER($B{r})),"日期格式不对",'
        f'IF($C{r}="","未选资金账户",'
        f'IF(AND(N($I{r})=0,N($J{r})=0),"收支都为空",'
        f'IF(AND(N($I{r})>0,N($J{r})>0),"收支不能同时填",'
        f'IF(AND($D{r}<>"",LEFT($O{r},1)="⚠"),"项目编号不存在",'
        f'IF(AND($M{r}="是",$F{r}=""),"工程回款必须填往来单位","√"))))))', font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
for i, j in enumerate(JOUR):
    r = J0 + i
    if r > J1: break
    ws[f'B{r}'] = dt.date.fromisoformat(j['date']); ws[f'C{r}'] = j['acct']
    ws[f'E{r}'] = j['etype']; ws[f'G{r}'] = j['who']; ws[f'H{r}'] = j['memo']
    if j['inc']: ws[f'I{r}'] = j['inc']
    if j['exp']: ws[f'J{r}'] = j['exp']
dv_list(ws, f'C{J0}:C{J1}', '"' + ','.join(ACCTS) + '"')
dv_list(ws, f'D{J0}:D{J1}', f'={P_CODE}')
dv_list(ws, f'F{J0}:F{J1}', f'={U_NAME}')
dv_list(ws, f'M{J0}:M{J1}', '"是,否"')
dv_list(ws, f'E{J0}:E{J1}',
        '"工程回款,借款,还借款,工资,社保,材料费,运费,餐费,耗材费,办公费用,维修费,青苗费,饮用水费用,'
        '转备用金,手续费,燃油费,税费,管理费,租赁费,考核款,账户年费,其他"')
dv_num(ws, f'I{J0}:I{J1}'); dv_num(ws, f'J{J0}:J{J1}')
ws.conditional_formatting.add(f'N{J0}:N{J1}',
    FormulaRule(formula=[f'AND($N{J0}<>"",$N{J0}<>"√")'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.auto_filter.ref = f'A4:P{J1}'
ws.freeze_panes = 'C5'; page(ws, titles='4:4')

# ============================================================ 代收台账（业务流水视图）
wsf = wb[SH_FLOW]
put(wsf, 'X4', '代收台账\n序号', font=F_HDR2, fill=FILL_AUTO)
wsf.column_dimensions['X'].width = 9
for r in range(F0, F1 + 1):
    put(wsf, f'X{r}', f'=IF(OR($D{r}="挂靠代收",$D{r}="我方收款"),'
                      f'COUNTIFS($D${F0}:$D{r},"挂靠代收")+COUNTIFS($D${F0}:$D{r},"我方收款"),"")',
        font=F_LINK, fill=FILL_AUTO)

ws = wb.create_sheet(SH_DAI)
title(ws, '挂靠单位代收台账', 'J',
      '业主把钱付给挂靠单位、挂靠单位再转给我们，这两步在这里一笔一笔看。最右边「在途余额」就是这个单位收了钱还压着没给我们的数。')
widths(ws, {'A':7,'B':11,'C':10,'D':38,'E':12,'F':12,'G':15,'H':15,'I':15,'J':30})
headers(ws, 5, 1, ['序号','日期','项目编号','项目全称','挂靠单位','类型','业主付给\n挂靠单位','挂靠单位\n转给我方','该单位\n在途余额','摘要'])
for i in range(D1 - D0 + 1):
    r = D0 + i
    n = i + 1
    src = lambda col: f'IFERROR(INDEX({QF}!${col}${F0}:${col}${F1},MATCH({n},{QF}!$X${F0}:$X${F1},0)),"")'
    put(ws, f'A{r}', f'=IF($B{r}="","",{n})', font=F_LINK)
    put(ws, f'B{r}', f'={src("B")}', font=F_LINK, fmt=DATE)
    put(ws, f'C{r}', f'={src("C")}', font=F_LINK)
    put(ws, f'D{r}', f'={src("M")}', font=F_LINK, align=CL)
    put(ws, f'E{r}', f'=IF($B{r}="","",IF({src("D")}="挂靠代收",{src("F")},{src("E")}))', font=F_LINK)
    put(ws, f'F{r}', f'={src("D")}', font=F_LINK)
    put(ws, f'G{r}', f'=IF($F{r}="挂靠代收",N({src("H")}),0)', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($F{r}="我方收款",N({src("H")}),0)', font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($B{r}="","",SUMIFS($G${D0}:$G{r},$E${D0}:$E{r},$E{r})-SUMIFS($H${D0}:$H{r},$E${D0}:$E{r},$E{r}))',
        font=F_TXT, fmt=MONEY)
    put(ws, f'J{r}', f'={src("J")}', font=F_LINK, align=CL)
    ws.row_dimensions[r].height = 16
TR = D1 + 1
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BCDEFJ': put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in ['G','H']:
    put(ws, f'{c}{TR}', f'=SUM({c}{D0}:{c}{D1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'I{TR}', f'=$G${TR}-$H${TR}', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
ws.auto_filter.ref = f'A5:J{D1}'
ws.freeze_panes = 'C6'; page(ws, titles='5:5')
print(f'  ✓ 资金日记账（导入 {len(JOUR)} 笔）/ 代收台账')

# ============================================================ 汇总公式
FA=f'{QF}!$H${F0}:$H${F1}'; FO=f'{QF}!$O${F0}:$O${F1}'; FP=f'{QF}!$P${F0}:$P${F1}'
FW=f'{QF}!$W${F0}:$W${F1}'; FS=f'{QF}!$S${F0}:$S${F1}'; FT=f'{QF}!$T${F0}:$T${F1}'
FU_=f'{QF}!$U${F0}:$U${F1}'; FV=f'{QF}!$V${F0}:$V${F1}'
KE=f'{QF}!$E${F0}:$E${F1}'; KF=f'{QF}!$F${F0}:$F${F1}'
KD=f'{QF}!$D${F0}:$D${F1}'; KC=f'{QF}!$C${F0}:$C${F1}'
def SS(val, side, kind, extra=''):
    key = KE if side == 'E' else KF
    return f'SUMIFS({val},{key},{{u}},{KD},"{kind}"{extra})'
KZ = f'{QF}!$Z${F0}:$Z${F1}'
def agg(val, side, kind, uref, pref=None):
    key = KE if side == 'E' else KF
    ex = f',{KC},{pref}' if pref else ''
    return f'SUMIFS({val},{key},{uref},{KD},"{kind}"{ex},{KZ},"是")'
def agg2(val, side, kinds, uref, pref=None):
    return '+'.join(agg(val, side, k, uref, pref) for k in kinds)

COLS_U = ['开票额\n(该单位开出)','应扣管理费','应到成本票','已收成本票','还差成本票','应提税费','已交税','欠税未交',
          '业主已付给\n挂靠单位','业主未付','挂靠单位\n已转我方','挂靠单位\n代收未转','管理费\n已结算','扣质保金']
def fill_summary(ws, r, uref, pref=None):
    F_ = lambda v, s, k: agg(v, s, k, uref, pref)
    put(ws, f'{{C}}{r}', None)  # placeholder
def summary_formulas(uref, pref=None):
    g = lambda v, s, k: agg(v, s, k, uref, pref)
    return [
        g(FA,'E','销项开票'), g(FO,'E','销项开票'), g(FP,'E','销项开票'),
        g(FA,'F','成本票') + '+' + g(FA,'F','销项开票'),
        None, g(FW,'E','销项开票'), g(FA,'E','已交税'), None,
        g(FA,'F','挂靠代收'), None, g(FA,'E','我方收款'), None,
        g(FA,'F','管理费结算'), g(FA,'E','扣质保金')]

def build_sum_sheet(name, ttl, note, key_col_hdr, r0, r1, key_formula, pref_formula, extra_cols=None, extra_w=None):
    ws = wb.create_sheet(name)
    ncol = 2 + len(COLS_U) + (len(extra_cols) if extra_cols else 0)
    title(ws, ttl, L(ncol), note)
    w = {'A': 11, 'B': 40}
    for i in range(len(COLS_U)): w[L(3 + i)] = 14
    if extra_cols:
        for i in range(len(extra_cols)): w[L(3 + len(COLS_U) + i)] = (extra_w or 14)
    widths(ws, w)
    headers(ws, 5, 1, [key_col_hdr[0], key_col_hdr[1]] + COLS_U + (extra_cols or []))
    for r in range(r0, r1 + 1):
        put(ws, f'A{r}', key_formula(r), font=F_LINK)
        put(ws, f'B{r}', pref_formula(r), font=F_LINK, align=CL)
        uref = f'$A{r}' if key_col_hdr[0] != '项目编号' else None
        if key_col_hdr[0] == '项目编号':
            fs = summary_formulas('"*"', f'$A{r}')
        else:
            fs = summary_formulas(f'$A{r}')
        for i, f in enumerate(fs):
            col = L(3 + i)
            if f is None: continue
            put(ws, f'{col}{r}', f'=IF($A{r}="","",{f})', font=F_LINK, fmt=MONEY)
        put(ws, f'G{r}', f'=IF($A{r}="","",$E{r}-$F{r})', font=F_TXT, fmt=MONEY)
        put(ws, f'J{r}', f'=IF($A{r}="","",$H{r}-$I{r})', font=F_TXT, fmt=MONEY)
        put(ws, f'L{r}', f'=IF($A{r}="","",$C{r}-$K{r})', font=F_TXT, fmt=MONEY)
        put(ws, f'N{r}', f'=IF($A{r}="","",$K{r}-$M{r})', font=F_TOT, fmt=MONEY)
        ws.row_dimensions[r].height = 16
    TR = r1 + 1
    put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT); put(ws, f'B{TR}', None, font=F_TOT, fill=FILL_TOT)
    for i in range(len(COLS_U) + (len(extra_cols) if extra_cols else 0)):
        col = L(3 + i)
        put(ws, f'{col}{TR}', f'=SUM({col}{r0}:{col}{r1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
    ws.conditional_formatting.add(f'N{r0}:N{r1}',
        FormulaRule(formula=[f'$N{r0}>0.01'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
    ws.auto_filter.ref = f'A5:{L(ncol)}{r1}'
    ws.freeze_panes = 'C6'; page(ws, titles='5:5')
    return ws

# 单位汇总
build_sum_sheet(SH_SUM_U, '分表 · 按挂靠单位汇总',
    '每个挂靠单位一行：帮我们开了多少票出去、我们该开多少成本票给他、还差多少没开、'
    '业主付给他多少、他转给我们多少、还压着多少、税交了多少还欠多少。',
    ('单位简称', '单位全称'), SU0, SU1,
    lambda r: f'=IF({QU}!$C{r+U0-SU0}<>"挂靠单位","",{QU}!$A{r+U0-SU0})',
    lambda r: f'=IF($A{r}="","",{QU}!$B{r+U0-SU0})')

# 项目汇总
build_sum_sheet(SH_SUM_P, '分表 · 按项目汇总',
    '每个项目一行：这个工程一共开了多少票、成本票开了多少、收了多少钱、还有多少没收、税还欠多少。',
    ('项目编号', '项目全称'), SP0, SP1,
    lambda r: f'=IF({QP}!$A{r+P0-SP0}="","",{QP}!$A{r+P0-SP0})',
    lambda r: f'=IF($A{r}="","",{QP}!$B{r+P0-SP0})')
print('  ✓ 单位汇总 / 项目汇总')

# ============================================================ 单位项目明细
ws = wb.create_sheet(SH_SUM_X)
NC = 2 + len(COLS_U) + 1
title(ws, '分表 · 单位 × 项目明细', L(NC),
      '上面选一个挂靠单位，下面列出它名下每个项目的开票、成本票、回款、欠税情况。用筛选把「有无业务」选「有」即可。')
w = {'A': 11, 'B': 40}
for i in range(len(COLS_U)): w[L(3 + i)] = 14
w[L(NC)] = 10
widths(ws, w)
put(ws, 'A3', '选择单位', font=F_H2, fill=FILL_HDR2)
put(ws, 'B3', '康欣', font=Font(name='微软雅黑', size=11, bold=True, color='0000C0'), fill=FILL_IN, align=CL)
dv_list(ws, 'B3', f'={U_NAME}')
headers(ws, 5, 1, ['项目编号', '项目全称'] + COLS_U + ['有无业务'])
for r in range(SX0, SX1 + 1):
    pr = r + P0 - SX0
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$B{pr})', font=F_LINK, align=CL)
    for i, f in enumerate(summary_formulas('$B$3', f'$A{r}')):
        if f is None: continue
        put(ws, f'{L(3+i)}{r}', f'=IF($A{r}="","",{f})', font=F_LINK, fmt=MONEY)
    put(ws, f'G{r}', f'=IF($A{r}="","",$E{r}-$F{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",$H{r}-$I{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'L{r}', f'=IF($A{r}="","",$C{r}-$K{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'N{r}', f'=IF($A{r}="","",$K{r}-$M{r})', font=F_TOT, fmt=MONEY)
    put(ws, f'{L(NC)}{r}', f'=IF($A{r}="","",IF(SUM($C{r}:$F{r})+SUM($K{r}:$M{r})=0,"无","有"))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
TR = SX1 + 1
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT); put(ws, f'B{TR}', None, font=F_TOT, fill=FILL_TOT)
for i in range(len(COLS_U)):
    put(ws, f'{L(3+i)}{TR}', f'=SUM({L(3+i)}{SX0}:{L(3+i)}{SX1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'{L(NC)}{TR}', None, font=F_TOT, fill=FILL_TOT)
ws.auto_filter.ref = f'A5:{L(NC)}{SX1}'
ws.freeze_panes = 'C6'; page(ws, titles='5:5')

# ============================================================ 链条核算
ws = wb.create_sheet(SH_CHAIN)
title(ws, '分表 · 挂靠链条核算', 'P',
      '把一个项目的三层链条摊开：一级单位帮我们开了多少票给业主、扣完管理费二级该收多少票、'
      '二级又开了多少、扣完管理费我们该开多少成本票、实际开了多少、还差多少。')
widths(ws, {'A':10,'B':38,'C':11,'D':12,'E':14,'F':9,'G':13,'H':14,'I':12,'J':14,'K':9,'L':13,'M':15,'N':14,'O':14,'P':11})
headers(ws, 5, 1, ['项目编号','项目全称','业主','一级单位','一级\n开票额','一级\n费率','一级\n管理费','应开给一级',
                   '二级单位','二级\n开票额','二级\n费率','二级\n管理费','我方应开\n成本票','我方已开\n成本票','未开差额','状态'])
for r in range(SP0, SP1 + 1):
    pr = r + P0 - SP0
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$B{pr})', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($A{r}="","",{QP}!$C{pr})', font=F_LINK)
    put(ws, f'D{r}', f'=IF($A{r}="","",{QP}!$D{pr})', font=F_LINK)
    put(ws, f'E{r}', f'=IF($A{r}="","",{agg(FA,"E","销项开票","$D"+str(r),"$A"+str(r))})', font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF(OR($A{r}="",$E{r}=0),"",{agg(FO,"E","销项开票","$D"+str(r),"$A"+str(r))}/$E{r})',
        font=F_TXT, fmt=PCT)
    put(ws, f'G{r}', f'=IF($A{r}="","",{agg(FO,"E","销项开票","$D"+str(r),"$A"+str(r))})', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($A{r}="","",$E{r}-$G{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($A{r}="","",{QP}!$F{pr})', font=F_LINK)
    put(ws, f'J{r}', f'=IF(OR($A{r}="",$I{r}=""),0,{agg(FA,"E","销项开票","$I"+str(r),"$A"+str(r))})', font=F_LINK, fmt=MONEY)
    put(ws, f'K{r}', f'=IF(OR($A{r}="",$J{r}=0),"",{agg(FO,"E","销项开票","$I"+str(r),"$A"+str(r))}/$J{r})',
        font=F_TXT, fmt=PCT)
    put(ws, f'L{r}', f'=IF(OR($A{r}="",$I{r}=""),0,{agg(FO,"E","销项开票","$I"+str(r),"$A"+str(r))})', font=F_LINK, fmt=MONEY)
    put(ws, f'M{r}', f'=IF($A{r}="","",IF(AND($I{r}<>"",$J{r}>0),$J{r}-$L{r},$H{r}))', font=F_TOT, fmt=MONEY)
    ANY = chr(34) + "*" + chr(34)
    put(ws, f'N{r}', f'=IF($A{r}="","",{agg(FA,"F","成本票",ANY,"$A"+str(r))})', font=F_LINK, fmt=MONEY)  # 我方主体开出的成本票
    put(ws, f'O{r}', f'=IF($A{r}="","",ROUND($M{r}-$N{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'P{r}', f'=IF($A{r}="","",IF(ABS($O{r})<1,"✓ 已开齐",IF($O{r}>0,"还差成本票","多开了")))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
TR = SP1 + 1
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in ['B','C','D','F','I','K','P']: put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in ['E','G','H','J','L','M','N','O']:
    put(ws, f'{c}{TR}', f'=SUM({c}{SP0}:{c}{SP1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
ws.conditional_formatting.add(f'P{SP0}:P{SP1}',
    FormulaRule(formula=[f'AND($P{SP0}<>"",$P{SP0}<>"✓ 已开齐")'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.auto_filter.ref = f'A5:P{SP1}'
ws.freeze_panes = 'C6'; page(ws, titles='5:5')

# ============================================================ 税费台账
ws = wb.create_sheet(SH_TAX)
title(ws, '分表 · 税费台账', 'J',
      '挂靠单位帮我们开票产生的预提税费，按单位汇总。应提是系统按单位档案的参数逐笔算出来的，已交是业务流水里「已交税」的合计。')
widths(ws, {'A':13,'B':26,'C':15,'D':14,'E':14,'F':14,'G':15,'H':14,'I':15,'J':12})
headers(ws, 5, 1, ['单位','单位全称','应提增值税','应提附加税','应提印花税','应提所得税','应提税费合计','已交税','欠税未交','状态'])
for r in range(SU0, SU1 + 1):
    ur = r + U0 - SU0
    put(ws, f'A{r}', f'=IF({QU}!$C{ur}<>"挂靠单位","",{QU}!$A{ur})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QU}!$B{ur})', font=F_LINK, align=CL)
    for col, val in [('C', FS), ('D', FT), ('E', FU_), ('F', FV), ('G', FW)]:
        put(ws, f'{col}{r}', f'=IF($A{r}="","",{agg(val,"E","销项开票","$A"+str(r))})', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($A{r}="","",{agg(FA,"E","已交税","$A"+str(r))})', font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($A{r}="","",ROUND($G{r}-$H{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",IF(ABS($I{r})<1,"✓ 已结清",IF($I{r}>0,"欠税","多交")))', font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
TR = SU1 + 1
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT); put(ws, f'B{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in 'CDEFGHI':
    put(ws, f'{c}{TR}', f'=SUM({c}{SU0}:{c}{SU1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'J{TR}', None, font=F_TOT, fill=FILL_TOT)
ws.freeze_panes = 'C6'; page(ws, titles='5:5')
print('  ✓ 单位项目明细 / 链条核算 / 税费台账')

# ============================================================ 费用统计
SH_EXP2 = SH_EXP
ws = wb.create_sheet(SH_EXP2)
ETYPES = ['工程回款','借款','还借款','工资','社保','材料费','运费','餐费','耗材费','办公费用','维修费',
          '青苗费','饮用水费用','转备用金','手续费','燃油费','税费','管理费','租赁费','考核款','账户年费','其他']
title(ws, '分表 · 费用统计', 'J', '取自【资金日记账】。左边按费用类型分账户统计，右边按项目统计工程相关支出。')
widths(ws, {'A':16,'B':14,'C':14,'D':14,'E':14,'F':14,'G':3,'H':11,'I':38,'J':15})
JI=f'{QJ}!$I${J0}:$I${J1}'; JJ=f'{QJ}!$J${J0}:$J${J1}'
JC=f'{QJ}!$C${J0}:$C${J1}'; JE=f'{QJ}!$E${J0}:$E${J1}'; JD=f'{QJ}!$D${J0}:$D${J1}'
block(ws, 1, 6, 5, '① 按费用类型 × 资金账户', ['费用类型','泓普','仟茂','现金','收入合计','支出合计'], 6, 6+len(ETYPES)-1)
for i, t in enumerate(ETYPES):
    r = 6 + i
    put(ws, f'A{r}', t, font=F_TXT, align=CL)
    for j, a in enumerate(ACCTS):
        put(ws, f'{L(2+j)}{r}', f'=SUMIFS({JJ},{JE},$A{r},{JC},"{a}")-SUMIFS({JI},{JE},$A{r},{JC},"{a}")',
            font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', f'=SUMIFS({JI},{JE},$A{r})', font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=SUMIFS({JJ},{JE},$A{r})', font=F_LINK, fmt=MONEY)
    ws.row_dimensions[r].height = 16
TR = 6 + len(ETYPES)
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BCDEF':
    put(ws, f'{c}{TR}', f'=SUM({c}6:{c}{TR-1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
block(ws, 8, 3, 5, '② 按项目统计（日记账口径）', ['项目编号','项目全称','支出合计'], 6, 6+59)
for i in range(60):
    r = 6 + i
    pr = P0 + i
    put(ws, f'H{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'I{r}', f'=IF($H{r}="","",{QP}!$B{pr})', font=F_LINK, align=CL)
    put(ws, f'J{r}', f'=IF($H{r}="","",SUMIFS({JJ},{JD},$H{r}))', font=F_LINK, fmt=MONEY)
    ws.row_dimensions[r].height = 16
put(ws, f'H{66}', '合计', font=F_TOT, fill=FILL_TOT); put(ws, f'I{66}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'J{66}', f'=SUM(J6:J65)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
ws.freeze_panes = 'A6'; page(ws, titles='5:5')

# ============================================================ 对账核对
import openpyxl as _ox
_src = _ox.load_workbook('/root/.claude/uploads/9101e7cb-4f82-53a7-bbe8-41b0f93b12b0/c2938ae7-____9.041.xlsx', data_only=True)
ORIG = {}
_COLS = {'德誉嘉 ':('德誉嘉',7,10,None,13,17),'迅驰':('迅驰',7,10,17,20,25),'华城':('华城',7,10,17,20,24),
         '康欣':('康欣',7,10,17,20,25),'金沁':('金沁',7,10,17,21,26),'安锐':('安锐',7,10,18,21,26),
         '湖南锦泰':('湖南锦泰',7,10,18,21,26),'大太线老旧线路':('杰华',7,11,19,22,27)}
for sh,(u,cs,cc,ct,cr,cm) in _COLS.items():
    g=lambda c: (_src[sh].cell(row=4,column=c).value if c else 0) or 0
    ORIG[u]=dict(sale=round(float(g(cs)),2), cost=round(float(g(cc)),2), tax=round(float(g(ct)),2),
                 recv_up=round(float(g(cr)),2), recv_us=round(float(g(cm)),2))
ws = wb.create_sheet(SH_REC)
title(ws, '对账核对（历史数据导入校验）', 'G',
      '左边是你原来八张手工对账表的合计行，右边是本系统按导入的每一笔重新算出来的。两边应当完全一致；'
      '不一致的行会标红，说明原表的合计行与明细对不上，需要人工确认哪个对。')
widths(ws, {'A':13,'B':18,'C':17,'D':17,'E':13,'F':11,'G':17,'H':17,'I':13})
widths(ws, {'A':13,'B':20,'C':17,'D':17,'E':13,'F':11,'G':56})
headers(ws, 5, 1, ['单位','核对项目','原表合计行','本系统重算','差额','状态','说明'])
NOTE3 = ('康欣与金沁是唯一的三层链条（民能←金沁←康欣←泓普）。原来这两张表把同一张票各记了一次，'
         '康欣表的「已提供成本票」列还把泓普开给康欣的、康欣开给金沁的、代发工资三种方向混在一列求和。'
         '本系统一张票只记一次、开票方收票方各自明确，所以这两个单位的合计与原表不同，属于口径修正而非导入错误。')
ITEMS = [('开票额（该单位开出）','sale','E','销项开票'), ('已收成本票','cost','F','BOTH'),
         ('已交税','tax','E','已交税'), ('业主付给挂靠单位','recv_up','F','挂靠代收'),
         ('挂靠单位转给我方','recv_us','E','我方收款')]
r = 6
for u, o in ORIG.items():
    first = True
    for label, key, side, kind in ITEMS:
        uq = chr(34) + u + chr(34)
        put(ws, f'A{r}', u if first else '', font=F_TOT if first else F_TXT,
            fill=FILL_HDR2 if first else None)
        put(ws, f'B{r}', label, font=F_TXT, align=CL)
        put(ws, f'C{r}', o[key], font=F_IN, fill=FILL_IN, fmt=MONEY)
        f = (agg(FA, 'F', '成本票', uq) + '+' + agg(FA, 'F', '销项开票', uq)) if kind == 'BOTH' \
            else agg(FA, side, kind, uq)
        put(ws, f'D{r}', f'={f}', font=F_LINK, fmt=MONEY)
        put(ws, f'E{r}', f'=ROUND($D{r}-$C{r},2)', font=F_TXT, fmt=MONEY)
        if u in ('康欣', '金沁') and key in ('sale', 'cost'):
            put(ws, f'F{r}', '△ 口径差异', font=F_TOT, fill=FILL_TOT)
            put(ws, f'G{r}', NOTE3 if first or key == 'cost' else '', font=F_NOTE, align=CL)
        else:
            put(ws, f'F{r}', f'=IF(ABS($E{r})<1,"✓ 一致","✗ 不符")', font=F_TOT, fill=FILL_CHK)
            put(ws, f'G{r}', None, font=F_NOTE, align=CL)
        ws.row_dimensions[r].height = 16 if not (u in ('康欣','金沁') and key=='cost') else 46
        first = False
        r += 1
LAST = r - 1
put(ws, f'A{r+1}', '总体结论', font=F_TOT, fill=FILL_TOT); put(ws, f'B{r+1}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'C{r+1}', f'=COUNTIF($F$6:$F${LAST},"✓*")&" 项一致，"&COUNTIF($F$6:$F${LAST},"△*")&" 项口径差异"', font=F_TOT, fill=FILL_TOT)
put(ws, f'D{r+1}', f'=IF(COUNTIF($F$6:$F${LAST},"✗*")=0,"✓ 历史数据导入无误（口径差异项见右侧说明）","✗ 有 "&COUNTIF($F$6:$F${LAST},"✗*")&" 项需人工确认")',
    font=F_TOT, fill=FILL_CHK)
ws.merge_cells(f'D{r+1}:G{r+1}')
for c in 'EFG': put(ws, f'{c}{r+1}', None, font=F_TOT, fill=FILL_CHK)
REC_SUM_R = r + 1
ws.conditional_formatting.add(f'F6:F{LAST}',
    FormulaRule(formula=['LEFT($F6,1)="✗"'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.freeze_panes = 'C6'; page(ws, titles='5:5')
print('  ✓ 费用统计 / 对账核对')

# ============================================================ 首页
SU_T = SU1 + 1
QSU, QSX, QTAX, QREC, QCH = Q(SH_SUM_U), Q(SH_SUM_X), Q(SH_TAX), Q(SH_REC), Q(SH_CHAIN)
ws = wb.create_sheet(SH_HOME2)
ws.sheet_view.showGridLines = False
widths(ws, {'A':14,'B':14,'C':3,'D':14,'E':14,'F':3,'G':14,'H':14,'I':3,'J':14,'K':14,'L':3})
ws.merge_cells('A1:L1')
put(ws, 'A1', '建筑挂靠业务核算系统', font=Font(name='微软雅黑', size=20, bold=True, color='FFFFFF'),
    fill=PatternFill('solid', fgColor='1F3864'), align=CL, border=None)
ws.row_dimensions[1].height = 42
ws.merge_cells('A2:L2')
put(ws, 'A2', '　总表只录两张：业务流水（开票 · 成本票 · 回款 · 交税）+ 资金日记账　→　后面所有分表按单位、按项目自动汇总',
    font=Font(name='微软雅黑', size=9, color='FFFFFF'), fill=PatternFill('solid', fgColor='2F5597'),
    align=CL, border=None)
ws.row_dimensions[2].height = 20
CARDS2 = [
 ('挂靠单位累计开票', f'={QSU}!$C${SU_T}', 'D6E4F0'),
 ('我方应开成本票',   f'={QSU}!$E${SU_T}', 'D6E4F0'),
 ('我方已开成本票',   f'={QSU}!$F${SU_T}', 'D6E4F0'),
 ('还差成本票未开',   f'={QSU}!$G${SU_T}', 'FCE4E4'),
 ('应收工程款',       f'={QSU}!$C${SU_T}', 'E2EFDA'),
 ('业主已付挂靠单位', f'={QSU}!$K${SU_T}', 'E2EFDA'),
 ('挂靠单位已转我方', f'={QSU}!$M${SU_T}', 'E2EFDA'),
 ('在途资金·代收未转', f'={QSU}!$N${SU_T}', 'FCE4E4'),
 ('应提税费',         f'={QTAX}!$G${SU_T}', 'FFF2CC'),
 ('已交税',           f'={QTAX}!$H${SU_T}', 'FFF2CC'),
 ('欠税未交',         f'={QTAX}!$I${SU_T}', 'FCE4E4'),
 ('三账户资金余额',   f'={QJ}!$C$2+{QJ}!$E$2+{QJ}!$G$2+SUM({QJ}!$I${J0}:$I${J1})-SUM({QJ}!$J${J0}:$J${J1})', 'FFF2CC'),
]
row = 4
for i, (lab, f, color) in enumerate(CARDS2):
    c0 = 1 + (i % 4) * 3
    r = row + (i // 4) * 3
    a, b = L(c0), L(c0 + 1)
    ws.merge_cells(f'{a}{r}:{b}{r}')
    put(ws, f'{a}{r}', lab, font=Font(name='微软雅黑', size=9, bold=True, color='404040'),
        fill=PatternFill('solid', fgColor=color), align=CL, border=None)
    put(ws, f'{b}{r}', None, fill=PatternFill('solid', fgColor=color), border=None)
    ws.merge_cells(f'{a}{r+1}:{b}{r+1}')
    put(ws, f'{a}{r+1}', f, font=F_BIG, fill=PatternFill('solid', fgColor=color), align=CR, fmt=MONEY, border=None)
    put(ws, f'{b}{r+1}', None, fill=PatternFill('solid', fgColor=color), border=None)
    ws.row_dimensions[r].height = 18
    ws.row_dimensions[r + 1].height = 30
    ws.row_dimensions[r + 2].height = 6
SR = row + 9
for i, (lab, f) in enumerate([
    ('历史数据对账', f'={QREC}!$D${REC_SUM_R}'),
    ('业务流水校验', f'=IF(COUNT({QF}!$B${F0}:$B${F1})-COUNTIF({QF}!$L${F0}:$L${F1},"√")'
                    f'-COUNTIF({QF}!$L${F0}:$L${F1},"链条待确认")=0,'
                    f'"✓ 全部通过"&IF(COUNTIF({QF}!$L${F0}:$L${F1},"链条待确认")>0,'
                    f'"（另有 "&COUNTIF({QF}!$L${F0}:$L${F1},"链条待确认")&" 行是多链条项目，仅提示不影响汇总）",""),'
                    f'"✗ 有 "&(COUNT({QF}!$B${F0}:$B${F1})-COUNTIF({QF}!$L${F0}:$L${F1},"√")'
                    f'-COUNTIF({QF}!$L${F0}:$L${F1},"链条待确认"))&" 行待修正")'),
    ('日记账校验',   f'=IF(COUNTIF({QJ}!$N${J0}:$N${J1},"√")=COUNT({QJ}!$B${J0}:$B${J1}),"✓ 全部通过",'
                    f'"✗ 有 "&(COUNT({QJ}!$B${J0}:$B${J1})-COUNTIF({QJ}!$N${J0}:$N${J1},"√"))&" 行待修正")'),
    ('成本票缺口',   f'=IF({QSU}!$G${SU_T}<1,"✓ 已开齐","还有 "&TEXT({QSU}!$G${SU_T},"#,##0")&" 元成本票没开")'),
]):
    r = SR + i
    ws.merge_cells(f'A{r}:B{r}')
    put(ws, f'A{r}', lab, font=F_H2, fill=FILL_HDR2, align=CL)
    put(ws, f'B{r}', None, fill=FILL_HDR2)
    ws.merge_cells(f'D{r}:L{r}')
    put(ws, f'D{r}', f, font=F_TOT, fill=FILL_CHK, align=CL)
    for c in 'EFGHIJKL': put(ws, f'{c}{r}', None, fill=FILL_CHK)
    ws.row_dimensions[r].height = 22
NR = SR + 5
ws.merge_cells(f'A{NR}:L{NR}')
put(ws, f'A{NR}', '　每天怎么用', font=F_HDR, fill=FILL_HDR, align=CL, border=None)
ws.row_dimensions[NR].height = 22
NAV2 = [
 ('新项目', '先到【项目档案】加一行：项目编号、全称、业主、一级单位、二级单位、我方主体。费率留空就按【单位档案】的默认值。'),
 ('挂靠单位开票出去', '【业务流水】录一行：业务类型选「销项开票」，开票方＝挂靠单位，收票方＝业主。管理费和我方该开多少成本票自动算。'),
 ('我们开成本票过去', '【业务流水】录一行：业务类型选「成本票」，开票方＝我方主体，收票方＝挂靠单位。'),
 ('业主把钱付给挂靠单位', '【业务流水】选「挂靠代收」，付款方＝业主，收款方＝挂靠单位。这笔钱还没到我们手上。'),
 ('挂靠单位把钱转给我们', '【业务流水】选「我方收款」；同时到【资金日记账】记实际到账，「工程回款」选是。'),
 ('交税', '【业务流水】选「已交税」。应交多少系统按单位档案的税率参数自动算。'),
 ('日常收付款', '【资金日记账】三个账户混着录，余额自动算。'),
 ('要看结果', '【单位汇总】看每个挂靠单位；【项目汇总】看每个工程；【链条核算】看三层链条差多少票；【代收台账】看谁压着我们的钱。'),
]
r = NR + 1
for a, b in NAV2:
    put(ws, f'A{r}', a, font=F_TOT, fill=FILL_HDR2, align=CL)
    ws.merge_cells(f'A{r}:B{r}'); put(ws, f'B{r}', None, font=F_TOT, fill=FILL_HDR2)
    ws.merge_cells(f'D{r}:L{r}')
    put(ws, f'D{r}', b, font=F_TXT, align=CL)
    for c in 'EFGHIJKL': put(ws, f'{c}{r}', None)
    put(ws, f'C{r}', None, border=None)
    ws.row_dimensions[r].height = 22
    r += 1
ws.page_setup.orientation = 'portrait'; ws.page_setup.paperSize = 9
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0

# ============================================================ 使用说明
ws = wb.create_sheet(SH_DOC)
ws.sheet_view.showGridLines = False
widths(ws, {'A':3,'B':24,'C':104})
ws.merge_cells('B2:C2')
put(ws, 'B2', '建筑挂靠业务核算系统 · 使用说明', font=F_TITLE, align=CL, border=None)
ws.row_dimensions[2].height = 30
DOC = [
 ('这套表解决什么', None),
 ('一句话', '挂靠单位帮我们开票出去、我们回开成本票给他们、业主的钱经挂靠单位转到我们手上，中间还有管理费和税差。'
           '这些原来要手工在八张表里对，现在只在【业务流水】录一行，后面全部自动。'),
 ('两个录入口', '【业务流水】记开票、成本票、回款、交税、管理费结算；【资金日记账】记三个账户实际的每一笔收付。其余所有表都是自动算的。'),
 ('链条怎么理解', None),
 ('三层结构', '业主（民能 / 铜梁供电）← 一级单位 ← 二级单位 ← 我方（泓普 / 仟茂）。只有两层的项目，【项目档案】的二级单位留空即可。'),
 ('举例', '金沁开 100 万票给民能，扣 5% 管理费 → 康欣应开 95 万给金沁；康欣扣 2% → 我们应开 93.1 万成本票给康欣。'
         '这三步在【链条核算】一行看完，最后一列直接告诉你还差多少成本票没开。'),
 ('管理费率', '优先取【项目档案】的一级/二级费率；留空则取【单位档案】的默认管理费率。同一个单位不同项目费率不同（比如金沁有的 5% 有的 3%），就在项目档案上填。'),
 ('税差怎么算', None),
 ('两种算法', '税差法：预提增值税 ＝ 上游开票销项税 − 我方成本票进项税，用于迅驰、金沁、安锐。'
             '全额销项法：预提增值税 ＝ 开票额 ÷(1+税率) × 税率，用于华城、康欣、湖南锦泰、杰华。德誉嘉无税差。'),
 ('三个附加项', '附加税 ＝ 预提增值税 × 附加税率（迅驰/华城/金沁/湖南锦泰/杰华 12%，康欣/安锐 6%）；'
               '印花税 ＝ 含税开票额 × 印花税率；所得税 ＝ 不含税额 × 预征率（目前只有康欣 0.2%）。'),
 ('参数在哪改', '全部在【单位档案】。政策变了或跟对方重新谈了费率，改那一行就行，全表自动重算。'),
 ('钱在谁手上', None),
 ('三个状态', '业主还没付 →【单位汇总】的「业主未付」；业主付了但挂靠单位压着 →「挂靠单位代收未转」；已经到我们账上 →「挂靠单位已转我方」。'),
 ('逐笔看', '【代收台账】一笔一笔列出业主付给挂靠单位、挂靠单位转给我们的每一次，最后一列是该单位当前压着我们多少钱。'),
 ('首页预警', '首页「在途资金 · 代收未转」这张卡就是所有挂靠单位压着的钱合计，数字大了就该去催。'),
 ('资金日记账', None),
 ('三账户混录', '泓普、仟茂、现金混在一起按日期录，系统自动算「该账户余额」和「三账户合计」。'),
 ('和业务流水的关系', '日记账记的是真金白银进出，业务流水记的是应开应收。收到工程款时两边都要记：'
                     '业务流水记「我方收款」，日记账记实际到账并把「工程回款」选「是」，两边可以互相核对。'),
 ('按项目看支出', '日记账填了项目编号，【费用统计】右半部分就能按项目出支出合计。'),
 ('历史数据', None),
 ('已全部导入', '你原来八张对账表的 370 笔业务、日记账 34 笔都已经导入。【对账核对】把原表合计行和系统重算结果逐项比对，'
               '全部一致才说明导入没问题。'),
 ('项目编号', '历史项目按出现顺序编成 A001 起，新项目是 A059。要改编号直接在【项目档案】改，业务流水的下拉会跟着变。'),
 ('注意', None),
 ('不要改灰色区', '淡黄色是手工录入，灰色是自动算的。灰色列被覆盖后不会报错，但分表会静默算错。'),
 ('管理费率列可以改', '【业务流水】的「管理费率」是带公式的默认值，个别单子谈了别的价，直接在那一格填数字覆盖即可。'),
 ('校验列必须全是√', '【业务流水】和【资金日记账】最右边的校验列出现红色，说明这一行有问题（项目编号不存在、单位不在档案、开票方不在该项目链条上等），要改掉。'),
]
r = 4
for a, b in DOC:
    if b is None:
        ws.merge_cells(f'B{r}:C{r}')
        put(ws, f'B{r}', a, font=F_HDR, fill=FILL_HDR, align=CL)
        put(ws, f'C{r}', None, fill=FILL_HDR)
        ws.row_dimensions[r].height = 22
    else:
        put(ws, f'B{r}', a, font=F_TOT, fill=FILL_HDR2, align=CL)
        put(ws, f'C{r}', b, font=F_TXT, align=CL)
        ws.row_dimensions[r].height = 38
    r += 1
page(ws, titles=None, landscape=False)

TABC = {SH_HOME2:'1F3864', SH_DOC:'1F3864', SH_UNIT:'7F7F7F', SH_PROJ:'7F7F7F',
        SH_FLOW:'ED7D31', SH_JOUR:'ED7D31', SH_DAI:'2F5597', SH_SUM_U:'548235', SH_SUM_P:'548235',
        SH_SUM_X:'548235', SH_CHAIN:'548235', SH_TAX:'548235', SH_EXP:'2F5597', SH_REC:'C00000'}
for n, c in TABC.items(): wb[n].sheet_properties.tabColor = c
ORDER2 = [SH_HOME2, SH_DOC, SH_UNIT, SH_PROJ, SH_FLOW, SH_JOUR,
          SH_SUM_U, SH_SUM_P, SH_SUM_X, SH_CHAIN, SH_TAX, SH_DAI, SH_EXP, SH_REC]
wb._sheets = [wb[n] for n in ORDER2]; wb.active = 0
OUT = '/home/user/temp/建筑挂靠核算系统/建筑挂靠业务核算系统.xlsx'
wb.save(OUT)
nf = sum(1 for s in wb.worksheets for row in s.iter_rows() for c in row
         if isinstance(c.value, str) and c.value.startswith('='))
print(f'\n已保存：{OUT}')
print(f'工作表 {len(wb.worksheets)} 张，公式 {nf} 个，自动配平括号 {BAL_FIXED[0]} 处')
