# -*- coding: utf-8 -*-
"""建筑挂靠业务核算系统 —— 生成脚本（v2：多年度 · 合同台账 · 往来款 · 发票缺口 · 单位子表）

跑法：cd 生成脚本 && python3 gk_build.py
依赖：gl_common.py（样式常量）、ev2.json（历史业务事件，由 parse_gk.py + enrich.py 产出）、jour.json
"""
import sys, os, json, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gl_common import *
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule

HERE = os.path.dirname(os.path.abspath(__file__))
D2 = json.load(open(os.path.join(HERE, 'ev2.json')))
EV = [e for e in D2['events'] if e['kind'] != '预提税费']
EV.sort(key=lambda e: (e['date'], e['src']))
PASS_ROWS, DED_ROWS, PARTNER = D2['pass_rows'], D2['ded_rows'], set(D2['partner'])
WAGE_ROWS = D2.get('wage', [])
JOUR = json.load(open(os.path.join(HERE, 'jour.json')))

# ---------------- 容量（按连续多年使用留量） ----------------
U0, U1 = 6, 45          # 单位档案   40 家
P0, P1 = 6, 405         # 项目档案  400 个
F0, F1 = 5, 2504        # 业务流水 2500 行
J0, J1 = 5, 2504        # 资金日记账 2500 行
HR = 5                  # 查询表 表头行
TR = 6                  # 查询表 合计行（放表头正下方，空行再多也不用翻到底）
Q0 = 7                  # 查询表 明细首行
QU_1 = Q0 + (U1 - U0)           # 按单位的查询表末行 = 46
QP_1 = Q0 + (P1 - P0)           # 按项目的查询表末行 = 406
D0, D1 = Q0, Q0 + 799           # 代收台账 800 行

SH_HOME2, SH_DOC = '首页', '使用说明'
SH_UNIT, SH_PROJ = '单位档案', '项目档案'
SH_FLOW, SH_JOUR, SH_DAI = '业务流水', '资金日记账', '代收台账'
SH_SUM_U, SH_SUM_P, SH_SUM_X = '单位汇总', '项目汇总', '单位项目明细'
SH_CHAIN, SH_TAX, SH_EXP, SH_REC = '链条核算', '税费台账', '费用统计', '对账核对'
SH_GAP, SH_CUR = '发票缺口', '往来台账'

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

# 业务类型
KINDS = ['销项开票', '成本票', '挂靠代收', '我方收款', '管理费结算', '扣质保金', '已交税', '退税',
         '工资扣抵', '其他应收收回', '其他应付发生', '其他应付扣税', '其他应付支付', '其他']
KIND_DV = '"' + ','.join(KINDS) + '"'
ITYPES = ['劳务票', '土建票', '安装票', '机械设备票', '材料票', '其他', '不适用']
ITYPE_DV = '"' + ','.join(ITYPES) + '"'
ACCTS = ['泓普', '仟茂', '现金']
ETYPE_FIX = {'账户费年费': '账户年费', '耗材费用': '耗材费'}   # 原日记账两处笔误，落不进费用统计
OPEN_BAL = {'泓普': 3656.12, '仟茂': 4754.58, '现金': 3981.49}

def col_idx(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


# ============================================================ 筛选带（所有查询表通用）
def filter_band(ws, lastcol, unit_default=None,
                note='留空＝全部期间；填了年度就按整年取数，另填起止日期则以起止为准'):
    """在第 3 行画统一的「年度 / 起止日期」筛选带，第 4 行（隐藏）放解析后的取数区间。
       返回 (起单元格, 止单元格)，公式里直接当绝对引用用。"""
    if unit_default is not None:
        cells = [('选择单位', 'A', 'B', unit_default, None), ('年度', 'C', 'D', None, '0'),
                 ('起始日期', 'E', 'F', None, DATE), ('截止日期', 'G', 'H', None, DATE)]
        y, s_, e_ = '$D$3', '$F$3', '$H$3'
        lab_c, disp_c = 'I', 'J'
    else:
        cells = [('年度', 'A', 'B', None, '0'), ('起始日期', 'C', 'D', None, DATE),
                 ('截止日期', 'E', 'F', None, DATE)]
        y, s_, e_ = '$B$3', '$D$3', '$F$3'
        lab_c, disp_c = 'G', 'H'
    for lab, lc, col, dflt, fmt in cells:
        put(ws, f'{lc}3', lab, font=F_H2, fill=FILL_HDR2)
        put(ws, f'{col}3', dflt, font=Font(name='微软雅黑', size=10, bold=True, color='0000C0'),
            fill=FILL_IN, fmt=fmt)
    put(ws, f'{lab_c}3', '当前取数', font=F_H2, fill=FILL_HDR2)
    for ci in range(ord(disp_c) - 64, col_idx(lastcol) + 1):
        put(ws, f'{L(ci)}3', None, font=F_TOT, fill=FILL_CHK, align=CL)
    ws.merge_cells(f'{disp_c}3:{lastcol}3')
    put(ws, f'{disp_c}3',
        f'=IF(AND({y}="",{s_}="",{e_}=""),"全部期间（业务流水共 "&COUNT({QF}!$B${F0}:$B${F1})&" 笔）",'
        f'TEXT($A$4,"yyyy-mm-dd")&"  至  "&TEXT($B$4,"yyyy-mm-dd"))',
        font=F_TOT, fill=FILL_CHK, align=CL)
    put(ws, 'A4', f'=IF({s_}<>"",{s_},IF({y}<>"",DATE({y},1,1),DATE(1900,1,1)))', font=F_NOTE, fmt=DATE)
    put(ws, 'B4', f'=IF({e_}<>"",{e_},IF({y}<>"",DATE({y},12,31),DATE(2199,12,31)))', font=F_NOTE, fmt=DATE)
    put(ws, 'C4', note, font=F_NOTE, align=CL, border=None)
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].hidden = True
    return '$A$4', '$B$4'


# ============================================================ 单位档案
ws = wb.create_sheet(SH_UNIT)
title(ws, '单位档案', 'S',
      '每个挂靠单位的管理费率和税费参数都在这里配。改这里，业务流水的自动计算立刻跟着变。'
      '同一家单位有两种管理费政策的（金沁框架 5%／劳务 3%、德誉嘉扣 8% 返 4%／直接扣 4%），'
      '在「第二档」三列填第二种，项目档案上按项目选用哪一档。')
widths(ws, {'A':13,'B':26,'C':11,'D':11,'E':10,'F':18,'G':11,'H':10,'I':10,'J':11,'K':9,'L':13,
            'M':11,'N':11,'O':11,'P':11,'Q':9,'R':8,'S':30})
headers(ws, 5, 1, ['单位简称','单位全称','类型','默认\n管理费率','默认\n返现率','第二档\n名称',
                   '第二档\n管理费率','第二档\n返现率','上游\n开票税率','我方回开\n票种税率','有无\n税差',
                   '增值税\n计算方式','增值税\n预征率','附加税率\n(按增值税)','印花税率\n(按开票额)',
                   '所得税\n预征率','过账\n单位','状态','备注'])
#        A         B                      C        D     E     F                        G     H   I     J     K    L        M   N      O        P     Q    R     S
UNITS = [
 ('民能','重庆民能实业有限公司','业主','','','','','','','','否','—','','','','','否','正常','工程发包方'),
 ('铜梁供电','国网重庆铜梁供电分公司','业主','','','','','','','','否','—','','','','','否','正常','工程发包方'),
 ('德誉嘉','德誉嘉（德誉佳）','挂靠单位',0.08,0.04,'2026-7-22起直接扣4%不返现',0.04,'',0.13,0.13,'否','—','','','','','否','正常',
  '默认档＝原政策：扣 8%、其中 4% 返现金（返现进【往来台账】其他应收款）；第二档＝新政策直接扣 4%。无税差'),
 ('迅驰','迅驰','挂靠单位',0.04,'','',' ','',0.09,0.03,'是','税差法','',0.12,0.000588,'','是','正常',
  '开9%专票给民能，我方只能开3%，税差补给对方；迅驰过账的钱属于别人，记「其他应付发生」'),
 ('华城','华城','挂靠单位',0.02,'','','','',0.03,0.01,'是','全额销项法','',0.12,0.0003,'','否','正常','开3%劳务票，我方回1%普票'),
 ('康欣','康欣','挂靠单位',0.02,'','','','',0.03,0.03,'是','全额销项法','',0.06,'',0.002,'否','正常','老项目不收管理费只过票；新项目直接开给民能下浮2%'),
 ('金沁','金沁','挂靠单位',0.05,'','劳务合同 3%',0.03,'',0.09,0.03,'是','税差法','',0.12,0.0006,'','否','正常',
  '框架合同扣 5%、劳务合同扣 3%。项目档案的「合同类型」选“劳务合同”就自动走第二档'),
 ('安锐','安锐','挂靠单位',0.02,'','','','',0.13,0.13,'是','税差法','',0.06,'','','否','正常','两端同为13%，税差即管理费部分'),
 ('湖南锦泰','湖南锦泰电力建设有限公司','挂靠单位',0.10,'','','','',0.09,0.09,'是','全额销项法','',0.12,'','','否','正常',''),
 ('杰华','重庆杰华电气有限公司','挂靠单位',0.10,'','','','',0.13,0.13,'是','全额销项法','',0.12,'','','否','正常',''),
 ('泓普','泓普（我方主体）','我方主体','','','','','',0.03,'','否','—','','','','','否','正常','主要开票主体'),
 ('仟茂','仟茂（我方主体）','我方主体','','','','','',0.13,'','否','—','','','','','否','正常','13%票种主体'),
 ('税局','税务机关','其他','','','','','','','','否','—','','','','','否','正常',''),
]
UCOLS = 'ABCDEFGHIJKLMNOPQRS'
PCT_U = 'DEGHIJMNOP'
for i in range(U1 - U0 + 1):
    r = U0 + i
    for c in UCOLS:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN,
            align=CL if c in 'BFS' else C, fmt=PCT if c in PCT_U else None)
    if i < len(UNITS):
        for c, v in zip(UCOLS, UNITS[i]):
            if v != '' and v != ' ': ws[f'{c}{r}'] = v
    ws.row_dimensions[r].height = 17
dv_list(ws, f'C{U0}:C{U1}', '"业主,挂靠单位,我方主体,其他"')
dv_list(ws, f'K{U0}:K{U1}', '"是,否"')
dv_list(ws, f'L{U0}:L{U1}', '"税差法,全额销项法,预征率法,—"')
dv_list(ws, f'Q{U0}:Q{U1}', '"是,否"')
dv_list(ws, f'R{U0}:R{U1}', '"正常,停用"')
put(ws, f'A{U1+2}',
    '税差法：预提增值税 ＝ 上游开票销项税 − 我方成本票进项税（迅驰、金沁、安锐）。'
    '全额销项法：预提增值税 ＝ 开票额 ÷(1+上游税率) × 上游税率（华城、康欣、湖南锦泰、杰华）。'
    '预征率法：开票额 ÷(1+上游税率) × 预征率。'
    '附加税＝预提增值税×附加税率；印花税＝含税开票额×印花税率；所得税＝不含税额×所得税预征率。以上全部逐笔核对过你原表的数。　　'
    '返现率：挂靠单位先按管理费率全额扣，再把其中一部分以现金返还我方 —— 返的这部分自动进【往来台账】其他应收款。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{U1+2}:S{U1+2}')
ws.freeze_panes = 'B6'; page(ws, titles='5:5')

# ============================================================ 项目档案（含合同管理台账）
ws = wb.create_sheet(SH_PROJ)
title(ws, '项目档案 · 合同管理台账', 'X',
      '一个项目一行，既是挂靠链条档案，也是合同台账：左边链条与费率，中间合同与审计金额，右边自动算开票进度。'
      '链条最多三层：业主 ← 一级单位 ← 二级单位 ← 我方，只有两层的二级单位留空即可。')
widths(ws, {'A':11,'B':44,'C':11,'D':12,'E':10,'F':11,'G':12,'H':9,'I':12,'J':9,'K':11,'L':15,
            'M':11,'N':11,'O':11,'P':9,'Q':14,'R':14,'S':13,'T':14,'U':10,'V':16,'W':14,'X':22})
headers(ws, 5, 1, ['项目编号','项目全称','项目类型','合同类型','费率档','业主','一级单位','一级\n费率',
                   '二级单位','二级\n费率','我方主体','合同编号','签订日期','开工日期','竣工日期','状态',
                   '合同金额','审计(结算)\n金额','审减额','累计开票','开票\n进度','施工地址','校验','备注'])
PCOL_IN  = list('ABCDEFGHIJKLMNOPQRVX')      # 手工录入
PCOL_AUTO = list('STUW')                      # 自动算

PROJ_ROWS = []
OWN = ('民能', '铜梁供电')
seen = {}
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
    cand = [(k, v) for k, v in info['L2'].items() if k[1] == l1]
    info['lv1'] = l1
    info['lv2'] = max(cand, key=lambda x: x[1])[0][0] if cand else ''
# 合同类型/费率档不靠项目名猜，按原表里这个项目实际扣的管理费率反推：
# 命中单位档案的「第二档」费率就是劳务合同，否则按框架合同（默认档）
TIER2 = {u[0]: u[6] for u in UNITS if isinstance(u[6], float) and u[6] > 0}
TIER1 = {u[0]: u[3] for u in UNITS if isinstance(u[3], float) and u[3] > 0}
rate_hit = {}
for e in EV:
    if e['kind'] != '销项开票' or not e['proj'] or not e['amt']: continue
    rate_hit.setdefault(e['proj'], {}).setdefault(e['payer'], []).append(
        round((e.get('mfee') or 0) / e['amt'], 4))
TIER2_NAME = {u[0]: (u[5] or '') for u in UNITS}
def tier_of(proj, l1):
    """→ (合同类型, 费率档)。第二档对金沁就是劳务合同，对德誉嘉是 2026-7-22 起的新政策，
       所以合同类型只在「第二档名称」里写了劳务时才跟着变。"""
    rs = [r for r in rate_hit.get(proj, {}).get(l1, []) if r > 0]
    if not rs or l1 not in TIER2: return '框架合同', '默认', ''
    top = max(set(rs), key=rs.count)
    if abs(top - TIER2[l1]) < 1e-4 and abs(top - TIER1.get(l1, -1)) > 1e-4:
        if '劳务' in TIER2_NAME.get(l1, ''):
            return '劳务合同', '第二档', ''
        return '框架合同', '第二档', f'{l1}按第二档「{TIER2_NAME.get(l1,"")}」计费，已按原表实际费率 {top:.0%} 定档'
    return '框架合同', '默认', ''
hist = list(seen.items())
ordered = hist[:58] + [(None, None)] + hist[58:]   # 新项目插在第 59 位 → 编号 A059
PCODE = {}
for i, (p, info) in enumerate(ordered):
    code = f'A{i+1:03d}'
    if p is None:
        NEW_CODE = code
        PROJ_ROWS.append(dict(code=code, name='（新项目：请填项目全称）', ptype='自营项目', ctype='框架合同',
                              tier='默认', tiermemo='', owner='民能', l1='', l2='', mine='泓普', memo='本次新增项目'))
        continue
    PCODE[p] = code
    multi = len(info['L1']) > 1
    _ct, _tier, _tm = tier_of(p, info['lv1'])
    PROJ_ROWS.append(dict(
        code=code, name=p,
        ptype='合伙项目' if p in PARTNER else '自营项目',
        ctype=_ct, tier=_tier, tiermemo=_tm,
        owner=info['owner'] or '民能', l1=info['lv1'], l2=info['lv2'], mine='泓普',
        memo=('该项目走了 ' + str(len(info['L1'])) + ' 条挂靠链：' + '、'.join(info['L1'].keys()) +
              '，档案只能填一条，链条核算以此为准；其余链条请看单位项目明细') if multi else ''))
PROJ_BY_CODE = {d['code']: d for d in PROJ_ROWS}

FA_ = f'{QF}!$I${F0}:$I${F1}'          # 金额
for i in range(P1 - P0 + 1):
    r = P0 + i
    for c in PCOL_IN:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if c in 'BVX' else C,
            fmt=PCT if c in 'HJ' else (DATE if c in 'MNO' else (MONEY if c in 'QR' else None)))
    if i < len(PROJ_ROWS):
        d = PROJ_ROWS[i]
        ws[f'A{r}'] = d['code']; ws[f'B{r}'] = d['name']; ws[f'C{r}'] = d['ptype']
        ws[f'D{r}'] = d['ctype']; ws[f'F{r}'] = d['owner']; ws[f'K{r}'] = d['mine']
        if d['l1']: ws[f'G{r}'] = d['l1']
        if d['l2']: ws[f'I{r}'] = d['l2']
        ws[f'P{r}'] = '在建'
        _mm = '；'.join(x for x in (d.get('tiermemo'), d['memo']) if x)
        if _mm: ws[f'X{r}'] = _mm
    put(ws, f'E{r}', f'=IF($A{r}="","",IF($D{r}="劳务合同","第二档","默认"))', font=F_IN, fill=FILL_IN)
    if i < len(PROJ_ROWS) and PROJ_ROWS[i]['tier'] != ('第二档' if PROJ_ROWS[i]['ctype'] == '劳务合同' else '默认'):
        ws[f'E{r}'] = PROJ_ROWS[i]['tier']
    put(ws, f'S{r}', f'=IF($A{r}="","",IF(N($Q{r})=0,"",ROUND(N($Q{r})-N($R{r}),2)))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'T{r}', f'=IF($A{r}="","",SUMIFS({FA_},{QF}!$D${F0}:$D${F1},"销项开票",'
                     f'{QF}!$C${F0}:$C${F1},$A{r},{QF}!$AH${F0}:$AH${F1},"是"))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'U{r}', f'=IF($A{r}="","",IF(N($R{r})>0,$T{r}/$R{r},IF(N($Q{r})>0,$T{r}/$Q{r},"")))',
        font=F_LINK, fill=FILL_AUTO, fmt=PCT)
    put(ws, f'W{r}',
        f'=IF($A{r}="","",IF(COUNTIF($A${P0}:$A${P1},$A{r})>1,"编号重复",'
        f'IF($B{r}="","没填项目全称",'
        f'IF(AND(N($Q{r})>0,N($R{r})>0,N($R{r})>N($Q{r})),"审计金额大于合同金额",'
        f'IF(OR($F{r}="",$G{r}=""),"待完善（业主/一级单位）","√")))))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 17
for c in 'FGIK': dv_list(ws, f'{c}{P0}:{c}{P1}', f'={U_NAME}')
dv_list(ws, f'C{P0}:C{P1}', '"自营项目,合伙项目"')
dv_list(ws, f'D{P0}:D{P1}', '"框架合同,劳务合同,施工合同,租赁合同,其他"')
dv_list(ws, f'E{P0}:E{P1}', '"默认,第二档"')
dv_list(ws, f'P{P0}:P{P1}', '"在建,已完工,已结清,暂停"')
dv_num(ws, f'Q{P0}:Q{P1}'); dv_num(ws, f'R{P0}:R{P1}')
ws.conditional_formatting.add(f'W{P0}:W{P1}',
    FormulaRule(formula=[f'AND($W{P0}<>"",$W{P0}<>"√",LEFT($W{P0},3)<>"待完善")'],
                fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.conditional_formatting.add(f'W{P0}:W{P1}',
    FormulaRule(formula=[f'LEFT($W{P0},3)="待完善"'], fill=FILL_IN, font=Font(color='9C6500')))
ws.auto_filter.ref = f'A5:X{P1}'
ws.freeze_panes = 'C6'; page(ws, titles='5:5')
print(f'  ✓ 单位档案 / 项目档案·合同台账（{len(PROJ_ROWS)} 个项目，'
      f'{sum(1 for d in PROJ_ROWS if d["ptype"]=="合伙项目")} 个合伙项目，新项目 {NEW_CODE}）')

# ============================================================ 业务流水（总表）
ws = wb.create_sheet(SH_FLOW)
title(ws, '业务流水（总表）', 'AH',
      '所有开票、成本票、回款、交税、管理费结算、往来款都在这里录一行。右边灰色区自动算管理费、应开成本票、'
      '该笔要预提的各项税费，以及其他应收（返现）与其他应付（过账 / 合伙）。后面所有分表按项目、按单位、按期间自动汇总。')
widths(ws, {'A':7,'B':11,'C':10,'D':13,'E':12,'F':12,'G':11,'H':12,'I':14,'J':9,'K':9,'L':34,'M':16,'N':15,
            'O':38,'P':8,'Q':9,'R':12,'S':14,'T':8,'U':8,'V':13,'W':12,'X':12,'Y':12,'Z':13,
            'AA':13,'AB':13,'AC':7,'AD':6,'AE':9,'AF':6,'AG':13,'AH':10})
inband(ws, 'A', 'N', 3); sysband(ws, 'O', 'AH', 3)
headers(ws, 4, 1, ['序号','日期','项目编号','业务类型','开票/\n付款方','收票/\n收款方','发票性质','票据类型',
                   '金额','管理费率','返现率','摘要','备注','校验'])
headers(ws, 4, 15, ['项目全称','层级','成本归类','管理费','应开成本票','上游\n税率','我方\n税率',
                    '预提增值税','预提附加税','预提印花税','预提所得税','预提税费合计',
                    '其他应收\n(返现)','其他应付\n(过账/合伙)','年度','代收\n标记','代收台账\n序号',
                    '末层','来源表'], fill=FILL_AUTO, font=F_HDR2)
put(ws, 'AH4', '计入汇总', font=F_HDR2, fill=FILL_HDR2)

U_D  = f'{QU}!$D${U0}:$D${U1}'; U_E  = f'{QU}!$E${U0}:$E${U1}'
U_G2 = f'{QU}!$G${U0}:$G${U1}'; U_H2 = f'{QU}!$H${U0}:$H${U1}'
U_I  = f'{QU}!$I${U0}:$I${U1}'; U_J  = f'{QU}!$J${U0}:$J${U1}'
U_K  = f'{QU}!$K${U0}:$K${U1}'; U_L  = f'{QU}!$L${U0}:$L${U1}'
U_N  = f'{QU}!$N${U0}:$N${U1}'; U_O  = f'{QU}!$O${U0}:$O${U1}'; U_P = f'{QU}!$P${U0}:$P${U1}'
P_B  = f'{QP}!$B${P0}:$B${P1}'; P_C  = f'{QP}!$C${P0}:$C${P1}'; P_E2 = f'{QP}!$E${P0}:$E${P1}'
P_G  = f'{QP}!$G${P0}:$G${P1}'; P_H  = f'{QP}!$H${P0}:$H${P1}'
P_I  = f'{QP}!$I${P0}:$I${P1}'; P_J  = f'{QP}!$J${P0}:$J${P1}'; P_K = f'{QP}!$K${P0}:$K${P1}'
KE = f'{QF}!$E${F0}:$E${F1}'; KF = f'{QF}!$F${F0}:$F${F1}'
KD = f'{QF}!$D${F0}:$D${F1}'; KC = f'{QF}!$C${F0}:$C${F1}'; KQ = f'{QF}!$Q${F0}:$Q${F1}'
KZ = f'{QF}!$AH${F0}:$AH${F1}'
FAH = f'{QF}!$AF${F0}:$AF${F1}'      # 末层标记

def lk(val, key, out, nf='""'): return f'IFERROR(INDEX({out},MATCH({val},{key},0)),{nf})'

for r in range(F0, F1 + 1):
    for c, fmt, tx in [('B', DATE, 0), ('C', None, 0), ('D', None, 0), ('E', None, 0), ('F', None, 0),
                       ('G', None, 0), ('H', None, 0), ('I', MONEY, 0), ('L', None, 1), ('M', None, 1)]:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if tx else C, fmt=fmt)
    put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-{F0-1})', font=F_LINK)
    put(ws, f'O{r}', f'=IF($C{r}="","",{lk(f"$C{r}", P_CODE, P_B, chr(34)+"⚠项目编号不存在"+chr(34))})',
        font=F_LINK, fill=FILL_AUTO, align=CL)
    put(ws, f'P{r}', f'=IF($E{r}="","",IF($E{r}={lk(f"$C{r}", P_CODE, P_G)},"一级",'
                     f'IF($E{r}={lk(f"$C{r}", P_CODE, P_I)},"二级",'
                     f'IF($E{r}={lk(f"$C{r}", P_CODE, P_K)},"我方","—"))))', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'Q{r}', f'=IF($H{r}="","",IF(OR($H{r}="劳务票",$H{r}="土建票",$H{r}="安装票"),"劳务类",'
                     f'IF($H{r}="机械设备票","机械类",IF($H{r}="材料票","材料类","其他"))))',
        font=F_LINK, fill=FILL_AUTO)
    # 默认档 / 第二档：项目档案的「费率档」说了算，第二档费率留空时退回默认
    second = f'AND({lk(f"$C{r}", P_CODE, P_E2)}="第二档",{lk(f"$E{r}", U_NAME, U_G2)}<>"")'
    base_rate = f'IF({second},N({lk(f"$E{r}", U_NAME, U_G2)}),N({lk(f"$E{r}", U_NAME, U_D)}))'
    base_reb  = f'IF({second},N({lk(f"$E{r}", U_NAME, U_H2)}),N({lk(f"$E{r}", U_NAME, U_E)}))'
    put(ws, f'J{r}', f'=IF($D{r}<>"销项开票","",IF($P{r}="一级",'
                     f'IF({lk(f"$C{r}", P_CODE, P_H)}="",{base_rate},{lk(f"$C{r}", P_CODE, P_H)}),'
                     f'IF($P{r}="二级",IF({lk(f"$C{r}", P_CODE, P_J)}="",{base_rate},{lk(f"$C{r}", P_CODE, P_J)}),'
                     f'{base_rate})))', font=F_IN, fill=FILL_IN, fmt=PCT)
    put(ws, f'K{r}', f'=IF($D{r}<>"销项开票","",{base_reb})', font=F_IN, fill=FILL_IN, fmt=PCT)
    put(ws, f'R{r}', f'=IF($D{r}<>"销项开票",0,ROUND(N($I{r})*N($J{r}),2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'S{r}', f'=IF($D{r}<>"销项开票",0,ROUND(N($I{r})-$R{r},2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'T{r}', f'=IF($E{r}="","",N({lk(f"$E{r}", U_NAME, U_I)}))', font=F_LINK, fill=FILL_AUTO, fmt=PCT)
    put(ws, f'U{r}', f'=IF($E{r}="","",N({lk(f"$E{r}", U_NAME, U_J)}))', font=F_LINK, fill=FILL_AUTO, fmt=PCT)
    put(ws, f'V{r}', f'=IF(OR($D{r}<>"销项开票",{lk(f"$E{r}", U_NAME, U_K)}<>"是",N($T{r})=0),0,'
                     f'IF({lk(f"$E{r}", U_NAME, U_L)}="税差法",'
                     f'ROUND(N($I{r})/(1+$T{r})*$T{r}-IF(N($U{r})=0,0,$S{r}/(1+$U{r})*$U{r}),2),'
                     f'ROUND(N($I{r})/(1+$T{r})*$T{r},2)))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'W{r}', f'=ROUND($V{r}*N({lk(f"$E{r}", U_NAME, U_N)}),2)', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'X{r}', f'=IF($D{r}<>"销项开票",0,ROUND(N($I{r})*N({lk(f"$E{r}", U_NAME, U_O)}),2))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'Y{r}', f'=IF(OR($D{r}<>"销项开票",N($T{r})=0),0,'
                     f'ROUND(N($I{r})/(1+$T{r})*N({lk(f"$E{r}", U_NAME, U_P)}),2))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'Z{r}', f'=ROUND($V{r}+$W{r}+$X{r}+$Y{r},2)', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'AA{r}', f'=IF($D{r}="销项开票",ROUND(N($I{r})*N($K{r}),2),'
                      f'IF($D{r}="其他应收收回",-N($I{r}),0))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'AB{r}', f'=IF($D{r}="其他应付发生",N($I{r}),'
                      f'IF(OR($D{r}="其他应付扣税",$D{r}="其他应付支付"),-N($I{r}),0))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'AC{r}', f'=IF($B{r}="","",IF(ISNUMBER($B{r}),YEAR($B{r}),"日期非法"))', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'AD{r}', f'=IF(AND(OR($D{r}="挂靠代收",$D{r}="我方收款"),ISNUMBER($B{r}),'
                      f'$B{r}>={QD}!$A$4,$B{r}<={QD}!$B$4),1,0)', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'AE{r}', f'=IF($AD{r}=0,"",SUM($AD${F0}:$AD{r}))', font=F_LINK, fill=FILL_AUTO)
    # 末层＝这一行的开票方，在同一项目里没有再从别人手上收到销项票 → 该由我方直接回成本票给他
    put(ws, f'AF{r}', f'=IF($D{r}<>"销项开票",0,'
                      f'IF(COUNTIFS({KC},$C{r},{KD},"销项开票",{KF},$E{r},{KZ},"是")=0,1,0))',
        font=F_LINK, fill=FILL_AUTO)
    put(ws, f'AG{r}', None, font=F_LINK, fill=FILL_AUTO)
    put(ws, f'N{r}',
        f'=IF($B{r}="","",'
        f'IF(NOT(ISNUMBER($B{r})),"日期格式不对",'
        f'IF($D{r}="","未选业务类型",'
        f'IF(AND($C{r}="",OR($D{r}="销项开票",$D{r}="成本票",$D{r}="工资扣抵",'
        f'$D{r}="挂靠代收",$D{r}="我方收款")),"未选项目编号",'
        f'IF(LEFT($O{r},1)="⚠","项目编号不存在",'
        f'IF($E{r}="","未选开票/付款方",'
        f'IF(ISNA(MATCH($E{r},{U_NAME},0)),"开票方不在单位档案",'
        f'IF($F{r}="","未选收票/收款方",'
        f'IF(ISNA(MATCH($F{r},{U_NAME},0)),"收款方不在单位档案",'
        f'IF(NOT(ISNUMBER($I{r})),"金额须为数字",'
        f'IF(AND(OR($D{r}="销项开票",$D{r}="成本票"),$H{r}=""),"未选票据类型",'
        f'IF(AND($D{r}="销项开票",$P{r}="—"),"链条待确认","√"))))))))))))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16

# ---- 导入历史事件 ----
KIND_MAP = {'销项开票':'销项开票','成本票':'成本票','挂靠单位代收':'挂靠代收','我方收款':'我方收款',
            '管理费结算':'管理费结算','扣质保金':'扣质保金','已交税':'已交税'}
MFEE_RATE = {}
for _e in EV:
    if _e['kind'] == '销项开票' and _e['amt']:
        _k = (_e['payer'], round(_e['amt'], 2))
        MFEE_RATE[_k] = max(MFEE_RATE.get(_k, 0.0), round((_e.get('mfee') or 0) / _e['amt'], 6))
imported = 0
for i, e in enumerate(EV):
    r = F0 + i
    if r > F1: raise SystemExit('业务流水容量不够')
    ws[f'B{r}'] = dt.date.fromisoformat(e['date'])
    ws[f'C{r}'] = PCODE.get(e['proj'], '')
    ws[f'D{r}'] = KIND_MAP.get(e['kind'], '其他')
    ws[f'E{r}'] = e['payer']; ws[f'F{r}'] = e['payee']
    ws[f'G{r}'] = e.get('inv', '')
    if e.get('itype'): ws[f'H{r}'] = e['itype']
    ws[f'I{r}'] = e['amt']
    ws[f'L{r}'] = e['memo'][:120]
    ws[f'M{r}'] = '历史导入'
    ws[f'AG{r}'] = e['src'].split('!')[0].strip()
    ws[f'AH{r}'] = e.get('count_in', '是')
    if e.get('count_in') == '否':
        ws[f'M{r}'] = '与 ' + str(e.get('dup_of', '')) + ' 为同一张票，只计一次'
    if e['kind'] == '销项开票' and e['amt']:
        # 原表同一张票在两个单位的表里各出现一次，管理费往往只记在其中一张上；
        # 按「开票方＋金额」在全部行里找那个非零费率，找不到就照原表写 0（康欣老项目只过票不收费）
        ws[f'J{r}'] = MFEE_RATE.get((e['payer'], round(e['amt'], 2)), 0.0)
    # 德誉嘉：原表逐笔标了哪几张票返 4%，按原表写死，其余单位留公式默认
    if e['kind'] == '销项开票' and e['unit'] == '德誉嘉' and e['amt']:
        ws[f'K{r}'] = round(e.get('rebate', 0) / e['amt'], 6) if e.get('rebate') else 0
    imported += 1

# ---- 追加：迅驰过账应付 + 应扣税费（原表「过账项目支付情况」AA/AB 两列） ----
NAME2CODE = {d['name']: d['code'] for d in PROJ_ROWS}
r = F0 + imported
extra = 0
for x in PASS_ROWS:
    ws[f'B{r}'] = dt.date(2026, 3, 27)
    ws[f'C{r}'] = NAME2CODE.get(x['proj'], '')
    ws[f'D{r}'] = '其他应付发生'
    ws[f'E{r}'] = '迅驰'; ws[f'F{r}'] = '迅驰'
    ws[f'H{r}'] = '不适用'
    ws[f'I{r}'] = x['amt']
    ws[f'L{r}'] = '迅驰过账项目应付工程款：' + x['memo'][:60]
    ws[f'M{r}'] = '历史导入·原表迅驰「过账应付工程款」列；实际收款对象请按实补填'
    ws[f'AG{r}'] = '迅驰'; ws[f'AH{r}'] = '是'
    r += 1; extra += 1
for x in WAGE_ROWS:
    ws[f'B{r}'] = dt.date(2026, 6, 30)
    ws[f'C{r}'] = NAME2CODE.get(x['proj'], '')
    ws[f'D{r}'] = '工资扣抵'
    ws[f'E{r}'] = '泓普'; ws[f'F{r}'] = '康欣'
    ws[f'H{r}'] = '劳务票'
    ws[f'I{r}'] = x['amt']
    ws[f'L{r}'] = '工资表顶抵的劳务成本，不再另开成本票'
    ws[f'M{r}'] = '历史导入·原总台账「劳务成本·工资扣抵」列'
    ws[f'AG{r}'] = '总台账'; ws[f'AH{r}'] = '是'
    r += 1; extra += 1
for x in DED_ROWS:
    ws[f'B{r}'] = dt.date(2026, 3, 27)
    ws[f'C{r}'] = ''
    ws[f'D{r}'] = '其他应付扣税'
    ws[f'E{r}'] = '泓普'; ws[f'F{r}'] = '迅驰'
    ws[f'H{r}'] = '不适用'
    ws[f'I{r}'] = x['amt']
    ws[f'L{r}'] = '过账应付里要扣除的税费：' + x['memo'][:60]
    ws[f'M{r}'] = '历史导入·原表迅驰「已付款」列'
    ws[f'AG{r}'] = '迅驰'; ws[f'AH{r}'] = '是'
    r += 1; extra += 1
for rr in range(F0, F1 + 1):
    if ws[f'AH{rr}'].value is None:
        put(ws, f'AH{rr}', '是', font=F_IN, fill=FILL_IN)
    else:
        put(ws, f'AH{rr}', None, font=F_IN, fill=FILL_IN)
dv_list(ws, f'AH{F0}:AH{F1}', '"是,否"')
dv_list(ws, f'C{F0}:C{F1}', f'={P_CODE}')
dv_list(ws, f'D{F0}:D{F1}', KIND_DV)
dv_list(ws, f'E{F0}:E{F1}', f'={U_NAME}')
dv_list(ws, f'F{F0}:F{F1}', f'={U_NAME}')
dv_list(ws, f'G{F0}:G{F1}', '"13%专票,9%专票,6%专票,3%专票,1%普票,3%普票,13%普票,不开票"')
dv_list(ws, f'H{F0}:H{F1}', ITYPE_DV)
dv_num(ws, f'I{F0}:I{F1}', 'greaterThanOrEqual', '-99999999')
ws.conditional_formatting.add(f'N{F0}:N{F1}',
    FormulaRule(formula=[f'AND($N{F0}<>"",$N{F0}<>"√",$N{F0}<>"链条待确认")'],
                fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.conditional_formatting.add(f'N{F0}:N{F1}',
    FormulaRule(formula=[f'$N{F0}="链条待确认"'], fill=FILL_IN, font=Font(color='9C6500')))
ws.auto_filter.ref = f'A4:AH{F1}'
ws.freeze_panes = 'C5'; page(ws, titles='4:4')
FLOW_USED = imported + extra
print(f'  ✓ 业务流水（历史 {imported} 笔 + 过账/扣税 {extra} 笔 = {FLOW_USED} 笔，容量 {F1-F0+1} 行）')

# ============================================================ 资金日记账
ws = wb.create_sheet(SH_JOUR)
title(ws, '资金日记账（泓普 · 仟茂 · 现金 混合录入）', 'Q')
widths(ws, {'A':7,'B':11,'C':10,'D':10,'E':13,'F':12,'G':10,'H':40,'I':13,'J':13,'K':14,'L':14,'M':11,
            'N':15,'O':34,'P':9,'Q':7})
inband(ws, 'A', 'J', 3); sysband(ws, 'K', 'Q', 3)
headers(ws, 4, 1, ['序号','日期','资金账户','项目编号','费用类型','往来单位','经办人','摘要','收入','支出'])
headers(ws, 4, 11, ['该账户余额','三账户合计','工程\n回款','校验','项目全称','年月','年度'],
        fill=FILL_AUTO, font=F_HDR2)
put(ws, 'H2', '三个账户混在一起按日期录，右边自动出该账户即时余额和三账户合计余额；'
              '填了项目编号就能按项目筛选，工程回款把「工程回款」列选「是」即可与业务流水勾稽。'
              '跨年度继续往下录即可，【费用统计】按年度或起止日期取数。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells('H2:Q2')
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
    put(ws, f'P{r}', f'=IF($B{r}="","",IF(ISNUMBER($B{r}),TEXT($B{r},"yyyy-mm"),"日期非法"))',
        font=F_LINK, fill=FILL_AUTO)
    put(ws, f'Q{r}', f'=IF($B{r}="","",IF(ISNUMBER($B{r}),YEAR($B{r}),"日期非法"))', font=F_LINK, fill=FILL_AUTO)
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
    ws[f'B{r}'] = dt.date.fromisoformat(j['date']); ws[f'C{r}'] = j['acct']
    ws[f'E{r}'] = ETYPE_FIX.get(j['etype'], j['etype'])
    ws[f'G{r}'] = j['who']; ws[f'H{r}'] = j['memo']
    if j['inc']: ws[f'I{r}'] = j['inc']
    if j['exp']: ws[f'J{r}'] = j['exp']
ETYPES = ['工程回款','货款','借款','还借款','工资','社保','材料费','运费','餐费','耗材费','办公费用','维修费',
          '青苗费','饮用水费用','转备用金','手续费','燃油费','税费','管理费','租赁费','考核款','账户年费',
          '其他应收收回','其他应付支付','其他']
dv_list(ws, f'C{J0}:C{J1}', '"' + ','.join(ACCTS) + '"')
dv_list(ws, f'D{J0}:D{J1}', f'={P_CODE}')
dv_list(ws, f'F{J0}:F{J1}', f'={U_NAME}')
dv_list(ws, f'M{J0}:M{J1}', '"是,否"')
dv_list(ws, f'E{J0}:E{J1}', '"' + ','.join(ETYPES) + '"')
dv_num(ws, f'I{J0}:I{J1}'); dv_num(ws, f'J{J0}:J{J1}')
ws.conditional_formatting.add(f'N{J0}:N{J1}',
    FormulaRule(formula=[f'AND($N{J0}<>"",$N{J0}<>"√")'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.auto_filter.ref = f'A4:Q{J1}'
ws.freeze_panes = 'C5'; page(ws, titles='4:4')

# ============================================================ 取数公式（带起止日期）
FB = f'{QF}!$B${F0}:$B${F1}'          # 日期
FI = f'{QF}!$I${F0}:$I${F1}'          # 金额
FR = f'{QF}!$R${F0}:$R${F1}'          # 管理费
FS = f'{QF}!$S${F0}:$S${F1}'          # 应开成本票
FV = f'{QF}!$V${F0}:$V${F1}'; FW = f'{QF}!$W${F0}:$W${F1}'
FX = f'{QF}!$X${F0}:$X${F1}'; FY = f'{QF}!$Y${F0}:$Y${F1}'; FZ = f'{QF}!$Z${F0}:$Z${F1}'
FAA = f'{QF}!$AA${F0}:$AA${F1}'; FAB = f'{QF}!$AB${F0}:$AB${F1}'
def agg(val, side, kind, uref, pref=None, dr=None, cls=None, last=False):
    """side: 'E' 开票/付款方  'F' 收票/收款方  None 不限；dr=(起,止) 加日期区间；
       last=True 只取链条末层的销项票（我方该直接回成本票的那一层）"""
    ex = ''
    if side: ex += f',{KE if side == "E" else KF},{uref}'
    if pref: ex += f',{KC},{pref}'
    if cls:  ex += f',{KQ},"{cls}"'
    if last: ex += f',{FAH},1'
    if dr:   ex += f',{FB},">="&{dr[0]},{FB},"<="&{dr[1]}'
    return f'SUMIFS({val},{KD},"{kind}"{ex},{KZ},"是")'

COLS_U = ['开票额\n(该单位开出)','应扣管理费','应到成本票','已收成本票','还差成本票','应提税费','已交税','欠税未交',
          '业主已付给\n挂靠单位','业主未付','挂靠单位\n已转我方','挂靠单位\n代收未转','管理费\n已结算','扣质保金']
def summary_formulas(uref, pref=None, dr=None):
    g = lambda v, s, k: agg(v, s, k, uref, pref, dr)
    return [g(FI,'E','销项开票'), g(FR,'E','销项开票'), g(FS,'E','销项开票'),
            g(FI,'F','成本票') + '+' + g(FI,'F','销项开票'),
            None, g(FZ,'E','销项开票'), g(FI,'E','已交税'), None,
            g(FI,'F','挂靠代收'), None, g(FI,'E','我方收款'), None,
            g(FI,'F','管理费结算'), g(FI,'E','扣质保金')]

def totals_row(ws, cols, r0, r1, row=TR, first_lab='合  计'):
    put(ws, f'A{row}', first_lab, font=F_TOT, fill=FILL_TOT)
    put(ws, f'B{row}', None, font=F_TOT, fill=FILL_TOT)
    for c in cols:
        put(ws, f'{c}{row}', f'=SUM({c}{r0}:{c}{r1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
    ws.row_dimensions[row].height = 20

# 单位档案加一列隐藏的「挂靠单位序号」，让按单位的查询表能连续排下来
wsu = wb[SH_UNIT]
put(wsu, f'T{HR}', '挂靠序号', font=F_HDR2, fill=FILL_AUTO)
for r in range(U0, U1 + 1):
    put(wsu, f'T{r}', f'=IF($C{r}<>"挂靠单位","",COUNTIF($C${U0}:$C{r},"挂靠单位"))',
        font=F_NOTE, fill=FILL_AUTO)
wsu.column_dimensions['T'].hidden = True
U_RANK = f'{QU}!$T${U0}:$T${U1}'

HIDE_NOTE = '空白行已隐藏。新增项目 / 换了查询区间后，点一下 数据 → 筛选 → 重新应用，行数就会跟着变'

SPARE = 15          # 折叠时在数据后面留几行空位，新增项目 / 单位不至于一填就看不见

def hide_tail(ws, r0, r1, keep_rows, filter_col=None, ref=None, spare=SPARE, vals=('有',)):
    """把没有数据的行折叠起来（查询表专用；录入表一律全部显示）"""
    keep_rows = set(keep_rows)
    if keep_rows:
        nxt = max(keep_rows) + 1
        keep_rows |= set(range(nxt, min(nxt + spare, r1 + 1)))
    for r in range(r0, r1 + 1):
        if r not in keep_rows:
            ws.row_dimensions[r].hidden = True
    if ref:
        ws.auto_filter.ref = ref
        if filter_col is not None:
            ws.auto_filter.add_filter_column(filter_col, list(vals), blank=False)

def detail_cols(ws, r, uref, pref, dr, last_col_letter, biz_col):
    """一行 14 列的通用汇总公式 + 有无业务"""
    for i, f in enumerate(summary_formulas(uref, pref, dr)):
        if f is None: continue
        put(ws, f'{L(3+i)}{r}', f'=IF($A{r}="","",{f})', font=F_LINK, fmt=MONEY)
    put(ws, f'G{r}', f'=IF($A{r}="","",$E{r}-$F{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",$H{r}-$I{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'L{r}', f'=IF($A{r}="","",$C{r}-$K{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'N{r}', f'=IF($A{r}="","",$K{r}-$M{r})', font=F_TOT, fmt=MONEY)
    put(ws, f'{biz_col}{r}', f'=IF($A{r}="","",IF(ROUND(SUM($C{r}:$F{r})+SUM($K{r}:$M{r}),2)=0,"无","有"))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16

SUM_W = {'A': 11, 'B': 40}
for i in range(len(COLS_U)): SUM_W[L(3 + i)] = 14
SUM_W['Q'] = 10
BIZ = 'Q'

# ============================================================ 单位汇总
ws = wb.create_sheet(SH_SUM_U)
title(ws, '分表 · 按挂靠单位汇总', 'Q',
      '每个挂靠单位一行：帮我们开了多少票出去、我们该开多少成本票给他、还差多少没开、'
      '业主付给他多少、他转给我们多少、还压着多少、税交了多少还欠多少。上面可按年度或起止日期取数。'
      '注意「还差成本票」是按单位口径逐层各算各的；全公司实际还要开多少票，以【发票缺口】为准。')
widths(ws, SUM_W)
DR = filter_band(ws, 'Q')
headers(ws, HR, 1, ['单位简称', '单位全称'] + COLS_U + ['有无业务'])
for i in range(U1 - U0 + 1):
    r = Q0 + i
    put(ws, f'A{r}', f'=IFERROR(INDEX({U_NAME},MATCH({i+1},{U_RANK},0)),"")', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",INDEX({QU}!$B${U0}:$B${U1},MATCH($A{r},{U_NAME},0)))',
        font=F_LINK, align=CL)
    detail_cols(ws, r, f'$A{r}', None, DR, 'Q', BIZ)
totals_row(ws, [L(3 + i) for i in range(len(COLS_U))], Q0, QU_1)
put(ws, f'{BIZ}{TR}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
ws.conditional_formatting.add(f'N{Q0}:N{QU_1}',
    FormulaRule(formula=[f'$N{Q0}>0.01'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
N_UNITS = sum(1 for u in UNITS if u[2] == '挂靠单位')
hide_tail(ws, Q0, QU_1, set(range(Q0, Q0 + N_UNITS)), col_idx(BIZ) - 1, f'A{HR}:Q{QU_1}')
ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')

# ============================================================ 项目汇总
ws = wb.create_sheet(SH_SUM_P)
title(ws, '分表 · 按项目汇总', 'Q',
      '每个项目一行：这个工程一共开了多少票、成本票开了多少、收了多少钱、还有多少没收、税还欠多少。'
      '上面可按年度或起止日期取数。')
widths(ws, SUM_W)
DR = filter_band(ws, 'Q')
headers(ws, HR, 1, ['项目编号', '项目全称'] + COLS_U + ['有无业务'])
for i in range(P1 - P0 + 1):
    r, pr = Q0 + i, P0 + i
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$B{pr})', font=F_LINK, align=CL)
    detail_cols(ws, r, '"*"', f'$A{r}', DR, 'Q', BIZ)
totals_row(ws, [L(3 + i) for i in range(len(COLS_U))], Q0, QP_1)
put(ws, f'{BIZ}{TR}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
ws.conditional_formatting.add(f'N{Q0}:N{QP_1}',
    FormulaRule(formula=[f'$N{Q0}>0.01'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
NP = len(PROJ_ROWS)
hide_tail(ws, Q0, QP_1, set(range(Q0, Q0 + NP)), col_idx(BIZ) - 1, f'A{HR}:Q{QP_1}')
ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')
print('  ✓ 单位汇总 / 项目汇总')

# ============================================================ 单位项目明细 + 每个单位一张子表
PAIR = {}
for e in EV:
    if e.get('count_in', '是') != '是': continue
    code = PCODE.get(e['proj'], '')
    if not code: continue
    for u in {e.get('payer'), e.get('payee')}:
        if u: PAIR.setdefault(u, set()).add(code)
HOLD_UNITS = [u[0] for u in UNITS if u[2] == '挂靠单位']

def unit_detail_sheet(name, ttl, note, unit_default, fixed_unit):
    ws = wb.create_sheet(name)
    title(ws, ttl, 'Q', note)
    widths(ws, SUM_W)
    if fixed_unit:
        put(ws, 'A3', '本表单位', font=F_H2, fill=FILL_HDR2)
        put(ws, 'B3', unit_default, font=Font(name='微软雅黑', size=11, bold=True, color='1F3864'),
            fill=FILL_HDR2, align=CL)
        for lab, lc, col, fmt in [('年度','C','D','0'), ('起始日期','E','F',DATE), ('截止日期','G','H',DATE)]:
            put(ws, f'{lc}3', lab, font=F_H2, fill=FILL_HDR2)
            put(ws, f'{col}3', None, font=Font(name='微软雅黑', size=10, bold=True, color='0000C0'),
                fill=FILL_IN, fmt=fmt)
        put(ws, 'I3', '当前取数', font=F_H2, fill=FILL_HDR2)
        for ci in range(10, 18): put(ws, f'{L(ci)}3', None, font=F_TOT, fill=FILL_CHK, align=CL)
        ws.merge_cells('J3:Q3')
        put(ws, 'J3', f'=IF(AND($D$3="",$F$3="",$H$3=""),"全部期间",'
                      f'TEXT($A$4,"yyyy-mm-dd")&"  至  "&TEXT($B$4,"yyyy-mm-dd"))',
            font=F_TOT, fill=FILL_CHK, align=CL)
        put(ws, 'A4', '=IF($F$3<>"",$F$3,IF($D$3<>"",DATE($D$3,1,1),DATE(1900,1,1)))', font=F_NOTE, fmt=DATE)
        put(ws, 'B4', '=IF($H$3<>"",$H$3,IF($D$3<>"",DATE($D$3,12,31),DATE(2199,12,31)))', font=F_NOTE, fmt=DATE)
        ws.row_dimensions[3].height = 22; ws.row_dimensions[4].hidden = True
        dr, uref = ('$A$4', '$B$4'), '$B$3'
    else:
        dr = filter_band(ws, 'Q', unit_default=unit_default)
        dv_list(ws, 'B3', f'={U_NAME}')
        uref = '$B$3'
    headers(ws, HR, 1, ['项目编号', '项目全称'] + COLS_U + ['有无业务'])
    for i in range(P1 - P0 + 1):
        r, pr = Q0 + i, P0 + i
        put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
        put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$B{pr})', font=F_LINK, align=CL)
        detail_cols(ws, r, uref, f'$A{r}', dr, 'Q', BIZ)
    totals_row(ws, [L(3 + i) for i in range(len(COLS_U))], Q0, QP_1)
    put(ws, f'{BIZ}{TR}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
    ws.conditional_formatting.add(f'N{Q0}:N{QP_1}',
        FormulaRule(formula=[f'$N{Q0}>0.01'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
    keep = set()
    if fixed_unit:
        codes = PAIR.get(unit_default, set())
        for i, d in enumerate(PROJ_ROWS):
            if d['code'] in codes: keep.add(Q0 + i)
    else:
        keep = set(range(Q0, Q0 + len(PROJ_ROWS)))
    hide_tail(ws, Q0, QP_1, keep, col_idx(BIZ) - 1, f'A{HR}:Q{QP_1}')
    ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')
    return ws

unit_detail_sheet(SH_SUM_X, '分表 · 单位 × 项目明细',
    '上面选一个挂靠单位，下面列出它名下每个项目的开票、成本票、回款、欠税情况。'
    '要给领导单独看某一家，直接翻后面那一家的专用子表（表名就是单位名）。',
    '康欣', fixed_unit=False)
SUB_SHEETS = []
for u in HOLD_UNITS:
    nm = f'{u}明细'
    unit_detail_sheet(nm, f'{u} · 项目明细（给领导看的单位专表）',
        f'只统计【{u}】名下的项目。可按年度或起止日期取数，空白行已折叠，直接打印即可。',
        u, fixed_unit=True)
    SUB_SHEETS.append(nm)
print(f'  ✓ 单位项目明细 + {len(SUB_SHEETS)} 张单位子表：{"、".join(SUB_SHEETS)}')

# ============================================================ 链条核算
ws = wb.create_sheet(SH_CHAIN)
title(ws, '分表 · 挂靠链条核算', 'R',
      '把一个项目的链条摊开：一级单位帮我们开了多少票给业主、扣完管理费二级该收多少票、二级又开了多少。'
      '「我方应开成本票」取的是这个项目所有链条最末一层的应开数 —— 同一个工程走了两三条挂靠链也不会算漏或算重。')
widths(ws, {'A':10,'B':36,'C':11,'D':12,'E':14,'F':9,'G':13,'H':14,'I':12,'J':14,'K':9,'L':13,
            'M':15,'N':14,'O':13,'P':14,'Q':11,'R':10})
DR = filter_band(ws, 'R')
headers(ws, HR, 1, ['项目编号','项目全称','业主','一级单位','一级\n开票额','一级\n费率','一级\n管理费','应开给一级',
                    '二级单位','二级\n开票额','二级\n费率','二级\n管理费','我方应开\n成本票','我方已开\n成本票',
                    '工资扣抵','未开差额','状态','有无业务'])
ANY = '"*"'
for i in range(P1 - P0 + 1):
    r, pr = Q0 + i, P0 + i
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$B{pr})', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($A{r}="","",{QP}!$F{pr})', font=F_LINK)
    put(ws, f'D{r}', f'=IF($A{r}="","",{QP}!$G{pr})', font=F_LINK)
    s1 = agg(FI, 'E', '销项开票', f'$D{r}', f'$A{r}', DR)
    m1 = agg(FR, 'E', '销项开票', f'$D{r}', f'$A{r}', DR)
    put(ws, f'E{r}', f'=IF($A{r}="","",{s1})', font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF(OR($A{r}="",$E{r}=0),"",$G{r}/$E{r})', font=F_TXT, fmt=PCT)
    put(ws, f'G{r}', f'=IF($A{r}="","",{m1})', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($A{r}="","",$E{r}-$G{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($A{r}="","",{QP}!$I{pr})', font=F_LINK)
    s2 = agg(FI, 'E', '销项开票', f'$I{r}', f'$A{r}', DR)
    m2 = agg(FR, 'E', '销项开票', f'$I{r}', f'$A{r}', DR)
    put(ws, f'J{r}', f'=IF(OR($A{r}="",$I{r}=""),0,{s2})', font=F_LINK, fmt=MONEY)
    put(ws, f'K{r}', f'=IF(OR($A{r}="",$J{r}=0),"",$L{r}/$J{r})', font=F_TXT, fmt=PCT)
    put(ws, f'L{r}', f'=IF(OR($A{r}="",$I{r}=""),0,{m2})', font=F_LINK, fmt=MONEY)
    put(ws, f'M{r}', f'=IF($A{r}="","",{agg(FS,None,"销项开票",None,f"$A{r}",DR,last=True)})',
        font=F_TOT, fmt=MONEY)
    put(ws, f'N{r}', f'=IF($A{r}="","",{agg(FI,"F","成本票",ANY,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'O{r}', f'=IF($A{r}="","",{agg(FI,None,"工资扣抵",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'P{r}', f'=IF($A{r}="","",ROUND($M{r}-$N{r}-$O{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'Q{r}', f'=IF($A{r}="","",IF(AND($E{r}=0,$J{r}=0,$N{r}=0),"无业务",'
                     f'IF(ABS($P{r})<1,"✓ 已开齐",IF($P{r}>0,"还差成本票","多开了"))))', font=F_TXT, fill=FILL_CHK)
    put(ws, f'R{r}', f'=IF($A{r}="","",IF(ROUND($E{r}+$J{r}+$N{r}+$O{r},2)=0,"无","有"))', font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
totals_row(ws, ['E','G','H','J','L','M','N','O','P'], Q0, QP_1)
for c in ['C','D','F','I','K','Q']: put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'R{TR}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
ws.conditional_formatting.add(f'Q{Q0}:Q{QP_1}',
    FormulaRule(formula=[f'AND($Q{Q0}<>"",$Q{Q0}<>"✓ 已开齐",$Q{Q0}<>"无业务")'],
                fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
hide_tail(ws, Q0, QP_1, set(range(Q0, Q0 + NP)), 17, f'A{HR}:R{QP_1}')
ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')

# ============================================================ 发票缺口（劳务票 / 机械票 差多少）
ws = wb.create_sheet(SH_GAP)
title(ws, '分表 · 发票缺口（劳务票 · 机械票）', 'P',
      '我本身开了多少票出去，对应的成本票准备了多少、还差多少 —— 劳务和机械分开算。'
      '「工资扣抵」是用工资表 / 代发工资顶掉的那部分劳务成本，从应开里扣。')
widths(ws, {'A':10,'B':36,'C':12,'D':14,'E':14,'F':13,'G':14,'H':14,'I':14,'J':14,
            'K':13,'L':13,'M':13,'N':14,'O':12,'P':10})
DR = filter_band(ws, 'P')
ws.merge_cells(f'D4:G4'); ws.merge_cells(f'H4:J4'); ws.merge_cells(f'K4:M4')
headers(ws, HR, 1, ['项目编号','项目全称','一级单位','劳务票\n应开','劳务票\n已开','劳务票\n工资扣抵',
                    '劳务票\n还差','机械票\n应开','机械票\n已开','机械票\n还差',
                    '其他票\n应开','其他票\n已开','其他票\n还差','合计\n还差','状态','有无业务'])
for i in range(P1 - P0 + 1):
    r, pr = Q0 + i, P0 + i
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$B{pr})', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($A{r}="","",{QP}!$G{pr})', font=F_LINK)
    # 应开＝该项目所有挂靠链「末层」那一张销项票的应开成本票（多链条不会算重也不会算漏）
    lab = [('D', 'E', '劳务类'), ('H', 'I', '机械类'), ('K', 'L', '其他')]
    for cdue, cdone, cls in lab:
        due = agg(FS, None, '销项开票', None, f'$A{r}', DR, cls=cls, last=True)
        done = agg(FI, None, '成本票', None, f'$A{r}', DR, cls=cls)
        put(ws, f'{cdue}{r}', f'=IF($A{r}="","",{due})', font=F_LINK, fmt=MONEY)
        put(ws, f'{cdone}{r}', f'=IF($A{r}="","",{done})', font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",{agg(FI,None,"工资扣抵",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'G{r}', f'=IF($A{r}="","",ROUND($D{r}-$E{r}-$F{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",ROUND($H{r}-$I{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'M{r}', f'=IF($A{r}="","",ROUND($K{r}-$L{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'N{r}', f'=IF($A{r}="","",ROUND($G{r}+$J{r}+$M{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'O{r}', f'=IF($A{r}="","",IF(ROUND($D{r}+$H{r}+$K{r},2)=0,"无开票",'
                     f'IF(ABS($N{r})<1,"✓ 已开齐",IF($N{r}>0,"还差票","多开了"))))', font=F_TXT, fill=FILL_CHK)
    put(ws, f'P{r}', f'=IF($A{r}="","",IF(ROUND($D{r}+$E{r}+$H{r}+$I{r}+$K{r}+$L{r},2)=0,"无","有"))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
totals_row(ws, list('DEFGHIJKLMN'), Q0, QP_1)
put(ws, f'C{TR}', None, font=F_TOT, fill=FILL_TOT); put(ws, f'O{TR}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'P{TR}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
ws.conditional_formatting.add(f'O{Q0}:O{QP_1}',
    FormulaRule(formula=[f'AND($O{Q0}<>"",$O{Q0}<>"✓ 已开齐",$O{Q0}<>"无开票")'],
                fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
hide_tail(ws, Q0, QP_1, set(range(Q0, Q0 + NP)), 15, f'A{HR}:P{QP_1}')
ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')

# ============================================================ 往来台账（其他应收 / 其他应付）
ws = wb.create_sheet(SH_CUR)
title(ws, '分表 · 往来台账（其他应收款 · 其他应付款）', 'M',
      '其他应收款＝挂靠单位按管理费率扣走后又答应返给我方的现金（德誉嘉扣 8% 返 4%）。'
      '其他应付款＝收到的过账款、合伙项目要转给别人的钱，先减掉我方已代垫的税费，剩下的才是真正要付的。')
widths(ws, {'A':12,'B':24,'C':15,'D':14,'E':14,'F':11,'G':3,'H':14,'I':14,'J':14,'K':14,'L':11,'M':10})
DR = filter_band(ws, 'M')
ws.merge_cells('C4:F4'); ws.merge_cells('H4:L4')
headers(ws, HR, 1, ['单位简称','单位全称','其他应收\n返现发生','其他应收\n已收回','其他应收\n余额','状态'])
put(ws, f'G{HR}', None, font=F_HDR, fill=FILL_HDR)
headers(ws, HR, 8, ['其他应付\n发生','其他应付\n减代扣税费','其他应付\n已支付','其他应付\n余额','状态','有无往来'],
        fill=FILL_HDR, font=F_HDR)
for i in range(U1 - U0 + 1):
    r = Q0 + i
    put(ws, f'A{r}', f'=IFERROR(INDEX({U_NAME},MATCH({i+1},{U_RANK},0)),"")', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",INDEX({QU}!$B${U0}:$B${U1},MATCH($A{r},{U_NAME},0)))',
        font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($A{r}="","",{agg(FAA,"E","销项开票",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'D{r}', f'=IF($A{r}="","",{agg(FI,"E","其他应收收回",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', f'=IF($A{r}="","",ROUND($C{r}-$D{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",IF(ROUND($C{r},2)=0,"—",IF(ABS($E{r})<1,"✓ 已收清","未收回")))',
        font=F_TXT, fill=FILL_CHK)
    put(ws, f'G{r}', None, border=None)
    put(ws, f'H{r}', f'=IF($A{r}="","",{agg(FI,"F","其他应付发生",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($A{r}="","",{agg(FI,"F","其他应付扣税",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",{agg(FI,"F","其他应付支付",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'K{r}', f'=IF($A{r}="","",ROUND($H{r}-$I{r}-$J{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'L{r}', f'=IF($A{r}="","",IF(ROUND($H{r},2)=0,"—",IF(ABS($K{r})<1,"✓ 已付清","未付清")))',
        font=F_TXT, fill=FILL_CHK)
    put(ws, f'M{r}', f'=IF($A{r}="","",IF(ROUND($C{r}+$D{r}+$H{r}+$I{r}+$J{r},2)=0,"无","有"))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
totals_row(ws, list('CDE') + list('HIJK'), Q0, QU_1)
for c in ['F','G','L']: put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'M{TR}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
ws.conditional_formatting.add(f'E{Q0}:E{QU_1}',
    FormulaRule(formula=[f'$E{Q0}>0.01'], fill=FILL_IN, font=Font(color='9C6500', bold=True)))
ws.conditional_formatting.add(f'K{Q0}:K{QU_1}',
    FormulaRule(formula=[f'$K{Q0}>0.01'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
hide_tail(ws, Q0, QU_1, set(range(Q0, Q0 + N_UNITS)), 12, f'A{HR}:M{QU_1}')

# ③ 其他应付款 · 按项目（合伙项目要转给别人的钱）
PR_H = QU_1 + 3            # 表头行
PR_T = PR_H + 1            # 合计行
PR_0 = PR_H + 2
PR_1 = PR_0 + (P1 - P0)
ws.merge_cells(f'A{PR_H-1}:M{PR_H-1}')
put(ws, f'A{PR_H-1}', '　③ 其他应付款 · 按项目（合伙项目扣完代垫税费后应转给别人的钱）',
    font=F_HDR, fill=FILL_HDR, align=CL)
headers(ws, PR_H, 1, ['项目编号','项目全称','项目类型','其他应付\n发生','减代扣\n税费','已支付',
                      '应付余额','状态','该项目\n我方已交税','','','','有无往来'])
for i in range(P1 - P0 + 1):
    r, pr = PR_0 + i, P0 + i
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$B{pr})', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($A{r}="","",{QP}!$C{pr})', font=F_LINK)
    put(ws, f'D{r}', f'=IF($A{r}="","",{agg(FI,None,"其他应付发生",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', f'=IF($A{r}="","",{agg(FI,None,"其他应付扣税",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",{agg(FI,None,"其他应付支付",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'G{r}', f'=IF($A{r}="","",ROUND($D{r}-$E{r}-$F{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($A{r}="","",IF(ROUND($D{r},2)=0,"—",IF(ABS($G{r})<1,"✓ 已付清","未付清")))',
        font=F_TXT, fill=FILL_CHK)
    put(ws, f'I{r}', f'=IF($A{r}="","",{agg(FI,None,"已交税",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    for c in 'JKL': put(ws, f'{c}{r}', None, font=F_TXT)
    put(ws, f'M{r}', f'=IF($A{r}="","",IF(ROUND($D{r}+$E{r}+$F{r},2)=0,"无","有"))', font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
put(ws, f'A{PR_T}', '合  计', font=F_TOT, fill=FILL_TOT); put(ws, f'B{PR_T}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'C{PR_T}', None, font=F_TOT, fill=FILL_TOT)
for c in list('DEFG') + ['I']:
    put(ws, f'{c}{PR_T}', f'=SUM({c}{PR_0}:{c}{PR_1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
for c in ['H','J','K','L']: put(ws, f'{c}{PR_T}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'M{PR_T}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
ws.conditional_formatting.add(f'G{PR_0}:G{PR_1}',
    FormulaRule(formula=[f'$G{PR_0}>0.01'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
PAY_CODES = {NAME2CODE.get(x['proj'], '') for x in PASS_ROWS}
keep3 = {PR_0 + i for i, d in enumerate(PROJ_ROWS) if d['code'] in PAY_CODES or d['ptype'] == '合伙项目'}
keep3 |= set(range(max(keep3) + 1, min(max(keep3) + 1 + SPARE, PR_1 + 1))) if keep3 else set()
for r in range(PR_0, PR_1 + 1):
    if r not in keep3: ws.row_dimensions[r].hidden = True
ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')
print('  ✓ 链条核算 / 发票缺口 / 往来台账')

# ============================================================ 税费台账
ws = wb.create_sheet(SH_TAX)
title(ws, '分表 · 税费台账', 'K',
      '挂靠单位帮我们开票产生的预提税费，按单位汇总。应提是系统按单位档案的参数逐笔算出来的，'
      '已交是业务流水里「已交税」的合计。上面可按年度或起止日期取数。')
widths(ws, {'A':13,'B':26,'C':15,'D':14,'E':14,'F':14,'G':15,'H':14,'I':15,'J':12,'K':10})
DR = filter_band(ws, 'K')
headers(ws, HR, 1, ['单位','单位全称','应提增值税','应提附加税','应提印花税','应提所得税','应提税费合计',
                    '已交税','欠税未交','状态','有无业务'])
for i in range(U1 - U0 + 1):
    r = Q0 + i
    put(ws, f'A{r}', f'=IFERROR(INDEX({U_NAME},MATCH({i+1},{U_RANK},0)),"")', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",INDEX({QU}!$B${U0}:$B${U1},MATCH($A{r},{U_NAME},0)))',
        font=F_LINK, align=CL)
    for col, val in [('C', FV), ('D', FW), ('E', FX), ('F', FY), ('G', FZ)]:
        put(ws, f'{col}{r}', f'=IF($A{r}="","",{agg(val,"E","销项开票",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($A{r}="","",{agg(FI,"E","已交税",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($A{r}="","",ROUND($G{r}-$H{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",IF(AND($G{r}=0,$H{r}=0),"—",'
                     f'IF(ABS($I{r})<1,"✓ 已结清",IF($I{r}>0,"欠税","多交"))))', font=F_TXT, fill=FILL_CHK)
    put(ws, f'K{r}', f'=IF($A{r}="","",IF(ROUND($G{r}+$H{r},2)=0,"无","有"))', font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
totals_row(ws, list('CDEFGHI'), Q0, QU_1)
put(ws, f'J{TR}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'K{TR}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
hide_tail(ws, Q0, QU_1, set(range(Q0, Q0 + N_UNITS)), 10, f'A{HR}:K{QU_1}')
ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')

# ============================================================ 代收台账
ws = wb.create_sheet(SH_DAI)
title(ws, '挂靠单位代收台账', 'J',
      '业主把钱付给挂靠单位、挂靠单位再转给我们，这两步在这里一笔一笔看。'
      '最右边「在途余额」就是这个单位收了钱还压着没给我们的数。上面可按年度或起止日期取数。')
widths(ws, {'A':7,'B':11,'C':10,'D':36,'E':12,'F':12,'G':15,'H':15,'I':15,'J':30})
DR = filter_band(ws, 'J')
headers(ws, HR, 1, ['序号','日期','项目编号','项目全称','挂靠单位','类型','业主付给\n挂靠单位',
                    '挂靠单位\n转给我方','该单位\n在途余额','摘要'])
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BCDEFJ': put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in ['G', 'H']:
    put(ws, f'{c}{TR}', f'=SUM({c}{D0}:{c}{D1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'I{TR}', f'=$G${TR}-$H${TR}', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
ws.row_dimensions[TR].height = 20
FAE = f'{QF}!$AE${F0}:$AE${F1}'
for i in range(D1 - D0 + 1):
    r, n = D0 + i, i + 1
    src = lambda col: f'IFERROR(INDEX({QF}!${col}${F0}:${col}${F1},MATCH({n},{FAE},0)),"")'
    put(ws, f'A{r}', f'=IF($B{r}="","",{n})', font=F_LINK)
    put(ws, f'B{r}', f'={src("B")}', font=F_LINK, fmt=DATE)
    put(ws, f'C{r}', f'={src("C")}', font=F_LINK)
    put(ws, f'D{r}', f'={src("O")}', font=F_LINK, align=CL)
    put(ws, f'E{r}', f'=IF($B{r}="","",IF({src("D")}="挂靠代收",{src("F")},{src("E")}))', font=F_LINK)
    put(ws, f'F{r}', f'={src("D")}', font=F_LINK)
    put(ws, f'G{r}', f'=IF($F{r}="挂靠代收",N({src("I")}),0)', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($F{r}="我方收款",N({src("I")}),0)', font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($B{r}="","",SUMIFS($G${D0}:$G{r},$E${D0}:$E{r},$E{r})'
                     f'-SUMIFS($H${D0}:$H{r},$E${D0}:$E{r},$E{r}))', font=F_TXT, fmt=MONEY)
    put(ws, f'J{r}', f'={src("L")}', font=F_LINK, align=CL)
    ws.row_dimensions[r].height = 16
N_DAI = sum(1 for e in EV if e['kind'] in ('挂靠单位代收', '我方收款') and e.get('count_in', '是') == '是')
hide_tail(ws, D0, D1, set(range(D0, D0 + N_DAI)), 5, f'A{HR}:J{D1}',
          spare=30, vals=('挂靠代收', '我方收款'))
ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')

# ============================================================ 费用统计
ws = wb.create_sheet(SH_EXP)
title(ws, '分表 · 费用统计', 'J',
      '取自【资金日记账】。左边按费用类型分账户统计，右边按项目统计工程相关支出。上面可按年度或起止日期取数。')
widths(ws, {'A':16,'B':14,'C':14,'D':14,'E':14,'F':14,'G':3,'H':11,'I':36,'J':15})
DRJ = filter_band(ws, 'J')
JI = f'{QJ}!$I${J0}:$I${J1}'; JJ = f'{QJ}!$J${J0}:$J${J1}'; JB = f'{QJ}!$B${J0}:$B${J1}'
JC = f'{QJ}!$C${J0}:$C${J1}'; JE = f'{QJ}!$E${J0}:$E${J1}'; JD = f'{QJ}!$D${J0}:$D${J1}'
DJ = f',{JB},">="&{DRJ[0]},{JB},"<="&{DRJ[1]}'
headers(ws, HR, 1, ['费用类型', '泓普', '仟茂', '现金', '收入合计', '支出合计'])
put(ws, f'G{HR}', None, font=F_HDR, fill=FILL_HDR)
headers(ws, HR, 8, ['项目编号', '项目全称', '支出合计'])
E_0 = Q0
E_1 = Q0 + len(ETYPES) - 1
for i, t in enumerate(ETYPES):
    r = E_0 + i
    put(ws, f'A{r}', t, font=F_TXT, align=CL)
    for j, a in enumerate(ACCTS):
        put(ws, f'{L(2+j)}{r}', f'=SUMIFS({JJ},{JE},$A{r},{JC},"{a}"{DJ})'
                                f'-SUMIFS({JI},{JE},$A{r},{JC},"{a}"{DJ})', font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', f'=SUMIFS({JI},{JE},$A{r}{DJ})', font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=SUMIFS({JJ},{JE},$A{r}{DJ})', font=F_LINK, fmt=MONEY)
    put(ws, f'G{r}', None, border=None)
    ws.row_dimensions[r].height = 16
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BCDEF':
    put(ws, f'{c}{TR}', f'=SUM({c}{E_0}:{c}{E_1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'G{TR}', None, border=None)
put(ws, f'H{TR}', '合  计', font=F_TOT, fill=FILL_TOT); put(ws, f'I{TR}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'J{TR}', f'=SUM(J{Q0}:J{QP_1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
ws.row_dimensions[TR].height = 20
for i in range(P1 - P0 + 1):
    r, pr = Q0 + i, P0 + i
    put(ws, f'H{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'I{r}', f'=IF($H{r}="","",{QP}!$B{pr})', font=F_LINK, align=CL)
    put(ws, f'J{r}', f'=IF($H{r}="","",SUMIFS({JJ},{JD},$H{r}{DJ}))', font=F_LINK, fmt=MONEY)
    if r > E_1: ws.row_dimensions[r].height = 16
for r in range(max(E_1, Q0 + NP + SPARE) + 1, QP_1 + 1):
    ws.row_dimensions[r].hidden = True
put(ws, f'A{QP_1+2}', HIDE_NOTE + '　｜　左表的费用类型来自【资金日记账】E 列下拉，要加类型先在那里加。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{QP_1+2}:J{QP_1+2}')
ws.freeze_panes = f'A{Q0}'; page(ws, titles=f'{HR}:{HR}')
print('  ✓ 税费台账 / 代收台账 / 费用统计')

# ============================================================ 对账核对（全期口径，不受筛选影响）
import openpyxl as _ox
_src = _ox.load_workbook('/root/.claude/uploads/9101e7cb-4f82-53a7-bbe8-41b0f93b12b0/c2938ae7-____9.041.xlsx',
                         data_only=True)
ORIG = {}
#            单位   销售开票额 应扣管理费 已收成本票 已交税 业主付给单位 转给我方
_COLS = {'德誉嘉 ':('德誉嘉',7,8,10,None,13,17),'迅驰':('迅驰',7,8,10,17,20,25),'华城':('华城',7,8,10,17,20,24),
         '康欣':('康欣',7,8,10,17,20,25),'金沁':('金沁',7,8,10,17,21,26),'安锐':('安锐',7,8,10,18,21,26),
         '湖南锦泰':('湖南锦泰',7,8,10,18,21,26),'大太线老旧线路':('杰华',7,9,11,19,22,27)}
for sh, (u, cs, cmf, cc, ct, cr, cm) in _COLS.items():
    g = lambda c: (_src[sh].cell(row=4, column=c).value if c else 0) or 0
    ORIG[u] = dict(sale=round(float(g(cs)), 2), cost=round(float(g(cc)), 2), tax=round(float(g(ct)), 2),
                   recv_up=round(float(g(cr)), 2), recv_us=round(float(g(cm)), 2),
                   mfee=round(float(g(cmf)), 2))
ws = wb.create_sheet(SH_REC)
title(ws, '对账核对（历史数据导入校验）', 'G',
      '左边是你原来八张手工对账表的合计行，右边是本系统按导入的每一笔重新算出来的（全期口径，不受各表筛选影响）。'
      '两边应当完全一致；不一致的行会标红。')
widths(ws, {'A':13,'B':20,'C':17,'D':17,'E':13,'F':11,'G':56})
headers(ws, HR, 1, ['单位','核对项目','原表合计行','本系统重算','差额','状态','说明'])
NOTE3 = ('康欣与金沁是唯一的三层链条（民能←金沁←康欣←泓普）。原来这两张表把同一张票各记了一次，'
         '康欣表的「已提供成本票」列还把泓普开给康欣的、康欣开给金沁的、代发工资三种方向混在一列求和。'
         '本系统一张票只记一次、开票方收票方各自明确，所以这两个单位的合计与原表不同，属于口径修正而非导入错误。')
ITEMS = [('开票额（该单位开出）','sale','E','销项开票'), ('应扣管理费','mfee','E','MFEE'),
         ('已收成本票','cost','F','BOTH'),
         ('已交税','tax','E','已交税'), ('业主付给挂靠单位','recv_up','F','挂靠代收'),
         ('挂靠单位转给我方','recv_us','E','我方收款')]
r = TR
for u, o in ORIG.items():
    first = True
    for label, key, side, kind in ITEMS:
        uq = chr(34) + u + chr(34)
        put(ws, f'A{r}', u if first else '', font=F_TOT if first else F_TXT,
            fill=FILL_HDR2 if first else None)
        put(ws, f'B{r}', label, font=F_TXT, align=CL)
        put(ws, f'C{r}', o[key], font=F_IN, fill=FILL_IN, fmt=MONEY)
        if kind == 'BOTH':
            f = agg(FI, 'F', '成本票', uq) + '+' + agg(FI, 'F', '销项开票', uq)
        elif kind == 'MFEE':
            f = agg(FR, 'E', '销项开票', uq)
        else:
            f = agg(FI, side, kind, uq)
        put(ws, f'D{r}', f'={f}', font=F_LINK, fmt=MONEY)
        put(ws, f'E{r}', f'=ROUND($D{r}-$C{r},2)', font=F_TXT, fmt=MONEY)
        if u in ('康欣', '金沁') and key in ('sale', 'cost'):
            put(ws, f'F{r}', '△ 口径差异', font=F_TOT, fill=FILL_TOT)
            put(ws, f'G{r}', NOTE3 if first or key == 'cost' else '', font=F_NOTE, align=CL)
        else:
            put(ws, f'F{r}', f'=IF(ABS($E{r})<1,"✓ 一致","✗ 不符")', font=F_TOT, fill=FILL_CHK)
            put(ws, f'G{r}', None, font=F_NOTE, align=CL)
        ws.row_dimensions[r].height = 16 if not (u in ('康欣', '金沁') and key == 'cost') else 46
        first = False
        r += 1
LAST = r - 1
# 新增三项：原表新口径的合计也一起对
EXTRA_CHK = [
    ('德誉嘉', '返管理费 4%（其他应收）',
     round(sum(e.get('rebate', 0) for e in EV), 2),
     f'=SUMIFS({FAA},{KD},"销项开票",{KE},"德誉嘉",{KZ},"是")'),
    ('迅驰', '过账应付工程款',
     round(sum(x['amt'] for x in PASS_ROWS), 2),
     f'=SUMIFS({FI},{KD},"其他应付发生",{KF},"迅驰",{KZ},"是")'),
    ('迅驰', '过账应扣税费',
     round(sum(x['amt'] for x in DED_ROWS), 2),
     f'=SUMIFS({FI},{KD},"其他应付扣税",{KF},"迅驰",{KZ},"是")'),
    ('康欣', '工资扣抵（劳务成本）',
     round(sum(x['amt'] for x in WAGE_ROWS), 2),
     f'=SUMIFS({FI},{KD},"工资扣抵",{KZ},"是")'),
]
for u, label, orig, f in EXTRA_CHK:
    put(ws, f'A{r}', u, font=F_TOT, fill=FILL_HDR2)
    put(ws, f'B{r}', label, font=F_TXT, align=CL)
    put(ws, f'C{r}', orig, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'D{r}', f, font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', f'=ROUND($D{r}-$C{r},2)', font=F_TXT, fmt=MONEY)
    put(ws, f'F{r}', f'=IF(ABS($E{r})<1,"✓ 一致","✗ 不符")', font=F_TOT, fill=FILL_CHK)
    put(ws, f'G{r}', '本次新增维度，与原表对应列的合计核对', font=F_NOTE, align=CL)
    ws.row_dimensions[r].height = 16
    r += 1
LAST = r - 1
put(ws, f'A{r+1}', '总体结论', font=F_TOT, fill=FILL_TOT); put(ws, f'B{r+1}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'C{r+1}', f'=COUNTIF($F${TR}:$F${LAST},"✓*")&" 项一致，"&COUNTIF($F${TR}:$F${LAST},"△*")&" 项口径差异"',
    font=F_TOT, fill=FILL_TOT)
put(ws, f'D{r+1}', f'=IF(COUNTIF($F${TR}:$F${LAST},"✗*")=0,"✓ 历史数据导入无误（口径差异项见右侧说明）",'
                   f'"✗ 有 "&COUNTIF($F${TR}:$F${LAST},"✗*")&" 项需人工确认")', font=F_TOT, fill=FILL_CHK)
ws.merge_cells(f'D{r+1}:G{r+1}')
for c in 'EFG': put(ws, f'{c}{r+1}', None, font=F_TOT, fill=FILL_CHK)
REC_SUM_R = r + 1
ws.conditional_formatting.add(f'F{TR}:F{LAST}',
    FormulaRule(formula=[f'LEFT($F{TR},1)="✗"'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.freeze_panes = f'C{TR}'; page(ws, titles=f'{HR}:{HR}')
print('  ✓ 对账核对')

# ============================================================ 首页
QSU, QSX, QTAX, QREC, QCH = Q(SH_SUM_U), Q(SH_SUM_X), Q(SH_TAX), Q(SH_REC), Q(SH_CHAIN)
QGAP, QCUR = Q(SH_GAP), Q(SH_CUR)
ws = wb.create_sheet(SH_HOME2)
ws.sheet_view.showGridLines = False
widths(ws, {'A':14,'B':14,'C':3,'D':14,'E':14,'F':3,'G':14,'H':14,'I':3,'J':14,'K':14,'L':3})
ws.merge_cells('A1:L1')
put(ws, 'A1', '建筑挂靠业务核算系统', font=Font(name='微软雅黑', size=20, bold=True, color='FFFFFF'),
    fill=PatternFill('solid', fgColor='1F3864'), align=CL, border=None)
ws.row_dimensions[1].height = 42
ws.merge_cells('A2:L2')
put(ws, 'A2', '　只录两张：业务流水（开票 · 成本票 · 回款 · 交税 · 往来款）+ 资金日记账　→　'
              '后面所有分表按单位、按项目、按期间自动汇总，可连续多年使用',
    font=Font(name='微软雅黑', size=9, color='FFFFFF'), fill=PatternFill('solid', fgColor='2F5597'),
    align=CL, border=None)
ws.row_dimensions[2].height = 20
CARDS2 = [
 ('挂靠单位累计开票', f'={QSU}!$C${TR}', 'D6E4F0'),
 ('我方应开成本票',   f'={QSU}!$E${TR}', 'D6E4F0'),
 ('我方已开成本票',   f'={QSU}!$F${TR}', 'D6E4F0'),
 ('还差成本票未开',   f'={QGAP}!$N${TR}', 'FCE4E4'),
 ('劳务票还差',       f'={QGAP}!$G${TR}', 'FCE4E4'),
 ('机械票还差',       f'={QGAP}!$J${TR}', 'FCE4E4'),
 ('其他应收·待返现',  f'={QCUR}!$E${TR}', 'FFF2CC'),
 ('其他应付·待转付',  f'={QCUR}!$K${TR}', 'FCE4E4'),
 ('应收工程款',       f'={QSU}!$C${TR}', 'E2EFDA'),
 ('业主已付挂靠单位', f'={QSU}!$K${TR}', 'E2EFDA'),
 ('挂靠单位已转我方', f'={QSU}!$M${TR}', 'E2EFDA'),
 ('在途资金·代收未转', f'={QSU}!$N${TR}', 'FCE4E4'),
 ('应提税费',         f'={QTAX}!$G${TR}', 'FFF2CC'),
 ('已交税',           f'={QTAX}!$H${TR}', 'FFF2CC'),
 ('欠税未交',         f'={QTAX}!$I${TR}', 'FCE4E4'),
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
SR = row + 3 * ((len(CARDS2) + 3) // 4)
CHECKS = [
    ('历史数据对账', f'={QREC}!$D${REC_SUM_R}'),
    ('业务流水校验', f'=IF(COUNT({QF}!$B${F0}:$B${F1})-COUNTIF({QF}!$N${F0}:$N${F1},"√")'
                    f'-COUNTIF({QF}!$N${F0}:$N${F1},"链条待确认")=0,'
                    f'"✓ 全部通过"&IF(COUNTIF({QF}!$N${F0}:$N${F1},"链条待确认")>0,'
                    f'"（另有 "&COUNTIF({QF}!$N${F0}:$N${F1},"链条待确认")&" 行是多链条项目，仅提示不影响汇总）",""),'
                    f'"✗ 有 "&(COUNT({QF}!$B${F0}:$B${F1})-COUNTIF({QF}!$N${F0}:$N${F1},"√")'
                    f'-COUNTIF({QF}!$N${F0}:$N${F1},"链条待确认"))&" 行待修正")'),
    ('日记账校验',   f'=IF(COUNTIF({QJ}!$N${J0}:$N${J1},"√")=COUNT({QJ}!$B${J0}:$B${J1}),"✓ 全部通过",'
                    f'"✗ 有 "&(COUNT({QJ}!$B${J0}:$B${J1})-COUNTIF({QJ}!$N${J0}:$N${J1},"√"))&" 行待修正")'),
    ('项目档案校验', f'=IF(COUNTA({QP}!$A${P0}:$A${P1})-COUNTIF({QP}!$W${P0}:$W${P1},"√")'
                    f'-COUNTIF({QP}!$W${P0}:$W${P1},"待完善*")=0,'
                    f'"✓ 全部通过"&IF(COUNTIF({QP}!$W${P0}:$W${P1},"待完善*")>0,'
                    f'"（另有 "&COUNTIF({QP}!$W${P0}:$W${P1},"待完善*")&" 个项目的业主或一级单位还没填）",""),'
                    f'"✗ 有 "&(COUNTA({QP}!$A${P0}:$A${P1})-COUNTIF({QP}!$W${P0}:$W${P1},"√")'
                    f'-COUNTIF({QP}!$W${P0}:$W${P1},"待完善*"))&" 行待修正")'),
    ('成本票缺口',   f'=IF({QGAP}!$N${TR}<1,"✓ 已开齐","还有 "&TEXT({QGAP}!$N${TR},"#,##0")&" 元成本票没开"'
                    f'&"（劳务 "&TEXT({QGAP}!$G${TR},"#,##0")&"，机械 "&TEXT({QGAP}!$J${TR},"#,##0")&"）")'),
    ('往来款',       f'=IF(AND({QCUR}!$E${TR}<1,{QCUR}!$K${TR}<1),"✓ 已结清",'
                    f'"待返现 "&TEXT({QCUR}!$E${TR},"#,##0")&" 元，待转付 "&TEXT({QCUR}!$K${TR},"#,##0")&" 元")'),
]
for i, (lab, f) in enumerate(CHECKS):
    r = SR + i
    ws.merge_cells(f'A{r}:B{r}')
    put(ws, f'A{r}', lab, font=F_H2, fill=FILL_HDR2, align=CL)
    put(ws, f'B{r}', None, fill=FILL_HDR2)
    ws.merge_cells(f'D{r}:L{r}')
    put(ws, f'D{r}', f, font=F_TOT, fill=FILL_CHK, align=CL)
    for c in 'EFGHIJKL': put(ws, f'{c}{r}', None, fill=FILL_CHK)
    ws.row_dimensions[r].height = 22
NR = SR + len(CHECKS) + 1
ws.merge_cells(f'A{NR}:L{NR}')
put(ws, f'A{NR}', '　每天怎么用', font=F_HDR, fill=FILL_HDR, align=CL, border=None)
ws.row_dimensions[NR].height = 22
NAV2 = [
 ('新项目 / 新合同', '到【项目档案·合同台账】加一行：项目编号、全称、项目类型（自营/合伙）、合同类型、业主、一级/二级单位、'
                    '我方主体，再填合同金额和审计金额。费率留空就按【单位档案】走。'),
 ('挂靠单位开票出去', '【业务流水】录一行：业务类型「销项开票」，开票方＝挂靠单位，收票方＝业主，票据类型选劳务票还是机械设备票。'
                     '管理费、返现、应开成本票、各项预提税费全部自动。'),
 ('我们开成本票过去', '【业务流水】业务类型「成本票」，开票方＝我方主体，收票方＝挂靠单位，票据类型要和对应的销项票一致。'),
 ('用工资表顶劳务票', '【业务流水】业务类型「工资扣抵」，填项目编号和金额 —— 【发票缺口】里的劳务票应开就会相应减少。'),
 ('业主付钱给挂靠单位', '【业务流水】选「挂靠代收」，付款方＝业主，收款方＝挂靠单位。这笔钱还没到我们手上。'),
 ('挂靠单位转钱给我们', '【业务流水】选「我方收款」；同时到【资金日记账】记实际到账，「工程回款」选是。'),
 ('德誉嘉返 4% 现金', '开票那一行的「返现率」自动是 4%，钱一直挂在【往来台账】其他应收款上；'
                     '真收到现金时录一行「其他应收收回」冲掉。'),
 ('过账款 / 合伙分钱', '收到别人的过账款录「其他应付发生」；我方为这笔垫的税费录「其他应付扣税」；真转出去录「其他应付支付」。'
                      '【往来台账】③ 按项目算出「扣完税费后还该转给别人多少」。'),
 ('交税', '【业务流水】选「已交税」。应交多少系统按单位档案的税率参数自动算。'),
 ('日常收付款', '【资金日记账】三个账户混着录，跨年度继续往下录就行，余额自动算。'),
 ('要看结果', '【单位汇总】【项目汇总】【链条核算】【发票缺口】【往来台账】【税费台账】【费用统计】'
             '上面都有「年度 / 起止日期」，填了就只统计那一段。给领导看某一家，直接打印那家的单位专表。'),
]
r = NR + 1
for a, b in NAV2:
    put(ws, f'A{r}', a, font=F_TOT, fill=FILL_HDR2, align=CL)
    ws.merge_cells(f'A{r}:B{r}'); put(ws, f'B{r}', None, font=F_TOT, fill=FILL_HDR2)
    ws.merge_cells(f'D{r}:L{r}')
    put(ws, f'D{r}', b, font=F_TXT, align=CL)
    for c in 'EFGHIJKL': put(ws, f'{c}{r}', None)
    put(ws, f'C{r}', None, border=None)
    ws.row_dimensions[r].height = 30
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
 ('一句话', '挂靠单位帮我们开票出去、我们回开成本票给他们、业主的钱经挂靠单位转到我们手上，中间还有管理费、'
           '税差、返现和过账。原来要手工在八张表里对，现在只在【业务流水】录一行，后面全部自动。'),
 ('两个录入口', '【业务流水】记开票、成本票、回款、交税、管理费结算、往来款；【资金日记账】记三个账户实际的每一笔收付。'
               '其余所有表都是自动算的。'),
 ('连续多年怎么用', None),
 ('不用每年建新表', '业务流水和资金日记账各留了 2500 行，项目档案 400 个，按日期一直往下录就行。'
                  '每张查询表最上面都有「年度 / 起始日期 / 截止日期」：填 2026 就只看 2026 年，'
                  '填起止日期就只看那一段，三个都留空＝全部期间。'),
 ('空白行会折叠', '查询表只显示有数据的行，空白行已经隐藏起来。新增项目或换了查询区间之后，'
                 '点一下 数据 → 筛选 → 重新应用，行数就会跟着变。录入表（单位档案 / 项目档案 / 业务流水 / 资金日记账）'
                 '永远全部显示，方便往下录。'),
 ('合计在表头下面', '每张查询表的合计行就在表头正下方第 6 行，不用翻到底，筛选也不会影响它。'),
 ('管理费怎么定', None),
 ('两档费率', '【单位档案】每家可以配两档：默认管理费率 + 第二档管理费率。'
             '金沁＝框架合同 5%、劳务合同 3%；德誉嘉＝默认扣 8%（其中 4% 返现），第二档是 2026-7-22 起改成的直接扣 4%。'),
 ('按项目选档', '【项目档案】的「合同类型」选“劳务合同”，「费率档」自动跳到第二档；也可以手动改。'
               '项目档案上还能直接填一级/二级费率，填了就以它为准。'),
 ('单笔改价', '【业务流水】的「管理费率」「返现率」两列是带公式的默认值，个别单子谈了别的价，直接在那一格填数字覆盖。'),
 ('税差怎么算', None),
 ('两种算法', '税差法：预提增值税 ＝ 上游开票销项税 − 我方成本票进项税，用于迅驰、金沁、安锐。'
             '全额销项法：预提增值税 ＝ 开票额 ÷(1+税率) × 税率，用于华城、康欣、湖南锦泰、杰华。德誉嘉无税差。'),
 ('三个附加项', '附加税 ＝ 预提增值税 × 附加税率；印花税 ＝ 含税开票额 × 印花税率；'
               '所得税 ＝ 不含税额 × 预征率（目前只有康欣 0.2%）。参数全部在【单位档案】那一行，改了全表重算。'),
 ('发票缺口怎么看', None),
 ('票据类型必须选', '【业务流水】录销项开票和成本票时要选「票据类型」：劳务票 / 土建票 / 安装票 归劳务类，'
                  '机械设备票归机械类。没选会在校验列提示。'),
 ('缺口表', '【发票缺口】按项目分别算劳务票和机械票：应开多少、已开多少、工资顶掉多少、还差多少。'
           '「我开了多少票出去、劳务票差多少、机械票差多少」在这张表一眼看完。'),
 ('工资扣抵', '用工资表 / 代发工资顶掉的那部分劳务成本，在【业务流水】录「工资扣抵」，'
             '缺口表的劳务票还差就会相应减少。'),
 ('往来款怎么走', None),
 ('其他应收·返现', '挂靠单位先全额扣管理费，再把其中一部分现金返给我们（德誉嘉扣 8% 返 4%）。'
                 '开票那一行自动按「返现率」挂一笔应收；真收到钱录一行「其他应收收回」冲掉。'
                 '【往来台账】左半边就是每家还欠我们多少返现。'),
 ('其他应付·过账', '别人的钱从迅驰过账到我们账上，这笔钱不是我们的：录「其他应付发生」。'
                 '我方为这笔先垫的税费录「其他应付扣税」，真转出去录「其他应付支付」。'
                 '【往来台账】右半边＝还该转给谁多少；下半部分③按项目列出合伙项目扣完税费后要转出的钱。'),
 ('合伙项目', '【项目档案】的「项目类型」选“合伙项目”，【往来台账】③ 会单独列出来。'),
 ('合同台账', None),
 ('合同并进项目档案', '【项目档案】右半部分就是合同台账：合同编号、签订/开工/竣工日期、合同金额、审计（结算）金额，'
                    '审减额和累计开票、开票进度自动算。项目编号那一列是手工填的，不再有下拉，'
                    '重号会在「校验」列标红。'),
 ('给领导的单位专表', None),
 ('一家一张', '每个挂靠单位都有一张自己的明细子表（表名＝单位名＋明细），只统计这一家名下的项目，'
             '上面同样能按年度或起止日期取数，空白行已折叠，直接打印或导 PDF 给领导就行。'),
 ('钱在谁手上', None),
 ('三个状态', '业主还没付 →【单位汇总】的「业主未付」；业主付了但挂靠单位压着 →「挂靠单位代收未转」；'
             '已经到我们账上 →「挂靠单位已转我方」。逐笔看【代收台账】。'),
 ('注意', None),
 ('不要改灰色区', '淡黄色是手工录入，灰色是自动算的。灰色列被覆盖后不会报错，但分表会静默算错。'),
 ('校验列必须全是√', '【业务流水】【资金日记账】【项目档案】的校验列出现红色，说明这一行有问题，要改掉。'),
 ('历史数据', '原来八张对账表的业务、日记账 34 笔都已导入，德誉嘉返 4%、迅驰过账应付和应扣税费也一并进来了。'
             '【对账核对】把原表合计行和系统重算结果逐项比对。'),
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
        ws.row_dimensions[r].height = 44
    r += 1
page(ws, titles=None, landscape=False)

# ============================================================ 收尾
TABC = {SH_HOME2:'1F3864', SH_DOC:'1F3864', SH_UNIT:'7F7F7F', SH_PROJ:'7F7F7F',
        SH_FLOW:'ED7D31', SH_JOUR:'ED7D31', SH_DAI:'2F5597', SH_SUM_U:'548235', SH_SUM_P:'548235',
        SH_SUM_X:'548235', SH_CHAIN:'548235', SH_GAP:'7030A0', SH_CUR:'7030A0',
        SH_TAX:'548235', SH_EXP:'2F5597', SH_REC:'C00000'}
for n in SUB_SHEETS: TABC[n] = 'A9D08E'
for n, c in TABC.items(): wb[n].sheet_properties.tabColor = c
ORDER2 = ([SH_HOME2, SH_DOC, SH_UNIT, SH_PROJ, SH_FLOW, SH_JOUR,
           SH_SUM_U, SH_SUM_P, SH_SUM_X] + SUB_SHEETS +
          [SH_CHAIN, SH_GAP, SH_CUR, SH_TAX, SH_DAI, SH_EXP, SH_REC])
wb._sheets = [wb[n] for n in ORDER2]; wb.active = 0
OUT = os.environ.get('GK_OUT', os.path.join(HERE, '..', '建筑挂靠业务核算系统.xlsx'))
wb.save(OUT)
nf = sum(1 for s in wb.worksheets for row in s.iter_rows() for c in row
         if isinstance(c.value, str) and c.value.startswith('='))
print(f'\n已保存：{os.path.abspath(OUT)}')
print(f'工作表 {len(wb.worksheets)} 张，公式 {nf} 个，自动配平括号 {BAL_FIXED[0]} 处')
