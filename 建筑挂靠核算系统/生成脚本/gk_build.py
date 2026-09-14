# -*- coding: utf-8 -*-
"""建筑挂靠业务核算系统 —— 生成脚本（v2：多年度 · 合同台账 · 往来款 · 发票缺口 · 单位子表）

跑法：cd 生成脚本 && python3 gk_build.py
依赖：gl_common.py（样式常量）、ev2.json（历史业务事件，由 parse_gk.py + enrich.py 产出）、jour.json
"""
import sys, os, json, re, collections, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gl_common import *
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule

HERE = os.path.dirname(os.path.abspath(__file__))
D2 = json.load(open(os.path.join(HERE, 'ev2_913.json')))
EV = [e for e in D2['events'] if e['kind'] != '预提税费']
EV.sort(key=lambda e: (e['date'], e['src']))
PASS_ROWS, DED_ROWS, PARTNER = D2['pass_rows'], D2['ded_rows'], set(D2['partner'])
WAGE_ROWS = D2.get('wage', [])
PROJ_ORDER = D2['proj_order']           # 原表首现顺序 + 项目简称
RAW_ROWS = D2['raw_rows']               # 原表逐行留档，供【对账差异说明】比对
JOUR = json.load(open(os.path.join(HERE, 'jour.json')))

# ---------------- 容量（按连续多年使用留量） ----------------
# 留量按「够用 4~5 年、又不至于把表拖慢」定；不够了改这四个数重跑脚本即可
U0, U1 = 6, 45          # 单位档案   40 家
P0, P1 = 6, 305         # 项目档案  300 个（历史 63 个）
F0, F1 = 5, 2004        # 业务流水 2000 行（历史 373 行）
J0, J1 = 5, 2004        # 资金日记账 2000 行
HR = 5                  # 查询表 表头行
TR = 6                  # 查询表 合计行（放表头正下方，空行再多也不用翻到底）
Q0 = 7                  # 查询表 明细首行
QU_1 = Q0 + (U1 - U0)           # 按单位的查询表末行 = 46
QP_1 = Q0 + (P1 - P0)           # 按项目的查询表末行 = 406
D0, D1 = Q0, Q0 + 599           # 代收台账 600 行

SH_HOME2, SH_DOC = '首页', '使用说明'
SH_UNIT, SH_PROJ = '单位档案', '项目档案'
SH_FLOW, SH_JOUR, SH_DAI = '业务流水', '资金日记账', '代收台账'
SH_SUM_U, SH_SUM_P, SH_SUM_X = '单位汇总', '项目汇总', '单位项目明细'
SH_CHAIN, SH_TAX, SH_EXP = '链条核算', '税费台账', '费用统计'
SH_GAP, SH_CUR, SH_PRF = '发票缺口', '往来台账', '项目利润'
SH_DIFF = '对账差异说明'
SPLIT_ACCTS = ['泓普', '仟茂', '现金']          # 资金日记账拆分表
SH_SPLIT = {a: f'日记账-{a}' for a in SPLIT_ACCTS}

# ---------------- 业务流水列字典（改列序只动这一处，全表跟着走） ----------------
# 录入区 A~N：摘要按需求挪到「票据类型」后面（I 列），金额/管理费率/返现率顺次右移
# 税费录入区 O~R：税费按实际发生逐笔录入（原来按单位档案参数预提，跟原对账表对不上）
# 系统自动区 S 起
FC = dict(no='A', date='B', proj='C', kind='D', payer='E', payee='F', inv='G', itype='H',
          memo='I', amt='J', rate='K', reb='L', feeflag='M', note='N', chk='O',
          vat='P', add='Q', stamp='R', inc='S',
          sname='T', pname='U', tier='V', cls='W', mfee='X', due='Y',
          taxsum='Z', taxref='AA', ar='AB', ap='AC', year='AD', dflag='AE', dno='AF',
          last='AG', top='AH', src='AI', srow='AJ', cnt='AK', skey='AL', ocol='AM',
          ubel='AN')
FLOW_LAST = FC['ubel']
FEE_MODES = ['扣管理费', '不扣管理费', '不回成本票']
FEE_DV = '"' + ','.join(FEE_MODES) + '"'

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
FRNG = lambda k: f'{QF}!${FC[k]}${F0}:${FC[k]}${F1}'      # 业务流水某列的整列绝对引用

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
title(ws, '单位档案', 'N',
      '这里只放「这家单位是谁、税怎么算」。管理费率和返现率已经从这张表取消了 —— '
      '按你的要求改成在【业务流水】上逐笔手工填，后面所有表都从业务流水取数，'
      '再也不会出现「档案填一个率、原表实际是另一个率」对不上的情况。'
      '下面的税率参数只用来给【业务流水】的「税费参考试算」列提个醒，不再决定入账数：'
      '税费同样按实际发生逐笔录入。')
widths(ws, {'A':13,'B':26,'C':11,'D':10,'E':11,'F':9,'G':13,'H':11,'I':11,'J':11,'K':11,
            'L':9,'M':8,'N':34})
headers(ws, 5, 1, ['单位简称','单位全称','类型','上游\n开票税率','我方回开\n票种税率','有无\n税差',
                   '增值税\n计算方式','增值税\n预征率','附加税率\n(按增值税)','印花税率\n(按开票额)',
                   '所得税\n预征率','过账\n单位','状态','备注'])
#        A         B                      C        D     E     F    G        H   I      J        K     L    M     N
UNITS = [
 ('民能','重庆民能实业有限公司','业主','','','否','—','','','','','否','正常','工程发包方'),
 ('铜梁供电','国网重庆铜梁供电分公司','业主','','','否','—','','','','','否','正常','工程发包方'),
 ('德誉嘉','德誉嘉（德誉佳）','挂靠单位',0.13,0.13,'否','—','','','','','否','正常',
  '原政策扣 8%、其中 4% 返现金（返现自动进【往来台账】其他应收款）；2026-7-22 起改成直接扣 4% 不返现。'
  '两种政策都在【业务流水】逐笔填费率，不用再选档'),
 ('迅驰','迅驰','挂靠单位',0.09,0.03,'是','税差法','',0.12,0.000588,'','是','正常',
  '开9%专票给民能，我方只能开3%，税差补给对方；迅驰过账的钱属于别人，记「其他应付发生」'),
 ('华城','华城','挂靠单位',0.03,0.01,'是','全额销项法','',0.12,0.0003,'','否','正常',
  '劳务开3%、设备票开13%，一家两个票种，所以税费一律按原表逐笔录，不按这里的率算'),
 ('金沁','金沁','挂靠单位',0.09,0.03,'是','税差法','',0.12,0.0006,'','否','正常',
  '框架合同扣 5%、劳务合同扣 3%，哪一笔用哪个率直接在【业务流水】填'),
 ('湖南锦泰','湖南锦泰电力建设有限公司','挂靠单位',0.09,0.09,'是','全额销项法','',0.12,'','','否','正常',''),
 ('康欣','康欣','挂靠单位',0.03,0.03,'是','全额销项法','',0.06,'',0.002,'否','正常',
  '老项目不收管理费只过票（费率填 0）；新项目直接开给民能下浮2%'),
 ('安锐','安锐','挂靠单位',0.13,0.13,'是','税差法','',0.06,'','','否','正常','两端同为13%，税差即管理费部分'),
 ('杰华','重庆杰华电气有限公司','挂靠单位',0.13,0.13,'是','全额销项法','',0.12,'','','否','正常',''),
 ('泓普','泓普（我方主体）','我方主体',0.03,'','否','—','','','','','否','正常','主要开票主体'),
 ('仟茂','仟茂（我方主体）','我方主体',0.13,'','否','—','','','','','否','正常','13%票种主体'),
 ('税局','税务机关','其他','','','否','—','','','','','否','正常',''),
]
UCOLS = 'ABCDEFGHIJKLMN'
PCT_U = 'DEHIJK'
for i in range(U1 - U0 + 1):
    r = U0 + i
    for c in UCOLS:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN,
            align=CL if c in 'BN' else C, fmt=PCT if c in PCT_U else None)
    if i < len(UNITS):
        for c, v in zip(UCOLS, UNITS[i]):
            if v != '' and v != ' ': ws[f'{c}{r}'] = v
    ws.row_dimensions[r].height = 17
dv_list(ws, f'C{U0}:C{U1}', '"业主,挂靠单位,我方主体,其他"')
dv_list(ws, f'F{U0}:F{U1}', '"是,否"')
dv_list(ws, f'G{U0}:G{U1}', '"税差法,全额销项法,预征率法,—"')
dv_list(ws, f'L{U0}:L{U1}', '"是,否"')
dv_list(ws, f'M{U0}:M{U1}', '"正常,停用"')
put(ws, f'A{U1+2}',
    '管理费率 / 返现率已取消：请到【业务流水】的「管理费率」「返现率」两列逐笔填（淡黄色格子）。'
    '历史 3xx 行已按你那份《对账明细》逐笔写死，一分没动。　　'
    '这里的税率参数只驱动【业务流水】最右边那一列「税费参考试算」，供录入时对照，不参与任何汇总：'
    '税差法＝上游销项税−我方成本票进项税；全额销项法＝开票额÷(1+上游税率)×上游税率；'
    '预征率法＝开票额÷(1+上游税率)×预征率。真正入账的税费在【业务流水】O~R 四列按实际逐笔录。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{U1+2}:N{U1+2}')
ws.freeze_panes = 'B6'; page(ws, titles='5:5')

# ============================================================ 项目档案（含合同管理台账）
ws = wb.create_sheet(SH_PROJ)
title(ws, '项目档案 · 合同管理台账', 'V',
      '一个项目一行，既是挂靠链条档案，也是合同台账。行的先后＝这个项目在你那份《对账明细》里第一次出现的先后，'
      '对原表可以一行对一行往下走。C 列「项目简称」是从原表摘要「…：平双线」里取出来的，'
      '后面所有表（业务流水、各种汇总、单位明细）显示的都是简称，不再刷屏显示几十个字的全称。'
      '管理费率不在这张表了 —— 按需求改成在【业务流水】逐笔填。')
widths(ws, {'A':10,'B':46,'C':18,'D':11,'E':12,'F':11,'G':12,'H':12,'I':11,'J':15,
            'K':11,'L':11,'M':11,'N':9,'O':14,'P':14,'Q':13,'R':14,'S':10,'T':16,'U':16,'V':24})
headers(ws, 5, 1, ['项目编号','项目全称','项目简称','项目类型','合同类型','业主','一级单位','二级单位',
                   '我方主体','合同编号','签订日期','开工日期','竣工日期','状态',
                   '合同金额','审计(结算)\n金额','审减额','累计开票','开票\n进度','施工地址','校验','备注'])
PCOL_IN   = list('ABCDEFGHIJKLMNOPTV')      # 手工录入
PCOL_AUTO = list('QRSU')                    # 自动算

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

# 行序＝原表首现顺序（parse_gk2 产出的 proj_order），保证跟你手上那份对账明细一行对一行
for i, po in enumerate(PROJ_ORDER):
    p = po['full']
    info = seen.get(p, {'units': [], 'owner': '', 'L1': {}, 'lv1': '', 'lv2': ''})
    multi = len(info.get('L1', {})) > 1
    memo = []
    if multi:
        memo.append('该项目走了 ' + str(len(info['L1'])) + ' 条挂靠链：' + '、'.join(info['L1'].keys()) +
                    '，档案只能填一条，链条核算以此为准；其余链条请看单位项目明细')
    if po.get('short_dupe'):
        memo.append('简称与别的项目重名，建议手工改一个更好认的简称')
    memo.append('原表首现：' + po['first_sheet'].strip() + ' 表第 ' + str(po['first_row']) + ' 行')
    PROJ_ROWS.append(dict(code=f'A{i+1:03d}', name=p, short=po['short'],
                          ptype='合伙项目' if p in PARTNER else '自营项目',
                          ctype='框架合同', owner=info.get('owner') or '民能',
                          l1=info.get('lv1', ''), l2=info.get('lv2', ''), mine='泓普',
                          memo='；'.join(memo)))
NEW_CODE = f'A{len(PROJ_ROWS)+1:03d}'
PROJ_ROWS.append(dict(code=NEW_CODE, name='（新项目：请填项目全称）', short='（新项目）',
                      ptype='自营项目', ctype='框架合同', owner='民能', l1='', l2='', mine='泓普',
                      memo='本次预留的空白项目行'))
PROJ_BY_CODE = {d['code']: d for d in PROJ_ROWS}
PCODE = {d['name']: d['code'] for d in PROJ_ROWS if d['name']}
PSHORT = {d['code']: d['short'] for d in PROJ_ROWS}

for i in range(P1 - P0 + 1):
    r = P0 + i
    for c in PCOL_IN:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if c in 'BCTV' else C,
            fmt=DATE if c in 'KLM' else (MONEY if c in 'OP' else None))
    if i < len(PROJ_ROWS):
        d = PROJ_ROWS[i]
        ws[f'A{r}'] = d['code']; ws[f'B{r}'] = d['name']; ws[f'C{r}'] = d['short']
        ws[f'D{r}'] = d['ptype']; ws[f'E{r}'] = d['ctype']; ws[f'F{r}'] = d['owner']
        ws[f'I{r}'] = d['mine']
        if d['l1']: ws[f'G{r}'] = d['l1']
        if d['l2']: ws[f'H{r}'] = d['l2']
        ws[f'N{r}'] = '在建'
        if d['memo']: ws[f'V{r}'] = d['memo']
    put(ws, f'Q{r}', f'=IF($A{r}="","",IF(N($O{r})=0,"",ROUND(N($O{r})-N($P{r}),2)))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'R{r}', f'=IF($A{r}="","",SUMIFS({FRNG("amt")},{FRNG("kind")},"销项开票",'
                     f'{FRNG("proj")},$A{r},{FRNG("cnt")},"是"))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'S{r}', f'=IF($A{r}="","",IF(N($P{r})>0,$R{r}/$P{r},IF(N($O{r})>0,$R{r}/$O{r},"")))',
        font=F_LINK, fill=FILL_AUTO, fmt=PCT)
    put(ws, f'U{r}',
        f'=IF($A{r}="","",IF(COUNTIF($A${P0}:$A${P1},$A{r})>1,"编号重复",'
        f'IF($B{r}="","没填项目全称",'
        f'IF($C{r}="","没填项目简称",'
        f'IF(COUNTIF($C${P0}:$C${P1},$C{r})>1,"简称重名",'
        f'IF(AND(N($O{r})>0,N($P{r})>0,N($P{r})>N($O{r})),"审计金额大于合同金额",'
        f'IF(OR($F{r}="",$G{r}=""),"待完善（业主/一级单位）","√")))))))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 17
for c in 'FGHI': dv_list(ws, f'{c}{P0}:{c}{P1}', f'={U_NAME}')
dv_list(ws, f'D{P0}:D{P1}', '"自营项目,合伙项目"')
dv_list(ws, f'E{P0}:E{P1}', '"框架合同,劳务合同,施工合同,租赁合同,其他"')
dv_list(ws, f'N{P0}:N{P1}', '"在建,已完工,已结清,暂停"')
dv_num(ws, f'O{P0}:O{P1}'); dv_num(ws, f'P{P0}:P{P1}')
ws.conditional_formatting.add(f'U{P0}:U{P1}',
    FormulaRule(formula=[f'AND($U{P0}<>"",$U{P0}<>"√",LEFT($U{P0},3)<>"待完善")'],
                fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.conditional_formatting.add(f'U{P0}:U{P1}',
    FormulaRule(formula=[f'LEFT($U{P0},3)="待完善"'], fill=FILL_IN, font=Font(color='9C6500')))
ws.auto_filter.ref = f'A5:V{P1}'
ws.freeze_panes = 'D6'; page(ws, titles='5:5')
print(f'  ✓ 单位档案 / 项目档案·合同台账（{len(PROJ_ROWS)} 个项目，按原表首现顺序，'
      f'{sum(1 for d in PROJ_ROWS if d["ptype"]=="合伙项目")} 个合伙项目，预留新项目 {NEW_CODE}）')

# ============================================================ 业务流水（总表）
ws = wb.create_sheet(SH_FLOW)
C_ = FC          # 下面一律用 C_['xxx'] 取列字母，以后再调列序只改上面那张字典
title(ws, '业务流水（总表）', FLOW_LAST,
      '所有开票、成本票、回款、交税、管理费结算、往来款都在这里录一行。'
      '这一轮按你的要求，「管理费率」「返现率」和右边四列税费全部改成手工填的淡黄色格子 —— '
      '不再由单位档案的参数推算，一律以实际为准，后面所有表都从这张表取数。'
      '「计费方式」那一列说明这笔票要不要扣管理费、要不要我方回成本票，'
      '只过票不收费的选「不扣管理费」，合伙方内部分成、根本不用回票的选「不回成本票」。')
widths(ws, {'A':7,'B':11,'C':10,'D':13,'E':12,'F':12,'G':11,'H':12,'I':34,'J':14,'K':9,'L':9,
            'M':13,'N':16,'O':15,'P':12,'Q':12,'R':12,'S':12,
            'T':18,'U':38,'V':8,'W':9,'X':12,'Y':14,'Z':13,'AA':13,
            'AB':13,'AC':13,'AD':7,'AE':6,'AF':9,'AG':6,'AH':7,'AI':10,'AJ':7,'AK':10,
            'AL':10,'AM':12,'AN':12})
inband(ws, 'A', C_['chk'], 3)
ws.merge_cells(f'{C_["vat"]}3:{C_["inc"]}3')
put(ws, f'{C_["vat"]}3', '税费录入区 · 按实际发生逐笔填（只有销项开票行要填）',
    font=F_HDR2, fill=FILL_HDR2, align=C)
sysband(ws, C_['sname'], FLOW_LAST, 3)
headers(ws, 4, 1, ['序号','日期','项目编号','业务类型','开票/\n付款方','收票/\n收款方','发票性质','票据类型',
                   '摘要','金额','管理费率','返现率','计费方式','备注','校验'])
headers(ws, 4, col_idx(C_['vat']), ['预提增值税','预提附加税','预提印花税','预提所得税'],
        fill=FILL_IN, font=F_HDR2)
headers(ws, 4, col_idx(C_['sname']),
        ['项目简称','项目全称','层级','成本归类','管理费','应开成本票',
         '税费合计','税费\n参考试算','其他应收\n(返现)','其他应付\n(过账/合伙)','年度','代收\n标记',
         '代收台账\n序号','末层','对业主\n开票','来源表','源行'], fill=FILL_AUTO, font=F_HDR2)
put(ws, f'{C_["cnt"]}4', '计入汇总', font=F_HDR2, fill=FILL_HDR2)
put(ws, f'{C_["skey"]}4', '取数键', font=F_HDR2, fill=FILL_AUTO)
put(ws, f'{C_["ocol"]}4', '原表来源列', font=F_HDR2, fill=FILL_AUTO)
put(ws, f'{C_["ubel"]}4', '归属单位表', font=F_HDR2, fill=FILL_AUTO)

U_D  = f'{QU}!$D${U0}:$D${U1}'          # 上游开票税率
U_E  = f'{QU}!$E${U0}:$E${U1}'          # 我方回开票种税率
U_F  = f'{QU}!$F${U0}:$F${U1}'          # 有无税差
U_G  = f'{QU}!$G${U0}:$G${U1}'          # 增值税计算方式
U_I  = f'{QU}!$I${U0}:$I${U1}'          # 附加税率
U_J  = f'{QU}!$J${U0}:$J${U1}'          # 印花税率
U_K  = f'{QU}!$K${U0}:$K${U1}'          # 所得税预征率
U_TYPE = f'{QU}!$C${U0}:$C${U1}'
P_B  = f'{QP}!$B${P0}:$B${P1}'          # 项目全称
P_SN = f'{QP}!$C${P0}:$C${P1}'          # 项目简称
P_G  = f'{QP}!$G${P0}:$G${P1}'          # 一级单位
P_H  = f'{QP}!$H${P0}:$H${P1}'          # 二级单位
P_I  = f'{QP}!$I${P0}:$I${P1}'          # 我方主体
KE = FRNG('payer'); KF = FRNG('payee'); KD = FRNG('kind'); KC = FRNG('proj')
KQ = FRNG('cls');   KZ = FRNG('cnt')
FAH = FRNG('last')          # 末层标记
FTOP = FRNG('top')          # 对业主开票标记
FSRC = FRNG('src')          # 来源表

def lk(val, key, out, nf='""'): return f'IFERROR(INDEX({out},MATCH({val},{key},0)),{nf})'

A_ = C_['amt']; R_ = C_['rate']; B_ = C_['reb']; FG = C_['feeflag']
for r in range(F0, F1 + 1):
    for c, fmt, tx in [('B', DATE, 0), ('C', None, 0), ('D', None, 0), ('E', None, 0), ('F', None, 0),
                       ('G', None, 0), ('H', None, 0), (C_['memo'], None, 1), (A_, MONEY, 0),
                       (R_, PCT, 0), (B_, PCT, 0), (FG, None, 0), (C_['note'], None, 1),
                       (C_['vat'], MONEY, 0), (C_['add'], MONEY, 0),
                       (C_['stamp'], MONEY, 0), (C_['inc'], MONEY, 0)]:
        put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if tx else C, fmt=fmt)
    put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-{F0-1})', font=F_LINK)
    put(ws, f'{C_["sname"]}{r}',
        f'=IF($C{r}="","",{lk(f"$C{r}", P_CODE, P_SN, chr(34)+"⚠编号不存在"+chr(34))})',
        font=F_LINK, fill=FILL_AUTO, align=CL)
    put(ws, f'{C_["pname"]}{r}',
        f'=IF($C{r}="","",{lk(f"$C{r}", P_CODE, P_B, chr(34)+"⚠项目编号不存在"+chr(34))})',
        font=F_LINK, fill=FILL_AUTO, align=CL)
    put(ws, f'{C_["tier"]}{r}',
        f'=IF($E{r}="","",IF($E{r}={lk(f"$C{r}", P_CODE, P_G)},"一级",'
        f'IF($E{r}={lk(f"$C{r}", P_CODE, P_H)},"二级",'
        f'IF($E{r}={lk(f"$C{r}", P_CODE, P_I)},"我方","—"))))', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'{C_["cls"]}{r}',
        f'=IF($H{r}="","",IF(OR($H{r}="劳务票",$H{r}="土建票",$H{r}="安装票"),"劳务类",'
        f'IF($H{r}="机械设备票","机械类",IF($H{r}="材料票","材料类","其他"))))',
        font=F_LINK, fill=FILL_AUTO)
    # 管理费 / 应开成本票：完全按这一行手工填的费率和计费方式算，不再回头看任何档案
    put(ws, f'{C_["mfee"]}{r}',
        f'=IF($D{r}<>"销项开票",0,IF(${FG}{r}="不扣管理费",0,ROUND(N(${A_}{r})*N(${R_}{r}),2)))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'{C_["due"]}{r}',
        f'=IF($D{r}<>"销项开票",0,IF(${FG}{r}="不回成本票",0,'
        f'ROUND(N(${A_}{r})-${C_["mfee"]}{r},2)))', font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'{C_["taxsum"]}{r}',
        f'=ROUND(N(${C_["vat"]}{r})+N(${C_["add"]}{r})+N(${C_["stamp"]}{r})+N(${C_["inc"]}{r}),2)',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    # 税费参考试算：只是给录入的人对照用，不参与任何汇总
    _up = f'N({lk(f"$E{r}", U_NAME, U_D)})'
    _our = f'N({lk(f"$E{r}", U_NAME, U_E)})'
    _vat = (f'IF(OR($D{r}<>"销项开票",{lk(f"$E{r}", U_NAME, U_F)}<>"是",{_up}=0),0,'
            f'IF({lk(f"$E{r}", U_NAME, U_G)}="税差法",'
            f'ROUND(N(${A_}{r})/(1+{_up})*{_up}-IF({_our}=0,0,${C_["due"]}{r}/(1+{_our})*{_our}),2),'
            f'ROUND(N(${A_}{r})/(1+{_up})*{_up},2)))')
    put(ws, f'{C_["taxref"]}{r}',
        f'=IF($D{r}<>"销项开票","",ROUND(({_vat})*(1+N({lk(f"$E{r}", U_NAME, U_I)}))'
        f'+N(${A_}{r})*N({lk(f"$E{r}", U_NAME, U_J)})'
        f'+IF({_up}=0,0,N(${A_}{r})/(1+{_up})*N({lk(f"$E{r}", U_NAME, U_K)})),2))',
        font=F_NOTE, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'{C_["ar"]}{r}', f'=IF($D{r}="销项开票",ROUND(N(${A_}{r})*N(${B_}{r}),2),'
                              f'IF($D{r}="其他应收收回",-N(${A_}{r}),0))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'{C_["ap"]}{r}', f'=IF($D{r}="其他应付发生",N(${A_}{r}),'
                              f'IF(OR($D{r}="其他应付扣税",$D{r}="其他应付支付"),-N(${A_}{r}),0))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'{C_["year"]}{r}', f'=IF($B{r}="","",IF(ISNUMBER($B{r}),YEAR($B{r}),"日期非法"))',
        font=F_LINK, fill=FILL_AUTO)
    put(ws, f'{C_["dflag"]}{r}',
        f'=IF(AND(OR($D{r}="挂靠代收",$D{r}="我方收款"),ISNUMBER($B{r}),'
        f'$B{r}>={QD}!$A$4,$B{r}<={QD}!$B$4),1,0)', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'{C_["dno"]}{r}',
        f'=IF(${C_["dflag"]}{r}=0,"",SUM(${C_["dflag"]}${F0}:${C_["dflag"]}{r}))',
        font=F_LINK, fill=FILL_AUTO)
    # 末层＝这一行的开票方，在同一项目里没有再从别人手上收到销项票 → 该由我方直接回成本票给他
    put(ws, f'{C_["last"]}{r}', f'=IF($D{r}<>"销项开票",0,'
                                f'IF(COUNTIFS({KC},$C{r},{KD},"销项开票",{KF},$E{r},{KZ},"是")=0,1,0))',
        font=F_LINK, fill=FILL_AUTO)
    put(ws, f'{C_["top"]}{r}', f'=IF($D{r}<>"销项开票",0,'
                               f'IF({lk(f"$F{r}", U_NAME, U_TYPE)}="业主",1,0))',
        font=F_LINK, fill=FILL_AUTO)
    put(ws, f'{C_["src"]}{r}', None, font=F_LINK, fill=FILL_AUTO)
    put(ws, f'{C_["srow"]}{r}', None, font=F_NOTE, fill=FILL_AUTO)
    put(ws, f'{C_["ocol"]}{r}', None, font=F_NOTE, fill=FILL_AUTO)
    # 归属单位表＝这一笔算在哪家挂靠单位的对账表上。历史行照原表的来源表；
    # 新录的行自动认开票方（开票方不是挂靠单位就认收票方）
    put(ws, f'{C_["ubel"]}{r}',
        f'=IF(${C_["src"]}{r}<>"",${C_["src"]}{r},'
        f'IF({lk(f"$E{r}", U_NAME, U_TYPE)}="挂靠单位",$E{r},'
        f'IF({lk(f"$F{r}", U_NAME, U_TYPE)}="挂靠单位",$F{r},"")))',
        font=F_NOTE, fill=FILL_AUTO)
    # 取数键：来源表#该表内第几笔 —— 8 张单位竖版明细靠它一行一行取数
    put(ws, f'{C_["skey"]}{r}',
        f'=IF(${C_["src"]}{r}="","",${C_["src"]}{r}&"#"&'
        f'COUNTIF(${C_["src"]}${F0}:${C_["src"]}{r},${C_["src"]}{r}))',
        font=F_NOTE, fill=FILL_AUTO)
    put(ws, f'{C_["chk"]}{r}',
        f'=IF($B{r}="","",'
        f'IF(NOT(ISNUMBER($B{r})),"日期格式不对",'
        f'IF($D{r}="","未选业务类型",'
        f'IF(AND($C{r}="",OR($D{r}="销项开票",$D{r}="成本票",$D{r}="工资扣抵",'
        f'$D{r}="挂靠代收",$D{r}="我方收款")),"未选项目编号",'
        f'IF(LEFT(${C_["pname"]}{r},1)="⚠","项目编号不存在",'
        f'IF($E{r}="","未选开票/付款方",'
        f'IF(ISNA(MATCH($E{r},{U_NAME},0)),"开票方不在单位档案",'
        f'IF($F{r}="","未选收票/收款方",'
        f'IF(ISNA(MATCH($F{r},{U_NAME},0)),"收款方不在单位档案",'
        f'IF(NOT(ISNUMBER(${A_}{r})),"金额须为数字",'
        f'IF(AND(OR($D{r}="销项开票",$D{r}="成本票"),$H{r}=""),"未选票据类型",'
        f'IF(AND($D{r}="销项开票",${FG}{r}=""),"没选计费方式",'
        f'IF(AND($D{r}="销项开票",${FG}{r}="扣管理费",ROUND(N(${R_}{r}),6)=0),'
        f'"选了扣管理费但费率是 0",'
        f'IF(AND($D{r}="销项开票",${C_["tier"]}{r}="—"),"链条待确认","√"))))))))))))))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16

# ---- 导入历史事件 ----
KIND_MAP = {'销项开票':'销项开票','成本票':'成本票','挂靠单位代收':'挂靠代收','我方收款':'我方收款',
            '管理费结算':'管理费结算','扣质保金':'扣质保金','已交税':'已交税'}
# 原对账明细的表名 → 单位简称（只有杰华那张表名字不一样）
SHEET2UNIT = {'德誉嘉': '德誉嘉', '迅驰': '迅驰', '华城': '华城', '金沁': '金沁',
              '湖南锦泰': '湖南锦泰', '康欣': '康欣', '安锐': '安锐', '杰华电气': '杰华'}
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
    ws[f'{C_["memo"]}{r}'] = e['memo'][:120]
    ws[f'{A_}{r}'] = e['amt']
    ws[f'{C_["note"]}{r}'] = '历史导入'
    ws[f'{C_["src"]}{r}'] = SHEET2UNIT.get(e['src'].split('!')[0].strip(),
                                             e['src'].split('!')[0].strip())
    ws[f'{C_["srow"]}{r}'] = int(e['src'].split('!')[1])
    ws[f'{C_["cnt"]}{r}'] = e.get('count_in', '是')
    ws[f'{C_["ocol"]}{r}'] = '已开成本票' if e.get('from_cost_col') else '销售开票金额'
    if e.get('count_in') == '否':
        ws[f'{C_["note"]}{r}'] = '与 ' + str(e.get('dup_of', '')) + ' 为同一张票，只计一次'
    if e['kind'] == '销项开票':
        # 费率、返现率、税费一律照原表这一行写死（原来按单位档案参数算，跟原表对不上）
        ws[f'{R_}{r}'] = round(e.get('mrate', 0.0), 6)
        ws[f'{B_}{r}'] = round(e['rebate'] / e['amt'], 6) if (e.get('rebate') and e['amt']) else 0
        if e.get('from_cost_col'):
            # 这一行是从原表「已开成本票」列还原出来的票，本身不产生「我方应回成本票」的义务
            ws[f'{FG}{r}'] = '不回成本票'
            ws[f'{C_["note"]}{r}'] = '历史导入·原表「已开成本票」列还原；不另生成应回成本票'
        elif abs(e.get('mrate', 0.0)) < 1e-9:
            # 原表这一笔一分管理费没扣：要么只过票不收费，要么合伙内部连成本票都不用回
            ws[f'{FG}{r}'] = '不回成本票' if abs(float(e.get('cost_due') or 0)) < 0.005 else '不扣管理费'
        else:
            ws[f'{FG}{r}'] = '扣管理费'
        for col, key in (('vat', 'tax_v'), ('add', 'tax_s'), ('stamp', 'tax_y'), ('inc', 'tax_i')):
            v = round(float(e.get(key) or 0), 2)
            if v: ws[f'{C_[col]}{r}'] = v
    imported += 1

# ---- 追加：迅驰过账应付 + 应扣税费 + 工资扣抵 ----
r = F0 + imported
extra = 0
def _extra(date, code, kind, payer, payee, itype, memo, amt, note, src, srow):
    global r, extra
    ws[f'B{r}'] = date
    ws[f'C{r}'] = code
    ws[f'D{r}'] = kind
    ws[f'E{r}'] = payer; ws[f'F{r}'] = payee
    ws[f'H{r}'] = itype
    ws[f'{C_["memo"]}{r}'] = memo
    ws[f'{A_}{r}'] = amt
    ws[f'{C_["note"]}{r}'] = note
    ws[f'{C_["src"]}{r}'] = src; ws[f'{C_["srow"]}{r}'] = srow
    ws[f'{C_["cnt"]}{r}'] = '是'
    r += 1; extra += 1
for x in PASS_ROWS:
    _extra(dt.date(2026, 3, 27), PCODE.get(x['proj'], ''), '其他应付发生', '迅驰', '迅驰', '不适用',
           '迅驰过账项目应付工程款：' + x['memo'][:60], x['amt'],
           '历史导入·原表迅驰「过账应付工程款」列；实际收款对象请按实补填', '迅驰', x.get('row', ''))
for x in WAGE_ROWS:
    _extra(dt.date(2026, 6, 30), PCODE.get(x['proj'], ''), '工资扣抵', '泓普', '康欣', '劳务票',
           '工资表顶抵的劳务成本，不再另开成本票', x['amt'],
           '历史导入·原总台账「劳务成本·工资扣抵」列', '总台账', x.get('row', ''))
for x in DED_ROWS:
    _extra(dt.date(2026, 3, 27), '', '其他应付扣税', '泓普', '迅驰', '不适用',
           '过账应付里要扣除的税费：' + x['memo'][:60], x['amt'],
           '历史导入·原表迅驰「已付款」列', '迅驰', x.get('row', ''))
for rr in range(F0, F1 + 1):
    cell = f'{C_["cnt"]}{rr}'
    if ws[cell].value is None:
        put(ws, cell, '是', font=F_IN, fill=FILL_IN)
    else:
        put(ws, cell, None, font=F_IN, fill=FILL_IN)
dv_list(ws, f'{C_["cnt"]}{F0}:{C_["cnt"]}{F1}', '"是,否"')
dv_list(ws, f'C{F0}:C{F1}', f'={P_CODE}')
dv_list(ws, f'D{F0}:D{F1}', KIND_DV)
dv_list(ws, f'E{F0}:E{F1}', f'={U_NAME}')
dv_list(ws, f'F{F0}:F{F1}', f'={U_NAME}')
dv_list(ws, f'G{F0}:G{F1}', '"13%专票,9%专票,6%专票,3%专票,1%普票,3%普票,13%普票,不开票"')
dv_list(ws, f'H{F0}:H{F1}', ITYPE_DV)
dv_list(ws, f'{FG}{F0}:{FG}{F1}', FEE_DV)
dv_num(ws, f'{A_}{F0}:{A_}{F1}', 'greaterThanOrEqual', '-99999999')
for c in (C_['vat'], C_['add'], C_['stamp'], C_['inc']):
    dv_num(ws, f'{c}{F0}:{c}{F1}', 'greaterThanOrEqual', '-99999999')
CHK = C_['chk']
ws.conditional_formatting.add(f'{CHK}{F0}:{CHK}{F1}',
    FormulaRule(formula=[f'AND(${CHK}{F0}<>"",${CHK}{F0}<>"√",${CHK}{F0}<>"链条待确认")'],
                fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.conditional_formatting.add(f'{CHK}{F0}:{CHK}{F1}',
    FormulaRule(formula=[f'${CHK}{F0}="链条待确认"'], fill=FILL_IN, font=Font(color='9C6500')))
ws.conditional_formatting.add(f'{FG}{F0}:{FG}{F1}',
    FormulaRule(formula=[f'${FG}{F0}<>"扣管理费"'], fill=FILL_IN, font=Font(color='9C6500', bold=True)))
ws.auto_filter.ref = f'A4:{C_["cnt"]}{F1}'
ws.column_dimensions[C_['skey']].hidden = True
ws.column_dimensions[C_['srow']].hidden = True
ws.column_dimensions[C_['ocol']].hidden = True
ws.column_dimensions[C_['ubel']].hidden = True
ws.freeze_panes = 'C5'; page(ws, titles='4:4')
FLOW_USED = imported + extra
print(f'  ✓ 业务流水（历史 {imported} 笔 + 过账/工资/扣税 {extra} 笔 = {FLOW_USED} 笔，容量 {F1-F0+1} 行）')

# ============================================================ 资金日记账
ws = wb.create_sheet(SH_JOUR)
title(ws, '资金日记账（泓普 · 仟茂 · 现金 混合录入）', 'Q')
widths(ws, {'A':7,'B':11,'C':10,'D':10,'E':13,'F':12,'G':10,'H':40,'I':13,'J':13,'K':14,'L':14,'M':11,
            'N':15,'O':34,'P':9,'Q':7})
inband(ws, 'A', 'J', 3); sysband(ws, 'K', 'Q', 3)
headers(ws, 4, 1, ['序号','日期','资金账户','项目编号','费用类型','往来单位','经办人','摘要','收入','支出'])
headers(ws, 4, 11, ['该账户余额','三账户合计','工程\n回款','校验','项目简称','年月','年度'],
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
    put(ws, f'O{r}', f'=IF($D{r}="","",IFERROR(INDEX({P_SN},MATCH($D{r},{P_CODE},0)),"⚠编号不存在"))',
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
        f'IF(AND($M{r}="是",$F{r}=""),"工程回款必须填往来单位","√")))))))', font=F_TXT, fill=FILL_CHK)
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

# ============================================================ 资金日记账 · 三张拆分表
# 每个账户一张，谁管哪个账户就录哪一张；右边「复制到总表」接口把这张表按总表的列序排好，
# 整块复制 → 到总表选择性粘贴「数值」即可，不会碰到总表右边的公式列。
SJ_0, SJ_N = 5, 600
SJ_1 = SJ_0 + SJ_N - 1
SPLIT_SHEETS = []
for ai, acct in enumerate(SPLIT_ACCTS):
    nm = SH_SPLIT[acct]
    ws = wb.create_sheet(nm)
    title(ws, f'资金日记账 · {acct}（单账户录入表）', 'Z')
    widths(ws, {'A':7,'B':11,'C':10,'D':11,'E':12,'F':10,'G':38,'H':13,'I':13,'J':9,'K':16,
                'L':14,'M':14,'N':18,'O':3,
                'P':11,'Q':10,'R':10,'S':11,'T':12,'U':10,'V':38,'W':13,'X':13,'Y':3,'Z':10})
    put(ws, 'A2', '期初余额', font=F_H2, align=CR, border=None)
    put(ws, 'B2', f'={QJ}!${L(3 + ai * 2)}$2', font=F_TOT, fill=FILL_CHK, fmt=MONEY)
    put(ws, 'C2', f'期初跟总表第 2 行联动。只记【{acct}】这一个账户，左边淡黄色格子照常录，'
                  'L 列自动滚出本账户余额。录完把右边灰色「复制到总表」整块复制，'
                  '到【资金日记账】选择性粘贴「数值」。本表不参与任何汇总，不会重复计数。',
        font=F_NOTE, align=CL, border=None)
    ws.merge_cells('C2:Z2')
    inband(ws, 'A', 'N', 3)
    ws.merge_cells('P3:Z3')
    put(ws, 'P3', '复制到总表 · 灰色区自动排好，不要手工改', font=F_HDR2, fill=FILL_AUTO, align=C)
    headers(ws, 4, 1, ['序号','日期','项目编号','费用类型','往来单位','经办人','摘要','收入','支出',
                       '工程\n回款','备注','本账户余额','校验','项目简称'])
    headers(ws, 4, 16, ['日期','资金账户','项目编号','费用类型','往来单位','经办人','摘要','收入','支出'],
            fill=FILL_AUTO, font=F_HDR2)
    put(ws, 'Z4', '工程回款', font=F_HDR2, fill=FILL_AUTO)
    for i in range(SJ_N):
        r = SJ_0 + i
        for c, fmt, tx in [('B', DATE, 0), ('C', None, 0), ('D', None, 0), ('E', None, 0),
                           ('F', None, 0), ('G', None, 1), ('H', MONEY, 0), ('I', MONEY, 0),
                           ('J', None, 0), ('K', None, 1)]:
            put(ws, f'{c}{r}', None, font=F_IN, fill=FILL_IN, align=CL if tx else C, fmt=fmt)
        put(ws, f'A{r}', f'=IF($B{r}="","",ROW()-{SJ_0-1})', font=F_LINK)
        put(ws, f'L{r}', f'=IF($B{r}="","",$B$2+SUM($H${SJ_0}:$H{r})-SUM($I${SJ_0}:$I{r}))',
            font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
        put(ws, f'M{r}',
            f'=IF($B{r}="","",'
            f'IF(NOT(ISNUMBER($B{r})),"日期格式不对",'
            f'IF(AND(N($H{r})=0,N($I{r})=0),"收支都为空",'
            f'IF(AND(N($H{r})>0,N($I{r})>0),"收支不能同时填",'
            f'IF(AND($C{r}<>"",ISNA(MATCH($C{r},{P_CODE},0))),"项目编号不存在",'
            f'IF(AND($J{r}="是",$E{r}=""),"工程回款必须填往来单位","√"))))))',
            font=F_TXT, fill=FILL_CHK)
        put(ws, f'N{r}', f'=IF($C{r}="","",IFERROR(INDEX({P_SN},MATCH($C{r},{P_CODE},0)),"⚠编号不存在"))',
            font=F_LINK, fill=FILL_AUTO, align=CL)
        # 接口区：按总表 B~J 的列序排好
        put(ws, f'P{r}', f'=IF($B{r}="","",$B{r})', font=F_LINK, fill=FILL_AUTO, fmt=DATE)
        put(ws, f'Q{r}', f'=IF($B{r}="","","{acct}")', font=F_LINK, fill=FILL_AUTO)
        for tgt, srcc in (('R', 'C'), ('S', 'D'), ('T', 'E'), ('U', 'F'), ('V', 'G')):
            put(ws, f'{tgt}{r}', f'=IF($B{r}="","",${srcc}{r})', font=F_LINK, fill=FILL_AUTO,
                align=CL if tgt == 'V' else C)
        put(ws, f'W{r}', f'=IF($B{r}="","",IF(N($H{r})=0,"",N($H{r})))',
            font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
        put(ws, f'X{r}', f'=IF($B{r}="","",IF(N($I{r})=0,"",N($I{r})))',
            font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
        put(ws, f'Z{r}', f'=IF($B{r}="","",$J{r})', font=F_LINK, fill=FILL_AUTO)
        ws.row_dimensions[r].height = 16
    put(ws, f'A{SJ_1+2}',
        f'怎么用：① 左边淡黄色格子照常录【{acct}】的收支；② 选中 P{SJ_0}:X{SJ_1} 整块复制 → '
        f'打开【资金日记账】→ 点到第一行空白行的 B 列 → 右键「选择性粘贴 → 数值」；'
        f'③ 如果有工程回款要标记，再把 Z{SJ_0}:Z{SJ_1} 复制粘到总表 M 列对应位置。'
        '④ 粘完在总表上核对一眼「校验」列全是 √ 就完事。'
        '注意：本表不参与任何汇总，【费用统计】【项目利润】等一律只认总表，所以不会重复算。',
        font=F_NOTE, align=CL, border=None)
    ws.merge_cells(f'A{SJ_1+2}:Z{SJ_1+2}')
    dv_list(ws, f'C{SJ_0}:C{SJ_1}', f'={P_CODE}')
    dv_list(ws, f'D{SJ_0}:D{SJ_1}', '"' + ','.join(ETYPES) + '"')
    dv_list(ws, f'E{SJ_0}:E{SJ_1}', f'={U_NAME}')
    dv_list(ws, f'J{SJ_0}:J{SJ_1}', '"是,否"')
    dv_num(ws, f'H{SJ_0}:H{SJ_1}'); dv_num(ws, f'I{SJ_0}:I{SJ_1}')
    ws.conditional_formatting.add(f'M{SJ_0}:M{SJ_1}',
        FormulaRule(formula=[f'AND($M{SJ_0}<>"",$M{SJ_0}<>"√")'],
                    fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
    ws.column_dimensions['O'].width = 3
    ws.column_dimensions['Y'].width = 3
    ws.freeze_panes = 'C5'; page(ws, titles='4:4')
    ws.print_area = f'$A$1:$N${SJ_1}'
    SPLIT_SHEETS.append(nm)
print(f'  ✓ 资金日记账拆分表：{"、".join(SPLIT_SHEETS)}（各 {SJ_N} 行，带复制到总表的接口区）')

# ============================================================ 取数公式（带起止日期）
FB = FRNG('date')          # 日期
FI = FRNG('amt')           # 金额
FR = FRNG('mfee')          # 管理费
FS = FRNG('due')           # 应开成本票
FFLAG = FRNG('feeflag')    # 是否扣管理费
FV = FRNG('vat'); FW = FRNG('add')
FX = FRNG('stamp'); FY = FRNG('inc'); FZ = FRNG('taxsum')
FAA = FRNG('ar'); FAB = FRNG('ap')
def agg(val, side, kind, uref, pref=None, dr=None, cls=None, last=False, top=False,
        fee=None, src=None):
    """side: 'E' 开票/付款方  'F' 收票/收款方  None 不限；dr=(起,止) 加日期区间；
       last=True 只取链条末层的销项票（我方该直接回成本票的那一层）；
       top=True 只取直接开给业主的那一层（这个工程的收入口径）；
       fee='扣管理费'/'不扣管理费' 按这一笔到底扣没扣管理费再筛一道；
       src=来源表 只取原对账明细某一张表来的行"""
    ex = ''
    if side: ex += f',{KE if side == "E" else KF},{uref}'
    if pref: ex += f',{KC},{pref}'
    if cls:  ex += f',{KQ},"{cls}"'
    if last: ex += f',{FAH},1'
    if top:  ex += f',{FTOP},1'
    if fee:  ex += f',{FFLAG},"{fee}"'
    if src:  ex += f',{FSRC},{src}'
    if dr:   ex += f',{FB},">="&{dr[0]},{FB},"<="&{dr[1]}'
    return f'SUMIFS({val},{KD},"{kind}"{ex},{KZ},"是")'

COLS_U = ['开票额\n(该单位开出)','应扣管理费','应到成本票','其中·扣了\n管理费的','其中·没扣\n管理费的',
          '已收成本票','还差成本票','应提税费','已交税','欠税未交',
          '业主已付给\n挂靠单位','业主未付','挂靠单位\n已转我方','挂靠单位\n代收未转','管理费\n已结算','扣质保金']
SUM_LAST = L(2 + len(COLS_U))          # 最后一个金额列 = R
BIZ = L(3 + len(COLS_U))               # 有无业务 = S

def summary_formulas(uref, pref=None, dr=None):
    g = lambda v, s, k, **kw: agg(v, s, k, uref, pref, dr, **kw)
    return [g(FI,'E','销项开票'), g(FR,'E','销项开票'), g(FS,'E','销项开票'),
            g(FS,'E','销项开票', fee='扣管理费'), g(FS,'E','销项开票', fee='不扣管理费'),
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
put(wsu, f'O{HR}', '挂靠序号', font=F_HDR2, fill=FILL_AUTO)
for r in range(U0, U1 + 1):
    put(wsu, f'O{r}', f'=IF($C{r}<>"挂靠单位","",COUNTIF($C${U0}:$C{r},"挂靠单位"))',
        font=F_NOTE, fill=FILL_AUTO)
wsu.column_dimensions['O'].hidden = True
U_RANK = f'{QU}!$O${U0}:$O${U1}'

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
    """一行 16 列的通用汇总公式 + 有无业务"""
    for i, f in enumerate(summary_formulas(uref, pref, dr)):
        if f is None: continue
        put(ws, f'{L(3+i)}{r}', f'=IF($A{r}="","",{f})', font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($A{r}="","",$E{r}-$H{r})', font=F_TXT, fmt=MONEY)   # 还差成本票
    put(ws, f'L{r}', f'=IF($A{r}="","",$J{r}-$K{r})', font=F_TXT, fmt=MONEY)   # 欠税未交
    put(ws, f'N{r}', f'=IF($A{r}="","",$C{r}-$M{r})', font=F_TXT, fmt=MONEY)   # 业主未付
    put(ws, f'P{r}', f'=IF($A{r}="","",$M{r}-$O{r})', font=F_TOT, fmt=MONEY)   # 代收未转
    put(ws, f'{biz_col}{r}', f'=IF($A{r}="","",IF(ROUND(SUM($C{r}:$H{r})+SUM($M{r}:$O{r}),2)=0,"无","有"))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16

SUM_W = {'A': 11, 'B': 30}
for i in range(len(COLS_U)): SUM_W[L(3 + i)] = 14
SUM_W[BIZ] = 10

# ============================================================ 单位汇总
ws = wb.create_sheet(SH_SUM_U)
title(ws, '分表 · 按挂靠单位汇总', BIZ,
      '每个挂靠单位一行，行的先后跟你那份《对账明细》的表顺序一致（德誉嘉→迅驰→华城→金沁→湖南锦泰→康欣→安锐→杰华）。'
      '「应提税费 / 已交税 / 欠税未交」这一轮改成按原表逐笔录入的实际数，不再按参数预提 —— '
      '所以迅驰是 0（不欠税）、华城也跟原表一致了。'
      '「应到成本票」右边拆出「扣了管理费的 / 没扣管理费的」两列，只过票不收费的那部分一眼看得见。')
widths(ws, SUM_W)
DR = filter_band(ws, BIZ)
headers(ws, HR, 1, ['单位简称', '单位全称'] + COLS_U + ['有无业务'])
for i in range(U1 - U0 + 1):
    r = Q0 + i
    put(ws, f'A{r}', f'=IFERROR(INDEX({U_NAME},MATCH({i+1},{U_RANK},0)),"")', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",INDEX({QU}!$B${U0}:$B${U1},MATCH($A{r},{U_NAME},0)))',
        font=F_LINK, align=CL)
    detail_cols(ws, r, f'$A{r}', None, DR, SUM_LAST, BIZ)
    # 税费按「归属单位表」统计（＝这一笔算在哪家的对账表上），这样跟你原表那 8 张表逐家对得上；
    # 上面几列仍按开票方/收票方统计，好处是全公司合计不会把同一张票算两遍
    _dr = f',{FB},">="&{DR[0]},{FB},"<="&{DR[1]}'
    put(ws, f'J{r}', f'=IF($A{r}="","",SUMIFS({FZ},{FRNG("ubel")},$A{r}{_dr}))',
        font=F_LINK, fmt=MONEY)
    put(ws, f'K{r}', f'=IF($A{r}="","",SUMIFS({FI},{KD},"已交税",{FRNG("ubel")},$A{r}{_dr}))',
        font=F_LINK, fmt=MONEY)
totals_row(ws, [L(3 + i) for i in range(len(COLS_U))], Q0, QU_1)
put(ws, f'{BIZ}{TR}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
ws.conditional_formatting.add(f'L{Q0}:L{QU_1}',
    FormulaRule(formula=[f'$L{Q0}>0.01'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
N_UNITS = sum(1 for u in UNITS if u[2] == '挂靠单位')
hide_tail(ws, Q0, QU_1, set(range(Q0, Q0 + N_UNITS)), col_idx(BIZ) - 1, f'A{HR}:{BIZ}{QU_1}')
ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')

# ============================================================ 项目汇总
ws = wb.create_sheet(SH_SUM_P)
title(ws, '分表 · 按项目汇总', BIZ,
      '每个项目一行，行序＝项目档案的行序＝你那份《对账明细》里项目第一次出现的先后。'
      '「应到成本票」后面跟着两列：扣了管理费的那部分、没扣管理费的那部分 —— '
      '只过票不收费（管理费率填 0）的票会落到「没扣管理费的」那一列，不用再一笔笔翻。')
widths(ws, SUM_W)
DR = filter_band(ws, BIZ)
headers(ws, HR, 1, ['项目编号', '项目简称'] + COLS_U + ['有无业务'])
for i in range(P1 - P0 + 1):
    r, pr = Q0 + i, P0 + i
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$C{pr})', font=F_LINK, align=CL)
    detail_cols(ws, r, '"*"', f'$A{r}', DR, SUM_LAST, BIZ)
totals_row(ws, [L(3 + i) for i in range(len(COLS_U))], Q0, QP_1)
put(ws, f'{BIZ}{TR}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
ws.conditional_formatting.add(f'L{Q0}:L{QP_1}',
    FormulaRule(formula=[f'$L{Q0}>0.01'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.conditional_formatting.add(f'G{Q0}:G{QP_1}',
    FormulaRule(formula=[f'$G{Q0}>0.01'], fill=FILL_IN, font=Font(color='9C6500', bold=True)))
NP = len(PROJ_ROWS)
hide_tail(ws, Q0, QP_1, set(range(Q0, Q0 + NP)), col_idx(BIZ) - 1, f'A{HR}:{BIZ}{QP_1}')
ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')
print('  ✓ 单位汇总 / 项目汇总（应到成本票已拆「扣费 / 不扣费」两列）')

# ============================================================ 单位项目明细（选单位看它名下的项目）
PAIR = {}
for e in EV:
    if e.get('count_in', '是') != '是': continue
    code = PCODE.get(e['proj'], '')
    if not code: continue
    for u in {e.get('payer'), e.get('payee')}:
        if u: PAIR.setdefault(u, set()).add(code)
HOLD_UNITS = [u[0] for u in UNITS if u[2] == '挂靠单位']

ws = wb.create_sheet(SH_SUM_X)
title(ws, '分表 · 单位 × 项目明细', BIZ,
      '上面选一个挂靠单位，下面列出它名下每个项目的开票、成本票、回款、欠税情况。'
      '要给领导看逐笔的，翻后面那 8 张单位专表（表名就是单位名）—— 那几张是照你原对账表的样子竖着排的。')
widths(ws, SUM_W)
DR = filter_band(ws, BIZ, unit_default='康欣')
dv_list(ws, 'B3', f'={U_NAME}')
headers(ws, HR, 1, ['项目编号', '项目简称'] + COLS_U + ['有无业务'])
for i in range(P1 - P0 + 1):
    r, pr = Q0 + i, P0 + i
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$C{pr})', font=F_LINK, align=CL)
    detail_cols(ws, r, '$B$3', f'$A{r}', DR, SUM_LAST, BIZ)
totals_row(ws, [L(3 + i) for i in range(len(COLS_U))], Q0, QP_1)
put(ws, f'{BIZ}{TR}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
ws.conditional_formatting.add(f'L{Q0}:L{QP_1}',
    FormulaRule(formula=[f'$L{Q0}>0.01'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
hide_tail(ws, Q0, QP_1, set(range(Q0, Q0 + NP)), col_idx(BIZ) - 1, f'A{HR}:{BIZ}{QP_1}')
ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')
print('  ✓ 单位项目明细')

# ============================================================ 8 张单位竖版明细（给领导看）
# 表头照你那份《对账明细》的两层结构，一笔业务一行，按日期排；
# 数据源是【业务流水】里「来源表＝这家单位」的行，靠隐藏的「取数键」一行一行取。
SRC_SHEET = {u: u for u in ('德誉嘉', '迅驰', '华城', '金沁', '湖南锦泰', '康欣', '安锐', '杰华')}
SRC_XLS = {'德誉嘉': '德誉嘉 ', '迅驰': '迅驰', '华城': '华城', '金沁': '金沁',
           '湖南锦泰': '湖南锦泰', '康欣': '康欣', '安锐': '安锐', '杰华': '杰华电气'}
VD_0 = 7                       # 明细首行
VD_N = 250                     # 每张单位表留 250 行
VD_1 = VD_0 + VD_N - 1
VD_TOT = 6                     # 合计行（表头正下方）
VCOLS = [
    ('A', 7,  '序号', ''),            ('B', 11, '日 期', ''),
    ('C', 20, '项目简称', ''),        ('D', 40, '摘　　要', ''),
    ('E', 10, '业务\n性质', ''),
    ('F', 10, '发票\n性质', '开票情况'), ('G', 10, '收票\n单位', '开票情况'),
    ('H', 14, '销售开票金额', '开票情况'), ('I', 13, '应扣管理费', '开票情况'),
    ('J', 13, '应到成本票', '开票情况'), ('K', 13, '已到成本票', '开票情况'),
    ('L', 13, '剩余开票金额', '开票情况'),
    ('M', 12, '预收增值部', '交税情况'), ('N', 12, '预收附加税', '交税情况'),
    ('O', 12, '预收印花税', '交税情况'), ('P', 12, '预收所得税', '交税情况'),
    ('Q', 12, '应扣税费', '交税情况'), ('R', 12, '已交税', '交税情况'),
    ('S', 14, '业主付给\n挂靠单位', '回款情况'), ('T', 14, '挂靠单位\n转我方', '回款情况'),
    ('U', 13, '管理费结算', '回款情况'), ('V', 12, '扣质保金', '回款情况'),
    ('W', 22, '备　注', ''), ('X', 9, '在期间内', ''),
]
VD_LAST = 'X'
VD_MONEY = list('HIJKLMNOPQRSTUV')

def vfetch(key_cell, col_key):
    return f'IFERROR(INDEX({FRNG(col_key)},MATCH({key_cell},{FRNG("skey")},0)),"")'

SUB_SHEETS = []
for u in [x[0] for x in UNITS if x[2] == '挂靠单位']:
    sh = SRC_SHEET[u]
    nm = f'{u}明细'
    ws = wb.create_sheet(nm)
    title(ws, f'{u} · 对账明细（给领导看的逐笔明细）', VD_LAST,
          f'表头照你那份《对账明细》的「{u}」表来，原来横着排的一笔一笔改成竖着排，按日期顺序。'
          '明细全部显示不折叠；上面填年度或起止日期，第 6 行的合计只统计落在期间内的行（最后一列会标是/否）。'
          '数据全部来自【业务流水】里来源表＝本单位的行，在那边改一笔，这里立刻跟着变。')
    widths(ws, {c: w for c, w, _, _ in VCOLS})
    # 筛选带
    put(ws, 'A3', '本表单位', font=F_H2, fill=FILL_HDR2)
    put(ws, 'B3', u, font=Font(name='微软雅黑', size=11, bold=True, color='1F3864'),
        fill=FILL_HDR2, align=CL)
    for lab, lc, col, fmt in [('年度', 'C', 'D', '0'), ('起始日期', 'E', 'F', DATE),
                              ('截止日期', 'G', 'H', DATE)]:
        put(ws, f'{lc}3', lab, font=F_H2, fill=FILL_HDR2)
        put(ws, f'{col}3', None, font=Font(name='微软雅黑', size=10, bold=True, color='0000C0'),
            fill=FILL_IN, fmt=fmt)
    put(ws, 'I3', '当前取数', font=F_H2, fill=FILL_HDR2)
    for ci in range(10, col_idx(VD_LAST) + 1):
        put(ws, f'{L(ci)}3', None, font=F_TOT, fill=FILL_CHK, align=CL)
    ws.merge_cells(f'J3:{VD_LAST}3')
    put(ws, 'J3', '=IF(AND($D$3="",$F$3="",$H$3=""),"全部期间",'
                  'TEXT($A$4,"yyyy-mm-dd")&"  至  "&TEXT($B$4,"yyyy-mm-dd"))',
        font=F_TOT, fill=FILL_CHK, align=CL)
    put(ws, 'A4', '=IF($F$3<>"",$F$3,IF($D$3<>"",DATE($D$3,1,1),DATE(1900,1,1)))', font=F_NOTE, fmt=DATE)
    put(ws, 'B4', '=IF($H$3<>"",$H$3,IF($D$3<>"",DATE($D$3,12,31),DATE(2199,12,31)))', font=F_NOTE, fmt=DATE)
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].hidden = True
    # 两层表头（第 5 行分组、第 5/6 合并给单列，明细列名在第 5 行下一行）
    HR1, HR2 = 5, 5      # 单层表头即可，分组名放在第 5 行上方的第 4 行会被隐藏，故并入列名
    for c, w, nmc, grp in VCOLS:
        put(ws, f'{c}{HR2}', (grp + '\n' + nmc) if grp else nmc,
            font=F_HDR2, fill=FILL_HDR2 if grp else FILL_AUTO, align=C)
    ws.row_dimensions[HR2].height = 34
    key = lambda r: f'$B$3&"#"&$A{r}'
    for i in range(VD_N):
        r = VD_0 + i
        n = i + 1
        put(ws, f'A{r}', str(n), font=F_NOTE)
        ws[f'A{r}'] = n
        k = f'"{sh}#"&$A{r}'
        get = lambda ck: vfetch(k, ck)
        put(ws, f'B{r}', f'=IF({get("date")}="","",{get("date")})', font=F_LINK, fmt=DATE)
        put(ws, f'C{r}', f'=T({get("sname")})', font=F_LINK, align=CL)
        put(ws, f'D{r}', f'=T({get("memo")})', font=F_LINK, align=CL)
        put(ws, f'E{r}', f'=IF($B{r}="","",IFERROR(INDEX({QP}!$D${P0}:$D${P1},'
                         f'MATCH({get("proj")},{P_CODE},0)),""))', font=F_TXT)
        put(ws, f'F{r}', f'=T({get("inv")})', font=F_LINK)
        put(ws, f'G{r}', f'=T({get("payee")})', font=F_LINK)
        kd = get('kind'); am = f'N({get("amt")})'
        oc = get('ocol')
        put(ws, f'H{r}', f'=IF(AND({kd}="销项开票",{oc}<>"已开成本票"),{am},0)',
            font=F_LINK, fmt=MONEY)
        put(ws, f'I{r}', f'=N({get("mfee")})', font=F_LINK, fmt=MONEY)
        put(ws, f'J{r}', f'=N({get("due")})', font=F_LINK, fmt=MONEY)
        put(ws, f'K{r}', f'=IF({oc}="已开成本票",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'L{r}', f'=IF($B{r}="","",ROUND(SUM($J${VD_0}:$J{r})-SUM($K${VD_0}:$K{r}),2))',
            font=F_TXT, fmt=MONEY)
        for col, ck in (('M', 'vat'), ('N', 'add'), ('O', 'stamp'), ('P', 'inc')):
            put(ws, f'{col}{r}', f'=N({get(ck)})', font=F_LINK, fmt=MONEY)
        put(ws, f'Q{r}', f'=N({get("taxsum")})', font=F_TOT, fmt=MONEY)
        put(ws, f'R{r}', f'=IF({kd}="已交税",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'S{r}', f'=IF({kd}="挂靠代收",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'T{r}', f'=IF({kd}="我方收款",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'U{r}', f'=IF({kd}="管理费结算",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'V{r}', f'=IF({kd}="扣质保金",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'W{r}', f'=IF($B{r}="","","原表 {SRC_XLS[u].strip()} 表第 "&{vfetch(k,"srow")}&" 行"'
                         f'&IF({vfetch(k,"cnt")}="否","（与别的表同一张票，汇总只计一次）",""))',
            font=F_NOTE, align=CL)
        put(ws, f'X{r}', f'=IF($B{r}="","",IF(AND($B{r}>=$A$4,$B{r}<=$B$4),"是","否"))',
            font=F_TXT, fill=FILL_CHK)
        ws.row_dimensions[r].height = 16
    # 合计行：只统计落在期间内的行
    put(ws, f'A{VD_TOT}', '期间合计', font=F_TOT, fill=FILL_TOT)
    for c in 'BCDEFG':
        put(ws, f'{c}{VD_TOT}', None, font=F_TOT, fill=FILL_TOT)
    put(ws, f'D{VD_TOT}', f'=COUNTIFS($X${VD_0}:$X${VD_1},"是")&" 笔（本表共 "'
                          f'&COUNTIF($X${VD_0}:$X${VD_1},"是")+COUNTIF($X${VD_0}:$X${VD_1},"否")&" 笔）"',
        font=F_TOT, fill=FILL_TOT, align=CL)
    for c in VD_MONEY:
        if c == 'L':
            put(ws, f'L{VD_TOT}', f'=ROUND($J${VD_TOT}-$K${VD_TOT},2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
            continue
        put(ws, f'{c}{VD_TOT}', f'=ROUND(SUMIF($X${VD_0}:$X${VD_1},"是",{c}${VD_0}:{c}${VD_1}),2)',
            font=F_TOT, fill=FILL_TOT, fmt=MONEY)
    put(ws, f'W{VD_TOT}', '合计只算「在期间内＝是」的行', font=F_NOTE, fill=FILL_TOT, align=CL)
    put(ws, f'X{VD_TOT}', None, font=F_TOT, fill=FILL_TOT)
    ws.row_dimensions[VD_TOT].height = 20
    ws.conditional_formatting.add(f'A{VD_0}:{VD_LAST}{VD_1}',
        FormulaRule(formula=[f'$X{VD_0}="否"'], font=Font(color='A6A6A6')))
    # 没数据的行折起来，留 20 行空位
    nrow = sum(1 for e in EV if SHEET2UNIT.get(e['src'].split('!')[0].strip(),
                                                    e['src'].split('!')[0].strip()) == sh)
    for r in range(VD_0 + nrow + 20, VD_1 + 1):
        ws.row_dimensions[r].hidden = True
    ws.auto_filter.ref = f'A{HR2}:{VD_LAST}{VD_1}'
    ws.freeze_panes = f'C{VD_0}'
    page(ws, titles=f'{HR2}:{HR2}')
    ws.print_area = f'$A$1:${VD_LAST}${VD_1}'
    SUB_SHEETS.append(nm)
print(f'  ✓ {len(SUB_SHEETS)} 张单位竖版明细：{"、".join(SUB_SHEETS)}')

# ============================================================ 链条核算
ws = wb.create_sheet(SH_CHAIN)
title(ws, '分表 · 挂靠链条核算', 'R',
      '把一个项目的链条摊开：一级单位帮我们开了多少票给业主、扣完管理费二级该收多少票、二级又开了多少。'
      '「我方应开成本票」取的是这个项目所有链条最末一层的应开数 —— 同一个工程走了两三条挂靠链也不会算漏或算重。')
widths(ws, {'A':10,'B':36,'C':11,'D':12,'E':14,'F':9,'G':13,'H':14,'I':12,'J':14,'K':9,'L':13,
            'M':15,'N':14,'O':13,'P':14,'Q':11,'R':10})
DR = filter_band(ws, 'R')
headers(ws, HR, 1, ['项目编号','项目简称','业主','一级单位','一级\n开票额','一级\n费率','一级\n管理费','应开给一级',
                    '二级单位','二级\n开票额','二级\n费率','二级\n管理费','我方应开\n成本票','我方已开\n成本票',
                    '工资扣抵','未开差额','状态','有无业务'])
ANY = '"*"'
for i in range(P1 - P0 + 1):
    r, pr = Q0 + i, P0 + i
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$C{pr})', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($A{r}="","",{QP}!$F{pr})', font=F_LINK)
    put(ws, f'D{r}', f'=IF($A{r}="","",{QP}!$G{pr})', font=F_LINK)
    s1 = agg(FI, 'E', '销项开票', f'$D{r}', f'$A{r}', DR)
    m1 = agg(FR, 'E', '销项开票', f'$D{r}', f'$A{r}', DR)
    put(ws, f'E{r}', f'=IF($A{r}="","",{s1})', font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF(OR($A{r}="",$E{r}=0),"",$G{r}/$E{r})', font=F_TXT, fmt=PCT)
    put(ws, f'G{r}', f'=IF($A{r}="","",{m1})', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($A{r}="","",$E{r}-$G{r})', font=F_TXT, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($A{r}="","",{QP}!$H{pr})', font=F_LINK)
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
headers(ws, HR, 1, ['项目编号','项目简称','一级单位','劳务票\n应开','劳务票\n已开','劳务票\n工资扣抵',
                    '劳务票\n还差','机械票\n应开','机械票\n已开','机械票\n还差',
                    '其他票\n应开','其他票\n已开','其他票\n还差','合计\n还差','状态','有无业务'])
for i in range(P1 - P0 + 1):
    r, pr = Q0 + i, P0 + i
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$C{pr})', font=F_LINK, align=CL)
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
headers(ws, PR_H, 1, ['项目编号','项目简称','项目类型','其他应付\n发生','减代扣\n税费','已支付',
                      '应付余额','状态','该项目\n我方已交税','','','','有无往来'])
for i in range(P1 - P0 + 1):
    r, pr = PR_0 + i, P0 + i
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$C{pr})', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($A{r}="","",{QP}!$D{pr})', font=F_LINK)
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
PAY_CODES = {PCODE.get(x['proj'], '') for x in PASS_ROWS}
keep3 = {PR_0 + i for i, d in enumerate(PROJ_ROWS) if d['code'] in PAY_CODES or d['ptype'] == '合伙项目'}
keep3 |= set(range(max(keep3) + 1, min(max(keep3) + 1 + SPARE, PR_1 + 1))) if keep3 else set()
for r in range(PR_0, PR_1 + 1):
    if r not in keep3: ws.row_dimensions[r].hidden = True
ws.freeze_panes = f'C{Q0}'; page(ws, titles=f'{HR}:{HR}')
print('  ✓ 链条核算 / 发票缺口 / 往来台账')

# ============================================================ 项目利润（自营 / 合伙分开）
ws = wb.create_sheet(SH_PRF)
title(ws, '分表 · 按项目核算利润（自营项目 / 合伙项目分开小计）', 'R',
      '收入取「直接开给业主的那一层」开票额（一个工程走几条挂靠链就把几条的顶层加起来，不会重复计）；'
      '成本 ＝ 各层管理费 ＋ 应提税费 ＋ 日记账里记到这个项目名下的实际支出，德誉嘉返的 4% 现金冲减成本。'
      '合伙项目再减掉「应转合伙方」，剩下的才是归我方的。')
widths(ws, {'A':10,'B':32,'C':10,'D':11,'E':13,'F':13,'G':14,'H':13,'I':13,'J':11,'K':14,'L':9,
            'M':14,'N':14,'O':14,'P':14,'Q':11,'R':10})
DR = filter_band(ws, 'R')
ws.merge_cells('G4:J4'); ws.merge_cells('K4:L4'); ws.merge_cells('M4:N4'); ws.merge_cells('O4:P4')
headers(ws, HR, 1, ['项目编号','项目简称','项目类型','一级单位','合同金额','审计(结算)\n金额',
                    '业主端\n开票额','减:各层\n管理费','减:应提\n税费','加:返现',
                    '票面毛利','毛利率','减:项目\n实际支出','项目净利',
                    '减:应转他方\n(过账/合伙)','归属我方\n利润','状态','有无业务'])
PRF_0 = Q0 + 2                       # 6 全部合计 / 7 自营小计 / 8 合伙小计 / 9 起明细
PRF_1 = PRF_0 + (P1 - P0)
JJ2 = f'{QJ}!$J${J0}:$J${J1}'; JD2 = f'{QJ}!$D${J0}:$D${J1}'
JE2 = f'{QJ}!$E${J0}:$E${J1}'; JB2 = f'{QJ}!$B${J0}:$B${J1}'
DJ2 = f',{JB2},">="&{DR[0]},{JB2},"<="&{DR[1]}'
# 这些费用类型不算工程实际成本：要么已经在管理费/税费里算过，要么根本不是成本
NOT_COST = ['税费', '管理费', '借款', '还借款', '转备用金', '其他应付支付', '其他应收收回']
for i in range(P1 - P0 + 1):
    r, pr = PRF_0 + i, P0 + i
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$C{pr})', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($A{r}="","",{QP}!$D{pr})', font=F_TXT, fill=FILL_HDR2)
    put(ws, f'D{r}', f'=IF($A{r}="","",{QP}!$G{pr})', font=F_LINK)
    put(ws, f'E{r}', f'=IF($A{r}="","",N({QP}!$O{pr}))', font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",N({QP}!$P{pr}))', font=F_LINK, fmt=MONEY)
    put(ws, f'G{r}', f'=IF($A{r}="","",{agg(FI,None,"销项开票",None,f"$A{r}",DR,top=True)})',
        font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($A{r}="","",{agg(FR,None,"销项开票",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($A{r}="","",{agg(FZ,None,"销项开票",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",{agg(FAA,None,"销项开票",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'K{r}', f'=IF($A{r}="","",ROUND($G{r}-$H{r}-$I{r}+$J{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'L{r}', f'=IF(OR($A{r}="",$G{r}=0),"",$K{r}/$G{r})', font=F_TXT, fmt=PCT)
    cost = f'SUMIFS({JJ2},{JD2},$A{r}{DJ2})' + ''.join(
        f'-SUMIFS({JJ2},{JD2},$A{r},{JE2},"{t}"{DJ2})' for t in NOT_COST)
    put(ws, f'M{r}', f'=IF($A{r}="","",{cost})', font=F_LINK, fmt=MONEY)
    put(ws, f'N{r}', f'=IF($A{r}="","",ROUND($K{r}-$M{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'O{r}', f'=IF($A{r}="","",ROUND({agg(FI,None,"其他应付发生",None,f"$A{r}",DR)}'
                     f'-{agg(FI,None,"其他应付扣税",None,f"$A{r}",DR)},2))', font=F_LINK, fmt=MONEY)
    put(ws, f'P{r}', f'=IF($A{r}="","",ROUND($N{r}-$O{r},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'Q{r}', f'=IF($A{r}="","",IF(ROUND($G{r},2)=0,"未开票",'
                     f'IF(AND($O{r}>0,$O{r}>=$K{r}*0.5),"过账为主",'
                     f'IF($P{r}<0,"亏损",IF($L{r}<0.05,"薄利","正常")))))', font=F_TXT, fill=FILL_CHK)
    put(ws, f'R{r}', f'=IF($A{r}="","",IF(ROUND($G{r}+$H{r}+$M{r}+$O{r},2)=0,"无","有"))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16
MON = list('EFGHIJKMNOP')
for row, lab, cond in [(TR, '合  计', None), (TR + 1, '自营项目 小计', '自营项目'),
                       (TR + 2, '合伙项目 小计', '合伙项目')]:
    fill = FILL_TOT if cond is None else FILL_HDR2
    put(ws, f'A{row}', lab, font=F_TOT, fill=fill)
    for c in ['B', 'C', 'D', 'Q', 'R']: put(ws, f'{c}{row}', None, font=F_TOT, fill=fill)
    for c in MON:
        f = (f'=SUM({c}{PRF_0}:{c}{PRF_1})' if cond is None
             else f'=SUMIF($C${PRF_0}:$C${PRF_1},"{cond}",{c}${PRF_0}:{c}${PRF_1})')
        put(ws, f'{c}{row}', f, font=F_TOT, fill=fill, fmt=MONEY)
    put(ws, f'L{row}', f'=IF($G{row}=0,"",$K{row}/$G{row})', font=F_TOT, fill=fill, fmt=PCT)
    ws.row_dimensions[row].height = 20
put(ws, f'Q{TR}', f'=IF(COUNTA({QJ}!$D${J0}:$D${J1})=0,'
                  f'"⚠ 日记账还没有一行填项目编号，「项目实际支出」整列是 0",'
                  f'"日记账已有 "&COUNTA({QJ}!$D${J0}:$D${J1})&" 行填了项目编号")',
    font=F_NOTE, fill=FILL_TOT, align=CL)
put(ws, f'R{TR}', HIDE_NOTE, font=F_NOTE, fill=FILL_TOT, align=CL)
put(ws, f'B{TR+1}', '＝项目档案「项目类型」选自营项目的那些', font=F_NOTE, fill=FILL_HDR2, align=CL)
put(ws, f'B{TR+2}', '＝项目档案「项目类型」选合伙项目的那些', font=F_NOTE, fill=FILL_HDR2, align=CL)
ws.conditional_formatting.add(f'P{PRF_0}:P{PRF_1}',
    FormulaRule(formula=[f'AND($A{PRF_0}<>"",$P{PRF_0}<0)'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.conditional_formatting.add(f'C{PRF_0}:C{PRF_1}',
    FormulaRule(formula=[f'$C{PRF_0}="合伙项目"'], fill=FILL_IN, font=Font(color='9C6500', bold=True)))
hide_tail(ws, PRF_0, PRF_1, set(range(PRF_0, PRF_0 + NP)), col_idx('R') - 1, f'A{HR}:R{PRF_1}')
put(ws, f'A{PRF_1+2}',
    '口径：业主端开票额只取「对业主开票」那一层，多条挂靠链各自的顶层都算、中间层不重复计。'
    '应提税费是按单位档案参数逐笔算出来的（不管交没交）。项目实际支出取【资金日记账】里填了本项目编号的支出，'
    '已扣掉 ' + '、'.join(NOT_COST) + ' 这几类（要么已在管理费/税费里算过，要么不是工程成本）。'
    '「应转他方」＝【往来台账】里这个项目的其他应付发生额减去我方代垫税费；'
    '这一列吃掉大半毛利的项目状态标「过账为主」（比如迅驰过账的那 8 个机械费项目，'
    '钱本来就不是我们的，原表把该扣的 2,437.03 元税费记在单位层面没落到项目上，所以这几行会带个小负数）。　　'
    '注意：原日记账「项目」那一列填的是「铜梁项目」「光伏项目」这类工地/部门标注，不是项目编号，'
    '所以没有硬塞进项目编号列 —— 要按项目算工程成本，在【资金日记账】D 列把项目编号选上即可，这一列立刻出数。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{PRF_1+2}:R{PRF_1+2}')
ws.freeze_panes = f'C{PRF_0}'; page(ws, titles=f'{HR}:{HR}')
print('  ✓ 项目利润（自营 / 合伙分开小计）')

# ============================================================ 税费台账
ws = wb.create_sheet(SH_TAX)
title(ws, '分表 · 税费台账', 'K',
      '按单位汇总的税费。这一轮起，应提税费不再按单位档案的参数推算，而是【业务流水】右边'
      '「税费录入区」四列按实际逐笔录进去的；历史行照你那份《对账明细》9.13 版逐笔写死，'
      '所以这张表和【单位汇总】的欠税跟原表逐家一分不差。统计口径是「这一笔算在哪家单位的'
      '对账表上」（业务流水最右边隐藏的「归属单位表」列），不是按开票方 —— 因为原表就是一家一张表。'
      '上面可按年度或起止日期取数。')
widths(ws, {'A':13,'B':26,'C':15,'D':14,'E':14,'F':14,'G':15,'H':14,'I':15,'J':12,'K':10})
DR = filter_band(ws, 'K')
headers(ws, HR, 1, ['单位','单位全称','应提增值税','应提附加税','应提印花税','应提所得税','应提税费合计',
                    '已交税','欠税未交','状态','有无业务'])
for i in range(U1 - U0 + 1):
    r = Q0 + i
    put(ws, f'A{r}', f'=IFERROR(INDEX({U_NAME},MATCH({i+1},{U_RANK},0)),"")', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",INDEX({QU}!$B${U0}:$B${U1},MATCH($A{r},{U_NAME},0)))',
        font=F_LINK, align=CL)
    _dr = f',{FB},">="&{DR[0]},{FB},"<="&{DR[1]}'
    for col, val in [('C', FV), ('D', FW), ('E', FX), ('F', FY), ('G', FZ)]:
        put(ws, f'{col}{r}', f'=IF($A{r}="","",SUMIFS({val},{FRNG("ubel")},$A{r}{_dr}))',
            font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($A{r}="","",SUMIFS({FI},{KD},"已交税",{FRNG("ubel")},$A{r}{_dr}))',
        font=F_LINK, fmt=MONEY)
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
headers(ws, HR, 1, ['序号','日期','项目编号','项目简称','挂靠单位','类型','业主付给\n挂靠单位',
                    '挂靠单位\n转给我方','该单位\n在途余额','摘要'])
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BCDEFJ': put(ws, f'{c}{TR}', None, font=F_TOT, fill=FILL_TOT)
for c in ['G', 'H']:
    put(ws, f'{c}{TR}', f'=SUM({c}{D0}:{c}{D1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'I{TR}', f'=$G${TR}-$H${TR}', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
ws.row_dimensions[TR].height = 20
FAE = FRNG('dno')
for i in range(D1 - D0 + 1):
    r, n = D0 + i, i + 1
    src = lambda k: f'IFERROR(INDEX({FRNG(k)},MATCH({n},{FAE},0)),"")'
    put(ws, f'A{r}', f'=IF($B{r}="","",{n})', font=F_LINK)
    put(ws, f'B{r}', f'={src("date")}', font=F_LINK, fmt=DATE)
    put(ws, f'C{r}', f'={src("proj")}', font=F_LINK)
    put(ws, f'D{r}', f'={src("sname")}', font=F_LINK, align=CL)
    put(ws, f'E{r}', f'=IF($B{r}="","",IF({src("kind")}="挂靠代收",{src("payee")},{src("payer")}))',
        font=F_LINK)
    put(ws, f'F{r}', f'={src("kind")}', font=F_LINK)
    put(ws, f'G{r}', f'=IF($F{r}="挂靠代收",N({src("amt")}),0)', font=F_LINK, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($F{r}="我方收款",N({src("amt")}),0)', font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($B{r}="","",SUMIFS($G${D0}:$G{r},$E${D0}:$E{r},$E{r})'
                     f'-SUMIFS($H${D0}:$H{r},$E${D0}:$E{r},$E{r}))', font=F_TXT, fmt=MONEY)
    put(ws, f'J{r}', f'={src("memo")}', font=F_LINK, align=CL)
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
headers(ws, HR, 8, ['项目编号', '项目简称', '支出合计'])
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
    put(ws, f'I{r}', f'=IF($H{r}="","",{QP}!$C{pr})', font=F_LINK, align=CL)
    put(ws, f'J{r}', f'=IF($H{r}="","",SUMIFS({JJ},{JD},$H{r}{DJ}))', font=F_LINK, fmt=MONEY)
    if r > E_1: ws.row_dimensions[r].height = 16
for r in range(max(E_1, Q0 + NP + SPARE) + 1, QP_1 + 1):
    ws.row_dimensions[r].hidden = True
put(ws, f'A{QP_1+2}', HIDE_NOTE + '　｜　左表的费用类型来自【资金日记账】E 列下拉，要加类型先在那里加。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{QP_1+2}:J{QP_1+2}')
ws.freeze_panes = f'A{Q0}'; page(ws, titles=f'{HR}:{HR}')
print('  ✓ 税费台账 / 代收台账 / 费用统计')


# ============================================================ 对账差异说明
# 逐家逐列拿你那份《对账明细》9.13 版的合计行跟系统算出来的数比，差在哪、为什么差，一张表说清楚
ws = wb.create_sheet(SH_DIFF)
title(ws, '对账差异说明（系统 ↔ 原《对账明细》9.13 版）', 'G')
put(ws, 'A2', '左边是你原表每张表最上面那行合计，右边是系统按「来源表＝这张表」算出来的同口径数。'
              '差额绝对值小于 0.05 元的算对上了（原表本身有四舍五入尾差）。'
              '真正对不上的只有一处，下面第二块逐笔列了出来。这张表不参与任何计算，纯粹给你核对用。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells('A2:G2')
widths(ws, {'A': 12, 'B': 18, 'C': 16, 'D': 16, 'E': 13, 'F': 12, 'G': 60})
DF_H = 4
headers(ws, DF_H, 1, ['单位', '对账项目', '原表合计', '系统合计', '差额', '结论', '说明'])
SRC_OF = dict(SRC_XLS)                 # 单位简称 → 原工作簿里的表名
SRC_KEY = {k: k for k in SRC_OF}       # 业务流水「来源表」列存的是单位简称
def dsum(valkey, sh, kind=None, ocol=None):
    ex = ''
    if kind: ex += f',{FRNG("kind")},"{kind}"'
    if ocol: ex += f',{FRNG("ocol")},"{ocol}"'
    return f'SUMIFS({FRNG(valkey)},{FRNG("src")},"{sh}"{ex})'
DF_ITEMS = [
    ('销售开票金额',     lambda sh: dsum('amt', sh, '销项开票', '销售开票金额')),
    ('应扣管理费',       lambda sh: dsum('mfee', sh, '销项开票', '销售开票金额')),
    ('应到成本票',       lambda sh: dsum('due', sh, '销项开票', '销售开票金额')),
    ('已到成本票',       lambda sh: dsum('amt', sh, None, '已开成本票')),
    ('应扣税费',         lambda sh: dsum('taxsum', sh)),
    ('已交税',           lambda sh: dsum('amt', sh, '已交税')),
    ('业主付给挂靠单位', lambda sh: dsum('amt', sh, '挂靠代收')),
    ('挂靠单位转我方',   lambda sh: dsum('amt', sh, '我方收款')),
    ('管理费结算',       lambda sh: dsum('amt', sh, '管理费结算')),
    ('扣质保金',         lambda sh: dsum('amt', sh, '扣质保金')),
]
# 原表合计行逐列的列号（8 张表各不相同，逐张核对过）
ORIG_COL = {
 '德誉嘉 ': dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10,
                业主付给挂靠单位=13, 挂靠单位转我方=17, 管理费结算=16),
 '迅驰':   dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=16, 已交税=17,
                业主付给挂靠单位=20, 挂靠单位转我方=25, 管理费结算=23, 扣质保金=24),
 '华城':   dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=16, 已交税=17,
                业主付给挂靠单位=20, 挂靠单位转我方=24, 管理费结算=23),
 '金沁':   dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=16, 已交税=17,
                业主付给挂靠单位=21, 挂靠单位转我方=26, 管理费结算=25, 扣质保金=24),
 '湖南锦泰': dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=17, 已交税=18,
                业主付给挂靠单位=21, 挂靠单位转我方=26, 管理费结算=24, 扣质保金=25),
 '康欣':   dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=16, 已交税=17,
                业主付给挂靠单位=20, 挂靠单位转我方=25, 管理费结算=24, 扣质保金=23),
 '安锐':   dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=17, 已交税=18,
                业主付给挂靠单位=21, 挂靠单位转我方=26, 管理费结算=24, 扣质保金=25),
 '杰华电气': dict(销售开票金额=7, 应扣管理费=9, 应到成本票=10, 已到成本票=11, 应扣税费=18, 已交税=19,
                业主付给挂靠单位=22, 挂靠单位转我方=27, 管理费结算=25, 扣质保金=26),
}
import openpyxl as _op
_ref = _op.load_workbook(os.path.join(HERE, '..', '参考', '原对账明细_9.13.xlsx'), data_only=True)
NOTE_OK = '对上了'
r = DF_H + 1
DF_0 = r
for u in [x[0] for x in UNITS if x[2] == '挂靠单位']:
    sh = SRC_OF[u]
    _ws = _ref[sh]
    for lab, fn in DF_ITEMS:
        col = ORIG_COL[sh].get(lab)
        ov = _ws.cell(row=4, column=col).value if col else None
        ov = round(float(ov), 2) if isinstance(ov, (int, float)) else None
        sysf = fn(SRC_KEY[u])
        if ov is None:
            # 原表这张表没有这一列：只有系统算出非零才列出来
            put(ws, f'A{r}', u, font=F_TXT); put(ws, f'B{r}', lab, font=F_TXT, align=CL)
            put(ws, f'C{r}', '原表无此列', font=F_NOTE)
            put(ws, f'D{r}', f'={sysf}', font=F_LINK, fmt=MONEY)
            put(ws, f'E{r}', '—', font=F_NOTE)
            put(ws, f'F{r}', f'=IF(ROUND(N($D{r}),2)=0,"两边都没有","原表没这一列")',
                font=F_TXT, fill=FILL_CHK)
            put(ws, f'G{r}', f'原表「{sh.strip()}」这张表没有这一列', font=F_NOTE, align=CL)
            r += 1; continue
        put(ws, f'A{r}', u, font=F_TXT)
        put(ws, f'B{r}', lab, font=F_TXT, align=CL)
        put(ws, f'C{r}', ov, font=F_TOT, fmt=MONEY)
        put(ws, f'D{r}', f'={sysf}', font=F_LINK, fmt=MONEY)
        put(ws, f'E{r}', f'=ROUND(N($D{r})-N($C{r}),2)', font=F_TOT, fmt=MONEY)
        put(ws, f'F{r}', f'=IF(ABS(N($E{r}))<0.05,"✓ 对上","✗ 有差异")', font=F_TXT, fill=FILL_CHK)
        if u == '康欣' and lab == '已到成本票':
            put(ws, f'G{r}',
                '原表把康欣开给金沁的 7 张票在同一行的「销售开票金额」和「已提供成本票」两列各记了一次；'
                '系统一张票只记一次，这 7 张已经在金沁表里作为金沁收到的成本票统计过了。逐笔见下方第二块。',
                font=F_NOTE, align=CL)
        else:
            put(ws, f'G{r}', f'=IF(ABS(N($E{r}))<0.05,"{NOTE_OK}（尾差是原表四舍五入）","请点开【业务流水】按来源表筛一下")',
                font=F_NOTE, align=CL)
        ws.row_dimensions[r].height = 16
        r += 1
DF_1 = r - 1
ws.conditional_formatting.add(f'A{DF_0}:G{DF_1}',
    FormulaRule(formula=[f'$F{DF_0}="✗ 有差异"'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))

# ---- 第二块：造成差异的逐笔清单 ----
DUP = [x for x in RAW_ROWS
       if x['sheet'] == '康欣' and abs(x['cost_done']) > 0.004
       and abs(x['cost_done'] - x['sale']) <= 0.004]
r2 = DF_1 + 3
put(ws, f'A{r2}', '造成差异的逐笔清单 · 康欣「已提供成本票」', font=F_H2, fill=FILL_HDR2)
ws.merge_cells(f'A{r2}:G{r2}')
r2 += 1
headers(ws, r2, 1, ['原表行', '项目', '摘要', '销售开票金额', '已提供成本票', '重复金额', '处理方式'])
hr2 = r2
r2 += 1
for x in DUP:
    put(ws, f'A{r2}', f"康欣!{x['row']}", font=F_TXT)
    put(ws, f'B{r2}', PSHORT.get(PCODE.get(x['proj'], ''), ''), font=F_TXT, align=CL)
    put(ws, f'C{r2}', x['memo'][:60], font=F_TXT, align=CL)
    put(ws, f'D{r2}', x['sale'], font=F_LINK, fmt=MONEY)
    put(ws, f'E{r2}', x['cost_done'], font=F_LINK, fmt=MONEY)
    put(ws, f'F{r2}', x['cost_done'], font=F_TOT, fmt=MONEY)
    put(ws, f'G{r2}', '系统只记一次：在【金沁明细】里作为金沁收到的成本票', font=F_NOTE, align=CL)
    ws.row_dimensions[r2].height = 16
    r2 += 1
put(ws, f'A{r2}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'BC': put(ws, f'{c}{r2}', None, font=F_TOT, fill=FILL_TOT)
for c in 'DEF':
    put(ws, f'{c}{r2}', f'=SUM({c}{hr2+1}:{c}{r2-1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'G{r2}', '＝上面康欣「已到成本票」那一行的差额', font=F_NOTE, fill=FILL_TOT, align=CL)
r2 += 2
put(ws, f'A{r2}', '本次新增维度核对', font=F_H2, fill=FILL_HDR2)
ws.merge_cells(f'A{r2}:G{r2}')
r2 += 1
headers(ws, r2, 1, ['单位', '核对项目', '原表合计', '系统合计', '差额', '结论', '说明'])
hr3 = r2
r2 += 1
EXTRA_CHK = [
    ('德誉嘉', '返管理费 4%（其他应收）', round(sum(e.get('rebate', 0) for e in EV), 2),
     f'=SUMIFS({FAA},{KD},"销项开票",{KE},"德誉嘉",{KZ},"是")', '原表「返管理费4%」那一列的合计'),
    ('迅驰', '过账应付工程款', round(sum(x['amt'] for x in PASS_ROWS), 2),
     f'=SUMIFS({FI},{KD},"其他应付发生",{KF},"迅驰",{KZ},"是")', '原表迅驰「过账应付工程款」列'),
    ('迅驰', '过账应扣税费', round(sum(x['amt'] for x in DED_ROWS), 2),
     f'=SUMIFS({FI},{KD},"其他应付扣税",{KF},"迅驰",{KZ},"是")', '原表迅驰「已付款」列'),
    ('康欣', '工资扣抵（劳务成本）', round(sum(x['amt'] for x in WAGE_ROWS), 2),
     f'=SUMIFS({FI},{KD},"工资扣抵",{KZ},"是")', '原总台账「劳务成本·工资扣抵」列'),
]
for u, lab, ov, f, note in EXTRA_CHK:
    put(ws, f'A{r2}', u, font=F_TXT)
    put(ws, f'B{r2}', lab, font=F_TXT, align=CL)
    put(ws, f'C{r2}', ov, font=F_TOT, fmt=MONEY)
    put(ws, f'D{r2}', f, font=F_LINK, fmt=MONEY)
    put(ws, f'E{r2}', f'=ROUND(N($D{r2})-N($C{r2}),2)', font=F_TOT, fmt=MONEY)
    put(ws, f'F{r2}', f'=IF(ABS(N($E{r2}))<0.05,"✓ 对上","✗ 有差异")', font=F_TXT, fill=FILL_CHK)
    put(ws, f'G{r2}', note, font=F_NOTE, align=CL)
    ws.row_dimensions[r2].height = 16
    r2 += 1
EX_1 = r2 - 1
ws.conditional_formatting.add(f'A{hr3+1}:G{EX_1}',
    FormulaRule(formula=[f'$F{hr3+1}="✗ 有差异"'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
r2 += 1
put(ws, f'A{r2}', '总体结论', font=F_TOT, fill=FILL_TOT)
put(ws, f'B{r2}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'C{r2}', f'=COUNTIF($F${DF_0}:$F${DF_1},"✓*")+COUNTIF($F${hr3+1}:$F${EX_1},"✓*")&" 项对上"',
    font=F_TOT, fill=FILL_TOT)
put(ws, f'D{r2}',
    f'=IF(COUNTIF($F${DF_0}:$F${DF_1},"✗*")+COUNTIF($F${hr3+1}:$F${EX_1},"✗*")=0,"✓ 与原表完全一致",'
    f'"有 "&COUNTIF($F${DF_0}:$F${DF_1},"✗*")+COUNTIF($F${hr3+1}:$F${EX_1},"✗*")'
    f'&" 项差异（已在下方逐笔说明）")', font=F_TOT, fill=FILL_CHK)
ws.merge_cells(f'D{r2}:G{r2}')
for c in 'EFG': put(ws, f'{c}{r2}', None, font=F_TOT, fill=FILL_CHK)
DIF_SUM_R = r2
r2 += 2
put(ws, f'A{r2}', '另外说明：本轮把「应提税费」从按单位档案参数预提改成按原表逐笔录入，'
                  '所以【单位汇总】的欠税一栏现在跟原表完全一致 —— 迅驰 0（不欠税）、'
                  '华城 10,129.25、金沁 17,778.54、康欣 16,304.66、湖南锦泰 −2,610.25（多交）、'
                  '安锐 0、杰华 15,759.04、德誉嘉 原表没有税费列所以是 0。'
                  '以前是按「一家单位一个税率」推算的，华城同时开 3% 劳务票和 13% 设备票，'
                  '一个率算不出两种票，这就是原来对不上的根子。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{r2}:G{r2}')
ws.freeze_panes = f'C{DF_0}'
page(ws, titles=f'{DF_H}:{DF_H}', landscape=False)
print(f'  ✓ 对账差异说明（{DF_1-DF_0+1} 项逐列核对，{len(DUP)} 笔逐笔差异）')

# ============================================================ 首页
QSU, QSX, QTAX, QDIF, QCH = Q(SH_SUM_U), Q(SH_SUM_X), Q(SH_TAX), Q(SH_DIFF), Q(SH_CHAIN)
QGAP, QCUR, QPRF = Q(SH_GAP), Q(SH_CUR), Q(SH_PRF)
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
 ('我方已开成本票',   f'={QSU}!$H${TR}', 'D6E4F0'),
 ('还差成本票未开',   f'={QGAP}!$N${TR}', 'FCE4E4'),
 ('劳务票还差',       f'={QGAP}!$G${TR}', 'FCE4E4'),
 ('机械票还差',       f'={QGAP}!$J${TR}', 'FCE4E4'),
 ('其他应收·待返现',  f'={QCUR}!$E${TR}', 'FFF2CC'),
 ('其他应付·待转付',  f'={QCUR}!$K${TR}', 'FCE4E4'),
 ('应收工程款',       f'={QSU}!$C${TR}', 'E2EFDA'),
 ('业主已付挂靠单位', f'={QSU}!$M${TR}', 'E2EFDA'),
 ('挂靠单位已转我方', f'={QSU}!$O${TR}', 'E2EFDA'),
 ('在途资金·代收未转', f'={QSU}!$P${TR}', 'FCE4E4'),
 ('自营项目 归属利润', f'={QPRF}!$P${TR+1}', 'E2EFDA'),
 ('合伙项目 归属利润', f'={QPRF}!$P${TR+2}', 'FFF2CC'),
 ('全部项目 票面毛利', f'={QPRF}!$K${TR}', 'D6E4F0'),
 ('项目实际支出',     f'={QPRF}!$M${TR}', 'FCE4E4'),
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
    ('对原表核对', f'={QDIF}!$D${DIF_SUM_R}'),
    ('业务流水校验', f'=IF(COUNT({QF}!$B${F0}:$B${F1})-COUNTIF({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1},"√")'
                    f'-COUNTIF({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1},"链条待确认")=0,'
                    f'"✓ 全部通过"&IF(COUNTIF({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1},"链条待确认")>0,'
                    f'"（另有 "&COUNTIF({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1},"链条待确认")&" 行是多链条项目，仅提示不影响汇总）",""),'
                    f'"✗ 有 "&(COUNT({QF}!$B${F0}:$B${F1})-COUNTIF({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1},"√")'
                    f'-COUNTIF({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1},"链条待确认"))&" 行待修正")'),
    ('日记账校验',   f'=IF(COUNTIF({QJ}!$N${J0}:$N${J1},"√")=COUNT({QJ}!$B${J0}:$B${J1}),"✓ 全部通过",'
                    f'"✗ 有 "&(COUNT({QJ}!$B${J0}:$B${J1})-COUNTIF({QJ}!$N${J0}:$N${J1},"√"))&" 行待修正")'),
    ('项目档案校验', f'=IF(COUNTA({QP}!$A${P0}:$A${P1})-COUNTIF({QP}!$U${P0}:$U${P1},"√")'
                    f'-COUNTIF({QP}!$U${P0}:$U${P1},"待完善*")=0,'
                    f'"✓ 全部通过"&IF(COUNTIF({QP}!$U${P0}:$U${P1},"待完善*")>0,'
                    f'"（另有 "&COUNTIF({QP}!$U${P0}:$U${P1},"待完善*")&" 个项目的业主或一级单位还没填）",""),'
                    f'"✗ 有 "&(COUNTA({QP}!$A${P0}:$A${P1})-COUNTIF({QP}!$U${P0}:$U${P1},"√")'
                    f'-COUNTIF({QP}!$U${P0}:$U${P1},"待完善*"))&" 行待修正")'),
    ('成本票缺口',   f'=IF({QGAP}!$N${TR}<1,"✓ 已开齐","还有 "&TEXT({QGAP}!$N${TR},"#,##0")&" 元成本票没开"'
                    f'&"（劳务 "&TEXT({QGAP}!$G${TR},"#,##0")&"，机械 "&TEXT({QGAP}!$J${TR},"#,##0")&"）")'),
    ('项目利润',     f'="全部 "&TEXT({QPRF}!$P${TR},"#,##0")&" 元（毛利率 "'
                    f'&TEXT({QPRF}!$L${TR},"0.0%")&"）；其中自营 "&TEXT({QPRF}!$P${TR+1},"#,##0")'
                    f'&" 元、合伙 "&TEXT({QPRF}!$P${TR+2},"#,##0")&" 元"'
                    f'&IF(COUNTIF({QPRF}!$Q${TR+3}:$Q${PRF_1},"亏损")>0,'
                    f'"；有 "&COUNTIF({QPRF}!$Q${TR+3}:$Q${PRF_1},"亏损")&" 个项目算下来是亏的","")'),
    ('表格容量',     f'=IF(AND(COUNT({QF}!$B${F0}:$B${F1})<{int((F1-F0+1)*0.9)},'
                    f'COUNT({QJ}!$B${J0}:$B${J1})<{int((J1-J0+1)*0.9)},'
                    f'COUNTA({QP}!$A${P0}:$A${P1})<{int((P1-P0+1)*0.9)}),'
                    f'"✓ 够用（业务流水 "&COUNT({QF}!$B${F0}:$B${F1})&"/{F1-F0+1} 行，'
                    f'日记账 "&COUNT({QJ}!$B${J0}:$B${J1})&"/{J1-J0+1} 行，'
                    f'项目 "&COUNTA({QP}!$A${P0}:$A${P1})&"/{P1-P0+1} 个）",'
                    f'"⚠ 快满了：业务流水 "&COUNT({QF}!$B${F0}:$B${F1})&"/{F1-F0+1}，'
                    f'日记账 "&COUNT({QJ}!$B${J0}:$B${J1})&"/{J1-J0+1}，'
                    f'项目 "&COUNTA({QP}!$A${P0}:$A${P1})&"/{P1-P0+1}，请让人把容量调大")'),
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
                    '我方主体，再填合同金额和审计金额。别忘了填「项目简称」，后面所有表显示的都是简称。'),
 ('挂靠单位开票出去', '【业务流水】录一行：业务类型「销项开票」，开票方＝挂靠单位，收票方＝业主，票据类型选劳务票还是机械设备票。'
                     '再填这一笔的管理费率、返现率，选一下「计费方式」，右边税费四列按实际填。'
                     '管理费、应开成本票自动算。'),
 ('我们开成本票过去', '【业务流水】业务类型「成本票」，开票方＝我方主体，收票方＝挂靠单位，票据类型要和对应的销项票一致。'),
 ('用工资表顶劳务票', '【业务流水】业务类型「工资扣抵」，填项目编号和金额 —— 【发票缺口】里的劳务票应开就会相应减少。'),
 ('业主付钱给挂靠单位', '【业务流水】选「挂靠代收」，付款方＝业主，收款方＝挂靠单位。这笔钱还没到我们手上。'),
 ('挂靠单位转钱给我们', '【业务流水】选「我方收款」；同时到【资金日记账】记实际到账，「工程回款」选是。'),
 ('德誉嘉返 4% 现金', '开票那一行的「返现率」自动是 4%，钱一直挂在【往来台账】其他应收款上；'
                     '真收到现金时录一行「其他应收收回」冲掉。'),
 ('过账款 / 合伙分钱', '收到别人的过账款录「其他应付发生」；我方为这笔垫的税费录「其他应付扣税」；真转出去录「其他应付支付」。'
                      '【往来台账】③ 按项目算出「扣完税费后还该转给别人多少」。'),
 ('交税', '【业务流水】选「已交税」记实际交的；该提多少税在开票那一行右边「税费录入区」四列按实际填。'),
 ('日常收付款', '【资金日记账】三个账户混着录，跨年度继续往下录就行，余额自动算。'
             '想一个账户一个账户分开录的，用【日记账-泓普】【日记账-仟茂】【日记账-现金】三张拆分表，'
             '录完把右边接口区整块复制、选择性粘贴数值到总表。'),
 ('要看结果', '【单位汇总】【项目汇总】【项目利润】【链条核算】【发票缺口】【往来台账】【税费台账】【费用统计】'
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
 ('不用每年建新表', '业务流水和资金日记账各留了 2000 行，项目档案 300 个，够用四五年，按日期一直往下录就行。'
                  '首页「表格容量」那一行会盯着，用到九成会提醒。'
                  '每张查询表最上面都有「年度 / 起始日期 / 截止日期」：填 2026 就只看 2026 年，'
                  '填起止日期就只看那一段，三个都留空＝全部期间。'),
 ('空白行会折叠', '查询表只显示有数据的行，空白行已经隐藏起来。新增项目或换了查询区间之后，'
                 '点一下 数据 → 筛选 → 重新应用，行数就会跟着变。录入表（单位档案 / 项目档案 / 业务流水 / 资金日记账）'
                 '永远全部显示，方便往下录。'),
 ('合计在表头下面', '每张查询表的合计行就在表头正下方第 6 行，不用翻到底，筛选也不会影响它。'),
 ('管理费怎么定（这一轮改了）', None),
 ('费率改成逐笔手填', '【单位档案】的「默认管理费率 / 默认返现率 / 第二档」五列已经取消。'
                    '现在管理费率和返现率直接在【业务流水】那两个淡黄色格子里按这一笔的实际情况填，'
                    '后面所有表一律从业务流水取数。这样就不会再出现「档案填一个率、原表实际是另一个率」对不上的情况。'
                    '历史 385 行已按你那份《对账明细》9.13 版逐笔写死，一分没动。'),
 ('计费方式三选一', '【业务流水】新增「计费方式」列：'
                  '① 扣管理费 —— 正常业务，管理费＝金额×费率，应开成本票＝金额−管理费；'
                  '② 不扣管理费 —— 只借通道过票、一分不收（康欣老项目那种），应开成本票＝全额；'
                  '③ 不回成本票 —— 合伙方内部分成、根本不用我方回票，应开成本票＝0。'
                  '以前「费率填 0」和「忘了填」长得一模一样，现在分得清了，忘选会在校验列提示。'),
 ('应到成本票拆两列', '【项目汇总】【单位汇总】的「应到成本票」右边跟着两列：'
                    '「其中·扣了管理费的」和「其中·没扣管理费的」，只过票不收费那部分一眼看得见，'
                    '不用再一笔笔翻。没扣管理费那一列大于 0 会标黄。'),
 ('税费也改成按实际录', '【业务流水】右边「税费录入区」四列（预提增值税 / 附加 / 印花 / 所得）改成手工填，'
                      '历史行按你原表逐笔写死。以前是按【单位档案】「一家单位一个税率」推算的，'
                      '可华城同时开 3% 劳务票和 13% 设备票，一个率算不出两种票 —— 这就是原来欠税对不上的根子。'
                      '现在【单位汇总】的欠税跟原表完全一致：迅驰 0（不欠税）、华城 10,129.25、金沁 17,778.54、'
                      '康欣 16,304.66、湖南锦泰 −2,610.25（多交了）、安锐 0、杰华 15,759.04。'
                      '最右边「税费参考试算」那一列还是按单位档案的税率参数算，只给录入时对照，不参与任何汇总。'),
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
 ('项目赚不赚钱', None),
 ('项目利润表', '【项目利润】一个项目一行：收入取「直接开给业主的那一层」开票额（一个工程走几条挂靠链，'
               '几条的顶层加起来就是这个工程的合同额，中间层不会重复计）；成本 ＝ 各层管理费 ＋ 应提税费 ＋ '
               '【资金日记账】里记到这个项目名下的实际支出；德誉嘉返的 4% 现金冲减成本。'),
 ('自营和合伙分开', '表头下面三行合计：全部合计、自营项目小计、合伙项目小计，一眼看出两类各赚多少。'
                  '合伙项目那一列底色是黄的，也可以直接按「项目类型」筛。'
                  '合伙项目还要再减一道「应转他方」（＝往来台账里这个项目的其他应付发生额减我方代垫税费），'
                  '减完才是归我方的利润。状态列标「过账为主」的，是钱本来就不属于我们的过账项目。'),
 ('实际支出要先挂项目', '原日记账「项目」列填的是「铜梁项目」「光伏项目」这类工地/部门标注，不是项目编号，'
                      '导入时没有硬塞进项目编号列（塞了会和项目档案对不上）。'
                      '所以【项目利润】的「项目实际支出」现在是 0，'
                      '在【资金日记账】D 列把项目编号选上，这一列就跟着出数。'),
 ('哪些不算工程成本', '日记账里记到项目名下、但费用类型是税费 / 管理费 / 借款 / 还借款 / 转备用金 / '
                    '其他应付支付 / 其他应收收回的，不计入项目实际支出 —— 要么已经在管理费和税费里算过一次，'
                    '要么根本不是工程成本。'),
 ('给领导的单位专表（这一轮改了）', None),
 ('一家一张 · 改成竖版逐笔', '8 张单位专表（德誉嘉明细 / 迅驰明细 …）表头照你那份《对账明细》的样子做，'
                          '原来「一个项目一行」的横向汇总改成「一笔业务一行」按日期竖着排，'
                          '一行一行能跟原表对得上。明细全部显示不折叠；'
                          '上面填年度或起止日期，第 6 行的「期间合计」只统计落在期间内的行'
                          '（最后一列会标是/否，不在期间内的行显示成灰色）。'),
 ('数据从哪来', '每张单位专表取的是【业务流水】里「来源表＝这家单位」的行 —— '
               '也就是原来在这家单位那张对账表上记过的每一笔。在业务流水改一笔，这里立刻跟着变。'),
 ('对账差异说明', '最后一张【对账差异说明】把 8 家单位 × 10 个金额列共 80 项，'
                '逐项拿原表合计行和系统算出来的数比一遍，差额小于 0.05 元的算对上（原表本身有四舍五入尾差）。'
                '目前只有一处真差异：康欣「已提供成本票」差 356,218.09 元 —— '
                '原表把康欣开给金沁的 7 张票在同一行的两列里各记了一次，系统一张票只记一次，'
                '那 7 张已经在金沁表里作为金沁收到的成本票统计过。这 7 笔在差异表第二块逐笔列了出来。'),
 ('项目简称', '【项目档案】新增「项目简称」列，简称是从原表摘要「…：平双线」冒号后面那段取的；'
             '重名的自动带上第一次出现的那家单位，比如「平双线(德誉嘉)」「平双线(华城)」。'
             '业务流水和所有汇总表显示的都是简称，不再刷屏显示几十个字的全称。'
             '项目档案的行序＝项目在原对账明细里第一次出现的先后，对原表可以一行对一行往下走。'),
 ('日记账拆分表', '【日记账-泓普】【日记账-仟茂】【日记账-现金】三张单账户录入表，'
                '谁管哪个账户就录哪一张，左边照常填、L 列自动滚本账户余额。'
                '录完把右边灰色的「复制到总表」整块复制，到【资金日记账】里选择性粘贴「数值」：'
                '主块 9 列粘到总表 B 列那一格，右边单独那列「工程回款」粘到总表 M 列。'
                '拆分表只是录入草稿，所有汇总一律只认总表，不会重复算。'),
 ('钱在谁手上', None),
 ('三个状态', '业主还没付 →【单位汇总】的「业主未付」；业主付了但挂靠单位压着 →「挂靠单位代收未转」；'
             '已经到我们账上 →「挂靠单位已转我方」。逐笔看【代收台账】。'),
 ('注意', None),
 ('不要改灰色区', '淡黄色是手工录入，灰色是自动算的。灰色列被覆盖后不会报错，但分表会静默算错。'),
 ('校验列必须全是√', '【业务流水】【资金日记账】【项目档案】的校验列出现红色，说明这一行有问题，要改掉。'),
 ('历史数据', '原来八张对账表的业务、日记账 34 笔都已导入，德誉嘉返 4%、迅驰过账应付和应扣税费也一并进来了。'
             '【对账差异说明】把原表 8 张表的合计行和系统重算结果逐项比对，80 项里只有 1 项是真差异。'),
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
        SH_SUM_X:'548235', SH_CHAIN:'548235', SH_GAP:'7030A0', SH_CUR:'7030A0', SH_PRF:'BF8F00',
        SH_TAX:'548235', SH_EXP:'2F5597'}
for n in SUB_SHEETS: TABC[n] = 'A9D08E'
for n in SPLIT_SHEETS: TABC[n] = 'F4B183'
TABC[SH_DIFF] = 'C00000'
for n, c in TABC.items(): wb[n].sheet_properties.tabColor = c
ORDER2 = ([SH_HOME2, SH_DOC, SH_UNIT, SH_PROJ, SH_FLOW, SH_JOUR] + SPLIT_SHEETS +
          [SH_SUM_U, SH_SUM_P, SH_PRF, SH_SUM_X] + SUB_SHEETS +
          [SH_CHAIN, SH_GAP, SH_CUR, SH_TAX, SH_DAI, SH_EXP, SH_DIFF])
wb._sheets = [wb[n] for n in ORDER2]; wb.active = 0
OUT = os.environ.get('GK_OUT', os.path.join(HERE, '..', '建筑挂靠业务核算系统.xlsx'))
wb.save(OUT)
nf = sum(1 for s in wb.worksheets for row in s.iter_rows() for c in row
         if isinstance(c.value, str) and c.value.startswith('='))
print(f'\n已保存：{os.path.abspath(OUT)}')
print(f'工作表 {len(wb.worksheets)} 张，公式 {nf} 个，自动配平括号 {BAL_FIXED[0]} 处')
