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

# ---------------- 9.24 客户回传版：新项目、新日记账、新录的业务流水以它为准 ----------------
from openpyxl import load_workbook as _lw
BACK = _lw(os.path.join(HERE, '..', '参考', '客户回传_9.24.xlsx'))

# 资金日记账：9.24 起以回传版「资金日记帐-新表」为准（出纳那本账复制过来的）
_wj = BACK['资金日记帐-新表']
JOUR_ACCTS = [v for v in (_wj[f'O{r}'].value for r in range(6, 40)) if v]
JOUR_ROWS = []
for _r in range(6, _wj.max_row + 1):
    _row = {c: _wj[f'{c}{_r}'].value for c in 'BCDEFGHIJKLM'}
    if all(v in (None, '') for v in _row.values()): continue
    JOUR_ROWS.append({c: v for c, v in _row.items() if v not in (None, '')})
JOUR_YEAR = _wj['B5'].value.replace('年', '') if _wj['B5'].value else '2026'

# 康欣原表「已提供成本票」列里有 7 笔，系统上一版没收进来（康欣已到成本票比原表少 356,218.09）：
# 6 笔是跟康欣约定「未开票按 3% 扣点子（不提供成本票）」的（原表 H68 那行备注列的就是这几笔），
# 1 笔是 2025.12.08 大太老旧路改造 203,652 那张。按原表照样记成「成本票」，发票性质填「不开票」，
# 单位明细里这一类会标红，一眼能分出来哪些是真收到的票、哪些是按 3% 扣点抵掉的
KX_3PCT = {37, 38, 54, 55, 56, 57}
for _e in list(EV):
    if _e['kind'] != '销项开票' or not _e['src'].startswith('康欣!'): continue
    _n = int(_e['src'].split('!')[1])
    if _n not in KX_3PCT | {66}: continue
    EV.append(dict(date=_e['date'], unit='康欣', proj=_e['proj'], src=_e['src'], kind='成本票',
                   payer='泓普', payee='康欣', amt=_e['amt'], inv='不开票' if _n in KX_3PCT else '',
                   itype=_e['itype'], from_cost_col=True, count_in='是', rebate=0.0, mfee=0, cost_due=0,
                   memo=_e['memo'] + ('｜未开票按3%扣点子（不提供成本票）' if _n in KX_3PCT
                                      else '｜原表记在「已提供成本票」列'),
                   note9='9.24 补录·原表康欣表第 %d 行「已提供成本票」列' % _n))
# 照原表回款区补 4 类数据（逐行对过原表，差异说明里都有交代）
def _ev(src, kind='销项开票'):
    hit = [e for e in EV if e['src'] == src and e['kind'] == kind]
    assert len(hit) == 1, (src, kind, len(hit))
    return hit[0]
def _add(src, date, kind, payer, payee, amt, memo, note9, proj):
    EV.append(dict(date=date, unit='', proj=proj, src=src, kind=kind, payer=payer, payee=payee, amt=amt,
                   inv='', itype='', count_in='是', rebate=0.0, mfee=0, cost_due=0, memo=memo, note9=note9))
# ① 迅驰原表第 69 行：应退税费 373.3 在「迅驰开票金额」列手填了一笔应收、在「已到账」列记了一笔收回，一进一出。
#    系统上一版只收了「已到账」那一半，应收挂靠方余额因此少了 373.30
_x69 = _ev('迅驰!69', '我方收款')
_add('迅驰!69', '2026-07-23', '代垫应收', '迅驰', '泓普', 373.3, _x69['memo'],
     '9.24 补录·原表迅驰表第 69 行「迅驰开票金额」列手填 373.3（应退税费；同一行「已到账」373.3 已在系统里）',
     _x69['proj'])
# ② 华城原表第 14 行：3.31 孙代兰替华城交的设备票税费 16,883.46，7.20 跟华会计说定「税费结算时一起原路退回」，
#    原表在「泓普回款情况·应收工程款」列手填了这笔
_add('华城!14', '2026-07-20', '代垫应收', '华城', '泓普', 16883.46,
     '2026.7.20号和华会计确定：3.31孙代兰替华城交的设备票税费16883.46，税费结算时一起原路退回',
     '9.24 补录·原表华城表第 14 行「泓普回款情况·应收工程款」列手填 16,883.46', _ev('华城!11', '已交税')['proj'])
# ③ 金沁原表「金沁回款情况·质保金」列：第 5 行 521.43、第 30~34 行按开票额 3% 扣的质保金
for _n in (5, 30, 31, 32, 33, 34):
    _s = _ev(f'金沁!{_n}')
    _amt = 521.43 if _n == 5 else round(_s['amt'] * 0.03 + 1e-9, 2)
    _add(f'金沁!{_n}', _s['date'], '业主扣质保金', '民能', '金沁', _amt, _s['memo'],
         f'9.24 补录·原表金沁表第 {_n} 行「金沁回款情况·质保金」列' + ('' if _n == 5 else '（开票额×3%）'),
         _s['proj'])
# ④ 杰华原表：合计行「未到账余额」是按应扣 10% 管理费算的（=应收成本票−已到账），可「扣管理费10%」列一格没填。
#    9.11 跟夏旭核对的工资表抵成本 110,075.15 正好等于扣完 10% 的应收成本票，说明这 10% 已经按扣了结算
_add('杰华电气!5', '2026-09-11', '管理费结算', '泓普', '杰华', 12230.57,
     '2026.9.11和夏旭核对：大太线应扣管理费10%（122305.72×10%）',
     '9.24 补录·原表杰华表 I5 应扣管理费 12,230.57，「扣管理费10%」列漏填（合计行余额是按扣了算的）',
     _ev('杰华电气!5')['proj'])
EV.sort(key=lambda e: (e['date'], e['src']))

# ---------------- 容量（按连续多年使用留量） ----------------
# 留量按「够用 4~5 年、又不至于把表拖慢」定；不够了改这四个数重跑脚本即可
U0, U1 = 6, 45          # 单位档案   40 家
P0, P1 = 6, 305         # 项目档案  300 个（历史 63 个）
F0, F1 = 5, 2004        # 业务流水 2000 行（历史 373 行）
JR0, JR1 = 6, 2005      # 资金日记账 2000 行（出纳日记账格式：第 4~5 行表头，第 6 行起数据）
N_ACCT = 10             # 基础设置里能放几个科目
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
          ubel='AN', peer='AO')
FLOW_LAST = FC['peer']
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
JRNG = lambda c: f'{QJ}!${c}${JR0}:${c}${JR1}'      # 资金日记账某列整列
U_NAME = f'{QU}!$A${U0}:$A${U1}'
P_CODE = f'{QP}!$A${P0}:$A${P1}'
FRNG = lambda k: f'{QF}!${FC[k]}${F0}:${FC[k]}${F1}'      # 业务流水某列的整列绝对引用

# 业务类型
# 9.24 新增两类（都是照原表回款区手工加的那几笔补出来的）：
#   代垫应收     —— 我方替挂靠单位垫的钱、对方答应退回来（华城 16,883.46 设备票税费、迅驰 373.3 应退税费），
#                   算进「应收挂靠方」；对方真退回来时录「我方收款」冲掉
#   业主扣质保金 —— 业主扣在挂靠单位手上还没放的质保金（金沁表「金沁回款情况·质保金」那一列），
#                   只在「欠业主未付款」那一组里单列出来看，不改变欠款余额
KINDS = ['销项开票', '成本票', '挂靠代收', '我方收款', '管理费结算', '扣质保金', '业主扣质保金', '代垫应收',
         '已交税', '退税', '工资扣抵', '其他应收收回', '其他应付发生', '其他应付扣税', '其他应付支付', '其他']
KIND_DV = '"' + ','.join(KINDS) + '"'
ITYPES = ['劳务票', '土建票', '安装票', '机械设备票', '材料票', '其他', '不适用']
ITYPE_DV = '"' + ','.join(ITYPES) + '"'

def col_idx(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


# ============================================================ 筛选带（所有查询表通用）
def filter_band(ws, lastcol, unit_default=None,
                note='留空＝全部期间；填了年度就按整年取数，另填起止日期则以起止为准；余额类的列按截止日期算',
                cnt_label='业务流水', cnt_rng=None):
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
        f'=IF(AND({y}="",{s_}="",{e_}=""),"全部期间（{cnt_label}共 "&COUNT({cnt_rng or f"{QF}!$B${F0}:$B${F1}"})&" 笔）",'
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
    '历史行已按你那份《对账明细》逐笔写死。　　'
    '这里的税率参数只驱动【业务流水】最右边那一列「税费参考试算」，供录入时对照，不参与任何汇总：'
    '税差法＝上游销项税−我方成本票进项税；全额销项法＝开票额÷(1+上游税率)×上游税率；'
    '预征率法＝开票额÷(1+上游税率)×预征率。真正入账的税费在【业务流水】P~S 四列（税费录入区）按实际逐笔录。',
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
# 9.24 回传版里你新加的项目（A056 起），整行照搬
_wp = BACK['项目档案']
_have = {d['code'] for d in PROJ_ROWS}
NEW_PROJ = []
for _r in range(6, _wp.max_row + 1):
    _c = _wp[f'A{_r}'].value
    if not _c or _c in _have or str(_wp[f'B{_r}'].value or '').startswith('（新项目'): continue
    _g = lambda col: _wp[f'{col}{_r}'].value
    PROJ_ROWS.append(dict(code=_c, name=_g('B'), short=_g('C'), ptype=_g('D') or '自营项目',
                          ctype=_g('E') or '框架合同', owner=_g('F') or '', l1=_g('G') or '',
                          l2=_g('H') or '', mine=_g('I') or '泓普', memo=_g('V') or '9.24 回传版新增',
                          extra={c: _g(c) for c in 'JKLMNOPT' if _g(c) not in (None, '')}))
    NEW_PROJ.append(_c)
NEW_CODE = f'A{len(PROJ_ROWS)+1:03d}'         # 下一个新项目该用的编号
PROJ_BY_CODE = {d['code']: d for d in PROJ_ROWS}
PCODE = {d['name']: d['code'] for d in PROJ_ROWS if d['name']}
PSHORT = {d['code']: d['short'] for d in PROJ_ROWS}

# ============================================================ 9.24 历史数据更正
# 这一轮按原《对账明细》9.13 版把 8 家单位的回款区逐行重新对了一遍（欠业主未付款 / 应收挂靠方两组），
# 顺带查出上一版还原时的几类错。单位合计都不变，改的是「挂在哪个项目上、谁付给谁、哪天」——
# 不改的话，项目汇总里新加的两个余额列会出现 A008 −90,000、A035 −102,000 这种怪数。
# 每一笔改动都在【业务流水】备注列写了「9.24 更正：…」，【对账差异说明】最下面有逐条清单。
_S = lambda e: e['src'].replace(' ', '')
def _evs(src, kind=None, amt=None):
    _k = {'挂靠单位代收': '挂靠代收'}          # 原始事件里「挂靠代收」叫「挂靠单位代收」
    return [e for e in EV if _S(e) == src and (kind is None or _k.get(e['kind'], e['kind']) == kind)
            and (amt is None or abs(e['amt'] - amt) < 0.005)]
def _one(src, kind, amt=None):
    h = _evs(src, kind, amt)
    assert len(h) == 1, ('9.24 更正找不到唯一一行', src, kind, amt, len(h))
    return h[0]
def _fix(e, why, **kw):
    e.update(kw)
    e['note9'] = (e['note9'] + '｜' if e.get('note9') else '') + '9.24 更正：' + why
def _code(e): return e.get('code') or PCODE.get(e['proj'], '')
FIX924 = []                                            # (单位, 原表行, 改动, 金额)
def _log(u, src, what, amt): FIX924.append((u, src, what, amt))

# ① 新建 3 个项目：原表这几行「项目全称」是空的，上一版顺着上一行把它们挂到了 A049 侣新线凤飞7社(金沁) 上
NEWP = [dict(code=f'A{len(PROJ_ROWS)+1+i:03d}', name=n, short=sn, ptype=pt, ctype='框架合同', owner='民能',
             l1=l1, l2=l2, mine='泓普', memo='9.24 新建：原表这个项目没写全称，上一版误挂在 A049 侣新线凤飞7社(金沁) 下')
        for i, (n, sn, pt, l1, l2) in enumerate([
            ('重庆铜梁10kV大太线老旧线路改造工程（金沁、康欣这条链；原表未写全称）', '大太线老旧改造(金沁)', '自营项目', '金沁', '康欣'),
            ('10kV少盘线哨楼村支线#1杆线路迁改工程（康欣直接开给民能；原表未写全称）', '少盘线哨楼村支线(康欣)', '合伙项目', '康欣', ''),
            ('维新万头种猪养殖场（金沁、康欣这条链；原表未写全称）', '维新万头种猪养殖场(金沁)', '合伙项目', '金沁', '康欣')])]
PROJ_ROWS.extend(NEWP)
C_DATAI, C_SHAO, C_WEIX = [d['code'] for d in NEWP]
for e in _evs('康欣!66') + _evs('金沁!59') + _evs('金沁!60') + _evs('金沁!61'):
    _fix(e, f'大太线老旧改造原表没写项目全称，改挂新建的 {C_DATAI}（原误挂 A049 凤飞7社）', code=C_DATAI)
_log('康欣/金沁', '康欣!66、金沁!59~61', f'大太线老旧改造 {len(_evs("康欣!66"))+len(_evs("金沁!59"))+len(_evs("金沁!60"))+len(_evs("金沁!61"))} 行 A049 → {C_DATAI}', 203652)
_fix(_one('康欣!64', '销项开票'), f'哨楼村支线原表没写项目全称，改挂新建的 {C_SHAO}（原误挂 A049）', code=C_SHAO)
_log('康欣', '康欣!64', f'哨楼村支线 A049 → {C_SHAO}', 10029.37)
_wx = _evs('康欣!65') + _evs('金沁!56') + _evs('金沁!57') + _evs('金沁!58')
for e in _wx:
    _fix(e, f'维新万头种猪养殖场原表没写项目全称，改挂新建的 {C_WEIX}（原误挂 A049）', code=C_WEIX)
_log('康欣/金沁', '康欣!65、金沁!56~58', f'维新万头种猪养殖场 {len(_wx)} 行 A049 → {C_WEIX}', 27709)
# ② 大太线那笔 197,500：金沁表第 61 行写明是 2026.7.28 民能代付，上一版按康欣表开票日记成了 2025-12-08
for e in _evs('康欣!66', '挂靠代收') + _evs('康欣!66', '我方收款'):
    _fix(e, '日期按金沁表第 61 行「2026.7.28民能代付民工工资：大太线老旧改造」改为 2026-07-28（原记成开票日 2025-12-08）',
         date='2026-07-28')
_log('康欣', '康欣!66', '民能代付 197,500（挂靠代收 + 我方收款）日期 2025-12-08 → 2026-07-28', 197500)
# ③ 康欣表第 39/40 行：金额其实是 7.15 康欣开给金沁的两张票（金沁表第 22/23 行、康欣表第 44/45 行摘要），
#    原表摘要误抄成了「6.16金沁开到民能」—— 上一版照抄，康欣明细里就多出两行跟金沁表一模一样的「金沁开到民能」
for n, part in ((39, '安装'), (40, '土建')):
    e = _one(f'康欣!{n}', '销项开票')
    _fix(e, f'原表摘要误抄成「2026.6.16金沁开到民能」，按康欣表第 {n+5} 行、金沁表第 {62-n} 行改为 7.15 康欣开给金沁',
         payer='康欣', payee='金沁', date='2026-07-15', memo=f'2026.7.15康欣开票到金沁：金湖东西（{part}）')
    _log('康欣', f'康欣!{n}', f'开票方/收票方 金沁→民能 改为 康欣→金沁，日期 6.16 → 7.15', e['amt'])
    _m = _one(f'金沁!{62-n}', '销项开票', e['amt'])   # 金沁表里同一张票记在「已提供成本票」列，跟康欣表重复，只计一次
    for m in [_m]:
        if m.get('count_in', '是') == '是':
            _fix(m, f'跟康欣表第 {n} 行是同一张票（康欣开给金沁），汇总只计一次', count_in='否')
# ④ 项目编号挂错（原表项目名里的笔误让同一个项目被拆成两个编号，或回款挂到了别的项目上）
for src, kind, amt, to, why in [
    ('康欣!28', '挂靠代收', 57991, 'A002', '原表第 28 行是「水羊线金鸡3(4)社」，上一版误挂 A004'),
    ('迅驰!37', '我方收款', 3101.57, 'A028', '3101.57＝虎泉线业主回款 3230.8×96%，原表 T23 记在虎泉线上，上一版误挂 A031'),
    ('华城!6', '成本票', 80000, 'A001', '原表项目名把 10kV 写成了 11kV，是 A001 平双线同一个项目'),
    ('华城!26', '成本票', 5744.48, 'A037', '金额等于第 20 行高丰3社的应开成本票，上一版误挂 A036'),
    ('华城!27', '成本票', 6045.34, 'A038', '原表项目名少了个「个」字，是 A038 同一个项目'),
    ('华城!28', '销项开票', 3295.75, 'A042', '原表项目名把 220kV 写成了 219kV（摘要写的是 220kV），并回 A042'),
    ('华城!28', '管理费结算', 65.91, 'A042', '同上'),
    ('安锐!15', '成本票', 25682.49, 'A004', '金额＋第 20 行＝26,493.52，正好是第 7 行池塘4社的应收成本票，原表摘要把项目写反了（原挂 A005）'),
    ('安锐!20', '成本票', 811.03, 'A004', '同上'),
    ('安锐!16', '成本票', 25532.51, 'A005', '金额＋第 21 行＝26,338.80，正好是第 8 行团山的应收成本票，原表摘要把项目写反了（原挂 A004）'),
    ('安锐!21', '成本票', 806.29, 'A005', '同上'),
    ('金沁!20', '挂靠代收', 30000, 'A032', '金沁的金湖东西线是 A032，上一版误挂 A008（德誉嘉那条链）'),
    ('金沁!20', '我方收款', 30000, 'A032', '同上'),
    ('金沁!21', '挂靠代收', 60000, 'A032', '同上'), ('金沁!21', '我方收款', 60000, 'A032', '同上'),
    ('金沁!24', '挂靠代收', 70000, 'A032', '金沁的金湖东西线是 A032，上一版误挂 A035（华城许家3社）'),
    ('金沁!24', '我方收款', 70000, 'A032', '同上'),
    ('金沁!25', '挂靠代收', 32000, 'A032', '同上'), ('金沁!25', '我方收款', 32000, 'A032', '同上'),
    ('金沁!41', '已交税', 144.04, 'A047', '摘要写的是群益7社，上一版误挂 A045'),
    ('金沁!37', '已交税', 1.71, 'A048', '摘要写的是平镇线新桥社，上一版挂在 A049'),
    ('金沁!39', '已交税', 3.8, 'A045', '摘要写的是侣新线铜安1社，上一版挂在 A049'),
    ('安锐!10', '销项开票', 5330.51, 'A007', '安锐开给德誉嘉的平滩所何银强，是德誉嘉那条链 A007，上一版误挂 A043（金沁链）'),
    ('安锐!12', '挂靠代收', 5330.51, 'A007', '同上')]:
    e = _one(src, kind, amt)
    old = _code(e)
    _fix(e, why + f'（{old} → {to}）', code=to)
    _log(src.split('!')[0], src, f'{kind} {old} → {to}', amt)
_fix(_one('安锐!12', '挂靠代收', 5330.51), '付款方按原表摘要「德誉佳打款给安锐」改为德誉嘉（原记民能）', payer='德誉嘉')
_log('安锐', '安锐!12', '挂靠代收 付款方 民能 → 德誉嘉', 5330.51)
# ⑤ 华城表第 24 行：上一版还原成了「民能开给华城的销项票」，其实是泓普开给华城许家3社的成本票
#    （金额＝第 18 行许家3社的应开成本票 3,725.89；摘要是从金沁表第 25 行误抄过来的）
_e24 = _one('华城!24', '销项开票', 3725.89)
_fix(_e24, '原表摘要误抄了金沁表「民能代付工资」，金额＝第 18 行许家3社应开成本票，改为泓普开给华城的成本票',
     kind='成本票', payer='泓普', payee='华城', code='A035', itype='劳务票', date='2026-07-24',
     memo='2026.7.24泓普开票（劳务）到华城：永嘉所许家3社（原表摘要误抄成「2026.7.28民能代付工资：金湖东西电缆改造（土建）」，日期按同一批成本票改成 7.24）')
_log('华城', '华城!24', '销项开票 民能→华城 改为 成本票 泓普→华城（A035 许家3社）', 3725.89)
# ⑥ 批量结算的管理费 / 税费按项目拆开（原表一格记的是好几个项目合起来的数，只挂在最后一个项目上）
def _split(src, kind, amt, parts, why):
    e = _one(src, kind, amt)
    i = EV.index(e)
    agg_ = collections.OrderedDict()
    for c, a in parts: agg_[c] = agg_.get(c, 0) + a
    items = list(agg_.items())
    pieces, tot = [], 0.0
    for j, (c, a) in enumerate(items):
        a = round(a, 2) if j < len(items) - 1 else round(amt - tot, 2)
        tot = round(tot + a, 2)
        p = dict(e); p.update(code=c, amt=a)
        p['memo'] = e['memo'] + f'｜{len(items)} 个项目合起来{kind} {amt:,.2f}，本行是这个项目那一份'
        p['note9'] = f'9.24 按项目拆分：原表这一格 {amt:,.2f} 是 {len(items)} 个项目合起来{why}，按各项目金额拆开，合计不变'
        pieces.append(p)
    EV[i:i+1] = pieces
    _log(src.split('!')[0], src, f'{kind} {amt:,.2f} 拆到 {items[0][0]}~{items[-1][0]} 共 {len(items)} 个项目', amt)
def _fee_parts(sh, a, b):
    return [(_code(x), x['mfee']) for x in EV if _S(x).split('!')[0] == sh and x['kind'] == '销项开票'
            and a <= int(_S(x).split('!')[1]) <= b and x.get('mfee')]
for sh, a, b, src, amt in [('德誉嘉', 5, 10, '德誉嘉!10', 22320.78), ('德誉嘉', 27, 32, '德誉嘉!32', 2829.79),
                           ('德誉嘉', 45, 53, '德誉嘉!53', 984.06), ('德誉嘉', 73, 78, '德誉嘉!78', 1099.58),
                           ('迅驰', 5, 10, '迅驰!10', 4504.21), ('迅驰', 19, 26, '迅驰!26', 1753.06),
                           ('迅驰', 48, 53, '迅驰!54', 706.05), ('华城', 18, 22, '华城!22', 488.78),
                           ('金沁', 30, 34, '金沁!35', 3528.48), ('安锐', 5, 9, '安锐!9', 2728.39)]:
    _split(src, '管理费结算', amt, _fee_parts(sh, a, b), '结算的管理费（＝各项目应扣管理费之和）')
def _tax(x): return sum((x.get(k) or 0) for k in ('tax_v', 'tax_s', 'tax_y', 'tax_i'))
def _tax_parts(sh, a, b):
    return [(_code(x), _tax(x)) for x in EV if _S(x).split('!')[0] == sh and x['kind'] == '销项开票'
            and a <= int(_S(x).split('!')[1]) <= b]
for sh, a, b, src, amt in [('康欣', 5, 9, '康欣!9', 7443.98), ('湖南锦泰', 5, 7, '湖南锦泰!9', 16389.36)]:
    _split(src, '管理费结算', amt, _fee_parts(sh, a, b), '结算的管理费（＝各项目应扣管理费之和）')
_split('康欣!10', '已交税', 12213.85, _tax_parts('康欣', 5, 9), '交的税（＝第 5~9 行应扣税费之和）')
_split('迅驰!11', '已交税', 6953.21, _tax_parts('迅驰', 5, 10), '交的税（＝第 5~10 行应扣税费之和）')
_split('迅驰!54', '已交税', 1679.77, _tax_parts('迅驰', 48, 53) + [(_code(_one('迅驰!67', '销项开票')), 0)],
       '交的税（第 48~53 行应扣税费，余下约 589.85 挂少盘线：原表是 216.53＋7.24 退回的 373.30＝589.83，差 0.02 是原表税费四舍五入）')
_j38 = [(_code(_one('金沁!31', '销项开票')), 3.59 * _one('金沁!31', '销项开票')['amt']
         / (_one('金沁!31', '销项开票')['amt'] + _one('金沁!32', '销项开票')['amt'])),
        (_code(_one('金沁!32', '销项开票')), 0)]
_split('金沁!38', '已交税', 3.59, _j38, '交的印花税（摘要写了铜安9社和群益7社两个项目，按开票额分）')
for src, kind, amt in [('金沁!35', '成本票', 30.84), ('金沁!35', '我方收款', 30.8), ('金沁!45', '已交税', 30.84)]:
    _fix(_one(src, kind, amt), '这一笔是第 30~34 行 5 个项目合起来的税差（摘要「印花税差额30.8」），金额太小没拆，只挂在 A049')
# 康欣原表 H68「未开票按3%扣点子」一共 9 笔：上面补录了 6 笔，原来已经在系统里的 3 笔也标成「不开票」
for src, amt in (('康欣!21', 55907.06), ('康欣!33', 12200), ('康欣!50', 66500)):
    e = _one(src, '成本票', amt)
    _fix(e, '原表 H68「未开票按3%扣点子」9 笔之一，发票性质改为「不开票」', inv='不开票',
         memo=e['memo'] + ('' if '扣点' in e['memo'] else '｜未开票按3%扣点子（不提供成本票）'))
_log('康欣', '康欣!21、33、50', '成本票 55,907.06 / 12,200 / 66,500 发票性质改「不开票」（原表 H68 扣点清单）', 134607.06)
# 园区光伏电力新建：原表康欣 B50 把「泓普7W」写成「泓普6W」，被拆成 A054、A044 两个编号，并回 A044
for src in ('康欣!50', '康欣!52', '康欣!53'):
    for e in _evs(src):
        if _code(e) == 'A054':
            _fix(e, '原表项目名把「泓普7W」写成了「泓普6W」，跟金沁那边是同一个项目，并回 A044（原挂 A054）', code='A044')
_log('康欣', '康欣!50、52、53', '园区光伏电力新建 A054 → A044（项目名笔误）', 70000)
_fix(_one('金沁!28', '销项开票', 66500), '跟康欣表第 50 行是同一张票（康欣开给金沁的电力新建），并项目以后汇总只计一次', count_in='否')
_hn = _tax_parts('湖南锦泰', 5, 7)
for src, amt in (('湖南锦泰!8', 3668.81), ('湖南锦泰!9', 14097.81)):
    _t = sum(a for _, a in _hn)
    _split(src, '已交税', amt, [(c, amt * a / _t) for c, a in _hn],
           '预交的税（3 个项目合起来交的，按各项目应扣税费的比例分；原表合计多交 2,610.25）')
_split('安锐!11', '已交税', 332.72,
       [(_code(x), (x.get('tax_v') or 0) + (x.get('tax_s') or 0) + (x.get('tax_y') or 0) + (x.get('tax_i') or 0))
        for x in EV if _S(x).split('!')[0] == '安锐' and x['kind'] == '销项开票' and 5 <= int(_S(x).split('!')[1]) <= 9],
       '交的税（＝第 5~9 行应扣税费之和）')
# ⑦ 康欣表第 58 行（凤飞7社）「泓普回款情况·应收金额」手填了 408.77，数值等于金沁表该票的应扣税费，原表没写说明。
#    照原表补一笔「代垫应收」，康欣应收挂靠方余额跟原表对上；如果这笔不该收，删掉这一行即可
_add('康欣!58', '2026-07-15', '代垫应收', '康欣', '泓普', 408.77, _one('康欣!58', '销项开票')['memo'],
     '9.24 补录·原表康欣表第 58 行「应收金额」列手填 408.77（＝金沁表凤飞7社那张票的应扣税费），请核实',
     _one('康欣!58', '销项开票')['proj'])
_log('康欣', '康欣!58', '补录 代垫应收 408.77（原表手填，请核实）', 408.77)
# ⑧ 康欣表第 67 行（稷博汇）管理费率 0.019999 是从 71.40 倒推出来的，改成 2%（金额不变）
_fix(_one('康欣!67', '销项开票'), '管理费率 0.019999 → 2%（原来是倒推出来的，金额不变）', mrate=0.02)
for code, to, why in (('A034', 'A001', '原表项目名把 10kV 写成了 11kV'), ('A040', 'A038', '原表项目名少了个「个」字'),
                      ('A041', 'A042', '原表项目名把 220kV 写成了 219kV'), ('A054', 'A044', '原表项目名把「泓普7W」写成了「泓普6W」')):
    d = [x for x in PROJ_ROWS if x['code'] == code][0]
    d['memo'] = f'9.24 停用：{why}，跟 {to} 是同一个项目，业务已全部并入 {to}，这个编号不要再用'
    d['status'] = '暂停'
    d['short'] = d['short'] + '·停用'
for d in PROJ_ROWS:
    if d['code'] == 'A001': d['short'] = '平双线'
    if d['code'] == 'A035': d['l2'] = ''
    if d['code'] == 'A007': d['l2'] = '安锐'
EV.sort(key=lambda e: (e['date'], e['src']))
PROJ_BY_CODE = {d['code']: d for d in PROJ_ROWS}
PCODE = {d['name']: d['code'] for d in PROJ_ROWS if d['name']}
PSHORT = {d['code']: d['short'] for d in PROJ_ROWS}
NEW_CODE = f'A{len(PROJ_ROWS)+1:03d}'
print(f'  ✓ 9.24 历史数据更正 {len(FIX924)} 条，新建项目 {C_DATAI}/{C_SHAO}/{C_WEIX}')

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
        ws[f'N{r}'] = d.get('status', '在建')
        if d['memo']: ws[f'V{r}'] = d['memo']
        for c, v in d.get('extra', {}).items(): ws[f'{c}{r}'] = v
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
      f'{sum(1 for d in PROJ_ROWS if d["ptype"]=="合伙项目")} 个合伙项目，9.24 新增 {"、".join(NEW_PROJ)}，下一个编号 {NEW_CODE}）')

# ============================================================ 业务流水（总表）
ws = wb.create_sheet(SH_FLOW)
C_ = FC          # 下面一律用 C_['xxx'] 取列字母，以后再调列序只改上面那张字典
title(ws, '业务流水（总表）', FLOW_LAST,
      '所有开票、成本票、回款、交税、管理费结算、往来款都在这里录一行（每种情况怎么录、红冲怎么录、录错怎么改，看【操作流程】）。'
      '不要删整行、不要插行，多录的行清空内容即可。'
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
put(ws, f'{C_["peer"]}4', '收票方是挂靠单位', font=F_HDR2, fill=FILL_AUTO)

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
    # 原表来源列：历史行照原表写死；新录的行按业务类型自动认 —— 成本票算「已开成本票」、其余算「销售开票金额」。
    # 上一版新录的行这一格是空的，结果【单位汇总】的开票额、明细表的「已到成本票」都漏掉了新录的票
    put(ws, f'{C_["ocol"]}{r}', f'=IF($D{r}="","",IF($D{r}="成本票","已开成本票","销售开票金额"))',
        font=F_NOTE, fill=FILL_AUTO)
    # 归属单位表＝这一笔算在哪家挂靠单位的对账表上。历史行照原表的来源表；
    # 新录的行自动认开票方（开票方不是挂靠单位就认收票方）
    # 9.24：按业务类型定方向 —— 成本票、挂靠代收、业主扣质保金、管理费结算、工资扣抵是「收的那一家」的事，
    # 先认收票/收款方；其余（开票、转我方、代垫、交税…）先认开票/付款方。
    # 原来一律先认开票/付款方，康欣开给金沁的成本票、上一层挂靠单位付给下一层的钱都会记到付钱那一家头上
    _eh = f'{lk(f"$E{r}", U_NAME, U_TYPE)}="挂靠单位"'
    _fh = f'{lk(f"$F{r}", U_NAME, U_TYPE)}="挂靠单位"'
    put(ws, f'{C_["ubel"]}{r}',
        f'=IF(${C_["src"]}{r}<>"",${C_["src"]}{r},'
        f'IF(OR($D{r}="成本票",$D{r}="挂靠代收",$D{r}="业主扣质保金",$D{r}="管理费结算",$D{r}="工资扣抵"),'
        f'IF({_fh},$F{r},IF({_eh},$E{r},"")),'
        f'IF({_eh},$E{r},IF({_fh},$F{r},""))))',
        font=F_NOTE, fill=FILL_AUTO)
    # 取数键：归属单位#该单位第几笔 —— 8 张单位竖版明细靠它一行一行取数。
    # 用「归属单位表」而不是「来源表」：来源表只有历史行有值，你新录的行来源表是空的，
    # 用来源表当键的话新行永远不会出现在单位明细里。
    # 这一笔销项票是不是开给另一家挂靠单位的（那样它对收票方来说就是一张成本票）。
    # 【项目汇总】【单位项目明细】统计「已收成本票」时要用它来筛，
    # 否则收票方条件写成 "*" 会把每一张销项票都当成本票再加一遍 —— 整整翻一倍。
    put(ws, f'{C_["peer"]}{r}',
        f'=IF($D{r}<>"销项开票","",'
        f'IF({lk(f"$F{r}", U_NAME, U_TYPE)}="挂靠单位","是","否"))',
        font=F_NOTE, fill=FILL_AUTO)
    put(ws, f'{C_["skey"]}{r}',
        f'=IF(${C_["ubel"]}{r}="","",${C_["ubel"]}{r}&"#"&'
        f'COUNTIF(${C_["ubel"]}${F0}:${C_["ubel"]}{r},${C_["ubel"]}{r}))',
        font=F_NOTE, fill=FILL_AUTO)
    _et = lk(f"$E{r}", U_NAME, U_TYPE); _ft = lk(f"$F{r}", U_NAME, U_TYPE)
    put(ws, f'{C_["chk"]}{r}',
        f'=IF(AND($B{r}="",$D{r}="",${A_}{r}=""),"",'
        f'IF($B{r}="","未填日期",'
        f'IF(NOT(ISNUMBER($B{r})),"日期格式不对",'
        f'IF($D{r}="","未选业务类型",'
        f'IF(AND($C{r}="",OR($D{r}="销项开票",$D{r}="成本票",$D{r}="工资扣抵",'
        f'$D{r}="挂靠代收",$D{r}="我方收款")),"未选项目编号",'
        f'IF(LEFT(${C_["pname"]}{r},1)="⚠","项目编号不存在",'
        f'IF($E{r}="","未选开票/付款方",'
        f'IF(ISNA(MATCH($E{r},{U_NAME},0)),"开票方不在单位档案",'
        f'IF($F{r}="","未选收票/收款方",'
        f'IF(ISNA(MATCH($F{r},{U_NAME},0)),"收款方不在单位档案",'
        f'IF(RIGHT(${C_["sname"]}{r},3)="·停用","项目编号已停用（看项目档案备注，换成并入的那个编号）",'
        # 方向检查：9.23 华城那几行就是业务类型选错了（开票录成挂靠代收、成本票录成我方收款），以前校验照样 √
        f'IF(AND($D{r}="销项开票",{_et}="我方主体"),"泓普开给挂靠单位的票请选「成本票」",'
        f'IF(AND($D{r}="销项开票",{_et}="业主"),"开票方不能是业主（方向反了？）",'
        f'IF(AND($D{r}="挂靠代收",{_et}="挂靠单位",{_ft}="业主"),"挂靠单位开给业主的票请选「销项开票」",'
        f'IF(AND($D{r}="挂靠代收",{_ft}<>"挂靠单位"),"挂靠代收的收款方要选收钱的挂靠单位",'
        f'IF(AND($D{r}="我方收款",{_et}="我方主体"),"泓普开给挂靠单位的票请选「成本票」",'
        f'IF(AND($D{r}="我方收款",{_ft}<>"我方主体"),"我方收款的收款方要选泓普/仟茂",'
        f'IF(AND($D{r}="成本票",{_et}<>"我方主体",{_et}<>"挂靠单位"),"成本票的开票方应是泓普/仟茂",'
        f'IF(AND($D{r}<>"销项开票",OR(N(${R_}{r})<>0,${FG}{r}<>"")),"只有销项开票才填管理费率 / 计费方式（业务类型是不是选错了？）",'
        f'IF(NOT(ISNUMBER(${A_}{r})),"金额须为数字",'
        f'IF(AND(OR($D{r}="销项开票",$D{r}="成本票"),$H{r}=""),"未选票据类型",'
        f'IF(AND($D{r}="销项开票",${FG}{r}=""),"没选计费方式",'
        f'IF(AND($D{r}="销项开票",${FG}{r}="扣管理费",ROUND(N(${R_}{r}),6)=0),'
        f'"选了扣管理费但费率是 0",'
        f'IF(AND(${C_["ubel"]}{r}="",OR($D{r}="销项开票",$D{r}="成本票",$D{r}="挂靠代收",$D{r}="我方收款",'
        f'$D{r}="管理费结算",$D{r}="扣质保金",$D{r}="业主扣质保金",$D{r}="代垫应收",$D{r}="已交税",$D{r}="工资扣抵")),'
        f'"两边都不是挂靠单位，算不到哪家头上",'
        f'IF(AND($D{r}="销项开票",${C_["ocol"]}{r}="已开成本票",${FG}{r}<>"不回成本票"),'
        f'"这是收票方账上的镜像行，计费方式只能是不回成本票",'
        f'IF(AND($D{r}="销项开票",${C_["tier"]}{r}="—"),"链条待确认","√"))))))))))))))))))))))))))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16

# ---- 导入历史事件 ----
KIND_MAP = {'销项开票':'销项开票','成本票':'成本票','挂靠单位代收':'挂靠代收','我方收款':'我方收款',
            '管理费结算':'管理费结算','扣质保金':'扣质保金','已交税':'已交税',
            # 德誉嘉表最右边那三列「合伙项目应付款」原来没列进来，
            # 4 笔 66,983.97 全落成了「其他」，往来台账和对账差异说明都统计不到
            '其他应付发生':'其他应付发生','其他应付支付':'其他应付支付',
            '代垫应收':'代垫应收','业主扣质保金':'业主扣质保金'}
# 原对账明细的表名 → 单位简称（只有杰华那张表名字不一样）
SHEET2UNIT = {'德誉嘉': '德誉嘉', '迅驰': '迅驰', '华城': '华城', '金沁': '金沁',
              '湖南锦泰': '湖南锦泰', '康欣': '康欣', '安锐': '安锐', '杰华电气': '杰华'}
imported = 0
for i, e in enumerate(EV):
    r = F0 + i
    if r > F1: raise SystemExit('业务流水容量不够')
    ws[f'B{r}'] = dt.date.fromisoformat(e['date'])
    ws[f'C{r}'] = e.get('code') or PCODE.get(e['proj'], '')
    ws[f'D{r}'] = KIND_MAP.get(e['kind'], '其他')
    ws[f'E{r}'] = e['payer']; ws[f'F{r}'] = e['payee']
    if e.get('inv'): ws[f'G{r}'] = e['inv']
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
    if e.get('note9'):
        ws[f'{C_["note"]}{r}'] = e['note9']
    if e['kind'] == '销项开票':
        # 费率、返现率、税费一律照原表这一行写死（原来按单位档案参数算，跟原表对不上）
        ws[f'{R_}{r}'] = round(e.get('mrate', 0.0), 6)
        ws[f'{B_}{r}'] = round(e['rebate'] / e['amt'], 6) if (e.get('rebate') and e['amt']) else 0
        if e.get('from_cost_col'):
            # 这一行是从原表「已开成本票」列还原出来的票，本身不产生「我方应回成本票」的义务
            ws[f'{FG}{r}'] = '不回成本票'
            ws[f'{C_["note"]}{r}'] = ('历史导入·原表「已开成本票」列还原；不另生成应回成本票'
                                       + ('｜' + e['note9'] if e.get('note9') else ''))
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
# ---- 9.24 回传版：你改过的两处历史行 + 新录的 9 月业务（录错的地方按原意更正，备注列写明） ----
def _find(src, srow, kind, amt):
    hit = [rr for rr in range(F0, r) if ws[f'{C_["src"]}{rr}'].value == src
           and ws[f'{C_["srow"]}{rr}'].value == srow and ws[f'D{rr}'].value == kind
           and abs((ws[f'{A_}{rr}'].value or 0) - amt) < 0.005]
    assert len(hit) == 1, (src, srow, kind, amt, hit)
    return hit[0]
_rr = _find('康欣', 35, '挂靠代收', 16500)          # 你把付款方从民能改成了铜梁供电，照你的改
ws[f'E{_rr}'] = '铜梁供电'
ws[f'{C_["note"]}{_rr}'] = '9.24 照回传版：付款方由民能改为铜梁供电'
PATCH_ROWS = {'payer': _rr}
# 回传版把康欣第 9 行的管理费结算 7,443.98 删了（它不是重复，是同一张原表行里的管理费结算）—— 恢复，并按项目拆成 5 行
_kx9 = [rr for rr in range(F0, r) if ws[f'{C_["src"]}{rr}'].value == '康欣' and ws[f'{C_["srow"]}{rr}'].value == 9
        and ws[f'D{rr}'].value == '管理费结算']
assert abs(sum(ws[f'{A_}{rr}'].value for rr in _kx9) - 7443.98) < 0.005, _kx9
for rr in _kx9:
    ws[f'{C_["note"]}{rr}'] = '9.24 恢复：回传版把这笔管理费结算 7,443.98 删了（不是重复），已恢复｜' + str(ws[f'{C_["note"]}{rr}'].value)
PATCH_ROWS['kx_fee'] = _kx9[0]
_rr = _find('德誉嘉', 47, '销项开票', 1622.53)      # 这一行保持原样（7.22 那张票确实开过），红冲另起两行
PATCH_ROWS['dyj_old'] = _rr
ws[f'{C_["note"]}{_rr}'] = '9.24：按你说的「红冲更正发票」处理 —— 这张票 9.20 已红冲、重开 1,662.53（见日期 2026-09-20 的那两行），本行保持原样不要改'
_dyj_rate = ws[f'{R_}{_rr}'].value
NEW924 = []
def _new(date, code, kind, payer, payee, inv, itype, memo, amt, rate=None, reb=None, flag=None,
         note=None, vat=None):
    global r
    ws[f'B{r}'] = date; ws[f'C{r}'] = code; ws[f'D{r}'] = kind
    ws[f'E{r}'] = payer; ws[f'F{r}'] = payee
    if inv: ws[f'G{r}'] = inv
    if itype: ws[f'H{r}'] = itype
    ws[f'{C_["memo"]}{r}'] = memo; ws[f'{A_}{r}'] = amt
    if rate is not None: ws[f'{R_}{r}'] = rate
    if reb is not None: ws[f'{B_}{r}'] = reb
    if flag: ws[f'{FG}{r}'] = flag
    if note: ws[f'{C_["note"]}{r}'] = note
    NEW924.append(r); r += 1
_d920, _d923 = dt.date(2026, 9, 20), dt.date(2026, 9, 23)
_new(_d920, 'A017', '销项开票', '德誉嘉', '民能', '13%专票', '机械设备票',
     '2026.9.20德誉嘉红冲：冲销2026.7.22开到民能的发票（永嘉所义和8社）', -1622.53,
     rate=_dyj_rate, reb=0, flag='扣管理费',
     note='红冲：金额填负数，费率/计费方式跟被冲的那张票一模一样，管理费和应到成本票自动冲成负数')
_new(_d920, 'A017', '销项开票', '德誉嘉', '民能', '13%专票', '机械设备票',
     '2026.9.20德誉嘉重开发票到民能：永嘉所义和8社（更正后金额）', 1662.53,
     rate=0.04, reb=0, flag='扣管理费', note='红冲后重开的正确发票')
HC924 = [('A056', '虎峰所双桥11社', 462.92, 453.66), ('A057', '永嘉所柳树7社', 13200.28, 12936.27),
         ('A058', '永嘉所河中3.4社', 10736.77, 10522.03), ('A059', '永嘉所长寿3社', 11286.03, 11060.31)]
for code, sn, amt, cost in HC924:
    _new(_d923, code, '销项开票', '华城', '民能', '3%专票', '劳务票', f'2026.9.23华城-民能：{sn}', amt,
         rate=0.02, reb=0, flag='扣管理费',
         note=('9.24 更正：业务类型原录成「挂靠代收」，金额原录 356,218.09（按成本票 453.66÷98% 反推应为 462.92，请核对发票）'
               if code == 'A056' else '9.24 更正：业务类型原录成「' + ('挂靠代收' if code != 'A059' else '我方收款') + '」，华城开给民能的票应选「销项开票」'))
for code, sn, amt, cost in HC924:
    _new(_d923, code, '成本票', '泓普', '华城', '3%专票', '劳务票', f'2026.9.23泓普-华城：{sn}', cost,
         note='9.24 更正：业务类型原录成「我方收款」，泓普开给华城的票应选「成本票」；金额取两位小数')
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
ws.column_dimensions[C_['peer']].hidden = True
ws.freeze_panes = 'C5'; page(ws, titles='4:4')
FLOW_USED = imported + extra + len(NEW924)
print(f'  ✓ 业务流水（历史 {imported} 笔 + 过账/工资/扣税 {extra} 笔 + 9 月新录 {len(NEW924)} 笔 = {FLOW_USED} 笔，容量 {F1-F0+1} 行）')

# ============================================================ 资金日记账（9.24 起改用出纳日记账的格式）
# 版式照你给的「资金日记帐-新表」：B~M 是出纳那本账原样的 12 列，出纳录完整块复制过来粘贴即可；
# O 列「基础设置」就是科目（账户）清单，J 列「科目」下拉从这里取；
# 右边 P~S 是每个科目的期初 / 借 / 贷 / 余额，U~Y 是系统自动列（逐行余额、项目、校验）。
# 旧的「泓普 / 仟茂 / 现金 三账户」那张表整张删掉了，别的表全部改成从这张取数。
ws = wb.create_sheet(SH_JOUR)
ws.sheet_view.showGridLines = False
F_J = Font(name='微软雅黑', size=10)
F_JH = Font(name='微软雅黑', size=10, bold=True)
J_DATE = 'yyyy"年"m"月"d"日";@'
J_MONEY = '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ '
widths(ws, {'A': 3, 'B': 17.5, 'C': 8.8, 'D': 8, 'E': 12, 'F': 13, 'G': 14, 'H': 10, 'I': 60,
            'J': 9.8, 'K': 13.7, 'L': 13.7, 'M': 16, 'N': 0.6, 'O': 12, 'P': 13, 'Q': 13, 'R': 13,
            'S': 13, 'T': 2, 'U': 13, 'V': 16, 'W': 7, 'X': 6, 'Y': 18})
ws.column_dimensions['N'].hidden = True
ws.merge_cells('B1:M1')
put(ws, 'B1', '出纳日记账', font=Font(name='微软雅黑', size=25), align=C, border=None)
ws.row_dimensions[1].height = 42
ws.merge_cells('B2:C2'); ws.merge_cells('D2:J2')
put(ws, 'B2', '单位名称：', font=F_JH, align=CR, border=None)
# 原来 D2 是链到出纳那本账【使用说明】的外部链接（=[1]使用说明!D2），一换电脑就断，改成直接写单位名
put(ws, 'D2', '重庆泓普电力有限公司', font=F_JH, align=CL, border=None)
ws.merge_cells('O1:Y2')
put(ws, 'O1', '出纳在自己那本账里录完 → 把 B~M 列（日期到备注）整块复制 → 到这里点 B 列第一个空行 → '
              '选择性粘贴「数值」。科目必须是 O 列「基础设置」里有的；期初余额填 P 列。'
              '右边灰色列自动算，不用动。', font=F_NOTE, align=CL, border=None)
ws.row_dimensions[2].height = 22
ws.row_dimensions[3].height = 7
JH_IN = [('B', '日期'), ('C', '凭证\n种类'), ('D', '凭证\n编号'), ('E', '项目编号'), ('F', '费用类型'),
         ('G', '往来单位'), ('H', '报销人员'), ('I', '摘要'), ('J', '科目'), ('K', '借方'), ('L', '贷方'),
         ('M', '备注')]
JH_SET = [('O', '基础设置\n（科目）'), ('P', '期初余额'), ('Q', '借方合计'), ('R', '贷方合计'), ('S', '期末余额')]
JH_AUTO = [('U', '本科目余额'), ('V', '项目简称'), ('W', '年度'), ('X', '计入'), ('Y', '校验')]
for grp, fill, font in ((JH_IN, None, F_JH), (JH_SET, FILL_HDR2, F_HDR2), (JH_AUTO, FILL_AUTO, F_HDR2)):
    for c, t in grp:
        if c == 'B':
            put(ws, 'B4', t, font=font, align=C)
            put(ws, 'B5', f'{JOUR_YEAR}年', font=font, align=C)
            continue
        ws.merge_cells(f'{c}4:{c}5')
        put(ws, f'{c}4', t, font=font, fill=fill, align=C)
        put(ws, f'{c}5', None, font=font, fill=fill)
ws.row_dimensions[4].height = 17.25; ws.row_dimensions[5].height = 17.25
JA0, JA1 = JR0, JR0 + N_ACCT - 1              # 基础设置：科目清单 10 行
ACC_O = f'$O${JA0}:$O${JA1}'; ACC_P = f'$P${JA0}:$P${JA1}'
JCOL = lambda c: f'${c}${JR0}:${c}${JR1}'
for i in range(N_ACCT):
    r = JA0 + i
    put(ws, f'O{r}', JOUR_ACCTS[i] if i < len(JOUR_ACCTS) else None, font=F_IN, fill=FILL_IN)
    put(ws, f'P{r}', None, font=F_IN, fill=FILL_IN, fmt=MONEY)
    put(ws, f'Q{r}', f'=IF($O{r}="","",SUMIFS({JCOL("K")},{JCOL("J")},$O{r},{JCOL("X")},1))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'R{r}', f'=IF($O{r}="","",SUMIFS({JCOL("L")},{JCOL("J")},$O{r},{JCOL("X")},1))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'S{r}', f'=IF($O{r}="","",ROUND(N($P{r})+$Q{r}-$R{r},2))', font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
J_TOT = JA1 + 1
put(ws, f'O{J_TOT}', '合  计', font=F_TOT, fill=FILL_TOT)
for c in 'PQRS':
    put(ws, f'{c}{J_TOT}', f'=SUM({c}{JA0}:{c}{JA1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
_dup = f'SUMPRODUCT(({ACC_O}<>"")*(COUNTIF({ACC_O},{ACC_O}&"")>1))'
put(ws, f'O{J_TOT+1}', f'=IF({_dup}>0,"⚠ 科目清单里有重名的，余额会重复算，请删掉一个",'
                       f'IF(ROUND(SUMIFS({JCOL("K")},{JCOL("X")},1)-$Q${J_TOT},2)'
                       f'+ROUND(SUMIFS({JCOL("L")},{JCOL("X")},1)-$R${J_TOT},2)<>0,'
                       f'"⚠ 有 "&COUNTIF({JCOL("Y")},"科目不在基础设置里")&" 行的科目不在上面清单里，余额没算进去",'
                       f'IF(COUNT({ACC_P})=0,"⚠ P 列期初余额还没填：上面的余额只是本表录入以来的净发生额，不是账上真实余额",'
                       f'"✓ 每一笔都落到了上面的科目里")))',
    font=F_TOT, fill=FILL_CHK, align=CL)
put(ws, f'O{J_TOT+2}', '要加科目：在 O 列空格里填名字、P 列填期初（出纳账「期初余额表」里的数），J 列下拉马上多出来。',
    font=F_NOTE, align=CL, border=None)
put(ws, f'O{J_TOT+3}', '科目名字要跟出纳那本账写的一字不差，也不要重名。', font=F_NOTE, align=CL, border=None)
_dvo = DataValidation(type='custom', formula1=f'COUNTIF($O${JA0}:$O${JA1},O{JA0})=1', allow_blank=True,
                      showErrorMessage=True, errorStyle='stop', errorTitle='科目重名',
                      error='基础设置里已经有这个科目了，同一个科目只能写一次')
ws.add_data_validation(_dvo); _dvo.add(f'O{JA0}:O{JA1}')
SUBTOT = ['本月合计', '本年累计', '过次页', '承前页', '上年结转']
for r in range(JR0, JR1 + 1):
    for c, _ in JH_IN:
        put(ws, f'{c}{r}', None, font=F_J, align=CL if c in 'IM' else C,
            fmt=J_DATE if c == 'B' else (J_MONEY if c in 'KL' else ('@' if c in 'DEFGHI' else None)))
    _clean = f'SUBSTITUTE(SUBSTITUTE($I{r}," ",""),"　","")'
    _sub = 'OR(' + ','.join(f'{_clean}="{t}"' for t in SUBTOT) + ')'
    put(ws, f'X{r}', f'=IF(AND(ISNUMBER($B{r}),$J{r}<>"",OR(N($K{r})<>0,N($L{r})<>0),NOT({_sub})),1,0)',
        font=F_NOTE, fill=FILL_AUTO)
    put(ws, f'U{r}', f'=IF($X{r}<>1,"",ROUND(SUMIFS({ACC_P},{ACC_O},$J{r})'
                     f'+SUMIFS($K${JR0}:$K{r},$J${JR0}:$J{r},$J{r},$X${JR0}:$X{r},1)'
                     f'-SUMIFS($L${JR0}:$L{r},$J${JR0}:$J{r},$J{r},$X${JR0}:$X{r},1),2))',
        font=F_LINK, fill=FILL_AUTO, fmt=MONEY)
    put(ws, f'V{r}', f'=IF($E{r}="","",IFERROR(INDEX({P_SN},MATCH($E{r},{P_CODE},0)),"⚠编号不存在"))',
        font=F_LINK, fill=FILL_AUTO, align=CL)
    put(ws, f'W{r}', f'=IF(ISNUMBER($B{r}),YEAR($B{r}),"")', font=F_LINK, fill=FILL_AUTO)
    put(ws, f'Y{r}',
        f'=IF(AND($B{r}="",$J{r}="",N($K{r})=0,N($L{r})=0),"",'
        f'IF({_sub},"√ 合计/结转行（不计入）",'
        f'IF(NOT(ISNUMBER($B{r})),"日期格式不对",'
        f'IF($J{r}="","未填科目",'
        f'IF(ISNA(MATCH($J{r},{ACC_O},0)),"科目不在基础设置里",'
        f'IF(OR(AND($K{r}<>"",NOT(ISNUMBER($K{r}))),AND($L{r}<>"",NOT(ISNUMBER($L{r})))),"金额是文字，请改成数字",'
        f'IF(AND(N($K{r})=0,N($L{r})=0),"借贷都为空",'
        f'IF(AND(N($K{r})<>0,N($L{r})<>0),"借方贷方不能同时填",'
        f'IF(AND($E{r}<>"",LEFT($V{r},1)="⚠"),"项目编号不存在","√")))))))))',
        font=F_TXT, fill=FILL_CHK)
for i, row in enumerate(JOUR_ROWS):
    r = JR0 + i
    for c, v in row.items():
        ws[f'{c}{r}'] = v
dv_list(ws, f'J{JR0}:J{JR1}', f'={ACC_O}', msg='科目必须是右边「基础设置」里有的，要加科目先在 O 列加')
ws.conditional_formatting.add(f'Y{JR0}:Y{JR1}',
    FormulaRule(formula=[f'AND($Y{JR0}<>"",LEFT($Y{JR0},1)<>"√")'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))
ws.conditional_formatting.add(f'B{JR0}:M{JR1}',
    FormulaRule(formula=[f'LEFT($Y{JR0},2)="√ "'], font=Font(color='808080', italic=True)))
ws.conditional_formatting.add(f'U{JR0}:U{JR1}',
    FormulaRule(formula=[f'AND(ISNUMBER($U{JR0}),$U{JR0}<0)'], font=Font(color='C00000', bold=True)))
# 筛选只包出纳那 12 列：右边的科目余额块和公式列不跟着排序、不会被打乱
ws.auto_filter.ref = f'B5:M{JR1}'
ws.freeze_panes = f'A{JR0}'
page(ws, titles='4:5')
print(f'  ✓ 资金日记账（出纳日记账格式；科目 {len(JOUR_ACCTS)} 个：{"、".join(JOUR_ACCTS)}；'
      f'9 月 {len(JOUR_ROWS)} 笔，容量 {JR1-JR0+1} 行）')

# ============================================================ 取数公式（带起止日期）
FB = FRNG('date')          # 日期
FI = FRNG('amt')           # 金额
FR = FRNG('mfee')          # 管理费
FS = FRNG('due')           # 应开成本票
FFLAG = FRNG('feeflag')    # 是否扣管理费
FV = FRNG('vat'); FW = FRNG('add')
FX = FRNG('stamp'); FY = FRNG('inc'); FZ = FRNG('taxsum')
FAA = FRNG('ar'); FAB = FRNG('ap')
FPEER = FRNG('peer')
FUBEL = FRNG('ubel')
FOCOL = FRNG('ocol')
def agg(val, side, kind, uref, pref=None, dr=None, cls=None, last=False, top=False,
        fee=None, src=None, peer=None, ocol=None):
    """side: 'E' 开票/付款方  'F' 收票/收款方  None 不限；dr=(起,止) 加日期区间；
       last=True 只取链条末层的销项票（我方该直接回成本票的那一层）；
       top=True 只取直接开给业主的那一层（这个工程的收入口径）；
       fee='扣管理费'/'不扣管理费' 按这一笔到底扣没扣管理费再筛一道；
       src=来源表 只取原对账明细某一张表来的行；
       peer=True 只取「收票方也是挂靠单位」的销项票；
       ocol='销售开票金额'/'已开成本票' 按这一笔在原对账明细里落在哪一列筛"""
    # side='U' 走「归属单位表」——这一笔算在哪家单位的对账表上。
    # 8 张单位竖版明细用的就是这个口径，汇总表也一律跟它走，三张表才对得上。
    ex = ''
    if side == 'U':   ex += f',{FUBEL},{uref}'
    elif side:        ex += f',{KE if side == "E" else KF},{uref}'
    if pref: ex += f',{KC},{pref}'
    if cls:  ex += f',{KQ},"{cls}"'
    if last: ex += f',{FAH},1'
    if top:  ex += f',{FTOP},1'
    if fee:  ex += f',{FFLAG},"{fee}"'
    if src:  ex += f',{FSRC},{src}'
    if peer: ex += f',{FPEER},"是"'
    if ocol: ex += f',{FOCOL},"{ocol}"'
    if dr:   ex += f',{FB},">="&{dr[0]},{FB},"<="&{dr[1]}'
    # 按「归属单位表」取数时不再加跨表去重那道筛子：每张原表各算各的，
    # 一张票在对方表里也出现，是原表本来就有的重复。【对账差异说明】和
    # 8 张单位明细都是这么算的，汇总表必须跟它们一个口径，否则三张表对不上。
    tail = '' if side == 'U' else f',{KZ},"是"'
    return f'SUMIFS({val},{KD},"{kind}"{ex}{tail})'

COLS_U = ['开票额\n(该单位开出)','应扣管理费','应到成本票','其中·按净额\n(开票额−管理费)','其中·按全额\n(不扣管理费)',
          '已收成本票','还差成本票\n(截至截止日)','应提税费','已交税','欠税未交\n(截至截止日)',
          '业主已付给\n挂靠单位','欠业主\n未付款余额\n(截至截止日)',
          '应收挂靠方\n金额','管理费\n已结算','扣质保金','挂靠单位\n已转我方','应收挂靠方\n余额\n(截至截止日)',
          '代收未转\n(含对方扣的\n管理费质保金)']
SUM_LAST = L(2 + len(COLS_U))          # 最后一个金额列 = T
BIZ = L(3 + len(COLS_U))               # 有无业务 = U
# 列字母（detail_cols 和首页卡片按名字取，别再写死字母）
SU = {k: L(3 + i) for i, k in enumerate(['sale', 'mfee', 'due', 'due_net', 'due_full', 'done', 'gap',
                                          'tax', 'paid', 'owed', 'a_got', 'a_left',
                                          'b_ar', 'b_fee', 'b_bond', 'b_got', 'b_left', 'transit'])}

def summary_formulas(uref, pref=None, dr=None):
    g = lambda v, s, k, **kw: agg(v, s, k, uref, pref, dr, **kw)
    # 一律按「归属单位表」取数，跟 8 张单位竖版明细、【对账差异说明】完全同一个口径。
    # 开票额只认原表「销售开票金额」那一列，已收成本票只认「已开成本票」那一列 ——
    # 原来把「收票方是本单位的销项票」也算进已收成本票，会跟对方表里同一张票撞成两笔。
    # 9.24：回款区拆成两组，跟单位明细表一致 ——
    #   欠业主未付款余额 ＝ 开票额 − 业主已付给挂靠单位
    #   应收挂靠方余额   ＝ 应收挂靠方金额（开票额里去掉选了「不回成本票」的，加上代垫应收）
    #                      − 管理费已结算 − 扣质保金 − 挂靠单位已转我方
    return [g(FI,'U','销项开票', ocol='销售开票金额'),
            g(FR,'U','销项开票'), g(FS,'U','销项开票'),
            g(FS,'U','销项开票', fee='扣管理费'), g(FS,'U','销项开票', fee='不扣管理费'),
            f'SUMIFS({FI},{FUBEL},{uref},{FOCOL},"已开成本票"' +
            (f',{KC},{pref}' if pref else '') +
            (f',{FB},">="&{dr[0]},{FB},"<="&{dr[1]}' if dr else '') + ')',
            None, g(FZ,'U','销项开票'), g(FI,'U','已交税'), None,
            g(FI,'U','挂靠代收'), None,
            g(FI,'U','销项开票', ocol='销售开票金额') + '-' + g(FI,'U','销项开票', ocol='销售开票金额', fee='不回成本票')
            + '+' + g(FI,'U','代垫应收'),
            g(FI,'U','管理费结算'), g(FI,'U','扣质保金'), g(FI,'U','我方收款'), None, None]

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

def bal_formulas(uref, pref, end, owed_all=False):
    """余额列：截至截止日期的余额（从第一笔累计到截止日，不管起始日期），
       不是期间内发生额相减 —— 否则 2025 年开票、2026 年回款的会算出负数"""
    drE = ('DATE(1900,1,1)', end)
    g = lambda v, s_, k, **kw: agg(v, s_, k, uref, pref, drE, **kw)
    dE = f',{FB},"<="&{end}'
    pE = f',{KC},{pref}' if pref else ''
    sale = g(FI, 'U', '销项开票', ocol='销售开票金额')
    got = g(FI, 'U', '挂靠代收'); ours = g(FI, 'U', '我方收款')
    fee = g(FI, 'U', '管理费结算'); bond = g(FI, 'U', '扣质保金')
    b_ar = (f'{sale}-{g(FI, "U", "销项开票", ocol="销售开票金额", fee="不回成本票")}+{g(FI, "U", "代垫应收")}')
    done = f'SUMIFS({FI},{FUBEL},{uref},{FOCOL},"已开成本票"{pE}{dE})'
    tax = (f'SUMIFS({FZ},{FUBEL},{uref}{pE}{dE})' if owed_all else g(FZ, 'U', '销项开票'))
    return dict(gap=f'{g(FS, "U", "销项开票")}-{done}',
                owed=f'{tax}-{g(FI, "U", "已交税")}',
                a_left=f'{sale}-{got}',
                b_left=f'{b_ar}-{fee}-{bond}-{ours}',
                transit=f'{got}-{ours}')

def detail_cols(ws, r, uref, pref, dr, last_col_letter, biz_col, owed_all=False):
    """一行 18 列的通用汇总公式 + 有无业务：发生额按期间取；余额列（还差成本票、欠税未交、
       欠业主未付款余额、应收挂靠方余额、代收未转）按截止日期取"""
    for i, f in enumerate(summary_formulas(uref, pref, dr)):
        if f is None: continue
        put(ws, f'{L(3+i)}{r}', f'=IF($A{r}="","",{f})', font=F_LINK, fmt=MONEY)
    c = SU
    bf = bal_formulas(uref, pref, dr[1], owed_all)
    put(ws, f'{c["gap"]}{r}', f'=IF($A{r}="","",ROUND({bf["gap"]},2))', font=F_TXT, fmt=MONEY)
    put(ws, f'{c["owed"]}{r}', f'=IF($A{r}="","",ROUND({bf["owed"]},2))', font=F_TXT, fmt=MONEY)
    put(ws, f'{c["a_left"]}{r}', f'=IF($A{r}="","",ROUND({bf["a_left"]},2))',
        font=F_TOT, fill=PatternFill('solid', fgColor='FCE4D6'), fmt=MONEY)
    put(ws, f'{c["b_left"]}{r}', f'=IF($A{r}="","",ROUND({bf["b_left"]},2))',
        font=F_TOT, fill=PatternFill('solid', fgColor='E2EFDA'), fmt=MONEY)
    put(ws, f'{c["transit"]}{r}', f'=IF($A{r}="","",ROUND({bf["transit"]},2))', font=F_TXT, fmt=MONEY)
    put(ws, f'{biz_col}{r}', f'=IF($A{r}="","",IF(ROUND(SUMPRODUCT(ABS(${c["sale"]}{r}:${c["transit"]}{r})),2)=0,"无","有"))',
        font=F_TXT, fill=FILL_CHK)
    ws.row_dimensions[r].height = 16

SUM_W = {'A': 11, 'B': 30}
for i in range(len(COLS_U)): SUM_W[L(3 + i)] = 14
SUM_W[BIZ] = 10

# ============================================================ 单位汇总
ws = wb.create_sheet(SH_SUM_U)
title(ws, '分表 · 按挂靠单位汇总', BIZ,
      '每个挂靠单位一行，行的先后跟你那份《对账明细》的表顺序一致（德誉嘉→迅驰→华城→金沁→湖南锦泰→康欣→安锐→杰华）。'
      '9.24 起回款按原表分两组：「欠业主未付款余额」＝开票额 − 业主已付给挂靠单位；'
      '「应收挂靠方余额」＝应收挂靠方金额（开票额去掉选了不回成本票的、加上代垫应收）− 管理费已结算 − 扣质保金 − 挂靠单位已转我方。'
      '跟各家单位明细第 7 行的合计一一对得上。填了期间：发生额按期间算，带「截至截止日」的几列是到截止日期为止的余额。'
      '注意合计行是各家直接相加：链条项目上下两层各记一次（跟原表一样一家一张表），同一笔钱会算两遍；'
      '金沁那一行的应收挂靠方是金沁欠康欣的（原表金沁表第二组叫「康欣回款情况」）。')
widths(ws, SUM_W)
DR = filter_band(ws, BIZ)
headers(ws, HR, 1, ['单位简称', '单位全称'] + COLS_U + ['有无业务'])
ws.row_dimensions[HR].height = 48
for i in range(U1 - U0 + 1):
    r = Q0 + i
    put(ws, f'A{r}', f'=IFERROR(INDEX({U_NAME},MATCH({i+1},{U_RANK},0)),"")', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",INDEX({QU}!$B${U0}:$B${U1},MATCH($A{r},{U_NAME},0)))',
        font=F_LINK, align=CL)
    detail_cols(ws, r, f'$A{r}', None, DR, SUM_LAST, BIZ, owed_all=True)
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
      '只过票不收费（计费方式选「不扣管理费」）的票会落到「按全额」那一列，不用再一笔笔翻。'
      '按项目看两个余额时要注意：原表有几笔回款是好几个项目一起收的、只记在其中一个项目上（比如康欣 7.8 到款 303,566.03 记在 A006），'
      '那个项目会显示负数、其余几个显示正数，加起来是对的，单位合计不受影响。')
widths(ws, SUM_W)
DR = filter_band(ws, BIZ)
headers(ws, HR, 1, ['项目编号', '项目简称'] + COLS_U + ['有无业务'])
ws.row_dimensions[HR].height = 48
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
    code = e.get('code') or PCODE.get(e['proj'], '')
    if not code: continue
    for u in {e.get('payer'), e.get('payee')}:
        if u: PAIR.setdefault(u, set()).add(code)
HOLD_UNITS = [u[0] for u in UNITS if u[2] == '挂靠单位']

ws = wb.create_sheet(SH_SUM_X)
title(ws, '分表 · 单位 × 项目明细', BIZ,
      '上面选一个挂靠单位，下面列出它名下每个项目的开票、成本票、回款、欠税情况。'
      '要给领导看逐笔的，翻后面的单位专表（××明细）—— 那几张是照你原对账表的样子竖着排的。')
widths(ws, SUM_W)
DR = filter_band(ws, BIZ, unit_default='康欣')
dv_list(ws, 'B3', f'={U_NAME}')
headers(ws, HR, 1, ['项目编号', '项目简称'] + COLS_U + ['有无业务'])
ws.row_dimensions[HR].height = 48
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
# 表头照你那份《对账明细》的两层结构来：第 5 行是组名（开票情况 / 交税情况 / 回款情况 …），
# 第 6 行才是具体列名；一笔业务一行，按日期排。用不上的列组按单位隐藏，
# 每家打开就是它自己那张表的样子。
SRC_SHEET = {u: u for u in ('德誉嘉', '迅驰', '华城', '金沁', '湖南锦泰', '康欣', '安锐', '杰华')}
SRC_XLS = {'德誉嘉': '德誉嘉 ', '迅驰': '迅驰', '华城': '华城', '金沁': '金沁',
           '湖南锦泰': '湖南锦泰', '康欣': '康欣', '安锐': '安锐', '杰华': '杰华电气'}
VD_HG, VD_HC = 5, 6            # 两层表头：组名行 / 列名行
VD_TOT = 7                     # 期间合计行（跟原表第 4 行同位置）
VD_0 = 8                       # 明细首行
VD_N = 400
VD_1 = VD_0 + VD_N - 1
# (列, 宽, 组名, 列名, 类型)
# 9.24：回款区照原表改成两组 ——
#   欠业主未付款（原表「××回款情况」）：应收工程款 / 业主扣质保金 / 开票已回款 / 开票未到账
#   应收挂靠方  （原表「泓普回款情况」）：应收金额 / 扣管理费 / 质保金 / 已到账 / 未到账余额
# 组名行显示的标题每家不一样（照原表），见 GROUP_TITLE
G_A, G_B = '欠业主未付款', '应收挂靠方'
VCOLS = [
    ('A', 6,  '', '序号', 'no'),        ('B', 11, '', '日 期', 'date'),
    ('C', 20, '', '项目简称', 'sname'), ('D', 44, '', '摘　　要', 'memo'),
    ('E', 10, '', '业务\n性质', 'ptype'),
    ('F', 10, '开票情况', '发票\n性质', 'inv'),
    ('G', 10, '开票情况', '收票\n单位', 'payee'),
    ('H', 14, '开票情况', '销售开票金额', 'sale'),
    ('I', 13, '开票情况', '应扣管理费', 'mfee'),
    ('J', 13, '开票情况', '应到成本票', 'due'),
    ('K', 13, '开票情况', '已到成本票', 'done'),
    ('L', 12, '开票情况', '代发工资', 'wage'),
    ('M', 13, '开票情况', '剩余开票金额', 'rest'),
    ('N', 12, '交税情况', '预收增值部', 'vat'),
    ('O', 12, '交税情况', '预收附加税', 'add'),
    ('P', 12, '交税情况', '预收印花税', 'stamp'),
    ('Q', 12, '交税情况', '预收所得税', 'inc'),
    ('R', 12, '交税情况', '应扣税费', 'taxsum'),
    ('S', 12, '交税情况', '已交税', 'paid'),
    ('T', 12, '交税情况', '欠税未交', 'owed'),
    ('U', 14, G_A, '应收工程款', 'a_ar'),
    ('V', 12, G_A, '业主扣\n质保金', 'a_bond'),
    ('W', 13, G_A, '开票已回款', 'a_got'),
    ('X', 15, G_A, '开票未到账\n(欠业主未付款余额)', 'a_left'),
    ('Y', 14, G_B, '应收金额', 'b_ar'),
    ('Z', 12, G_B, '扣管理费', 'b_fee'),
    ('AA', 11, G_B, '质保金', 'b_bond'),
    ('AB', 13, G_B, '已到账', 'b_got'),
    ('AC', 15, G_B, '未到账余额\n(应收挂靠方余额)', 'b_left'),
    ('AD', 12, '返管理费', '返管理费', 'reb'),
    ('AE', 12, '返管理费', '已收款', 'rebgot'),
    ('AF', 12, '返管理费', '未收款', 'rebleft'),
    ('AG', 13, '过账/合伙应付款', '应付款', 'ap'),
    ('AH', 12, '过账/合伙应付款', '已付款', 'appaid'),
    ('AI', 12, '过账/合伙应付款', '未付款', 'apleft'),
    ('AJ', 24, '', '备　注', 'note'), ('AK', 9, '', '在期间内', 'inrange'),
    ('AL', 7, '', '行指针', 'ptr'),
]
VD_LAST = 'AL'
VC = {k: c for c, _, _, _, k in VCOLS}          # 类型 → 列字母
VD_MONEY = [c for c, _, _, _, k in VCOLS
            if k in ('sale', 'mfee', 'due', 'done', 'wage', 'rest', 'vat', 'add', 'stamp',
                     'inc', 'taxsum', 'paid', 'owed', 'a_ar', 'a_bond', 'a_got', 'a_left',
                     'b_ar', 'b_fee', 'b_bond', 'b_got', 'b_left',
                     'reb', 'rebgot', 'rebleft', 'ap', 'appaid', 'apleft')]
# 滚动余额列：本行余额 ＝ 从第一行累计到本行的（加项之和 − 减项之和）
VD_RUN = {'rest': (['due'], ['done']),
          'owed': (['taxsum'], ['paid']),
          'a_left': (['a_ar'], ['a_got']),
          'b_left': (['b_ar'], ['b_fee', 'b_bond', 'b_got']),
          'rebleft': (['reb'], ['rebgot']), 'apleft': (['ap'], ['appaid'])}
GROUP_COLS = {}
for c, _, g, _, _ in VCOLS:
    if g: GROUP_COLS.setdefault(g, []).append(c)
# 两组回款区在每家表上的标题（照原表第 2 行）
# 金沁原表第二组是「康欣回款情况」—— 记的是金沁该付给康欣的钱，不是欠我方的
GROUP_B_NOTE = {'金沁': '金沁欠康欣'}
AP_TITLE = {'德誉嘉': '合伙项目应付款', '迅驰': '过账项目支付情况'}
UNIT_A_BOND = {'金沁'}
UNIT_NO_B_BOND = {'德誉嘉', '华城'}
GROUP_TITLE = {
    '德誉嘉': ('德誉嘉回款情况', '泓普回款情况'), '迅驰': ('迅驰（销方）回款情况', '回款情况'),
    '华城': ('华城回款情况', '泓普回款情况'), '金沁': ('金沁回款情况', '康欣回款情况'),
    '湖南锦泰': ('工程款回款情况', '泓普回款情况'), '康欣': ('康欣回款情况', '泓普回款情况'),
    '安锐': ('工程款回款情况', '回款情况'), '杰华': ('工程款回款情况', '泓普回款情况')}

# 每行只定位一次（隐藏的 AG「行指针」），其余 30 列一律 INDEX(区间, 指针)，
# 不再每格算一次 MATCH —— 8 张表 × 400 行 × 30 列如果各算一次 MATCH，打开表会很慢
def vfetch(ptr, col_key):
    return f'IF({ptr}="","",INDEX({FRNG(col_key)},{ptr}))'

SUB_SHEETS = []
for u in [x[0] for x in UNITS if x[2] == '挂靠单位']:
    sh = SRC_SHEET[u]
    nm = f'{u}明细'
    ws = wb.create_sheet(nm)
    # 9.24 起：取哪家单位的数只看 B3「本表单位」，标题也跟着 B3 —— 以后新增挂靠单位，
    # 复制任意一张明细、改表名、B3 选新单位就是它的专表（步骤写在【操作流程】第八部分）
    title(ws, f'{u} · 对账明细（给领导看的逐笔明细）', VD_LAST,
          '表头照你那份《对账明细》的单位表来：第 5 行是组名、第 6 行是列名，'
          '一笔业务一行，按业务流水的录入顺序（历史行按日期）；回款区照原表分两组 ——「欠业主未付款」和「应收挂靠方」，两个余额逐行滚动。'
          '摘要前面带【业务类型】的是同一张原表行拆出来的回款 / 管理费，不是重复。明细全部显示不折叠；'
          '上面填年度或起止日期，第 7 行的合计只统计落在期间内的行（最后一列标是/否，'
          '不在期间内的行是灰的）。数据全部来自【业务流水】里归属到 B3「本表单位」的行，'
          '在那边改一笔，这里立刻跟着变 —— 包括你以后新录的。')
    ws['A1'] = '=$B$3&" · 对账明细（给领导看的逐笔明细）"'
    widths(ws, {c: w for c, w, _, _, _ in VCOLS})
    # 筛选带
    put(ws, 'A3', '本表单位', font=F_H2, fill=FILL_HDR2)
    put(ws, 'B3', u, font=Font(name='微软雅黑', size=11, bold=True, color='1F3864'),
        fill=FILL_HDR2, align=CL)
    dv_list(ws, 'B3', f'={U_NAME}', msg='从单位档案里选；新单位先到【单位档案】加一行、类型选「挂靠单位」')
    for lab, lc, col, fmt in [('年度', 'C', 'D', '0'), ('起始日期', 'E', 'F', DATE),
                              ('截止日期', 'G', 'H', DATE)]:
        put(ws, f'{lc}3', lab, font=F_H2, fill=FILL_HDR2)
        put(ws, f'{col}3', None, font=Font(name='微软雅黑', size=10, bold=True, color='0000C0'),
            fill=FILL_IN, fmt=fmt)
    put(ws, 'I3', '当前取数', font=F_H2, fill=FILL_HDR2)
    for ci in range(10, col_idx(VD_LAST) + 1):
        put(ws, f'{L(ci)}3', None, font=F_TOT, fill=FILL_CHK, align=CL)
    ws.merge_cells(f'J3:{VD_LAST}3')
    # 本表单位没选、不是挂靠单位、跟表名对不上（复制了表忘了改 B3）都在这里提示
    _sn = 'IFERROR(MID(CELL("filename",$A$1),FIND("]",CELL("filename",$A$1))+1,99),"")'
    put(ws, 'J3', '=IF($B$3="","⚠ 请在 B3 选本表单位",'
                  f'IF(COUNTIFS({U_NAME},$B$3,{U_TYPE},"挂靠单位")=0,'
                  '"⚠ 「"&$B$3&"」在【单位档案】里类型不是「挂靠单位」，这张表取不到数",'
                  f'IF(AND({_sn}<>"",ISERROR(FIND($B$3,{_sn}))),'
                  '"⚠ 表名跟本表单位「"&$B$3&"」对不上：复制过来的表要在 B3 选新单位",'
                  'IF(AND($D$3="",$F$3="",$H$3=""),"全部期间",'
                  'TEXT($A$4,"yyyy-mm-dd")&"  至  "&TEXT($B$4,"yyyy-mm-dd")))))',
        font=F_TOT, fill=FILL_CHK, align=CL)
    put(ws, 'A4', '=IF($F$3<>"",$F$3,IF($D$3<>"",DATE($D$3,1,1),DATE(1900,1,1)))', font=F_NOTE, fmt=DATE)
    put(ws, 'B4', '=IF($H$3<>"",$H$3,IF($D$3<>"",DATE($D$3,12,31),DATE(2199,12,31)))', font=F_NOTE, fmt=DATE)
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].hidden = True
    # 两层表头
    gt = {G_A: f'{GROUP_TITLE[u][0]}（欠业主未付款）',
          G_B: f'{GROUP_TITLE[u][1]}（{GROUP_B_NOTE.get(u, "应收挂靠方")}）',
          '过账/合伙应付款': AP_TITLE.get(u, '过账/合伙应付款')}
    for c, _, g, nmc, _ in VCOLS:
        put(ws, f'{c}{VD_HC}', nmc, font=F_HDR2, fill=FILL_HDR2, align=C)
        if not g:
            put(ws, f'{c}{VD_HG}', nmc, font=F_HDR2, fill=FILL_HDR2, align=C)
            ws.merge_cells(f'{c}{VD_HG}:{c}{VD_HC}')
        else:
            put(ws, f'{c}{VD_HG}', None, font=F_HDR2, fill=FILL_HDR2, align=C)
    for g, cols in GROUP_COLS.items():
        gfill = {G_A: PatternFill('solid', fgColor='FCE4D6'), G_B: PatternFill('solid', fgColor='E2EFDA')}.get(g, FILL_HDR2)
        for c in cols:
            ws[f'{c}{VD_HG}'].fill = gfill; ws[f'{c}{VD_HC}'].fill = gfill
        put(ws, f'{cols[0]}{VD_HG}', gt.get(g, g), font=F_HDR2, fill=gfill, align=C)
        if len(cols) > 1: ws.merge_cells(f'{cols[0]}{VD_HG}:{cols[-1]}{VD_HG}')
    ws.row_dimensions[VD_HG].height = 20
    ws.row_dimensions[VD_HC].height = 30
    PTR = VC['ptr']
    for i in range(VD_N):
        r = VD_0 + i
        n = i + 1
        ws[f'A{r}'] = n
        put(ws, f'A{r}', None, font=F_NOTE)
        put(ws, f'{PTR}{r}', f'=IFERROR(MATCH($B$3&"#"&$A{r},{FRNG("skey")},0),"")',
            font=F_NOTE, border=None)
        ptr = f'${PTR}{r}'
        get = lambda ck: vfetch(ptr, ck)
        kd = get('kind'); am = f'N({get("amt")})'; oc = get('ocol')
        is_sale = f'AND({kd}="销项开票",{oc}<>"已开成本票")'
        put(ws, f'B{r}', f'=IF({get("date")}="","",{get("date")})', font=F_LINK, fmt=DATE)
        put(ws, f'C{r}', f'=T({get("sname")})', font=F_LINK, align=CL)
        # 摘要：不是开票的那几类前面带上【业务类型】—— 同一张原表行拆出来的
        # 开票、业主付款、转我方、管理费结算几行摘要一模一样，不带类型看着像重复
        _kd = f'INDEX({FRNG("kind")},{ptr})'
        put(ws, f'D{r}', f'=IF({ptr}="","",IF(OR({_kd}="",{_kd}="销项开票"),"","【"&{_kd}&"】")'
                         f'&T(INDEX({FRNG("memo")},{ptr})))', font=F_LINK, align=CL)
        put(ws, f'E{r}', f'=IF($B{r}="","",IFERROR(INDEX({QP}!$D${P0}:$D${P1},'
                         f'MATCH({get("proj")},{P_CODE},0)),""))', font=F_TXT)
        put(ws, f'F{r}', f'=T({get("inv")})', font=F_LINK)
        put(ws, f'G{r}', f'=T({get("payee")})', font=F_LINK)
        put(ws, f'H{r}', f'=IF({is_sale},{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'I{r}', f'=N({get("mfee")})', font=F_LINK, fmt=MONEY)
        put(ws, f'J{r}', f'=N({get("due")})', font=F_LINK, fmt=MONEY)
        put(ws, f'K{r}', f'=IF({oc}="已开成本票",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'L{r}', f'=IF({kd}="工资扣抵",{am},0)', font=F_LINK, fmt=MONEY)
        for col, ck in (('N', 'vat'), ('O', 'add'), ('P', 'stamp'), ('Q', 'inc')):
            put(ws, f'{col}{r}', f'=N({get(ck)})', font=F_LINK, fmt=MONEY)
        put(ws, f'R{r}', f'=N({get("taxsum")})', font=F_TOT, fmt=MONEY)
        put(ws, f'S{r}', f'=IF({kd}="已交税",{am},0)', font=F_LINK, fmt=MONEY)
        # ---- 欠业主未付款：业主该付多少、已经付给挂靠单位多少 ----
        put(ws, f'{VC["a_ar"]}{r}', f'=IF({is_sale},{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'{VC["a_bond"]}{r}', f'=IF({kd}="业主扣质保金",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'{VC["a_got"]}{r}', f'=IF({kd}="挂靠代收",{am},0)', font=F_LINK, fmt=MONEY)
        # ---- 应收挂靠方：挂靠单位该给我方多少（选了「不回成本票」的票不算）、扣了什么、已到多少 ----
        put(ws, f'{VC["b_ar"]}{r}', f'=IF(AND({is_sale},{get("feeflag")}<>"不回成本票"),{am},0)'
                                    f'+IF({kd}="代垫应收",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'{VC["b_fee"]}{r}', f'=IF({kd}="管理费结算",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'{VC["b_bond"]}{r}', f'=IF({kd}="扣质保金",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'{VC["b_got"]}{r}', f'=IF({kd}="我方收款",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'{VC["reb"]}{r}', f'=N({get("ar")})', font=F_LINK, fmt=MONEY)
        put(ws, f'{VC["rebgot"]}{r}', f'=IF({kd}="其他应收收回",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'{VC["ap"]}{r}', f'=IF({kd}="其他应付发生",{am},0)', font=F_LINK, fmt=MONEY)
        put(ws, f'{VC["appaid"]}{r}', f'=IF(OR({kd}="其他应付支付",{kd}="其他应付扣税"),{am},0)',
            font=F_LINK, fmt=MONEY)
        for key, (plus, minus) in VD_RUN.items():
            f = '+'.join(f'SUM(${VC[k]}${VD_0}:${VC[k]}{r})' for k in plus)
            f += ''.join(f'-SUM(${VC[k]}${VD_0}:${VC[k]}{r})' for k in minus)
            put(ws, f'{VC[key]}{r}', f'=IF($B{r}="","",ROUND({f},2))',
                font=F_TOT if key in ('a_left', 'b_left') else F_TXT, fmt=MONEY)
        _nt = f'T({vfetch(ptr,"note")})'
        put(ws, f'{VC["note"]}{r}', f'=IF($B{r}="","",IF({vfetch(ptr,"srow")}="",IF({_nt}="","新录","新录｜"&{_nt}),'
                                    f'"原表 {SRC_XLS[u].strip()} 表第 "&{vfetch(ptr,"srow")}&" 行"'
                                    f'&IF(OR({_nt}="",LEFT({_nt},4)="历史导入"),"","｜"&{_nt})))',
            font=F_NOTE, align=CL)
        put(ws, f'{VC["inrange"]}{r}', f'=IF($B{r}="","",IF(AND($B{r}>=$A$4,$B{r}<=$B$4),"是","否"))',
            font=F_TXT, fill=FILL_CHK)
        ws.row_dimensions[r].height = 16
    # 合计行：只统计落在期间内的行；余额列直接写差额
    IR = VC['inrange']
    put(ws, f'A{VD_TOT}', '期间合计', font=F_TOT, fill=FILL_TOT)
    for c in 'BCEFG': put(ws, f'{c}{VD_TOT}', None, font=F_TOT, fill=FILL_TOT)
    put(ws, f'D{VD_TOT}', f'=COUNTIF(${IR}${VD_0}:${IR}${VD_1},"是")&" 笔（本表共 "'
                          f'&COUNTIF(${IR}${VD_0}:${IR}${VD_1},"是")+COUNTIF(${IR}${VD_0}:${IR}${VD_1},"否")'
                          f'&" 笔）"', font=F_TOT, fill=FILL_TOT, align=CL)
    RUNCOL = {VC[key]: pm for key, pm in VD_RUN.items()}
    _B = f'$B${VD_0}:$B${VD_1}'
    for c in VD_MONEY:
        if c in RUNCOL:
            # 余额列＝截至截止日期的余额（从第一笔累计到截止日），不是期间内的发生额相减 ——
            # 否则填了年度，2025 年开票、2026 年回款的项目会算出很大的负数
            plus, minus = RUNCOL[c]
            f = '+'.join(f'SUMIFS(${VC[k]}${VD_0}:${VC[k]}${VD_1},{_B},"<="&$B$4)' for k in plus)
            f += ''.join(f'-SUMIFS(${VC[k]}${VD_0}:${VC[k]}${VD_1},{_B},"<="&$B$4)' for k in minus)
            put(ws, f'{c}{VD_TOT}', f'=ROUND({f},2)', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
        else:
            put(ws, f'{c}{VD_TOT}',
                f'=ROUND(SUMIF(${IR}${VD_0}:${IR}${VD_1},"是",{c}${VD_0}:{c}${VD_1}),2)',
                font=F_TOT, fill=FILL_TOT, fmt=MONEY)
    put(ws, f'{VC["note"]}{VD_TOT}', '发生额只算「在期间内＝是」的行；余额列（剩余开票金额、欠税未交、两个余额、未收款、未付款）是截至截止日期的余额',
        font=F_NOTE, fill=FILL_TOT, align=CL)
    put(ws, f'{IR}{VD_TOT}', None, font=F_TOT, fill=FILL_TOT)
    put(ws, f'{PTR}{VD_TOT}', None, font=F_TOT, fill=FILL_TOT)
    ws.row_dimensions[VD_TOT].height = 20
    ws.conditional_formatting.add(f'A{VD_0}:{VD_LAST}{VD_1}',
        FormulaRule(formula=[f'${IR}{VD_0}="否"'], font=Font(color='A6A6A6')))
    # 按 3% 扣点子抵掉、实际没开的成本票（发票性质＝不开票）标红，跟原表一样一眼分得出来
    ws.conditional_formatting.add(f'D{VD_0}:D{VD_1} F{VD_0}:F{VD_1} K{VD_0}:K{VD_1}',
        FormulaRule(formula=[f'AND($F{VD_0}="不开票",$K{VD_0}<>0)'], font=Font(color='C00000', bold=True)))
    # 红冲（负数开票）整行标红
    ws.conditional_formatting.add(f'D{VD_0}:D{VD_1} H{VD_0}:H{VD_1}',
        FormulaRule(formula=[f'AND(ISNUMBER($H{VD_0}),$H{VD_0}<0)'], font=Font(color='C00000', bold=True)))
    # 这家单位用不上的列组直接隐藏，打开就是它自己那张表的样子
    mine = [e for e in EV if SHEET2UNIT.get(e['src'].split('!')[0].strip(),
                                            e['src'].split('!')[0].strip()) == sh]
    has_reb = any(e.get('rebate') for e in mine)
    has_ap = any(e['kind'].startswith('其他应付') for e in mine) or u == '迅驰'
    has_tax = any((e.get('tax_v') or 0) + (e.get('tax_s') or 0) + (e.get('tax_y') or 0)
                  + (e.get('tax_i') or 0) for e in mine)
    # 只隐藏这家单位原表里没有、历史上也一笔没有的整组（返管理费、过账/合伙应付款、交税）。
    # 回款两组、代发工资这几列一律显示 —— 以后录了扣质保金、工资扣抵，数就看得见
    hide = []
    if not has_reb: hide += GROUP_COLS['返管理费']
    if not has_ap: hide += GROUP_COLS['过账/合伙应付款']
    if not has_tax: hide += GROUP_COLS['交税情况']
    for c in hide: ws.column_dimensions[c].hidden = True
    ws.column_dimensions[PTR].hidden = True
    nrow = len(mine)
    ws.auto_filter.ref = f'A{VD_HC}:{IR}{VD_1}'
    ws.freeze_panes = f'C{VD_0}'
    page(ws, titles=f'{VD_HG}:{VD_HC}')
    ws.print_area = f'$A$1:${IR}${VD_1}'
    SUB_SHEETS.append(nm)
print(f'  ✓ {len(SUB_SHEETS)} 张单位竖版明细（两层表头，用不上的列组按单位隐藏）：{"、".join(SUB_SHEETS)}')

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
      '其他应收款＝挂靠单位按管理费率扣走后又答应返给我方的现金（德誉嘉 2026-7-22 以前的票扣 8% 返 4%，以后改成直接扣 4% 不返）。'
      '其他应付款＝收到的过账款、合伙项目要转给别人的钱，先减掉我方已代垫的税费，剩下的才是真正要付的。')
widths(ws, {'A':12,'B':24,'C':15,'D':14,'E':14,'F':11,'G':3,'H':14,'I':14,'J':14,'K':14,'L':11,'M':10})
DR = filter_band(ws, 'M')
ws.merge_cells('C4:F4'); ws.merge_cells('H4:L4')
headers(ws, HR, 1, ['单位简称','单位全称','其他应收\n返现发生','其他应收\n已收回','其他应收余额\n(截至截止日)','状态'])
ws.row_dimensions[HR].height = 34
put(ws, f'G{HR}', None, font=F_HDR, fill=FILL_HDR)
headers(ws, HR, 8, ['其他应付\n发生','其他应付\n减代扣税费','其他应付\n已支付','其他应付余额\n(截至截止日)','状态','有无往来'],
        fill=FILL_HDR, font=F_HDR)
for i in range(U1 - U0 + 1):
    r = Q0 + i
    put(ws, f'A{r}', f'=IFERROR(INDEX({U_NAME},MATCH({i+1},{U_RANK},0)),"")', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",INDEX({QU}!$B${U0}:$B${U1},MATCH($A{r},{U_NAME},0)))',
        font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($A{r}="","",{agg(FAA,"E","销项开票",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'D{r}', f'=IF($A{r}="","",{agg(FI,"E","其他应收收回",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    _drE = ('DATE(1900,1,1)', DR[1])       # 余额：截至截止日期
    put(ws, f'E{r}', f'=IF($A{r}="","",ROUND({agg(FAA,"E","销项开票",f"$A{r}",None,_drE)}'
                     f'-{agg(FI,"E","其他应收收回",f"$A{r}",None,_drE)},2))', font=F_TOT, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",IF(AND(ROUND($C{r},2)=0,ROUND($E{r},2)=0),"—",IF(ABS($E{r})<1,"✓ 已收清","未收回")))',
        font=F_TXT, fill=FILL_CHK)
    put(ws, f'G{r}', None, border=None)
    put(ws, f'H{r}', f'=IF($A{r}="","",{agg(FI,"U","其他应付发生",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'I{r}', f'=IF($A{r}="","",{agg(FI,"U","其他应付扣税",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",{agg(FI,"U","其他应付支付",f"$A{r}",None,DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'K{r}', f'=IF($A{r}="","",ROUND({agg(FI,"U","其他应付发生",f"$A{r}",None,_drE)}'
                     f'-{agg(FI,"U","其他应付扣税",f"$A{r}",None,_drE)}-{agg(FI,"U","其他应付支付",f"$A{r}",None,_drE)},2))',
        font=F_TOT, fmt=MONEY)
    put(ws, f'L{r}', f'=IF($A{r}="","",IF(AND(ROUND($H{r},2)=0,ROUND($K{r},2)=0),"—",IF(ABS($K{r})<1,"✓ 已付清","未付清")))',
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
                      '应付余额\n(截至截止日)','状态','该项目\n我方已交税','','','','有无往来'])
for i in range(P1 - P0 + 1):
    r, pr = PR_0 + i, P0 + i
    put(ws, f'A{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'B{r}', f'=IF($A{r}="","",{QP}!$C{pr})', font=F_LINK, align=CL)
    put(ws, f'C{r}', f'=IF($A{r}="","",{QP}!$D{pr})', font=F_LINK)
    put(ws, f'D{r}', f'=IF($A{r}="","",{agg(FI,None,"其他应付发生",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'E{r}', f'=IF($A{r}="","",{agg(FI,None,"其他应付扣税",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    put(ws, f'F{r}', f'=IF($A{r}="","",{agg(FI,None,"其他应付支付",None,f"$A{r}",DR)})', font=F_LINK, fmt=MONEY)
    _drE = ('DATE(1900,1,1)', DR[1])
    put(ws, f'G{r}', f'=IF($A{r}="","",ROUND({agg(FI,None,"其他应付发生",None,f"$A{r}",_drE)}'
                     f'-{agg(FI,None,"其他应付扣税",None,f"$A{r}",_drE)}-{agg(FI,None,"其他应付支付",None,f"$A{r}",_drE)},2))',
        font=F_TOT, fmt=MONEY)
    put(ws, f'H{r}', f'=IF($A{r}="","",IF(AND(ROUND($D{r},2)=0,ROUND($G{r},2)=0),"—",IF(ABS($G{r})<1,"✓ 已付清","未付清")))',
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
# 资金日记账（出纳日记账格式）：E 项目编号 · F 费用类型 · L 贷方（付出去的钱）· X 计入
JJ2 = JRNG('L'); JD2 = JRNG('E'); JE2 = JRNG('F'); JB2 = JRNG('B')
DJ2 = f',{JRNG("X")},1,{JB2},">="&{DR[0]},{JB2},"<="&{DR[1]}'
# 这些费用类型不算工程实际成本：要么已经在管理费/税费里算过，要么根本不是成本
NOT_COST = ['税费', '管理费', '借款', '还借款', '转备用金', '其他应付支付', '其他应收收回']
# 借方（收进来的钱）只有「退款」才冲减支出；这些收入类的借方一律不冲 —— 否则工程回款会把项目支出冲没
NOT_COST_DR = NOT_COST + ['工程回款', '货款', '考核款']
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
    JK2 = JRNG('K')
    cost = (f'SUMIFS({JJ2},{JD2},$A{r}{DJ2})-SUMIFS({JK2},{JD2},$A{r}{DJ2})' + ''.join(
        f'-SUMIFS({JJ2},{JD2},$A{r},{JE2},"{t}"{DJ2})' for t in NOT_COST)
        + ''.join(f'+SUMIFS({JK2},{JD2},$A{r},{JE2},"{t}"{DJ2})' for t in NOT_COST_DR))
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
put(ws, f'Q{TR}', f'=IF(COUNTA({JRNG("E")})=0,'
                  f'"⚠ 日记账还没有一行填项目编号，「项目实际支出」整列是 0",'
                  f'"日记账已有 "&COUNTA({JRNG("E")})&" 行填了项目编号")',
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
    '应提税费取【业务流水】税费录入区（P~S）逐笔填的数（不管交没交）。项目实际支出取【资金日记账】里填了本项目编号的支出，'
    '已扣掉 ' + '、'.join(NOT_COST) + ' 这几类（要么已在管理费/税费里算过，要么不是工程成本）。'
    '「应转他方」＝【往来台账】里这个项目的其他应付发生额减去我方代垫税费；'
    '这一列吃掉大半毛利的项目状态标「过账为主」（比如迅驰过账的那 8 个机械费项目，'
    '钱本来就不是我们的，原表把该扣的 2,437.03 元税费记在单位层面没落到项目上，所以这几行会带个小负数）。　　'
    '注意：项目实际支出认的是【资金日记账】E 列「项目编号」（填 A001 这种编号，跟项目档案一致），'
    '取的是贷方减借方（付出去的钱减退回来的钱；工程回款、货款、考核款这类收入不冲减），不含 ' + '、'.join(NOT_COST) + '；'
    '出纳那边这一列空着的，这里就是 0。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{PRF_1+2}:R{PRF_1+2}')
ws.freeze_panes = f'C{PRF_0}'; page(ws, titles=f'{HR}:{HR}')
print('  ✓ 项目利润（自营 / 合伙分开小计）')

# ============================================================ 税费台账
ws = wb.create_sheet(SH_TAX)
title(ws, '分表 · 税费台账', 'K',
      '按单位汇总的税费。这一轮起，应提税费不再按单位档案的参数推算，而是【业务流水】右边'
      '「税费录入区」四列按实际逐笔录进去的；历史行照你那份《对账明细》9.13 版逐笔写死，'
      '所以这张表和【单位汇总】的欠税跟原表逐家一分不差（填了期间：应提、已交按期间算，欠税是截至截止日期的余额）。统计口径是「这一笔算在哪家单位的'
      '对账表上」（业务流水最右边隐藏的「归属单位表」列），不是按开票方 —— 因为原表就是一家一张表。'
      '上面可按年度或起止日期取数。')
widths(ws, {'A':13,'B':26,'C':15,'D':14,'E':14,'F':14,'G':15,'H':14,'I':15,'J':12,'K':10})
DR = filter_band(ws, 'K')
headers(ws, HR, 1, ['单位','单位全称','应提增值税','应提附加税','应提印花税','应提所得税','应提税费合计',
                    '已交税','欠税未交\n(截至截止日)','状态','有无业务'])
ws.row_dimensions[HR].height = 34
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
    _de = f',{FB},"<="&{DR[1]}'          # 欠税是余额：从第一笔累计到截止日期，跟单位汇总同一个口径
    put(ws, f'I{r}', f'=IF($A{r}="","",ROUND(SUMIFS({FZ},{FRNG("ubel")},$A{r}{_de})'
                     f'-SUMIFS({FI},{KD},"已交税",{FRNG("ubel")},$A{r}{_de}),2))', font=F_TOT, fmt=MONEY)
    put(ws, f'J{r}', f'=IF($A{r}="","",IF(AND($G{r}=0,$H{r}=0,$I{r}=0),"—",'
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
# 9.24 起取自出纳日记账格式的【资金日记账】：F 费用类型 × J 科目，借方＝收入、贷方＝支出。
# 科目（账户）不再写死成泓普/仟茂/现金，表头直接取日记账「基础设置」那一列，加了科目这里自动多一列。
ws = wb.create_sheet(SH_EXP)
EX_LAST = 'Q'
title(ws, '分表 · 费用统计', EX_LAST,
      '取自【资金日记账】。左边按费用类型 × 科目统计净支出（贷方 − 借方），右边按项目统计支出。'
      '上面可按年度或起止日期取数。费用类型那一列是淡黄色的，出纳用了新的类型名在空格里补上就行；'
      '没补的也不会丢，统一落在最下面「上面没列到的类型」那一行。')
widths(ws, {'A': 16, **{L(2 + j): 12 for j in range(N_ACCT)}, 'L': 13, 'M': 13, 'N': 2,
            'O': 10, 'P': 30, 'Q': 14})
DRJ = filter_band(ws, EX_LAST, note='留空＝全部期间；填了年度就按整年取数，另填起止日期则以起止为准',
                  cnt_label='资金日记账', cnt_rng=JRNG('B'))
JK, JL, JB, JJC, JF, JE_ = JRNG('K'), JRNG('L'), JRNG('B'), JRNG('J'), JRNG('F'), JRNG('E')
DJ = f',{JRNG("X")},1,{JB},">="&{DRJ[0]},{JB},"<="&{DRJ[1]}'
headers(ws, HR, 1, ['费用类型'])
for j in range(N_ACCT):
    c = L(2 + j)
    put(ws, f'{c}{HR}', f'=IF({QJ}!$O${JR0 + j}="","",{QJ}!$O${JR0 + j})', font=F_HDR, fill=FILL_HDR)
headers(ws, HR, 12, ['收入合计\n(借方)', '支出合计\n(贷方)'])
put(ws, f'N{HR}', None, font=F_HDR, fill=FILL_HDR)
headers(ws, HR, 15, ['项目编号', '项目简称', '项目实际支出\n(贷−退款，不含税费/借还款等)'])
ws.row_dimensions[HR].height = 30
ETYPES = ['工程回款', '借款', '还借款', '工资', '社保', '福利费', '餐费', '车辆费用', '燃油费', '运费',
          '材料费', '耗材费用', '办公费用', '维修费', '青苗费', '饮用水费用', '医疗费', '租赁费', '手续费',
          '账户年费', '税费', '管理费', '考核款', '货款', '转备用金', '其他应收收回', '其他应付支付', '其他']
N_ET = 40
E_0, E_1 = Q0, Q0 + N_ET - 1
E_REST = E_1 + 1
for i in range(N_ET):
    r = E_0 + i
    put(ws, f'A{r}', ETYPES[i] if i < len(ETYPES) else None, font=F_IN, fill=FILL_IN, align=CL)
    for j in range(N_ACCT):
        c = L(2 + j)
        put(ws, f'{c}{r}', f'=IF(OR($A{r}="",{c}${HR}=""),"",SUMIFS({JL},{JF},$A{r},{JJC},{c}${HR}{DJ})'
                           f'-SUMIFS({JK},{JF},$A{r},{JJC},{c}${HR}{DJ}))', font=F_LINK, fmt=MONEY)
    put(ws, f'L{r}', f'=IF($A{r}="","",SUMIFS({JK},{JF},$A{r}{DJ}))', font=F_LINK, fmt=MONEY)
    put(ws, f'M{r}', f'=IF($A{r}="","",SUMIFS({JL},{JF},$A{r}{DJ}))', font=F_LINK, fmt=MONEY)
    put(ws, f'N{r}', None, border=None)
    ws.row_dimensions[r].height = 16
put(ws, f'A{E_REST}', '上面没列到的类型', font=F_TOT, fill=FILL_IN, align=CL)
for j in range(N_ACCT):
    c = L(2 + j)
    put(ws, f'{c}{E_REST}', f'=IF({c}${HR}="","",ROUND(SUMIFS({JL},{JJC},{c}${HR}{DJ})'
                            f'-SUMIFS({JK},{JJC},{c}${HR}{DJ})-SUM({c}{E_0}:{c}{E_1}),2))',
        font=F_TOT, fill=FILL_IN, fmt=MONEY)
put(ws, f'L{E_REST}', f'=ROUND(SUMIFS({JK}{DJ})-SUM(L{E_0}:L{E_1}),2)', font=F_TOT, fill=FILL_IN, fmt=MONEY)
put(ws, f'M{E_REST}', f'=ROUND(SUMIFS({JL}{DJ})-SUM(M{E_0}:M{E_1}),2)', font=F_TOT, fill=FILL_IN, fmt=MONEY)
put(ws, f'N{E_REST}', None, border=None)
ws.conditional_formatting.add(f'B{E_REST}:M{E_REST}',
    FormulaRule(formula=[f'AND(ISNUMBER(B{E_REST}),ABS(B{E_REST})>0.005)'], fill=FILL_WARN,
                font=Font(color='9C0006', bold=True)))
put(ws, f'A{TR}', '合  计', font=F_TOT, fill=FILL_TOT)
for j in range(N_ACCT + 2):
    c = L(2 + j)
    put(ws, f'{c}{TR}', f'=SUM({c}{E_0}:{c}{E_REST})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
put(ws, f'N{TR}', None, border=None)
put(ws, f'O{TR}', '合  计', font=F_TOT, fill=FILL_TOT); put(ws, f'P{TR}', None, font=F_TOT, fill=FILL_TOT)
put(ws, f'Q{TR}', f'=SUM(Q{Q0}:Q{QP_1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)
ws.row_dimensions[TR].height = 20
for i in range(P1 - P0 + 1):
    r, pr = Q0 + i, P0 + i
    put(ws, f'O{r}', f'=IF({QP}!$A{pr}="","",{QP}!$A{pr})', font=F_LINK)
    put(ws, f'P{r}', f'=IF($O{r}="","",{QP}!$C{pr})', font=F_LINK, align=CL)
    # 跟【项目利润】「项目实际支出」同一个口径：贷方 − 借方（退款冲减），不算税费 / 管理费 / 借还款这类
    _pc = (f'SUMIFS({JL},{JE_},$O{r}{DJ})-SUMIFS({JK},{JE_},$O{r}{DJ})'
           + ''.join(f'-SUMIFS({JL},{JE_},$O{r},{JF},"{t}"{DJ})' for t in NOT_COST)
           + ''.join(f'+SUMIFS({JK},{JE_},$O{r},{JF},"{t}"{DJ})' for t in NOT_COST_DR))
    put(ws, f'Q{r}', f'=IF($O{r}="","",ROUND({_pc},2))', font=F_LINK, fmt=MONEY)
    if r > E_REST: ws.row_dimensions[r].height = 16
# 科目不在基础设置里的：只进收入 / 支出合计，进不了任何一个科目列 —— 单独列一行提醒
E_ACC = E_REST + 1
put(ws, f'A{E_ACC}', '其中：科目不在基础设置里的', font=F_NOTE, align=CL)
for j in range(N_ACCT): put(ws, f'{L(2 + j)}{E_ACC}', None, font=F_NOTE)
put(ws, f'L{E_ACC}', f'=ROUND(SUMIFS({JK}{DJ})-SUMPRODUCT(SUMIFS({JK},{JJC},$B${HR}:$K${HR}{DJ})),2)',
    font=F_NOTE, fmt=MONEY)
put(ws, f'M{E_ACC}', f'=ROUND(SUMIFS({JL}{DJ})-SUMPRODUCT(SUMIFS({JL},{JJC},$B${HR}:$K${HR}{DJ})),2)',
    font=F_NOTE, fmt=MONEY)
ws.conditional_formatting.add(f'A{E_ACC}:M{E_ACC}',
    FormulaRule(formula=[f'OR(ABS(N($L{E_ACC}))>0.005,ABS(N($M{E_ACC}))>0.005)'], fill=FILL_WARN,
                font=Font(color='9C0006', bold=True)))
put(ws, f'A{QP_1+2}', '左表的科目表头取自【资金日记账】O 列「基础设置」；费用类型要跟出纳账里写的一字不差'
                      '（比如「耗材费用」和「耗材费」算两类），没列到的类型落在「上面没列到的类型」那一行。'
                      '右表「项目实际支出」跟【项目利润】同一个口径。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells(f'A{QP_1+2}:{EX_LAST}{QP_1+2}')
ws.freeze_panes = f'B{Q0}'; page(ws, titles=f'{HR}:{HR}')
print('  ✓ 税费台账 / 代收台账 / 费用统计')


# ============================================================ 对账差异说明
# 逐家逐列拿你那份《对账明细》9.13 版的合计行跟系统算出来的数比，差在哪、为什么差，一张表说清楚
ws = wb.create_sheet(SH_DIFF)
title(ws, '对账差异说明（系统 ↔ 原《对账明细》9.13 版）', 'G')
put(ws, 'A2', '左边是你原表每张表最上面那行合计，右边是系统按「来源表＝这张表」算出来的同口径数（只算原表来的行）。'
              '差额绝对值小于 0.05 元的算对上了（原表本身有四舍五入尾差）；标「△ 已查明」的是原表自己手填、漏填造成的，原因写在说明列。'
              '注意：单位明细第 7 行、单位汇总还含 9.13 以后新录的行，所以会比这里多 —— 华城多 9.23 那 4 张票 35,686.00，'
              '德誉嘉多 9.20 红冲重开的 40.00。最下面第二块逐笔列了 9.24 对业务流水的每一处改动。这张表不参与任何计算，纯粹给你核对用。',
    font=F_NOTE, align=CL, border=None)
ws.merge_cells('A2:G2')
widths(ws, {'A': 12, 'B': 18, 'C': 16, 'D': 16, 'E': 13, 'F': 12, 'G': 60})
DF_H = 4
headers(ws, DF_H, 1, ['单位', '对账项目', '原表合计', '系统合计', '差额', '结论', '说明'])
SRC_OF = dict(SRC_XLS)                 # 单位简称 → 原工作簿里的表名
SRC_KEY = {k: k for k in SRC_OF}       # 业务流水「来源表」列存的是单位简称
def dsum(valkey, sh, kind=None, ocol=None, fee=None):
    ex = ''
    if kind: ex += f',{FRNG("kind")},"{kind}"'
    if ocol: ex += f',{FRNG("ocol")},"{ocol}"'
    if fee:  ex += f',{FRNG("feeflag")},"{fee}"'
    return f'SUMIFS({FRNG(valkey)},{FRNG("src")},"{sh}"{ex})'
_sale = lambda sh: dsum('amt', sh, '销项开票', '销售开票金额')
_a_left = lambda sh: f'{_sale(sh)}-{dsum("amt", sh, "挂靠代收")}'
_b_ar = lambda sh: (f'{_sale(sh)}-{dsum("amt", sh, "销项开票", "销售开票金额", "不回成本票")}'
                    f'+{dsum("amt", sh, "代垫应收")}')
_b_left = lambda sh: (f'{_b_ar(sh)}-{dsum("amt", sh, "管理费结算")}-{dsum("amt", sh, "扣质保金")}'
                      f'-{dsum("amt", sh, "我方收款")}')
DF_ITEMS = [
    ('销售开票金额',     _sale),
    ('应扣管理费',       lambda sh: dsum('mfee', sh, '销项开票', '销售开票金额')),
    ('应到成本票',       lambda sh: dsum('due', sh, '销项开票', '销售开票金额')),
    ('已到成本票',       lambda sh: dsum('amt', sh, None, '已开成本票')),
    ('应扣税费',         lambda sh: dsum('taxsum', sh)),
    ('已交税',           lambda sh: dsum('amt', sh, '已交税')),
    ('欠业主·应收工程款', _sale),
    ('欠业主·业主扣质保金', lambda sh: dsum('amt', sh, '业主扣质保金')),
    ('欠业主·开票已回款', lambda sh: dsum('amt', sh, '挂靠代收')),
    ('欠业主未付款余额',  _a_left),
    ('应收挂靠方·应收金额', _b_ar),
    ('应收挂靠方·扣管理费', lambda sh: dsum('amt', sh, '管理费结算')),
    ('应收挂靠方·质保金',  lambda sh: dsum('amt', sh, '扣质保金')),
    ('应收挂靠方·已到账',  lambda sh: dsum('amt', sh, '我方收款')),
    ('应收挂靠方余额',    _b_left),
    ('代发工资',         lambda sh: dsum('amt', sh, '工资扣抵')),
    ('返管理费',         lambda sh: dsum('ar', sh, '销项开票')),
    ('合伙项目应付款',   lambda sh: dsum('amt', sh, '其他应付发生')),
]
# 原表合计行（第 4 行）逐列的列号（8 张表各不相同，逐张核对过；回款两组 9.24 由 4 个代理逐行核过）
_AB = lambda a_ar, a_bond, a_got, a_left, b_ar, b_fee, b_bond, b_got, b_left: {
    k: v for k, v in (('欠业主·应收工程款', a_ar), ('欠业主·业主扣质保金', a_bond), ('欠业主·开票已回款', a_got),
                      ('欠业主未付款余额', a_left), ('应收挂靠方·应收金额', b_ar), ('应收挂靠方·扣管理费', b_fee),
                      ('应收挂靠方·质保金', b_bond), ('应收挂靠方·已到账', b_got), ('应收挂靠方余额', b_left))
    if v}
ORIG_COL = {
 '德誉嘉 ': dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10,
                返管理费=19, 合伙项目应付款=22, **_AB(12, None, 13, 14, 15, 16, None, 17, 18)),
 '迅驰':   dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=16, 已交税=17,
                **_AB(19, None, 20, 21, 22, 23, 24, 25, 26)),
 '华城':   dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=16, 已交税=17,
                **_AB(19, None, 20, 21, 22, 23, None, 24, 25)),
 '金沁':   dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=16, 已交税=17,
                **_AB(19, 20, 21, 22, 23, 25, 24, 26, 27)),
 '湖南锦泰': dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=17, 已交税=18,
                代发工资=11, **_AB(20, None, 21, 22, 23, 24, 25, 26, 27)),
 '康欣':   dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=16, 已交税=17,
                **_AB(19, None, 20, 21, 22, 24, 23, 25, 26)),
 '安锐':   dict(销售开票金额=7, 应扣管理费=8, 应到成本票=9, 已到成本票=10, 应扣税费=17, 已交税=18,
                代发工资=11, **_AB(20, None, 21, 22, 23, 24, 25, 26, 27)),
 '杰华电气': dict(销售开票金额=7, 应扣管理费=9, 应到成本票=10, 已到成本票=11, 应扣税费=18, 已交税=19,
                代发工资=12, **_AB(21, None, 22, 23, 24, 25, 26, 27, 28)),
}
# 已知的、查清了原因的差异 / 补录说明（逐行对过原表，写在「说明」列）
_SEG = '原表逐行余额是按项目分段重新起算的，只比合计行；系统按日期连续累计，最后一行就等于合计'
DF_NOTE = {
    ('康欣', '已到成本票'): '9.24 补录了原表「已提供成本票」列 7 笔（6 笔扣点 ＋ 2025.12.08 大太老旧路 203,652）。'
                           '原表 H68 列的「未开票按3%扣点子」一共 9 笔 287,173.15（×3%＝8,615.19，就是第 21 行已到账那笔），'
                           '9 笔发票性质都记「不开票」，康欣明细里标红',
    ('华城', '欠业主·应收工程款'): '原表这一列是手填的：第 18~22 行（7.20/7.22 开给民能的 5 张票，共 24,438.86）漏填，'
                                  '第 15 行把 6,701.38 写成 6,702.38（多 1.00）。系统逐笔取开票额，以系统为准',
    ('华城', '欠业主未付款余额'): '差额来源同上一行（原表第 18~22 行漏填、第 15 行多写 1.00）',
    ('华城', '应收挂靠方·应收金额'): '含原表第 14 行手填的代垫税费 16,883.46（9.24 按「代垫应收」补录）',
    ('迅驰', '应收挂靠方·应收金额'): '含原表第 69 行手填的应退税费 373.3（9.24 按「代垫应收」补录，同一行已到账 373.3 一进一出）',
    ('康欣', '应收挂靠方·应收金额'): '原表第 34、64、65 行（合伙、不回成本票）不计应收，系统同口径；'
                                  '第 58 行手填的 408.77 按「代垫应收」补录（数值＝金沁表该票应扣税费，原表没写说明，请核实）',
    ('金沁', '欠业主·业主扣质保金'): '9.24 按原表补录 6 笔「业主扣质保金」（第 5 行 521.43；第 30~34 行按开票额×3%，系统取到分）',
    ('金沁', '欠业主未付款余额'): '原表合计行＝应收−已回款（不扣质保金），系统同口径；原表逐行公式另扣了质保金，逐行跟合计行对不上',
    ('杰华', '应收挂靠方·扣管理费'): '原表「扣管理费10%」列一格没填，可合计行余额公式 AB4＝J4−Z4−AA4 用的是扣完 10% 的应收成本票；'
                                  '原表自己的逐行余额（AB9＝119,004.86）又没扣，原表前后不一致。9.24 按合计行口径补录了 9.11 管理费结算 12,230.57，'
                                  '这一列因此比原表多、余额跟合计行对上 —— 请跟杰华核实这 10% 是不是已经按扣了',
    ('杰华', '应收挂靠方余额'): '按原表合计行 AB4 口径（已扣 10% 管理费）；原表逐行余额是 119,004.86。请核实',
    ('湖南锦泰', '应收挂靠方余额'): '原表合计行按应扣管理费 16,389.33 算、实扣手填 16,389.36，尾差 0.03；' + _SEG,
    ('安锐', '欠业主·开票已回款'): '原表第 9 行手填的 2,728.39 正好等于 5 个项目应扣管理费之和，德誉嘉表里找不到这笔付款，'
                                 '像是把管理费记进了已回款；照原表保留，请核实',
}
import openpyxl as _op
_ref = _op.load_workbook(os.path.join(HERE, '..', '参考', '原对账明细_9.13.xlsx'), data_only=True)

def _colsum(_ws, col):
    return round(sum(float(_ws.cell(row=_r, column=col).value)
                     for _r in range(5, _ws.max_row + 1)
                     if isinstance(_ws.cell(row=_r, column=col).value, (int, float))), 2)
# 金沁表有两栏质保金：T 列是民能扣金沁的，X 列是再往下扣到我方这一层的。
# 系统记的是 X 列那一层，T 列多出来的部分单独列出来说明。
JQ_BOND_UP = round(_colsum(_ref['金沁'], 20) - _colsum(_ref['金沁'], 24), 2)
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
        _dn0 = DF_NOTE.get((u, lab))
        put(ws, f'F{r}', f'=IF(ABS(N($E{r}))<0.05,"✓ 对上",' + ('"△ 已查明"' if _dn0 else '"✗ 有差异"') + ')',
            font=F_TXT, fill=FILL_CHK)
        _dn = DF_NOTE.get((u, lab)) or (_SEG if lab in ('欠业主未付款余额', '应收挂靠方余额') else None)
        if _dn:
            put(ws, f'G{r}', f'=IF(ABS(N($E{r}))<0.05,"✓ ","差异原因：")&"{_dn}"', font=F_NOTE, align=CL)
        else:
            put(ws, f'G{r}', f'=IF(ABS(N($E{r}))<0.05,"{NOTE_OK}（尾差是原表四舍五入）","请点开【业务流水】按来源表筛一下")',
                font=F_NOTE, align=CL)
        ws.row_dimensions[r].height = 16
        r += 1
DF_1 = r - 1
ws.conditional_formatting.add(f'A{DF_0}:G{DF_1}',
    FormulaRule(formula=[f'$F{DF_0}="✗ 有差异"'], fill=FILL_WARN, font=Font(color='9C0006', bold=True)))

# ---- 第二块：9.24 这一轮对业务流水做的改动（逐笔） ----
ws.conditional_formatting.add(f'A{DF_0}:G{DF_1}',
    FormulaRule(formula=[f'$F{DF_0}="△ 已查明"'], fill=FILL_IN, font=Font(color='9C6500')))
r2 = DF_1 + 3
put(ws, f'A{r2}', '9.24 这一轮对【业务流水】做的改动 · 逐笔（备注列带「9.24」的行）', font=F_H2, fill=FILL_HDR2)
ws.merge_cells(f'A{r2}:G{r2}')
r2 += 1
headers(ws, r2, 1, ['流水行号', '归属单位', '日期', '业务类型', '金额', '项目', '改了什么 / 依据'])
hr2 = r2
r2 += 1
_wf = wb[SH_FLOW]
N924 = 0
for rr in range(F0, F0 + FLOW_USED):
    nt = str(_wf[f'{C_["note"]}{rr}'].value or '')
    if '9.24' not in nt and '红冲' not in nt: continue
    src = _wf[f'{C_["src"]}{rr}'].value
    put(ws, f'A{r2}', f'=HYPERLINK("#{SH_FLOW}!A{rr}","第 {rr} 行")', font=F_LINK)
    _pe, _pf = _wf[f'E{rr}'].value, _wf[f'F{rr}'].value
    put(ws, f'B{r2}', src if src else (_pe if _pe in HOLD_UNITS else _pf), font=F_TXT)
    put(ws, f'C{r2}', _wf[f'B{rr}'].value, font=F_TXT, fmt=DATE)
    put(ws, f'D{r2}', _wf[f'D{rr}'].value, font=F_TXT)
    put(ws, f'E{r2}', _wf[f'{A_}{rr}'].value, font=F_LINK, fmt=MONEY)
    put(ws, f'F{r2}', _wf[f'C{rr}'].value, font=F_TXT)
    put(ws, f'G{r2}', nt, font=F_NOTE, align=CL)
    ws.row_dimensions[r2].height = 30
    r2 += 1; N924 += 1
DUP = list(range(N924))
r2 += 2
put(ws, f'A{r2}', '本次新增维度核对', font=F_H2, fill=FILL_HDR2)
ws.merge_cells(f'A{r2}:G{r2}')
r2 += 1
headers(ws, r2, 1, ['单位', '核对项目', '原表合计', '系统合计', '差额', '结论', '说明'])
hr3 = r2
r2 += 1
EXTRA_CHK = [
    ('德誉嘉', '返管理费 4%（其他应收）', round(sum(e.get('rebate', 0) for e in EV), 2),
     f'=SUMIFS({FAA},{KD},"销项开票",{KE},"德誉嘉",{KZ},"是",{FSRC},"<>")', '原表「返管理费4%」那一列的合计'),
    ('迅驰', '过账应付工程款', round(sum(x['amt'] for x in PASS_ROWS), 2),
     f'=SUMIFS({FI},{KD},"其他应付发生",{FSRC},"迅驰")', '原表迅驰「过账应付工程款」列'),
    ('迅驰', '过账应扣税费', round(sum(x['amt'] for x in DED_ROWS), 2),
     f'=SUMIFS({FI},{KD},"其他应付扣税",{FSRC},"迅驰")', '原表迅驰「已付款」列'),
    ('康欣', '工资扣抵（劳务成本）', round(sum(x['amt'] for x in WAGE_ROWS), 2),
     f'=SUMIFS({FI},{KD},"工资扣抵",{FSRC},"总台账")', '原总台账「劳务成本·工资扣抵」列'),
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
    f'=IF(COUNTIF($F${DF_0}:$F${DF_1},"✗*")+COUNTIF($F${hr3+1}:$F${EX_1},"✗*")=0,"✓ 与原表一致",'
    f'"有 "&COUNTIF($F${DF_0}:$F${DF_1},"✗*")+COUNTIF($F${hr3+1}:$F${EX_1},"✗*")&" 项没对上，请看红色行")'
    f'&IF(COUNTIF($F${DF_0}:$F${DF_1},"△*")>0,"；另有 "&COUNTIF($F${DF_0}:$F${DF_1},"△*")'
    f'&" 项是原表自己手填漏填造成的差异，原因已写在说明列（黄色行）","")', font=F_TOT, fill=FILL_CHK)
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
put(ws, 'A2', '　只录三处：项目档案 + 业务流水（开票 · 成本票 · 回款 · 交税 · 往来款）+ 资金日记账（出纳账粘过来）　→　'
              '后面所有分表自动汇总。怎么录看【操作流程】',
    font=Font(name='微软雅黑', size=9, color='FFFFFF'), fill=PatternFill('solid', fgColor='2F5597'),
    align=CL, border=None)
ws.row_dimensions[2].height = 20
CARDS2 = [
 ('挂靠单位累计开票', f'={QSU}!${SU["sale"]}${TR}', 'D6E4F0'),
 ('我方应开成本票',   f'={QSU}!${SU["due"]}${TR}', 'D6E4F0'),
 ('我方已开成本票',   f'={QSU}!${SU["done"]}${TR}', 'D6E4F0'),
 ('还差成本票未开',   f'=SUMIF({QSU}!${SU["gap"]}${Q0}:${SU["gap"]}${QU_1},">0.5")', 'FCE4E4'),
 ('劳务票还差（按项目净额）', f'={QGAP}!$G${TR}', 'FCE4E4'),
 ('机械票还差（按项目净额）', f'={QGAP}!$J${TR}', 'FCE4E4'),
 ('其他应收·待返现',  f'={QCUR}!$E${TR}', 'FFF2CC'),
 ('其他应付·待转付',  f'={QCUR}!$K${TR}', 'FCE4E4'),
 ('欠业主未付款（各家相加）', f'={QSU}!${SU["a_left"]}${TR}', 'FCE4D6'),
 ('应收挂靠方（各家相加）', f'={QSU}!${SU["b_left"]}${TR}', 'E2EFDA'),
 ('挂靠单位已转出（各家相加）', f'={QSU}!${SU["b_got"]}${TR}', 'E2EFDA'),
 ('代收未转（各家相加）', f'={QSU}!${SU["transit"]}${TR}', 'FCE4E4'),
 ('自营项目 归属利润', f'={QPRF}!$P${TR+1}', 'E2EFDA'),
 ('合伙项目 归属利润', f'={QPRF}!$P${TR+2}', 'FFF2CC'),
 ('全部项目 票面毛利', f'={QPRF}!$K${TR}', 'D6E4F0'),
 ('项目实际支出',     f'={QPRF}!$M${TR}', 'FCE4E4'),
 ('应提税费',         f'={QTAX}!$G${TR}', 'FFF2CC'),
 ('已交税',           f'={QTAX}!$H${TR}', 'FFF2CC'),
 ('欠税未交',         f'={QTAX}!$I${TR}', 'FCE4E4'),
 ('资金余额（日记账各科目合计）', f'={QJ}!$S${J_TOT}', 'FFF2CC'),
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
    ('业务流水校验', f'=IF(SUMPRODUCT(({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1}<>"")'
                    f'*({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1}<>"√")'
                    f'*({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1}<>"链条待确认"))=0,'
                    f'"✓ 全部通过"&IF(COUNTIF({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1},"链条待确认")>0,'
                    f'"（另有 "&COUNTIF({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1},"链条待确认")&" 行是多链条项目，仅提示不影响汇总）",""),'
                    f'"✗ 有 "&SUMPRODUCT(({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1}<>"")'
                    f'*({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1}<>"√")'
                    f'*({QF}!${C_["chk"]}${F0}:${C_["chk"]}${F1}<>"链条待确认"))&" 行待修正（看业务流水 O 列红色的）")'),
    ('日记账校验',   f'=IF(LEFT({QJ}!$O${J_TOT+1},9)="⚠ 科目清单里有","✗ 科目清单里有重名的，资金余额会重复算，请到资金日记账 O 列删掉一个",'
                    f'IF(SUMPRODUCT(({JRNG("Y")}<>"")*(LEFT({JRNG("Y")},1)<>"√"))=0,'
                    f'"✓ 全部通过（"&COUNTIF({JRNG("Y")},"√")&" 笔，科目 "&COUNTA({QJ}!$O${JA0}:$O${JA1})&" 个）"'
                    f'&IF(LEFT({QJ}!$O${J_TOT+1},1)="⚠","；"&MID({QJ}!$O${J_TOT+1},3,60),""),'
                    f'"✗ 有 "&SUMPRODUCT(({JRNG("Y")}<>"")*(LEFT({JRNG("Y")},1)<>"√"))&" 行待修正（看日记账最右边校验列）"))'),
    ('项目档案校验', f'=IF(COUNTA({QP}!$A${P0}:$A${P1})-COUNTIF({QP}!$U${P0}:$U${P1},"√")'
                    f'-COUNTIF({QP}!$U${P0}:$U${P1},"待完善*")=0,'
                    f'"✓ 全部通过"&IF(COUNTIF({QP}!$U${P0}:$U${P1},"待完善*")>0,'
                    f'"（另有 "&COUNTIF({QP}!$U${P0}:$U${P1},"待完善*")&" 个项目的业主或一级单位还没填）",""),'
                    f'"✗ 有 "&(COUNTA({QP}!$A${P0}:$A${P1})-COUNTIF({QP}!$U${P0}:$U${P1},"√")'
                    f'-COUNTIF({QP}!$U${P0}:$U${P1},"待完善*"))&" 行待修正")'),
    # 按单位看（＝原表每张表的「剩余开票金额」）：只数还差的，多开的不拿来抵别家的
    ('成本票缺口',   f'=IF(SUMIF({QSU}!${SU["gap"]}${Q0}:${SU["gap"]}${QU_1},">0.5")<1,"✓ 已开齐",'
                    f'"还差 "&TEXT(SUMIF({QSU}!${SU["gap"]}${Q0}:${SU["gap"]}${QU_1},">0.5"),"#,##0.00")&" 元成本票没开，涉及 "'
                    f'&COUNTIF({QSU}!${SU["gap"]}${Q0}:${SU["gap"]}${QU_1},">0.5")&" 家单位（看【单位汇总】「还差成本票」列；'
                    f'按项目、按劳务/机械看【发票缺口】）")'),
    ('项目利润',     f'="全部 "&TEXT({QPRF}!$P${TR},"#,##0")&" 元（毛利率 "'
                    f'&TEXT({QPRF}!$L${TR},"0.0%")&"）；其中自营 "&TEXT({QPRF}!$P${TR+1},"#,##0")'
                    f'&" 元、合伙 "&TEXT({QPRF}!$P${TR+2},"#,##0")&" 元"'
                    f'&IF(COUNTIF({QPRF}!$Q${TR+3}:$Q${PRF_1},"亏损")>0,'
                    f'"；有 "&COUNTIF({QPRF}!$Q${TR+3}:$Q${PRF_1},"亏损")&" 个项目算下来是亏的","")'),
    ('表格容量',     f'=IF(AND(COUNT({QF}!$B${F0}:$B${F1})<{int((F1-F0+1)*0.9)},'
                    f'COUNT({JRNG("B")})<{int((JR1-JR0+1)*0.9)},'
                    f'COUNTA({QP}!$A${P0}:$A${P1})<{int((P1-P0+1)*0.9)}),'
                    f'"✓ 够用（业务流水 "&COUNT({QF}!$B${F0}:$B${F1})&"/{F1-F0+1} 行，'
                    f'日记账 "&COUNT({JRNG("B")})&"/{JR1-JR0+1} 行，'
                    f'项目 "&COUNTA({QP}!$A${P0}:$A${P1})&"/{P1-P0+1} 个）",'
                    f'"⚠ 快满了：业务流水 "&COUNT({QF}!$B${F0}:$B${F1})&"/{F1-F0+1}，'
                    f'日记账 "&COUNT({JRNG("B")})&"/{JR1-JR0+1}，'
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
 ('用工资表顶劳务票', '【业务流水】业务类型「工资扣抵」，开票/付款方＝泓普，收票/收款方＝挂靠单位，填项目编号和金额 —— '
                    '只有【发票缺口】里的劳务票还差会相应减少（单位明细的剩余开票金额跟原表一样不扣它）。'),
 ('业主付钱给挂靠单位', '【业务流水】选「挂靠代收」，付款方＝业主（或者上一层挂靠单位），收款方＝收钱的那家挂靠单位。这笔钱还没到我们手上。'),
 ('挂靠单位转钱给我们', '【业务流水】选「我方收款」（代发工资、税差抵扣的也算）；真到银行的那部分出纳照常记日记账。'),
 ('红冲 / 补成本票 / 改错', '原行不动，红冲另录负数 + 重开；补成本票另录一行「成本票」；千万别删整行。详见【操作流程】第四部分。'),
 ('德誉嘉返现', '2026-7-22 以前德誉嘉的票是扣 8% 返 4%（历史数据已录好，返现挂在【往来台账】其他应收款上）；'
               '以后的新票管理费率填 4%、返现率填 0。以前挂着的返现真收到时录一行「其他应收收回」冲掉。'),
 ('过账款 / 合伙分钱', '收到别人的过账款录「其他应付发生」；我方为这笔垫的税费录「其他应付扣税」；真转出去录「其他应付支付」。'
                      '这三类开票/付款方、收票/收款方都选那家过账单位（比如迅驰），真正的收款人写在摘要里。'
                      '【往来台账】③ 按项目算出「扣完税费后还该转给别人多少」。'),
 ('交税', '【业务流水】选「已交税」，开票/付款方选这笔税是替哪家挂靠单位交的（泓普代交的也选那家），收票/收款方选税局；'
        '该提多少税在开票那一行右边「税费录入区」四列按实际填。'),
 ('日常收付款', '出纳在自己那本《出纳日记账》里录，录完把 B~M 列复制，到【资金日记账】B 列第一个空行选择性粘贴「数值」。'
             '科目清单在 O 列「基础设置」，余额按科目自动算。'),
 ('要看结果', '【单位汇总】【项目汇总】【项目利润】【链条核算】【发票缺口】【往来台账】【税费台账】【费用统计】'
             '上面都有「年度 / 起止日期」，填了以后发生额只统计那一段，各种「余额」是截至截止日期的数。给领导看某一家，直接打印那家的单位专表。'),
 ('新增挂靠单位', '【单位档案】加一行（类型选「挂靠单位」），再复制一张「××明细」、改表名、B3 选新单位就行。详见【操作流程】第八部分。'),
]
r = NR + 1
for a, b in NAV2:
    put(ws, f'A{r}', a, font=F_TOT, fill=FILL_HDR2, align=CL)
    ws.merge_cells(f'A{r}:B{r}'); put(ws, f'B{r}', None, font=F_TOT, fill=FILL_HDR2)
    ws.merge_cells(f'D{r}:L{r}')
    put(ws, f'D{r}', b, font=F_TXT, align=CL)
    ws.row_dimensions[r].height = 30
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
 ('三个录入口', '【项目档案】加新项目；【业务流水】记开票、成本票、回款、交税、管理费结算、往来款；'
               '【资金日记账】是出纳账复制过来的（9.24 起改用出纳日记账的格式）。其余所有表都是自动算的。'
               '一步一步怎么录、各种场景的样例，看【操作流程】。'),
 ('连续多年怎么用', None),
 ('不用每年建新表', '业务流水和资金日记账各留了 2000 行，项目档案 300 个，够用四五年，按日期一直往下录就行。'
                  '首页「表格容量」那一行会盯着，用到九成会提醒。'
                  '每张查询表最上面都有「年度 / 起始日期 / 截止日期」：填 2026 就只看 2026 年，'
                  '填起止日期就只看那一段，三个都留空＝全部期间。'),
 ('空白行会折叠', '查询表只显示有数据的行，空白行已经隐藏起来。新增项目或换了查询区间之后，'
                 '点一下 数据 → 筛选 → 重新应用，行数就会跟着变。录入表（单位档案 / 项目档案 / 业务流水 / 资金日记账）'
                 '永远全部显示，方便往下录。'),
 ('合计在表头下面', '查询表的合计行在表头正下方（汇总表第 6 行，单位明细第 7 行），不用翻到底，筛选也不会影响它。'
                  '填了期间以后，发生额按期间算，各种「余额」按截止日期算（从第一笔累计到截止日）。'),
 ('管理费怎么定（这一轮改了）', None),
 ('费率改成逐笔手填', '【单位档案】的「默认管理费率 / 默认返现率 / 第二档」五列已经取消。'
                    '现在管理费率和返现率直接在【业务流水】那两个淡黄色格子里按这一笔的实际情况填，'
                    '后面所有表一律从业务流水取数。这样就不会再出现「档案填一个率、原表实际是另一个率」对不上的情况。'
                    '历史行已按你那份《对账明细》9.13 版逐笔写死。'),
 ('计费方式三选一', '【业务流水】新增「计费方式」列：'
                  '① 扣管理费 —— 正常业务，管理费＝金额×费率，应开成本票＝金额−管理费；'
                  '② 不扣管理费 —— 只借通道过票、一分不收（康欣老项目那种），应开成本票＝全额；'
                  '③ 不回成本票 —— 合伙方内部分成、根本不用我方回票，应开成本票＝0。'
                  '以前「费率填 0」和「忘了填」长得一模一样，现在分得清了，忘选会在校验列提示。'),
 ('应到成本票拆两列', '【项目汇总】【单位汇总】的「应到成本票」右边跟着两列：'
                    '「其中·按净额(开票额−管理费)」和「其中·按全额(不扣管理费)」—— '
                    '后一列说的是：这笔的应到成本票直接照开票额算，管理费另外结，不从成本票里扣。'
                    '只过票不收费的那部分也归在这一列，一眼看得见，'
                    '不用再一笔笔翻。没扣管理费那一列大于 0 会标黄。'),
 ('税费也改成按实际录', '【业务流水】右边「税费录入区」四列（预提增值税 / 附加 / 印花 / 所得）改成手工填，'
                      '历史行按你原表逐笔写死。以前是按【单位档案】「一家单位一个税率」推算的，'
                      '可华城同时开 3% 劳务票和 13% 设备票，一个率算不出两种票 —— 这就是原来欠税对不上的根子。'
                      '现在【单位汇总】的欠税跟原表完全一致：迅驰 0（不欠税）、华城 10,129.25、金沁 17,778.54、'
                      '康欣 16,304.66、湖南锦泰 −2,610.25（多交了）、安锐 0、杰华 15,759.04。'
                      '最右边「税费参考试算」那一列还是按单位档案的税率参数算，只给录入时对照，不参与任何汇总。'),
 ('单笔改价', '【业务流水】的「管理费率」「返现率」两列是手填的淡黄色格子，这一笔谈的是多少就填多少。'),
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
 ('其他应收·返现', '挂靠单位先全额扣管理费，再把其中一部分现金返给我们（德誉嘉 2026-7-22 以前扣 8% 返 4%，以后直接扣 4% 不返）。'
                 '开票那一行按手填的「返现率」挂一笔应收；真收到钱录一行「其他应收收回」冲掉。'
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
 ('实际支出要先挂项目', '【项目利润】的「项目实际支出」取【资金日记账】E 列「项目编号」填了的贷方金额，减去退款（借方；工程回款这类收入不减）。'
                      '出纳记账时这一列要填项目档案里的编号（A001 这种），填别的名字对不上、校验列会报「项目编号不存在」。'
                      '目前 9 月这 16 笔都没填项目编号，所以这一列暂时是 0。'),
 ('哪些不算工程成本', '日记账里记到项目名下、但费用类型是税费 / 管理费 / 借款 / 还借款 / 转备用金 / '
                    '其他应付支付 / 其他应收收回的，不计入项目实际支出 —— 要么已经在管理费和税费里算过一次，'
                    '要么根本不是工程成本。'),
 ('给领导的单位专表（这一轮改了）', None),
 ('一家一张 · 竖版逐笔', '每家挂靠单位一张专表（德誉嘉明细 / 迅驰明细 …，以后新增单位怎么加表见【操作流程】第八部分），表头照你那份《对账明细》的样子做，'
                          '「一笔业务一行」竖着排（按业务流水的录入顺序，历史行按日期），一行一行能跟原表对得上。'
                          '上面填年度或起止日期，第 7 行的发生额只统计落在期间内的行，余额列是截至截止日期的余额'
                          '（最后一列会标是/否，不在期间内的行显示成灰色）。'
                          '不是开票的行，摘要前面带【挂靠代收】【我方收款】这种类型标签，同一张原表行拆出来的几笔不会再看着像重复。'),
 ('两组回款（9.24 新增）', '照原表把回款区分成两组：「××回款情况（欠业主未付款）」＝应收工程款 / 业主扣质保金 / 开票已回款 / '
                        '开票未到账（欠业主未付款余额）；「泓普回款情况（应收挂靠方）」＝应收金额 / 扣管理费 / 质保金 / 已到账 / '
                        '未到账余额（应收挂靠方余额）。应收金额＝开票额 − 选了「不回成本票」的票 ＋ 代垫应收。'
                        '两个余额都是逐行滚动，最后一行就是截至今天的数；【单位汇总】【项目汇总】【单位项目明细】也加了这两个余额。'),
 ('数据从哪来', '每张单位专表取的是【业务流水】里「归属单位表＝这家单位」的行 —— 历史行就是原来记在这家单位那张对账表上的每一笔，'
               '新录的行按开票方 / 收款方自动归。在业务流水改一笔，这里立刻跟着变。'),
 ('对账差异说明', '【对账差异说明】把 8 家单位 × 18 个金额列逐项拿原表合计行和系统算出来的数比一遍，差额小于 0.05 元的算对上'
                '（原表本身有四舍五入尾差）。9.24 这一版：康欣「已提供成本票」已经补齐对上；'
                '剩下标「△ 已查明」的 3 项是原表自己手填漏填造成的（华城应收工程款漏填 5 张票、杰华扣管理费列没填），原因写在说明列。'
                '最下面逐笔列出了 9.24 这一轮对业务流水做的每一处改动。'),
 ('项目简称', '【项目档案】的「项目简称」是从原表摘要「…：平双线」冒号后面那段取的；'
             '重名的自动带上第一次出现的那家单位，比如「金湖东西线(德誉嘉)」「金湖东西线(迅驰)」。'
             '业务流水和所有汇总表显示的都是简称。'),
 ('资金日记账（9.24 改了）', '改用出纳日记账的格式：B~M 列跟出纳那本账一模一样，出纳录完复制过来、选择性粘贴「数值」即可。'
                        'J 列「科目」就是账户，清单在 O 列「基础设置」（现在 7 个：' + '、'.join(JOUR_ACCTS) + '），'
                        'P 列填期初余额，右边自动算每个科目的借方、贷方、余额，最右边是逐行余额和校验。'
                        '原来「泓普 / 仟茂 / 现金」三账户那张旧表和单独的《出纳资金日记账.xlsx》都已删掉，'
                        '【首页】【项目利润】【费用统计】全部改成从新表取数。'),
 ('钱在谁手上', None),
 ('四个数', '上游（业主或上一层挂靠单位）还没付给这家挂靠单位的 →「欠业主未付款余额」；'
           '这家挂靠单位还欠我们多少 →「应收挂靠方余额」；已经到我们手上 →「挂靠单位已转我方」；'
           '上游付了、扣掉管理费质保金后这家单位还压着没转的 →「挂靠单位代收未转」。逐笔看【代收台账】和单位明细。'),
 ('合计有重复', '链条项目（比如金沁开给民能、康欣再开给金沁；华城、安锐、杰华开给德誉嘉）在上下两层单位里各记一次 —— '
               '原表本来就是一家一张表各记各的。所以【单位汇总】合计行和首页那几张回款卡片是各家直接相加，同一笔钱会算两遍；'
               '要看准确的数，看单家那一行或单家明细。另外金沁表第二组原表叫「康欣回款情况」，记的是金沁该付给康欣的钱，不是欠我方的。'),
 ('注意', None),
 ('不要改灰色区', '淡黄色是手工录入，灰色是自动算的。灰色列被覆盖后不会报错，但分表会静默算错。'),
 ('校验列必须全是√', '【业务流水】【资金日记账】【项目档案】的校验列出现红色，说明这一行有问题，要改掉。'),
 ('不要删行', '业务流水、日记账都不要右键删除整行：多录的行清空内容就行。删行会让公式范围缩小、行号全乱。'),
 ('历史数据', '原来八张对账表的业务都已导入，德誉嘉返 4%、迅驰过账应付和应扣税费也一并进来了。'
             '9.24 又按原表回款区逐行核了一遍，补了康欣 7 笔成本票、4 类手填应收，纠正了一批项目编号、按项目拆开了批量结算的管理费，'
             '每一笔都在业务流水备注列写了「9.24」，【对账差异说明】最下面有清单。'),
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


# ============================================================ 操作流程（9.24 新增：一步一步怎么录）
SH_HOW = '操作流程'
ws = wb.create_sheet(SH_HOW)
ws.sheet_view.showGridLines = False
HOW_COLS = ['日期', '项目编号', '业务类型', '开票/付款方', '收票/收款方', '发票性质', '票据类型', '摘要',
            '金额', '管理费率', '返现率', '计费方式', '系统自动算出来的 / 说明']
widths(ws, {'A': 2, 'B': 16, 'C': 11, 'D': 9, 'E': 11, 'F': 10, 'G': 10, 'H': 9, 'I': 10, 'J': 34,
            'K': 11, 'L': 8, 'M': 7, 'N': 10, 'O': 40})
HL = 'O'
_fl = wb[SH_FLOW]
_nc = [rr for rr in range(F0, F0 + FLOW_USED) if _fl[f'{FG}{rr}'].value == '不回成本票']
N_NC_COSTCOL = sum(1 for rr in _nc if _fl[f'{C_["ocol"]}{rr}'].value == '已开成本票')
N_NC_OTHER = len(_nc) - N_NC_COSTCOL
hr = [1]
def _h1(t):
    r = hr[0] + 1
    ws.merge_cells(f'B{r}:{HL}{r}')
    put(ws, f'B{r}', t, font=Font(name='微软雅黑', size=12, bold=True, color='FFFFFF'), fill=FILL_HDR, align=CL)
    ws.row_dimensions[r].height = 24
    hr[0] = r
def _p(t, lab=None, h=None, bold=False):
    r = hr[0] + 1
    if lab:
        put(ws, f'B{r}', lab, font=F_TOT, fill=FILL_HDR2, align=CL)
        ws.merge_cells(f'C{r}:{HL}{r}')
        put(ws, f'C{r}', t, font=F_TOT if bold else F_TXT, align=CL)
    else:
        ws.merge_cells(f'B{r}:{HL}{r}')
        put(ws, f'B{r}', t, font=F_TOT if bold else F_TXT, align=CL, border=None)
    n = max(1, len(t) // (95 if lab else 110) + 1 + t.count('\n'))
    ws.row_dimensions[r].height = h or max(18, 15 * n)
    hr[0] = r
def _tbl(rows, note=None):
    r = hr[0] + 1
    put(ws, f'B{r}', '业务流水这样录 →', font=F_NOTE, align=CR, border=None)
    for i, t in enumerate(HOW_COLS):
        put(ws, f'{L(3+i)}{r}', t, font=F_HDR2, fill=FILL_HDR2, align=C)
    ws.row_dimensions[r].height = 20
    for row in rows:
        r += 1
        for i, v in enumerate(row):
            c = L(3 + i)
            fmt = DATE if i == 0 else (MONEY if i == 8 else ('0%' if i in (9, 10) and isinstance(v, (int, float)) else None))
            if i == 0 and isinstance(v, str): v = dt.date.fromisoformat(v)
            put(ws, f'{c}{r}', v, font=F_NOTE if i == 12 else F_IN, fill=None if i == 12 else FILL_IN,
                align=CL if i in (7, 12) else C, fmt=fmt)
        ws.row_dimensions[r].height = 30
    hr[0] = r
    if note: _p(note)
def _gap(): hr[0] += 1

ws.merge_cells(f'B2:{HL}2')
put(ws, 'B2', '操作流程 · 业务流水和日记账到底怎么录（照着做就不会错）', font=F_TITLE, align=CL, border=None)
ws.row_dimensions[2].height = 30
hr[0] = 3
_p('这张表只是说明，不参与任何计算。淡黄色的样例行是「在【业务流水】里应该怎么填」的示范，照抄格式、换成你自己的数就行。', h=18)

_h1('一、整本表只有三个地方要录，其余全部自动')
_p('新项目 / 新合同先到这里加一行：项目编号（接着往下编，现在下一个是 ' + NEW_CODE + '）、项目全称、项目简称、项目类型（自营/合伙）、'
   '业主、一级单位、我方主体。业务流水、日记账里的「项目编号」下拉都从这里取，没加的项目选不到。', '① 项目档案')
_p('所有跟挂靠单位有关的「票」和「钱」都在这里一笔一行：开票、回成本票、业主付款、挂靠单位转钱、管理费、质保金、交税、红冲……'
   '各家单位明细、单位汇总、项目汇总、发票缺口、税费台账全部从这里取数。', '② 业务流水')
_p('出纳在自己那本账里录，录完把 B~M 列复制过来粘贴（见第五部分）。费用统计、项目实际支出、首页资金余额从这里取数。', '③ 资金日记账')
_p('淡黄色格子＝手填；灰色格子＝公式自动算，千万别在灰色格子里填数或删除（删了不报错，但后面的表会静默算错）。'
   '每一行录完看最右边「校验」列：显示 √ 才算录对了；红色的提示照第七部分改。', '颜色规矩')

_h1('二、业务流水：每一笔都按这 8 步录')
for lab, t in [
    ('第 1 步', '在【业务流水】最下面第一个空行接着录。不要插行、不要删行、不要排序（要看某家单位的，点表头的筛选按钮筛，看完清除筛选）。'),
    ('第 2 步', '日期：填发票日期或钱到账的日期，格式 2026-9-23 这样（不要写成 2026.9.23，那是文字不是日期，校验会报「日期格式不对」）。'),
    ('第 3 步', '项目编号：下拉选。项目档案里没有的，先去项目档案加一行。交税、红冲这类也尽量选上项目，按项目看才准。'),
    ('第 4 步', '业务类型：按第三部分的对照表选，这一格选错了后面全错 —— 9.23 华城那几行就是把开票选成了「挂靠代收」「我方收款」，所以欠的成本票出不来。'),
    ('第 5 步', '开票/付款方、收票/收款方：下拉选单位简称。谁开票给谁、谁把钱付给谁，方向别反。'),
    ('第 6 步', '只有「销项开票」要填：发票性质、票据类型（劳务票/机械设备票…）、管理费率（2% 就填 2%）、返现率（没有填 0）、计费方式（三选一，见第四部分第 10 条）。'
               '「成本票」要填发票性质和票据类型，票据类型要跟对应那张销项票一致。'),
    ('第 7 步', '金额：一律填正数；只有红冲、退回填负数。只保留两位小数（453.6616 这种要写成 453.66）。'
               '销项开票那一行右边「税费录入区」四列按实际预提的填，没有就空着。'),
    ('第 8 步', '看最右边「校验」列是不是 √。备注列想写什么就写什么，不影响计算。')]:
    _p(t, lab)

_h1('三、业务类型对照表（16 种，选错类型是最常见的错）')
r = hr[0] + 1
for i, t in enumerate(['业务类型', '什么时候选', '', '', '', '开票/付款方', '收票/收款方', '金额填什么', '', '会进哪一列']):
    put(ws, f'{L(2+i)}{r}', t, font=F_HDR2, fill=FILL_HDR2, align=C)
ws.merge_cells(f'C{r}:F{r}'); ws.merge_cells(f'I{r}:J{r}'); ws.merge_cells(f'K{r}:{HL}{r}')
put(ws, f'K{r}', '会进哪一列（单位明细 / 单位汇总）', font=F_HDR2, fill=FILL_HDR2, align=C)
hr[0] = r
KIND_HELP = [
    ('销项开票', '挂靠单位开票出去（开给业主，或开给上一层挂靠单位）', '挂靠单位', '业主 / 上一层单位', '发票金额（红冲填负数）',
     '销售开票金额、应扣管理费、应到成本票；欠业主·应收工程款；应收挂靠方·应收金额'),
    ('成本票', '我方（泓普/仟茂）开成本票给挂靠单位；补开、按 3% 扣点抵掉的也选这个', '泓普 / 仟茂', '挂靠单位', '成本票金额',
     '已到成本票（剩余开票金额就是还欠的）'),
    ('挂靠代收', '业主（或上一层挂靠单位）把工程款付给了挂靠单位（钱还在挂靠单位手上）', '业主 / 上一层挂靠单位', '收钱的挂靠单位', '付的钱',
     '收钱那家的欠业主·开票已回款 → 欠业主未付款余额减少'),
    ('我方收款', '挂靠单位把钱转给了我们（包括用代发工资、税差抵扣的）', '挂靠单位', '泓普 / 仟茂', '到我方的钱',
     '应收挂靠方·已到账 → 应收挂靠方余额减少'),
    ('管理费结算', '管理费实际扣了 / 结了', '泓普', '挂靠单位', '这次结的管理费',
     '应收挂靠方·扣管理费'),
    ('扣质保金', '挂靠单位扣了我方的质保金', '挂靠单位', '泓普', '扣的质保金', '应收挂靠方·质保金'),
    ('业主扣质保金', '业主扣在挂靠单位手上、还没放的质保金（金沁表那一列）', '业主', '挂靠单位', '扣的质保金',
     '欠业主·业主扣质保金（只单列出来看，不改欠款余额）'),
    ('代垫应收', '我方替挂靠单位垫了钱（比如代交税费），对方答应退回', '挂靠单位', '泓普', '垫的钱（退回时填负数）',
     '应收挂靠方·应收金额（对方退回时再录一行「代垫应收」负数冲掉）'),
    ('已交税', '实际交了税（泓普代交的也选那家挂靠单位，谁交的写在摘要里）', '替哪家交的就选哪家', '税局', '交的税（退回来填负数）', '已交税 → 欠税未交减少'),
    ('退税', '税局退税（只做记录，不进任何汇总）', '税局', '挂靠单位', '退的税', '不进汇总'),
    ('工资扣抵', '用工资表 / 代发工资顶掉劳务成本、不另开成本票', '泓普', '挂靠单位', '顶掉的金额', '代发工资；只有【发票缺口】劳务票还差减少'),
    ('其他应收收回', '德誉嘉返的 4% 现金收到了', '挂靠单位', '泓普', '收到的返现', '返管理费·已收款'),
    ('其他应付发生', '别人的钱过账到我们这（过账 / 合伙分成）', '过账单位', '过账单位', '过账金额', '过账/合伙应付款·应付款；往来台账'),
    ('其他应付扣税', '过账款里我方代垫、要扣下来的税费', '泓普', '过账单位', '扣的税费', '过账/合伙应付款·已付款'),
    ('其他应付支付', '过账款真转出去了（真正的收款人写在摘要里）', '泓普', '过账单位', '转出去的钱', '过账/合伙应付款·已付款'),
    ('其他', '以上都不是的（尽量别用，汇总里不体现）', '', '', '', '不进汇总')]
for k, when, a, b, amt, where in KIND_HELP:
    r = hr[0] + 1
    put(ws, f'B{r}', k, font=F_TOT, fill=FILL_IN, align=C)
    ws.merge_cells(f'C{r}:F{r}'); put(ws, f'C{r}', when, font=F_TXT, align=CL)
    put(ws, f'G{r}', a, font=F_TXT, align=C); put(ws, f'H{r}', b, font=F_TXT, align=C)
    ws.merge_cells(f'I{r}:J{r}'); put(ws, f'I{r}', amt, font=F_TXT, align=CL)
    ws.merge_cells(f'K{r}:{HL}{r}'); put(ws, f'K{r}', where, font=F_NOTE, align=CL)
    ws.row_dimensions[r].height = 30
    hr[0] = r

_h1('四、常见场景 · 照着样例录')
_p('1. 挂靠单位开票给业主（最常见）', bold=True)
_tbl([('2026-09-23', 'A056', '销项开票', '华城', '民能', '3%专票', '劳务票', '2026.9.23华城-民能：虎峰所双桥11社', 462.92, 0.02, 0,
       '扣管理费', '系统算：管理费 9.26，应到成本票 453.66；华城明细「剩余开票金额」多出 453.66（这就是欠的成本票）')])
_p('2. 我们开成本票给挂靠单位', bold=True)
_tbl([('2026-09-23', 'A056', '成本票', '泓普', '华城', '3%专票', '劳务票', '2026.9.23泓普-华城：虎峰所双桥11社', 453.66, None, None,
       None, '华城明细「已到成本票」+453.66，剩余开票金额减 453.66（这个项目就开齐了）')],
     '一次没开齐也没关系：比如这次只开了 300.00，就录 300.00，剩下的 153.66 自动挂在「剩余开票金额 / 还差成本票」里，等补开时按下面第 4 条再录。')
_p('3. 成本票还没开（欠着）—— 不用录任何东西。只要上面第 1 条的「销项开票」录了，系统自动算出应到成本票；'
   '欠多少直接看：单位明细「剩余开票金额」、单位汇总「还差成本票」、【发票缺口】按项目和票种（劳务/机械）列出来的「还差」。'
   '你截图里「想把欠的录进去不行」，是因为那几行业务类型选成了「挂靠代收 / 我方收款」—— 已经帮你改成了销项开票 + 成本票，'
   '备注列写了「9.24 更正」。其中虎峰所双桥11社那张票金额原录 356,218.09，按它的成本票 453.66÷98% 反推改成了 462.92，请对照发票确认。',
   bold=False)
_p('4. 补成本票（之前欠的，现在补开了）：再录一行「成本票」，日期写补开那天，项目编号选原来那个项目，金额写这次补开的金额。'
   '可以分几次补，每次一行，剩余开票金额会一点点减到 0。不要回头去改原来那一行开票。')
_tbl([('2026-10-10', 'A056', '成本票', '泓普', '华城', '3%专票', '劳务票', '2026.10.10补开剩余成本票：虎峰所双桥11社', 153.66, None, None,
       None, '接上例：先开了 300.00，这次补开剩下的 153.66，剩余开票金额变成 0')])
_p('5. 跟挂靠单位约定「不提供成本票、按 3% 扣点子」（康欣那种）：也录一行「成本票」，发票性质选「不开票」，摘要写明「未开票按3%扣点子（不提供成本票）」，'
   '金额写抵掉的成本票金额 —— 这样还差成本票就清掉了，单位明细里这一行会显示成红色，跟真收到的票分得开。'
   '扣下来的那 3% 如果对方是用钱给的，再录一行「我方收款」。康欣原表 H68 列的扣点清单一共 9 笔 287,173.15，现在 9 笔都记成了「不开票」。')
_tbl([('2026-07-15', 'A045', '成本票', '泓普', '康欣', '不开票', '劳务票', '2026.7.15康欣开劳务费到金沁：侣新线铜安1社｜未开票按3%扣点子（不提供成本票）',
       21690.88, None, None, None, '康欣明细「已到成本票」这一格标红')])
_p('6. 业主付钱 → 挂靠单位 → 我们（两步分开录）', bold=True)
_tbl([('2026-10-08', 'A056', '挂靠代收', '民能', '华城', None, None, '2026.10.8民能付款到华城：虎峰所双桥11社', 462.92, None, None, None,
       '欠业主未付款余额 −462.92'),
      ('2026-10-12', 'A056', '我方收款', '华城', '泓普', None, None, '2026.10.12华城转泓普：虎峰所双桥11社', 453.66, None, None, None,
       '应收挂靠方余额 −453.66'),
      ('2026-10-12', 'A056', '管理费结算', '泓普', '华城', None, None, '2026.10.12结管理费：虎峰所双桥11社', 9.26, None, None, None,
       '应收挂靠方余额再 −9.26，这个项目就清了')],
     '一笔钱如果同时结了好几个项目，最好按项目分开录几行（原表一格记好几个项目的，这次已经帮你按项目拆开了，见对账差异说明）；'
     '实在分不开就挂在其中一个项目上，单位合计是对的，只是按项目看会有正有负。')
_p('7. 我方替挂靠单位垫了钱（比如代交税费），对方答应退：录「代垫应收」；对方真退回来时再录一行「代垫应收」金额填负数'
   '（摘要写「收回代垫」）。不要录成「我方收款」—— 那样会把「代收未转」也冲掉。', bold=False)
_tbl([('2026-07-20', 'A001', '代垫应收', '华城', '泓普', None, None, '2026.7.20和华会计确定：3.31替华城交的设备票税费，结算时退回', 16883.46,
       None, None, None, '应收挂靠方·应收金额 +16,883.46')])
_p('8. 红冲更正发票（德誉嘉 9.20 那种）：原来那一行一个字都不要改（7 月那张票确实开过，改了 7 月的报表就不对了）。另起两行：'
   '一行「销项开票」金额填负数把原票冲掉（费率、计费方式跟原票一模一样），再一行录重开的正确发票。'
   '两行日期都写红冲那天。管理费、应到成本票会自动先冲负再按新票算。', bold=False)
_tbl([('2026-09-20', 'A017', '销项开票', '德誉嘉', '民能', '13%专票', '机械设备票', '2026.9.20德誉嘉红冲：冲销7.22开到民能的发票（永嘉所义和8社）',
       -1622.53, 0.04, 0, '扣管理费', '负数：管理费、应到成本票自动冲成负数；明细里标红'),
      ('2026-09-20', 'A017', '销项开票', '德誉嘉', '民能', '13%专票', '机械设备票', '2026.9.20德誉嘉重开发票到民能：永嘉所义和8社（更正后金额）',
       1662.53, 0.04, 0, '扣管理费', '按新票重新算')],
     '原票右边税费录入区（P~S）填了税的，红冲那一行也要把这几格填成负数，重开那一行按新票填，不然欠税会多出一份。'
     '如果那张票对应的成本票也要红冲：同样再录一行「成本票」负数、一行重开的成本票。')
_p('9. 录错了怎么改：', bold=True)
_kx_row = PATCH_ROWS['kx_fee']
for lab, t in [('刚录的新行', '直接在那一行把错的格子改掉（日期、金额、单位、类型都可以改），改完看校验列是 √。'),
               ('历史行（备注写着「历史导入」或「9.24…」的）', '只能改日期、金额、项目编号、摘要，不要改开票方 / 收款方和业务类型：'
                                 '这些行背后有隐藏的「来源表」「原表来源列」，改单位不会换到别家，成本票和销项开票互改还会错乱（金额同时进应到和已到）。'
                                 '真要改，用第 8 条红冲的办法：录一行负数把原行冲掉，再按正确的单位 / 类型新录一行。'
                                 '镜像行（备注写着「原表「已开成本票」列还原」的销项开票，比如金沁表里的康欣开给金沁）要冲的话，'
                                 '录「成本票」康欣→金沁、金额填负数、计入汇总选「否」。'),
               ('已经对过账 / 给领导看过的月份', '别改原行，用第 8 条红冲的办法：负数冲掉、再录一行对的。这样以前那个月的数不会变。'),
               ('整行都是多录的', '只清空 B~N 列和 P~S 列（选中按 Delete），O 列和最右边灰色的列一格都别动；清空过的历史行也不要再拿来录新数据，'
                                 '新数据一律录到最下面的空行。千万别右键「删除行」：删行会把下面的行号打乱，公式范围也会缩小。'
                                 f'你回传的那份把第 20 行（康欣 2026.3.25 管理费结算 7,443.98）删了，康欣的管理费就少算了这一笔 —— '
                                 f'已经帮你恢复（新版在第 {_kx_row} 行起，按 5 个项目拆开了）。'),
               ('明细表里看着像重复的行', '同一张原表行里的「开票、业主付款、转我方、管理费结算」在系统里是分开的几笔，摘要一样，看着像重复，其实不是。'
                                       '现在单位明细的摘要前面会带【挂靠代收】【我方收款】这种类型标签，一眼能分出来。不要去业务流水删。')]:
    _p(t, lab)
_p('10. 「计费方式」三种怎么选，看到「不回成本票」要不要处理：', bold=True)
_nc_other = [rr for rr in _nc if _fl[f'{C_["ocol"]}{rr}'].value != '已开成本票']
for lab, t in [('扣管理费', '正常业务：管理费＝金额×管理费率，应到成本票＝金额−管理费。'),
               ('不扣管理费', '只借通道过票、不收管理费：管理费＝0，应到成本票＝全额（我方照样要回全额成本票）。'),
               ('不回成本票', '这张票本身就不需要我方回成本票：应到成本票＝0，也不算进「应收挂靠方」。历史数据里一共 '
                             f'{len(_nc)} 行是这个，分两类：'),
               (f'① {N_NC_COSTCOL} 行镜像行 · 不能改', '挂靠单位之间互开的票（比如康欣开给金沁）在收票那家原表里记在「已开成本票」列，'
                             '系统照样还原了一行（备注写着「历史导入·原表「已开成本票」列还原」）。它是收票那家收到的成本票，'
                             '开票那家另外还有一行真正的销项开票。这类行的计费方式一律不能改 —— 改了应到成本票会记到收票那家头上、还跟开票那家重复（校验列会报错）。'),
               (f'② {N_NC_OTHER} 行合伙项目 · 可以改', '业务流水第 ' + '、'.join(str(x) for x in _nc_other) + ' 行'
                             '（康欣平滩所、凤飞7社、哨楼村支线、维新，安锐平滩所），原表就没要求我方回票。确认其实要回票时，'
                             '把计费方式改成「扣管理费」（填费率）或「不扣管理费」，剩余开票金额 / 还差成本票马上出来。'),
               ('新录的票', '一般选「扣管理费」；只过票不收费的选「不扣管理费」；很少用到「不回成本票」。')]:
    _p(t, lab)
_p('11. 挂靠单位之间互开的票（比如康欣开给金沁）：录一行「销项开票」，开票方康欣、收票方金沁，它算在康欣明细里。'
   '如果金沁那边也要看到「收到了康欣这张成本票」，再录一行「成本票」开票方康欣、收票方金沁，最右边「计入汇总」选「否」—— '
   '这一行会自动算到金沁头上（成本票认收票方），「计入汇总＝否」让【发票缺口】【链条核算】【项目利润】这些全公司口径的表不重复算这张票；'
   '单位明细、单位汇总是一家一本账，照算。')
_p('12. 上一层挂靠单位付钱给下一层（比如德誉嘉付给安锐、金沁付给康欣）：录「挂靠代收」，付款方选上一层、收款方选收钱那家 —— '
   '会自动记到收钱那家头上，冲减它的「欠业主未付款余额」。金沁付给康欣的还要多录一行：金沁那张表第二组（原表叫「康欣回款情况」，'
   '记的是金沁该付康欣的钱）要冲掉，原表是记在「已到账」列 —— 系统里照历史的记法录「我方收款」金沁→泓普，摘要写明「付给康欣」。')
_p('13. 质保金：挂靠单位扣了我方的质保金录「扣质保金」；到期退回来时录一行「扣质保金」负数冲回，钱到账再录一行「我方收款」。'
   '业主扣在挂靠单位手上的录「业主扣质保金」；业主放了以后录「挂靠代收」，再录一行「业主扣质保金」负数冲掉。')
_p('14. 其他几类：德誉嘉 2026-7-22 以前的票是扣 8% 返 4%（历史已录好），以后的新票管理费率填 4%、返现率填 0；'
   '以前挂着的返现收到时录「其他应收收回」。过账款 —— 收到录「其他应付发生」、代垫的税费录「其他应付扣税」、转出去录「其他应付支付」，'
   '这三类开票/付款方、收票/收款方都选那家过账单位（比如迅驰），真正的收款人写在摘要里。'
   '工资表顶劳务成本 —— 录「工资扣抵」（泓普→挂靠单位），只在【发票缺口】里扣减劳务票还差。'
   '泓普替挂靠单位交的税 —— 录「已交税」，开票/付款方选那家挂靠单位。')

_h1('五、资金日记账：出纳的账怎么搬过来')
for lab, t in [('第 1 步', '出纳在自己那本《出纳日记账》里照常录（格式跟这边【资金日记账】B~M 列一模一样：日期、凭证种类、凭证编号、项目编号、费用类型、往来单位、报销人员、摘要、科目、借方、贷方、备注）。'),
               ('第 2 步', '录完选中出纳账里新录的那几行的 B~M 列，复制。'),
               ('第 3 步', '到本表【资金日记账】B 列第一个空行，右键 → 选择性粘贴 → 数值。不要整行粘、不要粘到 N 列以后（右边灰色是公式）。'),
               ('第 4 步', '看最右边「校验」列：全是 √ 就完事。常见提示：「科目不在基础设置里」＝J 列科目名字跟 O 列清单不一样（多了空格、写了别名）；'
                          '「日期格式不对」＝日期粘成了文字；「借方贷方不能同时填」。'),
               ('科目怎么加', f'O 列「基础设置」就是科目清单，现在是 {"、".join(JOUR_ACCTS)}，还能再加 {N_ACCT - len(JOUR_ACCTS)} 个。'
                            'P 列填每个科目的期初余额（出纳账「期初余额表」里的数），右边自动算借方合计、贷方合计、期末余额。'),
               ('项目编号', '日记账 E 列「项目编号」手填项目档案里的编号（A001 这种，这一列没有下拉，照项目档案抄），【项目利润】的「项目实际支出」才取得到；不填也行，只是项目利润那一列是 0。'),
               ('本月合计这类行', '出纳账里「本月合计 / 本年累计 / 过次页 / 承前页 / 上年结转」这种行粘过来也没关系，系统认得出，不会重复算（校验列显示灰色的「√ 合计/结转行（不计入）」）。'),
               ('期初余额', 'P 列期初余额现在还是空的，所以余额只是 9 月以来的净发生额。请把出纳账「期初余额表」里各科目的期初填进 P 列。'),
               ('原来那张旧日记账', '按你的要求已经删了（泓普/仟茂/现金三账户那张）。8 月那 34 笔是旧格式、旧科目（泓普/仟茂），没有搬过来；'
                                 '要的话在出纳账里按新科目录好再粘过来即可。')]:
    _p(t, lab)

_h1('六、月底对账看哪里')
for lab, t in [('某家单位欠多少', '打开那家的「××明细」看第 7 行。两组余额 ——「欠业主未付款余额」＝上游还没付给这家单位的；'
                                 '「应收挂靠方余额」＝这家单位还欠我们的（开票额 − 不回成本票的 + 代垫应收 − 管理费 − 质保金 − 已转我方）。'
                                 '上面填了截止日期，第 7 行的余额就是截至那天的数；不填就是截至今天。金沁那张的第二组原表叫「康欣回款情况」，是金沁欠康欣的。'),
               ('各家一起看', '【单位汇总】一家一行；【项目汇总】一个项目一行；【单位项目明细】上面选一家，看它名下每个项目。'
                           '合计行是各家直接相加，链条项目上下两层各记一次，会有重复，以单家为准。'),
               ('成本票欠多少', '【发票缺口】按项目、按劳务票 / 机械票分开列还差多少。'),
               ('税欠多少', '【税费台账】。'),
               ('跟原表对不对得上', '【对账差异说明】8 家 × 18 个金额逐项跟你原来那份《对账明细》9.13 版的合计行比；最下面列了 9.24 这一轮改动的每一笔。'),
               ('首页', '所有校验一行一个结论，全是 ✓ 就说明录得没问题。')]:
    _p(t, lab)

_h1('七、校验列的提示是什么意思')
for lab, t in [('未填日期', '这一行填了业务类型或金额，日期却空着。'),
               ('日期格式不对', '日期写成了 2026.9.23 这种文字，改成 2026-9-23。'),
               ('两边都不是挂靠单位，算不到哪家头上', '开票/付款方、收票/收款方至少一边要选挂靠单位（比如泓普代交的税，付款方选替哪家交的）。'),
               ('这是收票方账上的镜像行…', '历史上「已开成本票」列还原出来的行，计费方式只能是「不回成本票」，改回去（见第四部分第 10 条）。'),
               ('金额是文字，请改成数字（日记账）', '出纳账里的金额粘过来成了文字（比如带千分位的「1,000.00」），这一行不会算进余额。'),
               ('未选业务类型 / 未选项目编号', '对应那一格空着。'),
               ('项目编号不存在', '先去【项目档案】加这个项目。'),
               ('开票方不在单位档案 / 收款方不在单位档案', '单位名字要从下拉选，手打的对不上；新单位先去【单位档案】加。'),
               ('金额须为数字', '金额格里是文字（比如带了「元」字或空格）。'),
               ('未选票据类型', '销项开票、成本票要选劳务票 / 机械设备票…'),
               ('没选计费方式 / 选了扣管理费但费率是 0', '销项开票那一行的计费方式和管理费率要对上。'),
               ('链条待确认', '只是提示：这个项目走了好几条挂靠链，不影响汇总，可以不管。')]:
    _p(t, lab)

_h1('八、以后新增挂靠单位：单位档案加一行 + 复制一张明细表')
_p('新单位的业务照样只录【业务流水】。【单位汇总】【往来台账】【税费台账】【单位项目明细】会自动多出这一家，不用动；'
   '只有给领导看的「××明细」要照下面复制一张。下面拿新单位「海川」举例。')
for lab, t in [
    ('第 1 步', '到【单位档案】最下面第一个空行加一行：单位简称（以后所有表、所有下拉里显示的就是这个名字，比如「海川」）、单位全称、'
               '类型一定选「挂靠单位」（选成别的，哪张表都算不到它头上）、状态选「正常」。'
               '上游开票税率、我方回开票种税率这几格照实填，只影响业务流水最右边的「税费参考试算」。'),
    ('第 2 步', '复制一张明细表：在底下【杰华明细】的标签上点右键 →「移动或复制」→ 勾上「建立副本」→「下列选定工作表之前」选【单位汇总】→ 确定。'
               '复制哪一张都行，杰华明细最省事：它回款两组的组名是通用的（工程款回款情况 / 泓普回款情况），用不上的列组也已经收起来了。'),
    ('第 3 步', '双击新表的标签，改名成「海川明细」（表名里要带单位简称）。'),
    ('第 4 步', '点新表的 B3「本表单位」，下拉选「海川」。标题、逐笔明细、第 7 行合计马上就换成海川的。'
               'J3 出现 ⚠ 就按提示改：「请在 B3 选本表单位」＝B3 空着；「类型不是挂靠单位」＝第 1 步类型没选对；'
               '「表名跟本表单位对不上」＝表名和 B3 不是同一家。'),
    ('第 5 步（可选）', '第 5 行的组名想照对方的叫法，直接改那一格的字（比如「海川回款情况（欠业主未付款）」），不影响计算。'
                     '这家单位有返现（返管理费）或过账 / 合伙应付款的：选中 AC 到 AJ 这几列 → 右键「取消隐藏」，那两组列就出来了。'),
    ('第 6 步', '以后海川的业务照常录【业务流水】，开票/付款方、收票/收款方的下拉里已经有「海川」了。录完看海川明细第 7 行；'
               '【单位汇总】【往来台账】【税费台账】会多出海川一行（空白行是收起来的，没看到就点 数据 → 筛选 → 重新应用）；'
               '【单位项目明细】B3 也能选海川。'),
    ('B3 别漏改', '复制出来的表不改 B3，显示的还是被复制那家的数（J3 会提示表名对不上）。原来那 8 张明细的 B3 不要动。'),
    ('简称定了别改', '业务流水里记的是单位简称，改了简称，以前录的行就对不上这家单位了；全称、税率随时可以改。'),
    ('对账差异说明', '只核对原来 8 家跟你 9.13 那份《对账明细》，新单位不在里面，不用管。'),
    ('单位不做了', '单位档案里那一行不要删（删了以前的业务就找不到这家了），状态改「停用」做个记号就行；它的明细表可以留着看历史。')]:
    _p(t, lab)
page(ws, titles=None, landscape=True)
ws.print_area = f'$A$1:${HL}${hr[0]}'
print(f'  ✓ 操作流程（{hr[0]} 行，含 {len(KIND_HELP)} 种业务类型对照 + 14 个场景）')


# ============================================================ 收尾
TABC = {SH_HOME2:'1F3864', SH_DOC:'1F3864', SH_UNIT:'7F7F7F', SH_PROJ:'7F7F7F',
        SH_FLOW:'ED7D31', SH_JOUR:'ED7D31', SH_DAI:'2F5597', SH_SUM_U:'548235', SH_SUM_P:'548235',
        SH_SUM_X:'548235', SH_CHAIN:'548235', SH_GAP:'7030A0', SH_CUR:'7030A0', SH_PRF:'BF8F00',
        SH_TAX:'548235', SH_EXP:'2F5597'}
for n in SUB_SHEETS: TABC[n] = 'A9D08E'
TABC[SH_DIFF] = 'C00000'
TABC[SH_HOW] = 'C00000'
for n, c in TABC.items(): wb[n].sheet_properties.tabColor = c
# 表的先后照 9.24 回传版你自己排的顺序：明细放在日记账后面，迅驰、康欣、德誉嘉排最前
SUB_ORDER = [f'{u}明细' for u in ('迅驰', '康欣', '德誉嘉', '华城', '金沁', '湖南锦泰', '安锐', '杰华')]
assert sorted(SUB_ORDER) == sorted(SUB_SHEETS)
ORDER2 = ([SH_HOME2, SH_HOW, SH_DOC, SH_UNIT, SH_PROJ, SH_FLOW, SH_JOUR] + SUB_ORDER +
          [SH_SUM_U, SH_PRF, SH_SUM_X, SH_SUM_P] +
          [SH_CHAIN, SH_GAP, SH_CUR, SH_TAX, SH_DAI, SH_EXP, SH_DIFF])
wb._sheets = [wb[n] for n in ORDER2]; wb.active = 0
OUT = os.environ.get('GK_OUT', os.path.join(HERE, '..', '建筑挂靠业务核算系统.xlsx'))
wb.save(OUT)
nf = sum(1 for s in wb.worksheets for row in s.iter_rows() for c in row
         if isinstance(c.value, str) and c.value.startswith('='))
print(f'\n已保存：{os.path.abspath(OUT)}')
print(f'工作表 {len(wb.worksheets)} 张，公式 {nf} 个，自动配平括号 {BAL_FIXED[0]} 处')
