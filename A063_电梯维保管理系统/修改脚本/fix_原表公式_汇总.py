# -*- coding: utf-8 -*-
# 由 fix_原表公式.py 用 exec 调起，共用它的 wb / COLS / TEAMS / TOT_END 等变量
sys.path.insert(0, os.path.join(ROOT, '生成脚本'))
from a063_style import (put as sput, title as stitle, headers as sheaders, widths as swidths,
                        FILL_IN, FILL_AUTO, FILL_TOT, FILL_WARN, FILL_KPI, FILL_SEC,
                        F_IN, F_AUTO, F_TOT, F_WARN, F_NOTE, F_TXT, F_SEC, F_HDR, F_BIG,
                        C, CL_, CR, CT, NUM, PCT, C_MAIN, C_RPT, C_WARN, C_DASH, C_BASE)

MONEY = '#,##0.00;[Red]\\-#,##0.00;\\-'
if '汇总' in wb.sheetnames:
    del wb['汇总']
ws = wb.create_sheet('汇总', 0)
LBL = {'Sheet5': 'Sheet5（空白模板）'}


def TR(sn, k):
    c = COLS[sn][k]
    return f"'{sn}'!${c}$3:${c}${TOT_END}"


EA = lambda c: f'电梯档案!${c}$2:${c}${LIFT_NR}'

stitle(ws, '维 保 汇 总 · 梯子状态、数量、回款情况（全自动，不用填）', 'V',
       '★ 这张表只读不填。数字全部从《电梯档案》和各维保组的表里自动汇总，那边改了这边跟着变。\n'
       '★ 只有第 3 行三个淡黄格可以改：统计年度（④ 按年度用）、到期提醒天数（到期 / 年检 / 免保提醒都按它）、'
       '付款宽限天数（每期期初过了这几天还没收才算逾期，原表多数写「七日内付」，默认 7）。')
sput(ws, 'A3', '统计年度', font=F_TOT, fill=FILL_SEC, align=CR)
sput(ws, 'B3', 2026, font=F_BIG, fill=FILL_IN, fmt='0')
sput(ws, 'C3', '到期提醒天数', font=F_TOT, fill=FILL_SEC, align=CR)
sput(ws, 'D3', 30, font=F_BIG, fill=FILL_IN, fmt='0')
sput(ws, 'E3', '付款宽限天数', font=F_TOT, fill=FILL_SEC, align=CR)
sput(ws, 'F3', 7, font=F_BIG, fill=FILL_IN, fmt='0')
sput(ws, 'G3', '今天', font=F_TOT, fill=FILL_SEC, align=CR)
sput(ws, 'H3', '=TODAY()', font=F_TOT, fill=FILL_AUTO, fmt='yyyy-mm-dd')
ws.row_dimensions[3].height = 24
for nm, ref in [('统计年度', '汇总!$B$3'), ('提醒天数', '汇总!$D$3'), ('付款宽限天数', '汇总!$F$3')]:
    if nm in wb.defined_names:
        del wb.defined_names[nm]
    wb.defined_names.add(DefinedName(nm, attr_text=ref))


def sec(r, txt, color=C_MAIN, last='V'):
    sput(ws, f'A{r}', txt, font=F_HDR, fill=PatternFill('solid', fgColor=color), align=CL_)
    ws.merge_cells(f'A{r}:{last}{r}')
    ws.row_dimensions[r].height = 22


def hdr(r, cols, color=C_MAIN):
    for col, txt in cols:
        sput(ws, f'{col}{r}', txt, font=F_HDR, fill=PatternFill('solid', fgColor=color))
    ws.row_dimensions[r].height = 30


# ── ① 梯子情况 ──
sec(5, ' ① 梯子情况（来自《电梯档案》）：维保状态 / 梯型 / 维保人员 / 年检与免保提醒', C_BASE)
hdr(6, [('A', '维保状态'), ('B', '台量'), ('C', '占比'), ('D', '免保已到期或\n30天内到期'),
        ('F', '梯型'), ('G', '台量'), ('H', '占比'),
        ('J', '维保人员'), ('K', '台量'),
        ('M', '提醒'), ('N', '台数')], C_BASE)
TOTL = f'SUM({EA("N")})'
for i, st in enumerate(STATES):
    r = 7 + i
    sput(ws, f'A{r}', st, font=F_TXT, fill=FILL_AUTO, align=CL_)
    sput(ws, f'B{r}', f'=SUMIF({EA("W")},$A{r},{EA("N")})', font=F_IN, fill=FILL_AUTO, fmt=NUM)
    sput(ws, f'C{r}', f'=IFERROR($B{r}/$B$14,"")', font=F_IN, fill=FILL_AUTO, fmt=PCT)
    if st in ('安装免保', '技术免保', '质保期内'):
        sput(ws, f'D{r}', f'=COUNTIFS({EA("W")},$A{r},{EA("Y")},"▲免保已到期*")'
                          f'+COUNTIFS({EA("W")},$A{r},{EA("Y")},"免保剩*")', font=F_WARN, fill=FILL_AUTO, fmt=NUM)
    else:
        sput(ws, f'D{r}', None, font=F_IN, fill=FILL_AUTO)
sput(ws, 'A13', '还没填维保状态', font=F_WARN, fill=FILL_WARN, align=CL_)
sput(ws, 'B13', '=$B$14-SUM($B$7:$B$12)', font=F_WARN, fill=FILL_WARN, fmt=NUM)
sput(ws, 'C13', '=IFERROR($B13/$B$14,"")', font=F_WARN, fill=FILL_WARN, fmt=PCT)
sput(ws, 'D13', None, font=F_WARN, fill=FILL_WARN)
sput(ws, 'A14', '在册总台量', font=F_TOT, fill=FILL_TOT)
sput(ws, 'B14', f'={TOTL}', font=F_TOT, fill=FILL_TOT, fmt=NUM)
sput(ws, 'C14', '=IF($B$14=0,"",1)', font=F_TOT, fill=FILL_TOT, fmt=PCT)
sput(ws, 'D14', '=SUM($D$7:$D$12)', font=F_TOT, fill=FILL_TOT, fmt=NUM)

TYPES = [('客梯', f'SUMIF({EA("H")},"客梯",{EA("N")})'),
         ('客梯(无机房)', f'SUMIF({EA("H")},"客梯(无机房)",{EA("N")})+SUMIF({EA("H")},"客梯（无机房）",{EA("N")})'),
         ('货梯', f'SUMIF({EA("H")},"货梯",{EA("N")})'), ('扶梯', f'SUMIF({EA("H")},"扶梯",{EA("N")})'),
         ('人行道', f'SUMIF({EA("H")},"人行道",{EA("N")})'), ('杂物梯', f'SUMIF({EA("H")},"杂物梯",{EA("N")})')]
for i, (t, f) in enumerate(TYPES):
    r = 7 + i
    sput(ws, f'F{r}', t, font=F_TXT, fill=FILL_AUTO, align=CL_)
    sput(ws, f'G{r}', f'={f}', font=F_IN, fill=FILL_AUTO, fmt=NUM)
    sput(ws, f'H{r}', f'=IFERROR($G{r}/$G$14,"")', font=F_IN, fill=FILL_AUTO, fmt=PCT)
sput(ws, 'F13', '其他 / 梯型没填', font=F_WARN, fill=FILL_WARN, align=CL_)
sput(ws, 'G13', '=$G$14-SUM($G$7:$G$12)', font=F_WARN, fill=FILL_WARN, fmt=NUM)
sput(ws, 'H13', '=IFERROR($G13/$G$14,"")', font=F_WARN, fill=FILL_WARN, fmt=PCT)
sput(ws, 'F14', '合计', font=F_TOT, fill=FILL_TOT)
sput(ws, 'G14', f'={TOTL}', font=F_TOT, fill=FILL_TOT, fmt=NUM)
sput(ws, 'H14', '=IF($G$14=0,"",1)', font=F_TOT, fill=FILL_TOT, fmt=PCT)

sput(ws, 'J7', '已排维保人员', font=F_TXT, fill=FILL_AUTO, align=CL_)
sput(ws, 'K7', f'=SUMIF({EA("V")},"?*",{EA("N")})-$K$8', font=F_IN, fill=FILL_AUTO, fmt=NUM)
sput(ws, 'J8', '写着「未排」', font=F_WARN, fill=FILL_WARN, align=CL_)
sput(ws, 'K8', f'=SUMIF({EA("V")},"未排",{EA("N")})', font=F_WARN, fill=FILL_WARN, fmt=NUM)
sput(ws, 'J9', '维保人员没填', font=F_WARN, fill=FILL_WARN, align=CL_)
sput(ws, 'K9', '=$K$10-$K$7-$K$8', font=F_WARN, fill=FILL_WARN, fmt=NUM)
sput(ws, 'J10', '合计', font=F_TOT, fill=FILL_TOT)
sput(ws, 'K10', f'={TOTL}', font=F_TOT, fill=FILL_TOT, fmt=NUM)

REM = [('年检/限速器/载重 已过期', f'=COUNTIF({EA("Z")},"*已过期*")', True),
       ('年检/限速器/载重 快到期', f'=COUNTIFS({EA("Z")},"*剩*",{EA("Z")},"<>*已过期*")', True),
       ('免保已到期该转收费', f'=COUNTIF({EA("Y")},"▲免保已到期*")', True),
       ('免保快到期', f'=COUNTIF({EA("Y")},"免保剩*")', True),
       ('免保没填到期日', f'=COUNTIF({EA("Y")},"▲免保到期日没填")', True),
       ('档案在册台量', f'={TOTL}', False),
       ('各组合同台量合计', '=$C$25', False),
       ('档案比合同多出的台量', '=$N$12-$N$13', False)]
for i, (t, f, warn) in enumerate(REM):
    r = 7 + i
    sput(ws, f'M{r}', t, font=F_WARN if warn else F_TXT, fill=FILL_WARN if warn else FILL_AUTO, align=CL_)
    sput(ws, f'N{r}', f, font=F_WARN if warn else F_TOT, fill=FILL_WARN if warn else FILL_AUTO, fmt=NUM)
sput(ws, 'M15', '↑ 多出来的就是在档案里、但还没挂到任何一张维保组合同上的梯子（或合同台量没填）。',
     font=F_NOTE, align=CL_, border=None)
ws.merge_cells('M15:V15')

# ── ② 各维保组 合同与回款 ──
R2 = 17
sec(R2, ' ② 各维保组：合同与回款（每组一行）', C_RPT)
H2 = [('A', '维保组'), ('B', '项目\n行数'), ('C', '合同\n台量'), ('D', '合同\n未回'), ('E', '已到期'),
      ('F', '即将\n到期'), ('G', '免保\n项目'), ('H', '合同应收'), ('I', '合同外\n维修费'), ('J', '总应收'),
      ('K', '开票金额'), ('L', '收款金额'), ('M', '收现金额'), ('N', '收款不开票'), ('O', '已收合计'),
      ('P', '回款率'), ('Q', '开票应收额\n(开票未到账)'), ('R', '到账未开票\n(待补票)'), ('S', '未开票应收'),
      ('T', '应收余额\n(欠款)'), ('U', '截至今天应收\n(按付款周期)'), ('V', '逾期未收')]
hdr(R2 + 1, H2, C_RPT)
T0 = R2 + 2
for i, (sn, _L) in enumerate(TEAMS):
    r = T0 + i
    sput(ws, f'A{r}', LBL.get(sn, sn), font=F_TOT, fill=FILL_AUTO, align=CL_)
    F = lambda k: TR(sn, k)
    cells = [('B', f'=COUNTIF({F("key")},"?*")', NUM), ('C', f'=SUM({F("h")})', NUM),
             ('D', f'=COUNTIF({F("F")},"*合同未回*")', NUM), ('E', f'=COUNTIF({F("F")},"*已到期*")', NUM),
             ('F', f'=COUNTIF({F("F")},"*即将到期*")', NUM),
             ('G', f'=COUNTIF({F("zt")},"安装免保")+COUNTIF({F("zt")},"技术免保")+COUNTIF({F("zt")},"质保期内")', NUM),
             ('H', f'=SUM({F("j")})', MONEY), ('I', f'=SUM({F("k")})', MONEY), ('J', f'=SUM({F("l")})', MONEY),
             ('K', f'=SUM({F("n")})', MONEY), ('L', f'=SUM({F("p")})', MONEY), ('M', f'=SUM({F("q")})', MONEY),
             ('N', f'=SUM({F("x0")})', MONEY), ('O', f'=$L{r}+$M{r}+$N{r}', MONEY),
             ('P', f'=IFERROR($O{r}/$J{r},"")', PCT), ('Q', f'=SUM({F("s")})', MONEY),
             ('R', f'=SUM({F("t")})', MONEY), ('S', f'=SUM({F("u")})', MONEY), ('T', f'=SUM({F("v")})', MONEY),
             ('U', f'=SUM({F("jd")})', MONEY), ('V', f'=SUM({F("yq")})', MONEY)]
    for col, f, fmt in cells:
        warn = col in ('D', 'E', 'T', 'V')
        sput(ws, f'{col}{r}', f, font=F_WARN if warn else F_IN,
             fill=FILL_KPI if col in ('T',) else FILL_AUTO, fmt=fmt)
TT = T0 + len(TEAMS)
sput(ws, f'A{TT}', '合计', font=F_TOT, fill=FILL_TOT)
for col, _t in H2[1:]:
    if col == 'P':
        sput(ws, f'P{TT}', f'=IFERROR($O{TT}/$J{TT},"")', font=F_TOT, fill=FILL_TOT, fmt=PCT)
    else:
        sput(ws, f'{col}{TT}', f'=SUM({col}{T0}:{col}{TT-1})', font=F_TOT, fill=FILL_TOT,
             fmt=NUM if col in 'BCDEFG' else MONEY)
assert TT == 25, TT   # ① 里「各组合同台量合计」引用的是 $C$25

# ── ③ 收款五种情况 ──
R3 = TT + 3
sec(R3, ' ③ 收款五种情况（按维保组拆开）：钱是怎么收的、票是怎么开的', C_DASH)
tcols = [CL(3 + i) for i in range(len(TEAMS))]
TC, NC = CL(3 + len(TEAMS)), CL(4 + len(TEAMS))
hdr(R3 + 1, [('A', '情况'), ('B', '在组表里怎么填')] + [(c, LBL.get(sn, sn)) for c, (sn, _L) in zip(tcols, TEAMS)]
    + [(TC, '合计'), (NC, '涉及行数')], C_DASH)
CASES = [('① 不开票的现金账', '只填「收现金额」', 'q'),
         ('② 收款不开票（转账/微信/收据，不开票）', '只填「收款不开票」', 'x0'),
         ('③ 开票收款（票开了、钱也到了）', '「开票金额」和「收款金额」都填', 'kd'),
         ('④ 开票未到账（开票应收额）', '只填「开票金额」，钱没到', 's'),
         ('⑤ 到账未开票（待补票）', '只填「收款金额」，票还没开', 't'),
         ('　 未开票也未收（未开票应收）', '自动：还欠的钱里没开票的部分', 'u'),
         ('　 应收余额（最后还欠多少）', '自动：总应收 − 收款 − 收现 − 收款不开票', 'v')]
for i, (nm, how, k) in enumerate(CASES):
    r = R3 + 2 + i
    last = i >= 5
    sput(ws, f'A{r}', nm, font=F_TOT, fill=FILL_TOT if last else FILL_AUTO, align=CL_)
    sput(ws, f'B{r}', how, font=F_NOTE, fill=FILL_TOT if last else FILL_AUTO, align=CL_)
    for c, (sn, _L) in zip(tcols, TEAMS):
        sput(ws, f'{c}{r}', f'=SUM({TR(sn, k)})', font=F_IN, fill=FILL_TOT if last else FILL_AUTO, fmt=MONEY)
    sput(ws, f'{TC}{r}', f'=SUM({tcols[0]}{r}:{tcols[-1]}{r})', font=F_TOT, fill=FILL_KPI, fmt=MONEY)
    sput(ws, f'{NC}{r}', '=' + '+'.join(f'COUNTIF({TR(sn, k)},">0")' for sn, _L in TEAMS),
         font=F_IN, fill=FILL_AUTO, fmt=NUM)
R3E = R3 + 2 + len(CASES) - 1

# ── ④ 按年度 ──
R4 = R3E + 3
sec(R4, ' ④ 按年度看欠款（年度＝合同起始日所在年份；「统计年度」在第 3 行改）：26 年没收完的，27 年在「往年合同」里接着看', C_MAIN)
hdr(R4 + 1, [('A', '维保组'), ('B', '本年度合同\n总应收'), ('C', '本年度合同\n已收'), ('D', '本年度合同\n应收余额'),
             ('E', '往年合同\n应收余额(结转)'), ('F', '以后年度合同\n应收余额'), ('G', '没填合同期的\n应收余额'),
             ('H', '累计应收余额\n(全部)')], C_MAIN)
Y0 = R4 + 2
for i, (sn, _L) in enumerate(TEAMS):
    r = Y0 + i
    F = lambda k: TR(sn, k)
    sput(ws, f'A{r}', LBL.get(sn, sn), font=F_TOT, fill=FILL_AUTO, align=CL_)
    sput(ws, f'B{r}', f'=SUMIFS({F("l")},{F("nd")},统计年度)', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    sput(ws, f'C{r}', f'=SUMIFS({F("p")},{F("nd")},统计年度)+SUMIFS({F("q")},{F("nd")},统计年度)'
                      f'+SUMIFS({F("x0")},{F("nd")},统计年度)', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    sput(ws, f'D{r}', f'=SUMIFS({F("v")},{F("nd")},统计年度)', font=F_TOT, fill=FILL_AUTO, fmt=MONEY)
    sput(ws, f'E{r}', f'=SUMIFS({F("v")},{F("nd")},"<"&统计年度)', font=F_WARN, fill=FILL_AUTO, fmt=MONEY)
    sput(ws, f'F{r}', f'=SUMIFS({F("v")},{F("nd")},">"&统计年度)', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    sput(ws, f'H{r}', f'=SUM({F("v")})', font=F_TOT, fill=FILL_KPI, fmt=MONEY)
    sput(ws, f'G{r}', f'=ROUND($H{r}-$D{r}-$E{r}-$F{r},2)', font=F_IN, fill=FILL_AUTO, fmt=MONEY)
YT = Y0 + len(TEAMS)
sput(ws, f'A{YT}', '合计', font=F_TOT, fill=FILL_TOT)
for col in 'BCDEFGH':
    sput(ws, f'{col}{YT}', f'=SUM({col}{Y0}:{col}{YT-1})', font=F_TOT, fill=FILL_TOT, fmt=MONEY)

# ── ⑤ 核对与留意 ──
R5 = YT + 3
sec(R5, ' ⑤ 核对与留意', C_WARN)
hdr(R5 + 1, [('A', '项目'), ('B', '数值 / 口径一'), ('C', '口径二'), ('D', '结论')], C_WARN)
POS = '+'.join(f'SUMIF({TR(sn, "v")},">0")' for sn, _L in TEAMS)
CHK = [('应收余额合计 ＝ 总应收 − 已收合计', f'=$T${TT}', f'=ROUND($J${TT}-$O${TT},2)', True,
        '不平说明有行填了收款 / 金额，却没写使用单位和项目名称（这种行不会算进应收余额），或者数写在了第 %d 行以后。' % NR),
       ('开票应收额 ＋ 未开票应收 ＝ 还欠的钱', f'=ROUND($Q${TT}+$S${TT},2)', f'=ROUND({POS},2)', True, ''),
       ('五种情况表的应收余额 ＝ ② 表应收余额', f'={TC}{R3E}', f'=$T${TT}', True, ''),
       ('按年度拆开之后加回去 ＝ 累计应收余额', f'=ROUND($D${YT}+$E${YT}+$F${YT}+$G${YT},2)', f'=$H${YT}', True, ''),
       ('多收了的行（应收余额是负数）', '=' + '+'.join(f'COUNTIF({TR(sn, "v")},"<-0.005")' for sn, _L in TEAMS),
        None, False, '一般是单价 / 台量没填全，或者预收了下一期。'),
       ('没填单价却已经收了钱的行', '=' + '+'.join(f'COUNTIF({TR(sn, "w")},"未填单价，已收*")' for sn, _L in TEAMS),
        None, False, '补上单价、台量，应收余额才准。'),
       ('选了付款周期但没填合同起止的行', '=' + '+'.join(f'(COUNTIF({TR(sn, "zq")},"?*")-COUNTIFS({TR(sn, "zq")},"?*",{TR(sn, "tn")},">0"))' for sn, _L in TEAMS),
        None, False, '没有合同起止日，算不出分期到了第几期。')]
for i, (nm, a, b, is_chk, note) in enumerate(CHK):
    r = R5 + 2 + i
    sput(ws, f'A{r}', nm, font=F_TXT, fill=FILL_AUTO, align=CL_)
    sput(ws, f'B{r}', a, font=F_IN, fill=FILL_AUTO, fmt=MONEY if is_chk else NUM)
    sput(ws, f'C{r}', b, font=F_IN, fill=FILL_AUTO, fmt=MONEY)
    if is_chk:
        sput(ws, f'D{r}', f'=IF(ABS(N($B{r})-N($C{r}))<0.005,"✔ 平","★ 差 "&TEXT(N($B{r})-N($C{r}),"#,##0.00"))',
             font=F_TOT, fill=FILL_KPI)
    else:
        sput(ws, f'D{r}', f'=IF(N($B{r})=0,"✔ 没有","留意 "&N($B{r})&" 行")', font=F_TOT, fill=FILL_KPI)
    sput(ws, f'E{r}', note, font=F_NOTE, align=CL_, border=None)
    ws.merge_cells(f'E{r}:V{r}')
CK_END = R5 + 1 + len(CHK)

swidths(ws, {'A': 30, 'B': 16, 'C': 14, 'D': 14, 'E': 14, 'F': 14, 'G': 14, 'H': 14, 'I': 14,
             'J': 14, 'K': 13, 'L': 13, 'M': 13, 'N': 13, 'O': 13, 'P': 9, 'Q': 14, 'R': 14,
             'S': 13, 'T': 14, 'U': 14, 'V': 13})
ws.column_dimensions['M'].width = 24
ws.freeze_panes = 'B4'
ws.page_setup.orientation = 'landscape'
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.sheet_properties.tabColor = C_RPT
ws.sheet_view.showGridLines = False
print(f'  ✓ 汇总  ①梯子 ②各组回款(第{T0}-{TT}行) ③五种情况 ④按年度 ⑤核对(到第{CK_END}行)')

# ══════════════════════════════════════════════════════════════
# 填写说明
# ══════════════════════════════════════════════════════════════
if '填写说明' in wb.sheetnames:
    del wb['填写说明']
wd = wb.create_sheet('填写说明')
stitle(wd, '填 写 说 明 · 这次在原表上改了什么、收款区怎么填', 'H')
DOC = [
    ('sec', '一、原来的框架没动'),
    ('p', '《电梯档案》还是一台梯一行；各维保组还是一张表、一个合同（或一期）一行；第 2 行还是合计。'
          '原来的列一列没删，只是在「收款区」插了几列、在最右边追加了几列。浅蓝底表头＝公式自动算，浅黄底表头＝新加的要手填的。'),
    ('p', '公式已经铺到第 %d 行，往下接着填就行；不够了选中最后一行往下拖。' % NR),
    ('sec', '二、收款区：五种情况各填哪一格（最重要）'),
    ('p', '收款区现在是这样一排：开票日期｜开票金额｜收款日期｜收款金额｜收现金额｜收款不开票｜开票应收额｜到账未开票｜未开票应收｜应收余额｜回款情况｜收款备注。'
          '前面六格手填，后面除了「收款备注」都是公式。'),
    ('p', '① 不开票的现金账 → 填「收现金额」。'),
    ('p', '② 收款不开票（客户转账 / 微信 / 开收据，确定不开发票的）→ 填「收款不开票」。'),
    ('p', '③ 开票收款（票开了、钱也到了）→「开票金额」「收款金额」都填，两个数一样就两清。'),
    ('p', '④ 开票未到账 → 只填「开票金额」，「开票应收额」自动显示还差多少钱没到。'),
    ('p', '⑤ 到账未开票 → 只填「收款金额」，「到账未开票」自动显示还欠客户多少发票（待补票）；补开了票，把「开票金额」改上去就消了。'),
    ('p', '一句话：要开发票的钱记「收款金额」（不管是转账还是现金）；不开发票的现金记「收现金额」；不开发票的转账 / 微信 / 收据记「收款不开票」。'),
    ('p', '同一个合同分几次收的，金额就累加着填（可以直接写 =3600+3600），日期写最近一次，明细写在「收款备注」里。'
          '分期多的大合同，也可以照陆莹杨凯表里亿斯源那样，一期拆一行。'),
    ('p', '自动算出来的四个数：'),
    ('p', '　　应收余额（最后欠款）＝ 总应收 − 收款金额 − 收现金额 − 收款不开票；'),
    ('p', '　　开票应收额 ＝ 开票金额 − 收款金额（票开了钱没到的部分，最多不超过应收余额）；'),
    ('p', '　　未开票应收 ＝ 应收余额里还没开票的那部分；开票应收额 ＋ 未开票应收 ＝ 应收余额；'),
    ('p', '　　到账未开票 ＝ 收款金额 − 开票金额（钱到了票还没开，要补票）。'),
    ('p', '「回款情况」一格把这些用一句话说清楚，比如「欠 3,600.00（开票未到账 3,600.00）｜收现 500.00」、「✔ 已结清」、「多收 200.00」。'),
    ('sec', '三、合同没回来、合同到期'),
    ('p', '「到期提醒」列（F 列）重写了：'),
    ('p', '　　A 列写了「未见合同 / 已送出未回 / 未签」（含“未”或“送出”两个字都算），或者这一行只记了单位名、合同起止和单价都没填 → 显示【合同未回】；'),
    ('p', '　　合同止期已过 → 【已到期】；30 天内到期 → 【即将到期，剩余N天】（天数在《汇总》D3 改）；没填止期 → 【未填合同期】；'),
    ('p', '　　右边「维保状态」选了安装免保 / 技术免保 / 质保期内 → 后面再跟一个【安装免保】之类的标记。'),
    ('p', '原来的公式没填止期的空行也会显示【已到期】，现在修掉了。红底＝已到期或合同未回，黄底＝即将到期。'),
    ('sec', '四、合同款、合同外的款：分开登记，合并收款'),
    ('p', '原来就是这么设计的：「合同应收金额」＝ 单价 × 台量 × 年限（年限空着按 1 年）；「合同外维修费」手填（可以写 =250+330）；'
          '「总应收」＝ 两个相加。收款不分合同内外，统一冲总应收。原来有好几行漏了「总应收」公式，导致合计少算，这次全部补齐了。'),
    ('sec', '五、分季度 / 半月 / 半年付款的，回款怎么看'),
    ('p', '在最右边「付款周期」选：一次性 / 半年付 / 季度付 / 月付 / 半月付（原表备注里写了“半年请款一次”“季度付款”“七日内付”的，已经帮你预选好了，核对一下）。'),
    ('p', '只要填了合同起止日期，就自动算：'),
    ('p', '　　分期进度 —— 比如「季度付 第3/4期」；'),
    ('p', '　　截至今天应收 —— 合同应收金额 × 已到期期数 ÷ 总期数 ＋ 合同外维修费（合同外的按立即应收算）；'
          '每一期从期初起算，过了《汇总》F3「付款宽限天数」（默认 7 天，对应原表的「七日内付」）才算到期；'),
    ('p', '　　逾期未收 —— 截至今天应收 − 已收的钱，大于 0 标红，就是该催的。'),
    ('p', '这样大合同就算一年只记一行，也能看出现在应该收到多少、差了多少。《汇总》② 表最后两列是各组合计。'),
    ('sec', '六、26 年的欠款接到 27 年（连续多年用）'),
    ('p', '合同续签了，就在同一张组表里新加一行（使用单位、项目名称跟原来一模一样，合同起止写新的），旧的那一行别删、别改。'
          '这样：'),
    ('p', '　　新行的「往年欠款」会自动显示这个项目以前各期还没收回来的钱；'),
    ('p', '　　「项目累计欠款」是这个项目所有年份加起来一共还欠多少；'),
    ('p', '　　《汇总》④ 把「统计年度」改成 2027，就能看到 27 年合同的应收 / 已收 / 余额，以及 26 年及以前结转过来的欠款。'),
    ('p', '年度按「合同起始日」算：2026-09-07 开始的合同算 2026 年度。没填起止日的行，归在「没填合同期的」那一列。'),
    ('sec', '七、维保状态（安装免保 / 技术免保）'),
    ('p', '《电梯档案》W 列「维保状态」加了下拉：正常维保 / 安装免保 / 技术免保 / 质保期内 / 暂停维保 / 已解约。'
          '选了免保类的，在 X 列填免保到期日，Y 列会提醒「免保剩N天」「▲免保已到期，转收费」。'),
    ('p', 'Z 列是年检 / 限速器校验 / 载重试验提醒，原表里 2026.09、2027.8、2026.10.1、真日期这几种写法都能认（只写到月的按当月最后一天算）。'),
    ('p', '各维保组表最右边也有「维保状态」下拉，按合同标注，《汇总》② 表会统计免保项目数。'),
    ('sec', '八、顺手修掉的几处原表问题'),
    ('p', '① 到期提醒：没填止期的行（包括空行）原来显示【已到期】，现在不会了。'),
    ('p', '② 总应收：胡春平、陆莹杨凯等表有好些行没有「总应收」公式，合计比「合同应收」还少（截图里 77,400 对 68,980 就是这个原因），已全部补齐。'),
    ('p', '③ 开票应收额：胡春平表后面几行的公式减了「收款日期」，是错的，已统一改掉。'),
    ('p', '④ 应纳税额：原来写死「÷1.06」，改成跟着「税率」那一格走。'),
    ('p', '⑤ 冻结窗格统一成「锁住前三列和前两行」，往右翻收款区的时候单位名称还在。'),
]
r = 3
for kind, txt in DOC:
    if kind == 'sec':
        sput(wd, f'A{r}', txt, font=F_SEC, fill=FILL_SEC, align=CL_)
        wd.merge_cells(f'A{r}:H{r}')
        wd.row_dimensions[r].height = 24
    else:
        sput(wd, f'A{r}', txt, font=F_TXT, align=CT, border=None)
        wd.merge_cells(f'A{r}:H{r}')
        wd.row_dimensions[r].height = max(18, 17 * (1 + len(txt) // 62))
    r += 1
swidths(wd, {c: 18 for c in 'ABCDEFGH'})
wd.sheet_view.showGridLines = False
print('  ✓ 填写说明')

for s in wb.worksheets:
    s.sheet_view.tabSelected = (s.title == '汇总')
wb.active = 0
