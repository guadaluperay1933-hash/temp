# -*- coding: utf-8 -*-
"""研发费用归集与分摊台账 —— 泡棉生产 / 泡棉PET贴合 / 分切工艺（3C·新能源汽车）"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter as gcl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName

YEAR = 2026
NP = 10                       # 研发项目位数 RD01-RD10
MONTHS = ["%d-%02d" % (YEAR, m) for m in range(1, 13)]
DS, DE = 6, 245               # 四张月度台账的数据行区间（12月×20行）
OS_, OE = 4, 153              # 其他研发费用录入区间
ROS, ROE = 4, 53              # 人员名册 / 资产台账 数据行区间
IS_, IE = 6, 125              # 无形资产费用分配表 数据行区间

FN = "微软雅黑"
OUT = "/home/user/temp/研发费用归集与分摊台账_泡棉贴合分切_%d.xlsx" % YEAR

# ---------- 样式 ----------
C_TITLE  = "1F4E79"
C_GRP    = "2E75B6"
C_SUB    = "BDD7EE"
C_INPUT  = "FFF9E3"
C_CALC   = "F2F2F2"
C_TOTAL  = "DDEBF7"
C_NOTE   = "FFF2CC"

thin = Side(style="thin", color="9BA6B2")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

F_TITLE = Font(name=FN, size=15, bold=True, color="FFFFFF")
F_GRP   = Font(name=FN, size=10, bold=True, color="FFFFFF")
F_SUB   = Font(name=FN, size=10, bold=True, color="1F3864")
F_BODY  = Font(name=FN, size=10)
F_CALC  = Font(name=FN, size=10, color="404040")
F_BOLD  = Font(name=FN, size=10, bold=True)
F_NOTE  = Font(name=FN, size=9, color="843C0C")
F_RED   = Font(name=FN, size=9, color="C00000")
F_H2    = Font(name=FN, size=11, bold=True, color="1F4E79")

A_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
A_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
A_LT= Alignment(horizontal="left", vertical="top", wrap_text=True)
A_R = Alignment(horizontal="right", vertical="center")

FMT_H  = '0.00'          # 工时
FMT_M  = '#,##0.00'      # 金额
FMT_P  = '0.00%'         # 比例

fill = lambda c: PatternFill("solid", fgColor=c)

def title_bar(ws, last_col, text, note=None, note_row_h=30):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    c = ws.cell(1, 1, text); c.font = F_TITLE; c.fill = fill(C_TITLE); c.alignment = A_C
    ws.row_dimensions[1].height = 30
    if note is not None:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
        c = ws.cell(2, 1, note); c.font = F_NOTE; c.fill = fill(C_NOTE); c.alignment = A_L
        ws.row_dimensions[2].height = note_row_h

def box(ws, r1, c1, r2, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = BORDER

def setup_print(ws, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w

wb = openpyxl.Workbook()
wb.remove(wb.active)
# ============================================================
# 1. 使用说明
# ============================================================
ws = wb.create_sheet("使用说明")
widths(ws, {"A": 4, "B": 16, "C": 26, "D": 30, "E": 30, "F": 26, "G": 22})
title_bar(ws, 7, "研发费用归集与分摊台账  使用说明", None)
ws["A1"].value = "研发费用归集与分摊台账（%d年度）  使用说明" % YEAR

rows = [
    ("H2", "一、适用范围", None),
    ("T", "", "本台账适用于以泡棉发泡成型、泡棉与PET基材贴合（复合）、分切/模切为主要工艺，产品应用于3C电子与新能源汽车板块的企业。"),
    ("T", "", "针对研发人员、仪器设备与无形资产在研发与生产之间“混合使用”的情形，按工时法进行费用分配，形成《费用分配说明》及计算证据材料，"
            "满足高新技术企业认定、研发费用加计扣除等对“从事研发活动的人员（含外聘人员）和用于研发活动的仪器、设备、无形资产的费用分配说明"
            "（包括工作使用情况记录及费用分配计算证据材料）”的资料要求。"),
    ("H2", "二、核算原则", None),
    ("T", "", "① 按月核算、按年汇总：每月末由研发人员填报工时，财务据以计算分摊，年末自动汇总；"),
    ("T", "", "② 人员人工费用：当月某项目应分摊人工费 = 该人员当月总薪酬（含工资、社保、公积金、外聘劳务费）×（该项目研发工时 ÷ 该人员当月总工时）；"),
    ("T", "", "③ 仪器设备折旧/使用费：当月某项目应分摊设备费 = 该设备当月折旧额（或租赁费）×（该项目使用工时 ÷ 该设备当月总运行工时）；"),
    ("T", "", "④ 无形资产摊销：当月某项目应分摊的摊销额 = 该无形资产当月摊销额 ×（该项目使用工时 ÷ 该资产当月总使用工时），由研发按月填报无形资产使用工时；"),
    ("T", "", "⑤ 研发专用人员/设备/无形资产记录全额工时；研发生产共用的记录实际投入研发的工时；"),
    ("T", "", "⑥ 未分配至研发项目的部分，全额计入生产成本，做到“研发与生产各归其位、合计不重不漏”。"),
    ("T", "", "⑦ 数据来源：薪酬数据源自月度工资表；折旧数据源自月度固定资产折旧明细表；摊销数据源自无形资产摊销表。"),
    ("H2", "三、每月操作流程（5步）", None),
    ("S", "第1步  基础维护", "在【项目清单】【人员名册】【资产台账】中维护当期在研项目、研发人员（含外聘）、研发与生产共用的设备及无形资产。基础信息只需变动时维护。"),
    ("S", "第2步  工时填报", "研发部门于次月3日前完成《研发人员工时分配表》《仪器、设备工时分配表》《无形资产工时分配表》填报（签批用纸质表见【附表1】【附表2】【附表3】），"
                          "经项目负责人、部门负责人签字确认后交财务。签字件为“工作使用情况记录”原始证据，须装订留存。"),
    ("S", "第3步  数据录入", "财务将签批后的工时录入【研发人员工时分配表】与【仪器设备工时分配表】。注意“工时校验”列必须为0（红色即为不平）。"
                          "研发专用人员/设备的“非研发工时”应填0。"),
    ("S", "第4步  费用分摊", "在【研发人员工时费用分配表】录入当月薪酬构成；在【仪器设备工时费用分配表】核对当月折旧/租赁金额；"
                          "在【无形资产费用分配表】核对当月摊销额。三表的各项目分摊金额均按对应工时占比自动计算。"
                          "【其他研发费用】按需登记直接投入等费用（选填）。"),
    ("S", "第5步  汇总复核", "查看【月度归集汇总】与【年度汇总】（含表四勾稽校验，差异须为0），与研发支出辅助账、财务账核对一致；"
                          "同步更新【指标监测】中的收入与人数，跟踪三项高企指标；年末打印【费用分配说明】签章存档。"),
    ("H2", "四、本台账与《研发费用归集》制表要求的对应关系", None),
    ("S", "指标监测", "《各项指标监测》—— 科技人员占比≥10%、研发费用占销售收入比例（＜5000万≥5%／5000万~2亿≥4%／＞2亿≥3%）、高新收入占比≥60%。对应页签：【指标监测】"),
    ("S", "人员人工", "《研发人员工时分配表》研发按月填报 → 【研发人员工时分配表】；《研发人员工时费用分配表》财务按月按项目核算、按年汇总 → 【研发人员工时费用分配表】"),
    ("S", "仪器设备", "《仪器、设备工时分配表》研发按月填报 → 【仪器设备工时分配表】；《仪器、设备工时费用分配表》财务按月按项目核算、按年汇总 → 【仪器设备工时费用分配表】"),
    ("S", "无形资产", "《无形资产工时分配表》研发按月填报 → 【无形资产工时分配表】；《无形资产费用分配表》财务按月按项目核算、按年汇总 → 【无形资产费用分配表】"),
    ("H2", "五、单元格颜色约定", None),
]
r = 4
for kind, a, b in rows:
    if kind == "H2":
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        c = ws.cell(r, 1, a); c.font = F_H2; c.fill = fill(C_SUB); c.alignment = A_L
        ws.row_dimensions[r].height = 22
    elif kind == "T":
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        c = ws.cell(r, 2, b); c.font = F_BODY; c.alignment = A_LT
        ws.row_dimensions[r].height = 15 + 13 * (len(b) // 62)
    else:
        c = ws.cell(r, 2, a); c.font = F_BOLD; c.alignment = A_L; c.fill = fill(C_TOTAL); c.border = BORDER
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        c2 = ws.cell(r, 3, b); c2.font = F_BODY; c2.alignment = A_LT; c2.border = BORDER
        ws.row_dimensions[r].height = 16 + 13 * (len(b) // 66)
    r += 1

legend = [(C_INPUT, "浅黄底 = 需人工录入", "由研发/财务填写的原始数据"),
          (C_CALC,  "浅灰底 = 公式自动计算", "请勿手工覆盖，改动会破坏勾稽关系"),
          (C_TOTAL, "浅蓝底 = 合计/汇总行", "自动生成，用于与账面核对")]
for f, t1, t2 in legend:
    c = ws.cell(r, 2, ""); c.fill = fill(f); c.border = BORDER
    c = ws.cell(r, 3, t1); c.font = F_BOLD; c.alignment = A_L
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
    c = ws.cell(r, 4, t2); c.font = F_BODY; c.alignment = A_L
    ws.row_dimensions[r].height = 18
    r += 1

r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
c = ws.cell(r, 1, "六、留存备查资料清单（应对“费用分配说明”核查）"); c.font = F_H2; c.fill = fill(C_SUB); c.alignment = A_L
r += 1
keep = [
 "1. 本台账电子版（按年一个文件）及打印的【月度归集汇总】【年度汇总】；",
 "2. 经签批的《研发人员工时分配表》《仪器、设备工时分配表》《无形资产工时分配表》原件（附表1、附表2、附表3）；",
 "3. 《研发人员与设备费用分配说明》（本文件【费用分配说明】页，加盖公章）；",
 "4. 研发人员名册、劳动合同/外聘劳务合同、当月工资表及社保公积金缴纳凭证；",
 "5. 设备固定资产卡片、折旧计提表、租赁合同及付款凭证，无形资产摊销表；",
 "6. 研发项目立项decision文件、研发项目计划书、阶段性研发记录（试验记录、检测报告等）；",
 "7. 研发支出辅助账及与财务账、纳税申报表的核对说明。",
]
keep[5] = "6. 研发项目立项决议文件、项目计划书、阶段性研发记录（配方试验记录、贴合/分切工艺试验记录、检测报告等）；"
for t in keep:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c = ws.cell(r, 2, t); c.font = F_BODY; c.alignment = A_LT
    ws.row_dimensions[r].height = 15 + 13 * (len(t) // 62)
    r += 1
r += 1
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
c = ws.cell(r, 2, "提示：各录入表首行标有“★示例”的行为填写示范，正式使用前请整行删除，以免混入统计。")
c.font = F_RED; c.alignment = A_L
setup_print(ws, landscape=False)

# ============================================================
# 2. 项目清单
# ============================================================
PROJ = [
 ("RD01", "高回弹低压缩永久变形3C用聚氨酯泡棉开发",        "3C电子",      "发泡成型",   "2026-01-05", "2026-12-20"),
 ("RD02", "动力电池模组用高阻燃缓冲泡棉研发",              "新能源汽车",  "发泡成型",   "2026-01-05", "2026-12-20"),
 ("RD03", "泡棉与PET基材无溶剂贴合工艺开发",              "通用",        "PET贴合",    "2026-02-01", "2026-11-30"),
 ("RD04", "超薄泡棉高精度分切及毛边控制技术研究",          "3C电子",      "分切",       "2026-03-01", "2026-12-31"),
 ("RD05", "新能源汽车电池包用导热缓冲复合材料开发",        "新能源汽车",  "PET贴合",    "2026-03-01", "2027-02-28"),
 ("RD06", "3C电子用超薄黑色遮光泡棉制备技术",             "3C电子",      "发泡成型",   "", ""),
 ("RD07", "环保型水性胶泡棉贴合技术开发",                  "通用",        "PET贴合",    "", ""),
 ("RD08", "泡棉模切自动化在线检测技术研究",                "3C电子",      "模切",       "", ""),
 ("RD09", "", "", "", "", ""),
 ("RD10", "", "", "", "", ""),
]
ws = wb.create_sheet("项目清单")
widths(ws, {"A":6,"B":10,"C":44,"D":13,"E":13,"F":13,"G":13,"H":12,"I":11,"J":24})
title_bar(ws, 10, "二、研发项目清单（%d年度）" % YEAR,
          "填写说明：项目编号 RD01～RD10 为全表引用主键，请勿改动编号本身；项目名称、板块等可自由维护，编号顺序不可打乱。"
          "通常一个年度应有 5～10 个研发项目（至少 5 个），未使用的编号留空即可，不影响各表计算。")
hdr = ["序号","项目编号","研发项目名称","应用板块","主要工艺环节","立项日期","计划结题日期","项目负责人","项目状态","备注"]
for i, h in enumerate(hdr, 1):
    c = ws.cell(3, i, h); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C; c.border = BORDER
ws.row_dimensions[3].height = 30
for i, p in enumerate(PROJ):
    r = 4 + i
    ws.cell(r, 1, i + 1).font = F_BODY
    ws.cell(r, 2, p[0]).font = F_BOLD
    ws.cell(r, 3, p[1]).font = F_BODY
    ws.cell(r, 4, p[2]).font = F_BODY
    ws.cell(r, 5, p[3]).font = F_BODY
    ws.cell(r, 6, p[4]).font = F_BODY
    ws.cell(r, 7, p[5]).font = F_BODY
    ws.cell(r, 8, "").font = F_BODY
    ws.cell(r, 9, "在研" if p[1] else "").font = F_BODY
    ws.cell(r,10, "").font = F_BODY
    for cc in range(1, 11):
        cell = ws.cell(r, cc); cell.alignment = A_C if cc != 3 else A_L; cell.border = BORDER
        if cc >= 3: cell.fill = fill(C_INPUT)
    ws.row_dimensions[r].height = 22
box(ws, 3, 1, 13, 10)
ws.freeze_panes = "C4"
setup_print(ws)

# ============================================================
# 3. 人员名册
# ============================================================
STAFF = [
 ("YF001","（示例）张工","研发中心","配方工程师","在职","研发专用","是"),
 ("YF002","（示例）李工","研发中心","工艺工程师","在职","研发专用","是"),
 ("YF003","（示例）王工","生产部","贴合技师（研发生产共用）","在职","研发生产共用","是"),
 ("WP001","（示例）赵博","外聘专家","高分子材料顾问","外聘","研发生产共用","是"),
]
ws = wb.create_sheet("人员名册")
widths(ws, {"A":6,"B":11,"C":13,"D":15,"E":24,"F":11,"G":15,"H":13,"I":15,"J":24})
title_bar(ws, 10, "三、研发人员名册（含外聘人员）",
          "填写说明：①工号为工时表引用主键，须唯一且不可重复；②“人员使用性质”区分研发专用与研发生产共用——研发专用人员在工时表中记录全额工时，"
          "研发生产共用人员记录实际投入研发的工时；③外聘人员“人员类别”选“外聘”，其劳务费在【研发人员工时费用分配表】的“外聘劳务费”列填列；"
          "④“是否科技人员”用于高新技术企业科技人员占比（不低于10%）指标监测。")
hdr = ["序号","工号","姓名","所属部门","岗位/承担研发任务","人员类别","人员使用性质","是否科技人员","入职/合作日期","备注"]
for i, h in enumerate(hdr, 1):
    c = ws.cell(3, i, h); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C; c.border = BORDER
ws.row_dimensions[3].height = 30
for i in range(ROS, ROE + 1):
    idx = i - ROS
    ws.cell(i, 1, idx + 1).font = F_BODY
    if idx < len(STAFF):
        s = STAFF[idx]
        for j, v in enumerate(s, 2):
            ws.cell(i, j, v)
        ws.cell(i, 10, "★示例行，请删除").font = F_RED
    for cc in range(1, 11):
        cell = ws.cell(i, cc)
        if cell.font is None or cell.font.color is None or cell.font.color.rgb != "00C00000":
            cell.font = F_BODY
        cell.alignment = A_C if cc not in (5,10) else A_L
        cell.border = BORDER
        if cc >= 2: cell.fill = fill(C_INPUT)
    ws.row_dimensions[i].height = 20
for f1, col in (('"在职,外聘"', "F"), ('"研发专用,研发生产共用"', "G"), ('"是,否"', "H")):
    d = DataValidation(type="list", formula1=f1, allow_blank=True); ws.add_data_validation(d)
    d.add("%s%d:%s%d" % (col, ROS, col, ROE))
ws.freeze_panes = "D4"
setup_print(ws)

# ============================================================
# 4. 资产台账（设备 + 无形资产）
# ============================================================
ASSET = [
 ("SB001","（示例）连续发泡生产线","设备","FP-2000","发泡成型","2023-06-15",1800000,10,15000,"折旧","研发生产共用"),
 ("SB002","（示例）自动贴合复合机","设备","LM-1300","PET贴合","2024-03-20",960000,10,8000,"折旧","研发生产共用"),
 ("SB003","（示例）高精度分切机","设备","SL-800","分切","2024-09-10",620000,10,5166.67,"折旧","研发生产共用"),
 ("SB004","（示例）万能материал试验机","设备","UTM-50","检测试验","2025-01-12",180000,5,3000,"折旧","研发生产共用"),
 ("SB005","（示例）租入激光模切设备","设备","LD-600","模切","2026-01-01",0,0,12000,"租赁费","研发生产共用"),
 ("WX001","（示例）泡棉配方模拟软件","无形资产","V3.0","研发通用","2025-05-06",240000,5,4000,"摊销","研发专用"),
]
ASSET[3] = ("SB004","（示例）万能材料试验机","设备","UTM-50","检测试验","2025-01-12",180000,5,3000,"折旧","研发生产共用")
ws = wb.create_sheet("资产台账")
widths(ws, {"A":6,"B":11,"C":26,"D":11,"E":13,"F":13,"G":13,"H":14,"I":11,"J":15,"K":11,"L":15,"M":13,"N":22})
title_bar(ws, 14, "四、研发用仪器设备及无形资产台账（含研发生产共用资产）",
          "填写说明：资产编号为工时表引用主键，须唯一。“月折旧/摊销/租赁额”按财务实际计提数填列，【仪器设备工时费用分配表】表将自动取数（当月实际不同可在该表直接覆盖）。"
          "研发与生产共用的设备必须登记，其研发使用工时在【仪器设备工时分配表】中据实填报（研发专用设备记录全额运行工时）；"
          "无形资产（配方模拟软件、专利/专有技术等）同样登记，其使用工时在【无形资产工时分配表】中据实填报，摊销额按使用工时占比分摊。"
          "数据来源：月度固定资产折旧明细表 / 月度无形资产摊销表。")
hdr = ["序号","资产编号","资产名称","资产类别","规格型号","所属工序","取得日期","原值(元)","预计使用年限(年)",
       "月折旧/摊销/租赁额(元)","费用类型","资产使用性质","存放地点","备注"]
for i, h in enumerate(hdr, 1):
    c = ws.cell(3, i, h); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C; c.border = BORDER
ws.row_dimensions[3].height = 38
for i in range(ROS, ROE + 1):
    idx = i - ROS
    ws.cell(i, 1, idx + 1).font = F_BODY
    if idx < len(ASSET):
        a = ASSET[idx]
        for j, v in enumerate(a, 2):
            ws.cell(i, j, v)
        ws.cell(i, 14, "★示例行，请删除").font = F_RED
    for cc in range(1, 15):
        cell = ws.cell(i, cc)
        if not (cell.font and cell.font.color and cell.font.color.rgb == "00C00000"):
            cell.font = F_BODY
        cell.alignment = A_C if cc not in (3, 14) else A_L
        cell.border = BORDER
        if cc >= 2: cell.fill = fill(C_INPUT)
    ws.cell(i, 8).number_format = FMT_M
    ws.cell(i, 10).number_format = FMT_M
    ws.row_dimensions[i].height = 20
for f1, col in (('"设备,无形资产"', "D"), ('"发泡成型,熟化,PET贴合,分切,模切,检测试验,研发通用,其他"', "F"),
                ('"折旧,摊销,租赁费"', "K"), ('"研发专用,研发生产共用"', "L")):
    d = DataValidation(type="list", formula1=f1, allow_blank=True); ws.add_data_validation(d)
    d.add("%s%d:%s%d" % (col, ROS, col, ROE))
ws.freeze_panes = "D4"
setup_print(ws)
MONTH_DV = '"' + ",".join(MONTHS) + '"'

def proj_header(ws, left, group_label, right, pc0):
    """三行表头：固定列纵向合并 rows3-5；项目列 row3 组名 / row4 编号 / row5 名称"""
    n_left = len(left)
    for i, lab in enumerate(left, 1):
        ws.merge_cells(start_row=3, start_column=i, end_row=5, end_column=i)
        c = ws.cell(3, i, lab); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C
    ws.merge_cells(start_row=3, start_column=pc0, end_row=3, end_column=pc0 + NP - 1)
    c = ws.cell(3, pc0, group_label); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C
    for k in range(NP):
        col = pc0 + k
        c = ws.cell(4, col, "=项目清单!$B$%d" % (4 + k)); c.font = F_SUB; c.fill = fill(C_SUB); c.alignment = A_C
        c = ws.cell(5, col, '=IF(%s4="","",IFERROR(VLOOKUP(%s4,项目清单!$B$4:$C$13,2,0)&"",""))' % (gcl(col), gcl(col)))
        c.font = Font(name=FN, size=8, color="1F3864"); c.fill = fill(C_SUB); c.alignment = A_C
    for j, lab in enumerate(right):
        col = pc0 + NP + j
        ws.merge_cells(start_row=3, start_column=col, end_row=5, end_column=col)
        c = ws.cell(3, col, lab); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C
    last = pc0 + NP + len(right) - 1
    ws.row_dimensions[3].height = 26; ws.row_dimensions[4].height = 18; ws.row_dimensions[5].height = 34
    box(ws, 3, 1, 5, last)
    return last

# ============================================================
# 5. 研发人员工时分配表
# ============================================================
ws = wb.create_sheet("研发人员工时分配表")
PC = 8   # H列起为项目列
left  = ["年月\n(YYYY-MM)","工号","姓名","所属部门","人员类别","当月出勤\n总工时","非研发工时\n(生产及其他)"]
right = ["研发工时\n合计","工时校验\n(须为0)","研发工时\n占比","本月主要研发工作内容","备注"]
title_bar(ws, PC + NP + len(right) - 1,
  "五、《研发人员工时分配表》（研发部门按月填报，%d年度）" % YEAR,
  "录入说明：①由研发部门按月填报，依据经签批的《附表1-研发人员工时分配表(签批用)》逐人逐月录入；②“当月出勤总工时”为该人当月全部有效工作工时（含加班）；"
  "③研发专用人员记录全额工时（非研发工时填0），研发生产共用人员按实际投入分别记录研发工时与非研发（生产）工时；"
  "④“工时校验”须为0，显示红色表示总工时与明细不平，须更正后方可分摊；⑤同一人当月只填一行；⑥灰底列为公式自动计算，请勿覆盖。")
last = proj_header(ws, left, "各研发项目分配工时（小时）", right, PC)
widths(ws, {"A":11,"B":10,"C":11,"D":13,"E":10,"F":11,"G":12})
for k in range(NP): widths(ws, {gcl(PC + k): 9})
widths(ws, {gcl(PC+NP):10, gcl(PC+NP+1):10, gcl(PC+NP+2):10, gcl(PC+NP+3):30, gcl(PC+NP+4):18})

EX_P = [
 ("2026-09","YF001",176,16,{0:100,1:60},"完成低压变形泡棉配方第3轮小试及压缩永久变形验证"),
 ("2026-09","YF003",176,120,{2:36,4:20},"无溶剂贴合线速与胶量匹配试验（研发生产共用岗位）"),
 ("2026-09","WP001",80,0,{1:50,4:30},"外聘专家：阻燃体系选型指导与试验方案评审"),
]
for r in range(DS, DE + 1):
    ws.cell(r,3, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,人员名册!$B$4:$G$53,2,0)&"","名册中未登记"))' % (r, r))
    ws.cell(r,4, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,人员名册!$B$4:$G$53,3,0)&"",""))' % (r, r))
    ws.cell(r,5, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,人员名册!$B$4:$G$53,5,0)&"",""))' % (r, r))
    ws.cell(r,18,'=IF($A%d="","",SUM(%s%d:%s%d))' % (r, gcl(PC), r, gcl(PC+NP-1), r))
    ws.cell(r,19,'=IF($A%d="","",ROUND($F%d-$G%d-$R%d,2))' % (r, r, r, r))
    ws.cell(r,20,'=IF(OR($A%d="",$F%d=0),"",$R%d/$F%d)' % (r, r, r, r))
    idx = r - DS
    if idx < len(EX_P):
        e = EX_P[idx]
        ws.cell(r,1,e[0]); ws.cell(r,2,e[1]); ws.cell(r,6,e[2]); ws.cell(r,7,e[3])
        for k, v in e[4].items(): ws.cell(r, PC + k, v)
        ws.cell(r,21,e[5])
        ws.cell(r,22,"★示例行，请删除")
    ws.cell(r,23,'=IF($A%d="","",$A%d&"|"&$B%d)' % (r, r, r))
    for c in range(1, last + 1):
        cell = ws.cell(r, c); cell.border = BORDER
        cell.alignment = A_C if c not in (21,22) else A_L
        if c in (3,4,5,18,19,20):
            cell.font = F_CALC; cell.fill = fill(C_CALC)
        elif c == 22 and cell.value:
            cell.font = F_RED; cell.fill = fill(C_INPUT)
        else:
            cell.font = F_BODY; cell.fill = fill(C_INPUT)
        if c in (6,7,18,19) or PC <= c <= PC+NP-1: cell.number_format = FMT_H
        if c == 20: cell.number_format = FMT_P
    ws.row_dimensions[r].height = 18

ws.conditional_formatting.add("S%d:S%d" % (DS, DE),
    FormulaRule(formula=['AND($A%d<>"",$S%d<>0)' % (DS, DS)],
                fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FN, size=10, bold=True, color="9C0006")))
dv = DataValidation(type="list", formula1=MONTH_DV, allow_blank=True); ws.add_data_validation(dv)
dv.add("A%d:A%d" % (DS, DE))
dv2 = DataValidation(type="list", formula1="=人员工号列表", allow_blank=True); ws.add_data_validation(dv2)
dv2.add("B%d:B%d" % (DS, DE))
ws.cell(5,23,"匹配主键(勿删)").font = Font(name=FN, size=8, color="808080")
ws.column_dimensions["W"].hidden = True
ws.freeze_panes = "D6"; ws.auto_filter.ref = "A5:%s%d" % (gcl(last), DE)
setup_print(ws)

# ============================================================
# 6. 仪器设备工时分配表
# ============================================================
ws = wb.create_sheet("仪器设备工时分配表")
left  = ["年月\n(YYYY-MM)","资产编号","资产名称","资产类别","所属工序","当月总运行\n(使用)工时","非研发使用工时\n(生产及其他)"]
right = ["研发使用\n工时合计","工时校验\n(须为0)","研发使用\n占比","备注"]
title_bar(ws, PC + NP + len(right) - 1,
  "七、《仪器、设备工时分配表》（研发部门按月填报，%d年度）" % YEAR,
  "录入说明：①由研发部门/设备管理员按月填报，依据经签批的《附表2-仪器设备工时分配表(签批用)》录入；②“当月总运行(使用)工时”为该设备当月实际运行总工时"
  "（生产+研发），是分摊的分母，须与设备点检表、生产报工记录一致；③研发专用设备记录全额运行工时（非研发使用工时填0），研发生产共用设备按实际使用分别记录；"
  "④本表只登记“设备”类资产，无形资产不在此填报（其摊销在【无形资产费用分配表】中按人工工时比例分摊）；⑤同一设备当月只填一行；⑥“工时校验”须为0。")
last = proj_header(ws, left, "各研发项目使用工时（小时）", right, PC)
widths(ws, {"A":11,"B":11,"C":24,"D":11,"E":12,"F":13,"G":14})
for k in range(NP): widths(ws, {gcl(PC + k): 9})
widths(ws, {gcl(PC+NP):11, gcl(PC+NP+1):10, gcl(PC+NP+2):10, gcl(PC+NP+3):22})

EX_E = [
 ("2026-09","SB001",320,260,{0:40,1:20}),
 ("2026-09","SB002",300,240,{2:40,4:20}),
 ("2026-09","SB003",280,250,{3:30}),
 ("2026-09","SB005",120,0,{1:70,4:50}),
]
for r in range(DS, DE + 1):
    ws.cell(r,3, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,2,0)&"","台账中未登记"))' % (r, r))
    ws.cell(r,4, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,3,0)&"",""))' % (r, r))
    ws.cell(r,5, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,5,0)&"",""))' % (r, r))
    ws.cell(r,18,'=IF($A%d="","",SUM(%s%d:%s%d))' % (r, gcl(PC), r, gcl(PC+NP-1), r))
    ws.cell(r,19,'=IF($A%d="","",ROUND($F%d-$G%d-$R%d,2))' % (r, r, r, r))
    ws.cell(r,20,'=IF(OR($A%d="",$F%d=0),"",$R%d/$F%d)' % (r, r, r, r))
    idx = r - DS
    if idx < len(EX_E):
        e = EX_E[idx]
        ws.cell(r,1,e[0]); ws.cell(r,2,e[1]); ws.cell(r,6,e[2]); ws.cell(r,7,e[3])
        for k, v in e[4].items(): ws.cell(r, PC + k, v)
        ws.cell(r,21,"★示例行，请删除")
    ws.cell(r,22,'=IF($A%d="","",$A%d&"|"&$B%d)' % (r, r, r))
    for c in range(1, last + 1):
        cell = ws.cell(r, c); cell.border = BORDER
        cell.alignment = A_C if c not in (3,21) else A_L
        if c in (3,4,5,18,19,20):
            cell.font = F_CALC; cell.fill = fill(C_CALC)
        elif c == 21 and cell.value:
            cell.font = F_RED; cell.fill = fill(C_INPUT)
        else:
            cell.font = F_BODY; cell.fill = fill(C_INPUT)
        if c in (6,7,18,19) or PC <= c <= PC+NP-1: cell.number_format = FMT_H
        if c == 20: cell.number_format = FMT_P
    ws.row_dimensions[r].height = 18

ws.conditional_formatting.add("S%d:S%d" % (DS, DE),
    FormulaRule(formula=['AND($A%d<>"",$S%d<>0)' % (DS, DS)],
                fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FN, size=10, bold=True, color="9C0006")))
dv = DataValidation(type="list", formula1=MONTH_DV, allow_blank=True); ws.add_data_validation(dv); dv.add("A%d:A%d" % (DS, DE))
dv2 = DataValidation(type="list", formula1="=资产编号列表", allow_blank=True); ws.add_data_validation(dv2); dv2.add("B%d:B%d" % (DS, DE))
ws.cell(5,22,"匹配主键(勿删)").font = Font(name=FN, size=8, color="808080")
ws.column_dimensions["V"].hidden = True
ws.freeze_panes = "D6"; ws.auto_filter.ref = "A5:%s%d" % (gcl(last), DE)
setup_print(ws)
# ============================================================
# 7. 研发人员工时费用分配表
# ============================================================
ws = wb.create_sheet("研发人员工时费用分配表")
PC = 14  # N列起为项目列；X=小计 Y=生产成本 Z=数据校验 AA/AB=隐藏辅助列
left = ["年月\n(YYYY-MM)","工号","姓名","人员类别","工资薪金\n(元)","社保\n(单位承担,元)","公积金\n(单位承担,元)",
        "外聘劳务费\n(元)","其他人工支出\n(元)","当月薪酬合计\n(元)","当月总工时","研发工时合计","研发工时占比\n(分摊系数)"]
right = ["研发人工费\n小计(元)","计入生产成本的\n人工费(元)"]
title_bar(ws, PC + NP + 2,
 "六、《研发人员工时费用分配表》（财务按月按项目核算、按年汇总，%d年度）" % YEAR,
 "计算依据：当月某项目应分摊人工费 = 该人员当月总薪酬（含工资、社保、公积金、外聘劳务费）×（该项目研发工时 ÷ 该人员当月总工时）。"
 "数据来源：薪酬数据源自【月度工资表】及社保公积金缴费凭证。录入说明：只需录入 A~I 列（年月、工号及薪酬构成；外聘人员在“外聘劳务费”列填列）；"
 "工时及各项目分摊金额由《研发人员工时分配表》按“年月+工号”自动匹配取数。请关注末列“数据校验”，出现⚠提示须先更正工时分配表。")
last = proj_header(ws, left, "各研发项目分摊人工费用（元）", right, PC)
CHK_C = last + 1          # Z 数据校验
H1, H2 = last + 2, last + 3   # AA 匹配行号 / AB 重复计数
ws.merge_cells(start_row=3, start_column=CHK_C, end_row=5, end_column=CHK_C)
c = ws.cell(3, CHK_C, "数据校验"); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C
ws.cell(5, H1, "匹配行号(勿删)").font = Font(name=FN, size=8, color="808080")
ws.cell(5, H2, "重复计数(勿删)").font = Font(name=FN, size=8, color="808080")
widths(ws, {"A":11,"B":10,"C":11,"D":10,"E":12,"F":13,"G":13,"H":12,"I":13,"J":13,"K":11,"L":11,"M":12})
for k in range(NP): widths(ws, {gcl(PC + k): 12})
widths(ws, {gcl(PC+NP):13, gcl(PC+NP+1):15, gcl(CHK_C):20})
ws.column_dimensions[gcl(H1)].hidden = True
ws.column_dimensions[gcl(H2)].hidden = True

WSHT, KEYCOL = "研发人员工时分配表", "W"
AA, AB = gcl(H1), gcl(H2)
EX_S = [("2026-09","YF001",15000,3900,1800,0,1200),
        ("2026-09","YF003",9000,2340,1080,0,600),
        ("2026-09","WP001",0,0,0,30000,0)]
for r in range(DS, DE + 1):
    ws.cell(r,3, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,人员名册!$B$4:$G$53,2,0)&"","名册中未登记"))' % (r, r))
    ws.cell(r,4, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,人员名册!$B$4:$G$53,5,0)&"",""))' % (r, r))
    ws.cell(r,10,'=IF($A%d="","",ROUND(SUM(E%d:I%d),2))' % (r, r, r))
    ws.cell(r,H1,'=IF($A%d="",0,IFERROR(MATCH($A%d&"|"&$B%d,%s!$%s$%d:$%s$%d,0),0))'
            % (r, r, r, WSHT, KEYCOL, DS, KEYCOL, DE))
    ws.cell(r,H2,'=IF($A%d="",0,COUNTIFS(%s!$A$%d:$A$%d,$A%d,%s!$B$%d:$B$%d,$B%d))'
            % (r, WSHT, DS, DE, r, WSHT, DS, DE, r))
    ws.cell(r,CHK_C,'=IF($A%d="","",IF($%s%d=0,"⚠ 工时分配表无此人当月记录",IF($%s%d>1,"⚠ 工时分配表存在重复行","√ 正常")))'
            % (r, AA, r, AB, r))
    ws.cell(r,11,'=IF(OR($A%d="",$%s%d=0),0,INDEX(%s!$F$%d:$F$%d,$%s%d))' % (r, AA, r, WSHT, DS, DE, AA, r))
    ws.cell(r,12,'=IF(OR($A%d="",$%s%d=0),0,INDEX(%s!$R$%d:$R$%d,$%s%d))' % (r, AA, r, WSHT, DS, DE, AA, r))
    ws.cell(r,13,'=IF(OR($A%d="",$K%d=0),"",$L%d/$K%d)' % (r, r, r, r))
    for j in range(NP):
        src = gcl(8 + j)   # 研发人员工时分配表 H..Q
        ws.cell(r, PC + j, '=IF(OR($A%d="",$K%d=0),"",ROUND($J%d*INDEX(%s!$%s$%d:$%s$%d,$%s%d)/$K%d,2))'
                % (r, r, r, WSHT, src, DS, src, DE, AA, r, r))
    ws.cell(r, PC+NP,   '=IF($A%d="","",ROUND(SUM(%s%d:%s%d),2))' % (r, gcl(PC), r, gcl(PC+NP-1), r))
    ws.cell(r, PC+NP+1, '=IF($A%d="","",ROUND($J%d-%s%d,2))' % (r, r, gcl(PC+NP), r))
    idx = r - DS
    if idx < len(EX_S):
        e = EX_S[idx]
        ws.cell(r,1,e[0]); ws.cell(r,2,e[1])
        for j, v in enumerate(e[2:], 5): ws.cell(r, j, v)
    for c in range(1, CHK_C + 1):
        cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C
        if c in (1,2) or 5 <= c <= 9:
            cell.font = F_BODY; cell.fill = fill(C_INPUT)
        else:
            cell.font = F_CALC; cell.fill = fill(C_CALC)
        if 5 <= c <= 10 or PC <= c <= PC + NP + 1: cell.number_format = FMT_M
        if c in (11,12): cell.number_format = FMT_H
        if c == 13: cell.number_format = FMT_P
    ws.row_dimensions[r].height = 18
dv = DataValidation(type="list", formula1=MONTH_DV, allow_blank=True); ws.add_data_validation(dv); dv.add("A%d:A%d" % (DS, DE))
dv2 = DataValidation(type="list", formula1="=人员工号列表", allow_blank=True); ws.add_data_validation(dv2); dv2.add("B%d:B%d" % (DS, DE))
ws.conditional_formatting.add("%s%d:%s%d" % (gcl(CHK_C), DS, gcl(CHK_C), DE),
    FormulaRule(formula=['LEFT($%s%d,1)="⚠"' % (gcl(CHK_C), DS)],
                fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FN, size=10, bold=True, color="9C0006")))
ws.freeze_panes = "E6"; ws.auto_filter.ref = "A5:%s%d" % (gcl(CHK_C), DE)
setup_print(ws)

# ============================================================
# 8. 仪器设备工时费用分配表
# ============================================================
ws = wb.create_sheet("仪器设备工时费用分配表")
PC = 10  # J列起为项目列；T=小计 U=生产成本 V=数据校验 W/X=隐藏辅助列
left = ["年月\n(YYYY-MM)","资产编号","资产名称","资产类别","费用类型","当月折旧/\n租赁额(元)",
        "当月总运行\n(使用)工时","研发使用\n工时合计","研发使用占比\n(分摊系数)"]
right = ["研发费用\n小计(元)","计入生产成本的\n折旧费用(元)"]
title_bar(ws, PC + NP + 2,
 "八、《仪器、设备工时费用分配表》（财务按月按项目核算、按年汇总，%d年度）" % YEAR,
 "计算依据：当月某项目应分摊设备费 = 该设备当月折旧额（或租赁费）×（该项目使用工时 ÷ 该设备当月总运行工时）。"
 "数据来源：折旧数据源自【月度固定资产折旧明细表】。录入说明：只需录入 A、B 两列（年月、资产编号，仅登记“设备”类资产）；"
 "F列“当月折旧/租赁额”自动取自【资产台账】，当月实际计提数不同时可直接覆盖填写，但须与折旧明细表一致。请关注末列“数据校验”。")
last = proj_header(ws, left, "各研发项目分摊折旧/租赁费用（元）", right, PC)
CHK_C = last + 1
H1, H2 = last + 2, last + 3
ws.merge_cells(start_row=3, start_column=CHK_C, end_row=5, end_column=CHK_C)
c = ws.cell(3, CHK_C, "数据校验"); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C
ws.cell(5, H1, "匹配行号(勿删)").font = Font(name=FN, size=8, color="808080")
ws.cell(5, H2, "重复计数(勿删)").font = Font(name=FN, size=8, color="808080")
widths(ws, {"A":11,"B":11,"C":24,"D":11,"E":11,"F":14,"G":13,"H":12,"I":13})
for k in range(NP): widths(ws, {gcl(PC + k): 12})
widths(ws, {gcl(PC+NP):13, gcl(PC+NP+1):15, gcl(CHK_C):20})
ws.column_dimensions[gcl(H1)].hidden = True
ws.column_dimensions[gcl(H2)].hidden = True

ESHT, EKEY = "仪器设备工时分配表", "V"
AA, AB = gcl(H1), gcl(H2)
EX_D = [("2026-09","SB001"),("2026-09","SB002"),("2026-09","SB003"),("2026-09","SB005")]
for r in range(DS, DE + 1):
    ws.cell(r,3, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,2,0)&"","台账中未登记"))' % (r, r))
    ws.cell(r,4, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,3,0)&"",""))' % (r, r))
    ws.cell(r,5, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,10,0)&"",""))' % (r, r))
    ws.cell(r,6, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,9,0),0))' % (r, r))
    ws.cell(r,H1,'=IF($A%d="",0,IFERROR(MATCH($A%d&"|"&$B%d,%s!$%s$%d:$%s$%d,0),0))'
            % (r, r, r, ESHT, EKEY, DS, EKEY, DE))
    ws.cell(r,H2,'=IF($A%d="",0,COUNTIFS(%s!$A$%d:$A$%d,$A%d,%s!$B$%d:$B$%d,$B%d))'
            % (r, ESHT, DS, DE, r, ESHT, DS, DE, r))
    ws.cell(r,CHK_C,'=IF($A%d="","",IF($D%d="无形资产","⚠ 无形资产请在无形资产费用分配表登记",IF($%s%d=0,"⚠ 工时分配表无此资产当月记录",IF($%s%d>1,"⚠ 工时分配表存在重复行","√ 正常"))))'
            % (r, r, AA, r, AB, r))
    ws.cell(r,7, '=IF(OR($A%d="",$%s%d=0),0,INDEX(%s!$F$%d:$F$%d,$%s%d))' % (r, AA, r, ESHT, DS, DE, AA, r))
    ws.cell(r,8, '=IF(OR($A%d="",$%s%d=0),0,INDEX(%s!$R$%d:$R$%d,$%s%d))' % (r, AA, r, ESHT, DS, DE, AA, r))
    ws.cell(r,9, '=IF(OR($A%d="",$G%d=0),"",$H%d/$G%d)' % (r, r, r, r))
    for j in range(NP):
        src = gcl(8 + j)
        ws.cell(r, PC + j, '=IF(OR($A%d="",$G%d=0),"",ROUND($F%d*INDEX(%s!$%s$%d:$%s$%d,$%s%d)/$G%d,2))'
                % (r, r, r, ESHT, src, DS, src, DE, AA, r, r))
    ws.cell(r, PC+NP,   '=IF($A%d="","",ROUND(SUM(%s%d:%s%d),2))' % (r, gcl(PC), r, gcl(PC+NP-1), r))
    ws.cell(r, PC+NP+1, '=IF($A%d="","",ROUND($F%d-%s%d,2))' % (r, r, gcl(PC+NP), r))
    idx = r - DS
    if idx < len(EX_D):
        ws.cell(r,1,EX_D[idx][0]); ws.cell(r,2,EX_D[idx][1])
    for c in range(1, CHK_C + 1):
        cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C if c != 3 else A_L
        if c in (1,2):
            cell.font = F_BODY; cell.fill = fill(C_INPUT)
        elif c == 6:
            cell.font = F_BODY; cell.fill = fill(C_NOTE)
        else:
            cell.font = F_CALC; cell.fill = fill(C_CALC)
        if c == 6 or PC <= c <= PC + NP + 1: cell.number_format = FMT_M
        if c in (7,8): cell.number_format = FMT_H
        if c == 9: cell.number_format = FMT_P
    ws.row_dimensions[r].height = 18
ws.cell(3,6).comment = Comment("本列自动取自【资产台账】的“月折旧/摊销/租赁额”。\n当月实际计提金额与台账不同时（如新增、处置、停用），可直接在单元格覆盖填写实际数。\n数据来源：月度固定资产折旧明细表。", "财务")
dv = DataValidation(type="list", formula1=MONTH_DV, allow_blank=True); ws.add_data_validation(dv); dv.add("A%d:A%d" % (DS, DE))
dv2 = DataValidation(type="list", formula1="=资产编号列表", allow_blank=True); ws.add_data_validation(dv2); dv2.add("B%d:B%d" % (DS, DE))
ws.conditional_formatting.add("%s%d:%s%d" % (gcl(CHK_C), DS, gcl(CHK_C), DE),
    FormulaRule(formula=['LEFT($%s%d,1)="⚠"' % (gcl(CHK_C), DS)],
                fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FN, size=10, bold=True, color="9C0006")))
ws.freeze_panes = "E6"; ws.auto_filter.ref = "A5:%s%d" % (gcl(CHK_C), DE)
setup_print(ws)

# ============================================================
# 9. 其他研发费用（选填）
# ============================================================
CATS = ["直接投入费用","新产品设计费等","其他相关费用","委托外部研发费用"]
ws = wb.create_sheet("其他研发费用")
widths(ws, {"A":12,"B":12,"C":40,"D":18,"E":40,"F":15,"G":16,"H":22})
title_bar(ws, 8, "十一、其他研发费用登记表（直接投入/设计费/其他相关费用/委外研发）  ——  选填",
 "填写说明：本表用于登记除人工费、折旧摊销外可直接归属到项目的研发费用（如试验用泡棉原料、PET基材、胶粘剂、燃料动力、样品模具、"
 "委外检测与委托研发等），逐笔登记并注明凭证号，自动汇总至【月度归集汇总】。费用类别请从下拉菜单选择，口径与研发费用加计扣除政策一致。")
hdr = ["年月","项目编号","研发项目名称","费用类别","费用内容摘要","金额(元)","记账凭证号","备注"]
for i, h in enumerate(hdr, 1):
    c = ws.cell(3, i, h); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C; c.border = BORDER
ws.row_dimensions[3].height = 30
EX_O = [("2026-09","RD02","直接投入费用","无卤阻燃母粒及试验用聚醚多元醇（第3批小试）",12600,"记-09-0186"),
        ("2026-09","RD03","委托外部研发费用","委托XX大学开展无溶剂贴合剥离强度机理测试",30000,"记-09-0212"),
        ("2026-09","RD04","直接投入费用","分切刀具试验损耗及试切PET复合泡棉料带",8450,"记-09-0233")]
for r in range(OS_, OE + 1):
    ws.cell(r,3, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,项目清单!$B$4:$C$13,2,0)&"","清单中未登记"))' % (r, r))
    idx = r - OS_
    if idx < len(EX_O):
        e = EX_O[idx]
        ws.cell(r,1,e[0]); ws.cell(r,2,e[1]); ws.cell(r,4,e[2]); ws.cell(r,5,e[3]); ws.cell(r,6,e[4]); ws.cell(r,7,e[5])
        ws.cell(r,8,"★示例行，请删除")
    for c in range(1, 9):
        cell = ws.cell(r, c); cell.border = BORDER
        cell.alignment = A_C if c not in (3,5,8) else A_L
        if c == 3:
            cell.font = F_CALC; cell.fill = fill(C_CALC)
        elif c == 8 and cell.value:
            cell.font = F_RED; cell.fill = fill(C_INPUT)
        else:
            cell.font = F_BODY; cell.fill = fill(C_INPUT)
        if c == 6: cell.number_format = FMT_M
    ws.row_dimensions[r].height = 18
for f1, rng in ((MONTH_DV, "A"), ("=项目编号列表", "B"), ('"' + ",".join(CATS) + '"', "D")):
    d = DataValidation(type="list", formula1=f1, allow_blank=True); ws.add_data_validation(d)
    d.add("%s%d:%s%d" % (rng, OS_, rng, OE))
ws.freeze_panes = "A4"; ws.auto_filter.ref = "A3:H%d" % OE
setup_print(ws)
# ============================================================
# 8a.《无形资产工时分配表》（研发部门按月填报）
# ============================================================
ws = wb.create_sheet("无形资产工时分配表")
PC = 8
left  = ["年月\n(YYYY-MM)","资产编号","无形资产名称","资产类别","所属工序/用途","当月总使用\n工时","非研发使用工时\n(生产及其他)"]
right = ["研发使用\n工时合计","工时校验\n(须为0)","研发使用\n占比","备注"]
title_bar(ws, PC + NP + len(right) - 1,
  "九、《无形资产工时分配表》（研发部门按月填报，%d年度）" % YEAR,
  "录入说明：①由研发部门按月填报，依据经签批的《附表3-无形资产工时分配表(签批用)》录入；②“当月总使用工时”为该无形资产（如配方模拟软件、"
  "专利技术、专有技术等）当月实际使用总工时（生产+研发），是分摊的分母；③研发专用无形资产记录全额使用工时（非研发使用工时填0），"
  "研发生产共用的按实际使用分别记录；④同一资产当月只填一行；⑤“工时校验”须为0，显示红色表示不平，须更正后方可分摊。")
last = proj_header(ws, left, "各研发项目使用工时（小时）", right, PC)
widths(ws, {"A":11,"B":11,"C":24,"D":11,"E":14,"F":13,"G":14})
for k in range(NP): widths(ws, {gcl(PC + k): 9})
widths(ws, {gcl(PC+NP):11, gcl(PC+NP+1):10, gcl(PC+NP+2):10, gcl(PC+NP+3):22})

EX_W = [("2026-09","WX001",160,0,{0:60,1:60,4:40})]
for r in range(IS_, IE + 1):
    ws.cell(r,3, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,2,0)&"","台账中未登记"))' % (r, r))
    ws.cell(r,4, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,3,0)&"",""))' % (r, r))
    ws.cell(r,5, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,5,0)&"",""))' % (r, r))
    ws.cell(r,18,'=IF($A%d="","",SUM(%s%d:%s%d))' % (r, gcl(PC), r, gcl(PC+NP-1), r))
    ws.cell(r,19,'=IF($A%d="","",ROUND($F%d-$G%d-$R%d,2))' % (r, r, r, r))
    ws.cell(r,20,'=IF(OR($A%d="",$F%d=0),"",$R%d/$F%d)' % (r, r, r, r))
    ws.cell(r,22,'=IF($A%d="","",$A%d&"|"&$B%d)' % (r, r, r))
    idx = r - IS_
    if idx < len(EX_W):
        e = EX_W[idx]
        ws.cell(r,1,e[0]); ws.cell(r,2,e[1]); ws.cell(r,6,e[2]); ws.cell(r,7,e[3])
        for k, v in e[4].items(): ws.cell(r, PC + k, v)
        ws.cell(r,21,"★示例行，请删除")
    for c in range(1, last + 1):
        cell = ws.cell(r, c); cell.border = BORDER
        cell.alignment = A_C if c not in (3,21) else A_L
        if c in (3,4,5,18,19,20):
            cell.font = F_CALC; cell.fill = fill(C_CALC)
        elif c == 21 and cell.value:
            cell.font = F_RED; cell.fill = fill(C_INPUT)
        else:
            cell.font = F_BODY; cell.fill = fill(C_INPUT)
        if c in (6,7,18,19) or PC <= c <= PC+NP-1: cell.number_format = FMT_H
        if c == 20: cell.number_format = FMT_P
    ws.row_dimensions[r].height = 18
ws.conditional_formatting.add("S%d:S%d" % (IS_, IE),
    FormulaRule(formula=['AND($A%d<>"",$S%d<>0)' % (IS_, IS_)],
                fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FN, size=10, bold=True, color="9C0006")))
dv = DataValidation(type="list", formula1=MONTH_DV, allow_blank=True); ws.add_data_validation(dv); dv.add("A%d:A%d" % (IS_, IE))
dv2 = DataValidation(type="list", formula1="=资产编号列表", allow_blank=True); ws.add_data_validation(dv2); dv2.add("B%d:B%d" % (IS_, IE))
ws.cell(5,22,"匹配主键(勿删)").font = Font(name=FN, size=8, color="808080")
ws.column_dimensions["V"].hidden = True
ws.freeze_panes = "D6"; ws.auto_filter.ref = "A5:%s%d" % (gcl(last), IE)
setup_print(ws)

# ============================================================
# 8b.《无形资产费用分配表》（财务按月按项目核算、按年汇总）
# ============================================================
ws = wb.create_sheet("无形资产费用分配表")
PC = 10
left = ["年月\n(YYYY-MM)","资产编号","无形资产名称","资产类别","费用类型","当月摊销额\n(元)",
        "当月总使用\n工时","研发使用\n工时合计","研发使用占比\n(分摊系数)"]
right = ["研发摊销费用\n小计(元)","计入生产成本的\n摊销额(元)"]
title_bar(ws, PC + NP + 2,
 "十、《无形资产费用分配表》（财务按月按项目核算、按年汇总，%d年度）" % YEAR,
 "计算依据：当月某项目应分摊的无形资产摊销额 = 该无形资产当月摊销额 ×（该项目使用工时 ÷ 该资产当月总使用工时）。"
 "数据来源：摊销数据源自【月度无形资产摊销表】。录入说明：只需录入 A、B 两列（年月、资产编号）；F列“当月摊销额”自动取自【资产台账】，"
 "当月实际摊销数不同时可直接覆盖填写；工时及分摊金额由《无形资产工时分配表》按“年月+资产编号”自动匹配取数。")
last = proj_header(ws, left, "各研发项目分摊无形资产摊销额（元）", right, PC)
CHK_C = last + 1
H1, H2 = last + 2, last + 3
ws.merge_cells(start_row=3, start_column=CHK_C, end_row=5, end_column=CHK_C)
c = ws.cell(3, CHK_C, "数据校验"); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C
ws.cell(5, H1, "匹配行号(勿删)").font = Font(name=FN, size=8, color="808080")
ws.cell(5, H2, "重复计数(勿删)").font = Font(name=FN, size=8, color="808080")
widths(ws, {"A":11,"B":11,"C":24,"D":11,"E":11,"F":13,"G":13,"H":12,"I":13})
for k in range(NP): widths(ws, {gcl(PC + k): 12})
widths(ws, {gcl(PC+NP):14, gcl(PC+NP+1):15, gcl(CHK_C):22})
ws.column_dimensions[gcl(H1)].hidden = True
ws.column_dimensions[gcl(H2)].hidden = True

WSHT2, WKEY = "无形资产工时分配表", "V"
AA, AB = gcl(H1), gcl(H2)
EX_I = [("2026-09","WX001")]
for r in range(IS_, IE + 1):
    ws.cell(r,3, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,2,0)&"","台账中未登记"))' % (r, r))
    ws.cell(r,4, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,3,0)&"",""))' % (r, r))
    ws.cell(r,5, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,10,0)&"",""))' % (r, r))
    ws.cell(r,6, '=IF($B%d="","",IFERROR(VLOOKUP($B%d,资产台账!$B$4:$N$53,9,0),0))' % (r, r))
    ws.cell(r,H1,'=IF($A%d="",0,IFERROR(MATCH($A%d&"|"&$B%d,%s!$%s$%d:$%s$%d,0),0))'
            % (r, r, r, WSHT2, WKEY, IS_, WKEY, IE))
    ws.cell(r,H2,'=IF($A%d="",0,COUNTIFS(%s!$A$%d:$A$%d,$A%d,%s!$B$%d:$B$%d,$B%d))'
            % (r, WSHT2, IS_, IE, r, WSHT2, IS_, IE, r))
    ws.cell(r,CHK_C,'=IF($A%d="","",IF($D%d="设备","⚠ 设备请在仪器设备工时费用分配表登记",IF($%s%d=0,"⚠ 工时分配表无此资产当月记录",IF($%s%d>1,"⚠ 工时分配表存在重复行","√ 正常"))))'
            % (r, r, AA, r, AB, r))
    ws.cell(r,7, '=IF(OR($A%d="",$%s%d=0),0,INDEX(%s!$F$%d:$F$%d,$%s%d))' % (r, AA, r, WSHT2, IS_, IE, AA, r))
    ws.cell(r,8, '=IF(OR($A%d="",$%s%d=0),0,INDEX(%s!$R$%d:$R$%d,$%s%d))' % (r, AA, r, WSHT2, IS_, IE, AA, r))
    ws.cell(r,9, '=IF(OR($A%d="",$G%d=0),"",$H%d/$G%d)' % (r, r, r, r))
    for j in range(NP):
        src = gcl(8 + j)
        ws.cell(r, PC + j, '=IF(OR($A%d="",$G%d=0),"",ROUND($F%d*INDEX(%s!$%s$%d:$%s$%d,$%s%d)/$G%d,2))'
                % (r, r, r, WSHT2, src, IS_, src, IE, AA, r, r))
    ws.cell(r, PC+NP,   '=IF($A%d="","",ROUND(SUM(%s%d:%s%d),2))' % (r, gcl(PC), r, gcl(PC+NP-1), r))
    ws.cell(r, PC+NP+1, '=IF($A%d="","",ROUND($F%d-%s%d,2))' % (r, r, gcl(PC+NP), r))
    idx = r - IS_
    if idx < len(EX_I):
        ws.cell(r,1,EX_I[idx][0]); ws.cell(r,2,EX_I[idx][1])
    for c in range(1, CHK_C + 1):
        cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C if c != 3 else A_L
        if c in (1,2):
            cell.font = F_BODY; cell.fill = fill(C_INPUT)
        elif c == 6:
            cell.font = F_BODY; cell.fill = fill(C_NOTE)
        else:
            cell.font = F_CALC; cell.fill = fill(C_CALC)
        if c == 6 or PC <= c <= PC + NP + 1: cell.number_format = FMT_M
        if c in (7,8): cell.number_format = FMT_H
        if c == 9: cell.number_format = FMT_P
    ws.row_dimensions[r].height = 18
ws.cell(3,6).comment = Comment("自动取自【资产台账】的“月折旧/摊销/租赁额”。\n当月实际摊销额不同时可直接覆盖填写。\n数据来源：月度无形资产摊销表。", "财务")
dv = DataValidation(type="list", formula1=MONTH_DV, allow_blank=True); ws.add_data_validation(dv); dv.add("A%d:A%d" % (IS_, IE))
dv2 = DataValidation(type="list", formula1="=资产编号列表", allow_blank=True); ws.add_data_validation(dv2); dv2.add("B%d:B%d" % (IS_, IE))
ws.conditional_formatting.add("%s%d:%s%d" % (gcl(CHK_C), IS_, gcl(CHK_C), IE),
    FormulaRule(formula=['LEFT($%s%d,1)="⚠"' % (gcl(CHK_C), IS_)],
                fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FN, size=10, bold=True, color="9C0006")))
ws.freeze_panes = "E6"; ws.auto_filter.ref = "A5:%s%d" % (gcl(CHK_C), IE)
setup_print(ws)
# ============================================================
# 10. 月度归集汇总
# ============================================================
ws = wb.create_sheet("月度归集汇总")
widths(ws, {"A":11,"B":11,"C":40,"D":15,"E":14,"F":15,"G":14,"H":15,"I":14,"J":16,"K":16,"L":14,"M":14})
title_bar(ws, 13, "十二、研发费用月度归集汇总表（%d年度，按项目·按月）" % YEAR,
 "本表全部为公式自动汇总，请勿手工修改。费用类别口径与《研发费用加计扣除》一致，可直接与研发支出辅助账、"
 "“研发支出”科目明细账逐月核对。末两列同步列示当月归集的研发人员工时与设备使用工时，作为费用分配的量化依据。")
hdr = ["年月","项目编号","研发项目名称","人员人工费用","折旧费用","无形资产摊销","直接投入费用",
       "新产品设计费等","其他相关费用","委托外部研发费用","本月研发费用合计","研发人员工时\n(小时)","设备使用工时\n(小时)"]
for i, h in enumerate(hdr, 1):
    c = ws.cell(3, i, h); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C; c.border = BORDER
ws.row_dimensions[3].height = 34

LAB_S, LAB_E = DS, DE                  # 各明细表数据区
r = 4
sub_rows = []
for m in MONTHS:
    first = r
    for j in range(NP):
        pcol_lab  = gcl(14 + j)   # 研发人员工时费用分配表 N..W
        pcol_dep  = gcl(10 + j)   # 仪器设备工时费用分配表 J..S
        pcol_hr   = gcl(8 + j)    # 工时月报 H..Q
        ws.cell(r,1,m).font = F_BODY
        ws.cell(r,2,"=项目清单!$B$%d" % (4 + j))
        ws.cell(r,3,'=IF($B%d="","",IFERROR(VLOOKUP($B%d,项目清单!$B$4:$C$13,2,0)&"",""))' % (r, r))
        ws.cell(r,4,'=SUMIFS(研发人员工时费用分配表!$%s$%d:$%s$%d,研发人员工时费用分配表!$A$%d:$A$%d,$A%d)'
                % (pcol_lab, LAB_S, pcol_lab, LAB_E, LAB_S, LAB_E, r))
        ws.cell(r,5,'=SUMIFS(仪器设备工时费用分配表!$%s$%d:$%s$%d,仪器设备工时费用分配表!$A$%d:$A$%d,$A%d,仪器设备工时费用分配表!$D$%d:$D$%d,"设备")'
                % (pcol_dep, LAB_S, pcol_dep, LAB_E, LAB_S, LAB_E, r, LAB_S, LAB_E))
        ws.cell(r,6,'=SUMIFS(无形资产费用分配表!$%s$%d:$%s$%d,无形资产费用分配表!$A$%d:$A$%d,$A%d)'
                % (pcol_dep, IS_, pcol_dep, IE, IS_, IE, r))
        for k2, cat in enumerate(CATS):
            col = {0: 7, 1: 8, 2: 9, 3: 10}[k2]
            ws.cell(r, col, '=SUMIFS(其他研发费用!$F$%d:$F$%d,其他研发费用!$A$%d:$A$%d,$A%d,其他研发费用!$B$%d:$B$%d,$B%d,其他研发费用!$D$%d:$D$%d,"%s")'
                    % (OS_, OE, OS_, OE, r, OS_, OE, r, OS_, OE, cat))
        ws.cell(r,11,'=ROUND(SUM(D%d:J%d),2)' % (r, r))
        ws.cell(r,12,'=SUMIFS(研发人员工时分配表!$%s$%d:$%s$%d,研发人员工时分配表!$A$%d:$A$%d,$A%d)' % (pcol_hr, LAB_S, pcol_hr, LAB_E, LAB_S, LAB_E, r))
        ws.cell(r,13,'=SUMIFS(仪器设备工时分配表!$%s$%d:$%s$%d,仪器设备工时分配表!$A$%d:$A$%d,$A%d)' % (pcol_hr, LAB_S, pcol_hr, LAB_E, LAB_S, LAB_E, r))
        for c in range(1, 14):
            cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C if c != 3 else A_L
            cell.font = F_CALC if c != 1 else F_BODY
            cell.fill = fill(C_CALC)
            if 4 <= c <= 11: cell.number_format = FMT_M
            if c >= 12: cell.number_format = FMT_H
        ws.row_dimensions[r].height = 17
        r += 1
    # 月小计
    ws.cell(r,1,m); ws.cell(r,3,"本月合计")
    for c in range(4, 14):
        ws.cell(r, c, "=ROUND(SUM(%s%d:%s%d),2)" % (gcl(c), first, gcl(c), r - 1))
    for c in range(1, 14):
        cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C
        cell.font = F_BOLD; cell.fill = fill(C_TOTAL)
        if 4 <= c <= 11: cell.number_format = FMT_M
        if c >= 12: cell.number_format = FMT_H
    ws.row_dimensions[r].height = 18
    sub_rows.append(r)
    r += 1
GT = r
ws.cell(GT,3,"%d年度累计合计" % YEAR)
for c in range(4, 14):
    ws.cell(GT, c, "=ROUND(%s,2)" % "+".join("%s%d" % (gcl(c), x) for x in sub_rows))
for c in range(1, 14):
    cell = ws.cell(GT, c); cell.border = BORDER; cell.alignment = A_C
    cell.font = Font(name=FN, size=11, bold=True, color="FFFFFF"); cell.fill = fill(C_TITLE)
    if 4 <= c <= 11: cell.number_format = FMT_M
    if c >= 12: cell.number_format = FMT_H
ws.row_dimensions[GT].height = 22
MSUM_S, MSUM_E = 4, GT - 1
ws.freeze_panes = "D4"; ws.auto_filter.ref = "A3:M%d" % MSUM_E
setup_print(ws)

# ============================================================
# 11. 年度汇总
# ============================================================
ws = wb.create_sheet("年度汇总")
widths(ws, {"A":11,"B":38})
for i in range(3, 16): widths(ws, {gcl(i): 13})
widths(ws, {"O": 15})
title_bar(ws, 15, "十三、研发费用年度汇总表（%d年度）" % YEAR,
 "本表全部为公式自动汇总。表一按项目分月列示研发费用，表二按加计扣除费用类别列示，表三列示全年工时与费用强度，表四为勾稽校验（差异须为0）。")

def sec_title(row, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=15)
    c = ws.cell(row, 1, text); c.font = F_H2; c.fill = fill(C_SUB); c.alignment = A_L
    ws.row_dimensions[row].height = 22

def hdr_row(row, labels):
    for i, h in enumerate(labels, 1):
        c = ws.cell(row, i, h); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C; c.border = BORDER
    ws.row_dimensions[row].height = 26

# --- 表一 ---
sec_title(4, "表一  各研发项目分月度研发费用归集表（元）")
hdr_row(5, ["项目编号","研发项目名称"] + [m[-2:] + "月" for m in MONTHS] + ["全年合计"])
for j in range(NP):
    r = 6 + j
    ws.cell(r,1,"=项目清单!$B$%d" % (4 + j))
    ws.cell(r,2,'=IF($A%d="","",IFERROR(VLOOKUP($A%d,项目清单!$B$4:$C$13,2,0)&"",""))' % (r, r))
    for mi, m in enumerate(MONTHS):
        ws.cell(r, 3 + mi, '=SUMIFS(月度归集汇总!$K$%d:$K$%d,月度归集汇总!$A$%d:$A$%d,"%s",月度归集汇总!$B$%d:$B$%d,$A%d)'
                % (MSUM_S, MSUM_E, MSUM_S, MSUM_E, m, MSUM_S, MSUM_E, r))
    ws.cell(r,15,"=ROUND(SUM(C%d:N%d),2)" % (r, r))
    for c in range(1, 16):
        cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C if c != 2 else A_L
        cell.font = F_CALC; cell.fill = fill(C_CALC)
        if c >= 3: cell.number_format = FMT_M
    ws.row_dimensions[r].height = 18
r = 16
ws.cell(r,2,"合  计")
for c in range(3, 16): ws.cell(r, c, "=ROUND(SUM(%s6:%s15),2)" % (gcl(c), gcl(c)))
for c in range(1, 16):
    cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C
    cell.font = F_BOLD; cell.fill = fill(C_TOTAL)
    if c >= 3: cell.number_format = FMT_M
ws.row_dimensions[r].height = 20

# --- 表二 ---
sec_title(18, "表二  各研发项目全年研发费用构成表（元）")
hdr_row(19, ["项目编号","研发项目名称","人员人工费用","折旧费用","无形资产摊销","直接投入费用",
             "新产品设计费等","其他相关费用","委托外部研发费用","全年合计","占比"])
MAP = {3:"D", 4:"E", 5:"F", 6:"G", 7:"H", 8:"I", 9:"J"}
for j in range(NP):
    r = 20 + j
    ws.cell(r,1,"=项目清单!$B$%d" % (4 + j))
    ws.cell(r,2,'=IF($A%d="","",IFERROR(VLOOKUP($A%d,项目清单!$B$4:$C$13,2,0)&"",""))' % (r, r))
    for c, src in MAP.items():
        ws.cell(r, c, '=SUMIFS(月度归集汇总!$%s$%d:$%s$%d,月度归集汇总!$B$%d:$B$%d,$A%d)'
                % (src, MSUM_S, src, MSUM_E, MSUM_S, MSUM_E, r))
    ws.cell(r,10,"=ROUND(SUM(C%d:I%d),2)" % (r, r))
    ws.cell(r,11,"=IFERROR($J%d/$J$30,0)" % r)
    for c in range(1, 12):
        cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C if c != 2 else A_L
        cell.font = F_CALC; cell.fill = fill(C_CALC)
        if 3 <= c <= 10: cell.number_format = FMT_M
        if c == 11: cell.number_format = FMT_P
    ws.row_dimensions[r].height = 18
r = 30
ws.cell(r,2,"合  计")
for c in range(3, 11): ws.cell(r, c, "=ROUND(SUM(%s20:%s29),2)" % (gcl(c), gcl(c)))
ws.cell(r,11,"=IFERROR($J$30/$J$30,0)")
for c in range(1, 12):
    cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C
    cell.font = F_BOLD; cell.fill = fill(C_TOTAL)
    if 3 <= c <= 10: cell.number_format = FMT_M
    if c == 11: cell.number_format = FMT_P
ws.row_dimensions[r].height = 20

# --- 表三 ---
sec_title(32, "表三  各研发项目全年工时统计与费用强度")
hdr_row(33, ["项目编号","研发项目名称","全年研发人员工时\n(小时)","全年仪器设备\n使用工时(小时)",
             "全年无形资产\n使用工时(小时)","全年研发费用\n合计(元)","单位研发人员工时\n费用(元/小时)"])
ws.row_dimensions[33].height = 34
for j in range(NP):
    r = 34 + j
    hcol = gcl(8 + j)
    ws.cell(r,1,"=项目清单!$B$%d" % (4 + j))
    ws.cell(r,2,'=IF($A%d="","",IFERROR(VLOOKUP($A%d,项目清单!$B$4:$C$13,2,0)&"",""))' % (r, r))
    ws.cell(r,3,"=ROUND(SUM(研发人员工时分配表!$%s$%d:$%s$%d),2)" % (hcol, LAB_S, hcol, LAB_E))
    ws.cell(r,4,"=ROUND(SUM(仪器设备工时分配表!$%s$%d:$%s$%d),2)" % (hcol, LAB_S, hcol, LAB_E))
    ws.cell(r,5,"=ROUND(SUM(无形资产工时分配表!$%s$%d:$%s$%d),2)" % (hcol, IS_, hcol, IE))
    ws.cell(r,6,"=J%d" % (20 + j))
    ws.cell(r,7,"=IFERROR(ROUND($F%d/$C%d,2),0)" % (r, r))
    for c in range(1, 8):
        cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C if c != 2 else A_L
        cell.font = F_CALC; cell.fill = fill(C_CALC)
        if c in (3,4,5): cell.number_format = FMT_H
        if c in (6,7): cell.number_format = FMT_M
    ws.row_dimensions[r].height = 18
r = 44
ws.cell(r,2,"合  计")
for c in range(3, 7): ws.cell(r, c, "=ROUND(SUM(%s34:%s43),2)" % (gcl(c), gcl(c)))
ws.cell(r,7,"=IFERROR(ROUND($F44/$C44,2),0)")
for c in range(1, 8):
    cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C
    cell.font = F_BOLD; cell.fill = fill(C_TOTAL)
    if c in (3,4,5): cell.number_format = FMT_H
    if c in (6,7): cell.number_format = FMT_M
ws.row_dimensions[r].height = 20

# --- 表四 勾稽校验 ---
sec_title(46, "表四  勾稽校验（“差异”列须全部为 0，否则说明明细表存在漏录或月份/编号填写错误）")
hdr_row(47, ["校验项目","校验说明","明细表金额(元)","汇总表金额(元)","差异(元)"])
CHK = [
 ("人员人工费用", "《研发人员工时费用分配表》研发人工费小计合计  ⇔  年度汇总表二“人员人工费用”合计",
  "=ROUND(SUM(研发人员工时费用分配表!$X$%d:$X$%d),2)" % (LAB_S, LAB_E), "=C30"),
 ("仪器设备折旧费", "《仪器、设备工时费用分配表》研发费用小计合计  ⇔  年度汇总表二“折旧费用”合计",
  "=ROUND(SUM(仪器设备工时费用分配表!$T$%d:$T$%d),2)" % (LAB_S, LAB_E), "=D30"),
 ("无形资产摊销", "《无形资产费用分配表》分摊合计  ⇔  年度汇总表二“无形资产摊销”合计",
  "=ROUND(SUM(无形资产费用分配表!$T$%d:$T$%d),2)" % (IS_, IE), "=E30"),
 ("其他研发费用", "《其他研发费用登记表》金额合计  ⇔  年度汇总表二直接投入+设计费+其他相关+委外",
  "=ROUND(SUM(其他研发费用!$F$%d:$F$%d),2)" % (OS_, OE), "=ROUND(F30+G30+H30+I30,2)"),
 ("研发费用总额", "【月度归集汇总】年度累计合计  ⇔  年度汇总表一合计",
  "=月度归集汇总!$K$%d" % GT, "=O16"),
]
for i, (a, b, c1, c2) in enumerate(CHK):
    r = 48 + i
    ws.cell(r,1,a); ws.cell(r,2,b); ws.cell(r,3,c1); ws.cell(r,4,c2)
    ws.cell(r,5,"=ROUND($C%d-$D%d,2)" % (r, r))
    for c in range(1, 6):
        cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C if c != 2 else A_L
        cell.font = F_BODY if c <= 2 else F_CALC
        cell.fill = fill(C_CALC)
        if c >= 3: cell.number_format = FMT_M
    ws.row_dimensions[r].height = 20
ws.merge_cells(start_row=48, start_column=2, end_row=48, end_column=2)
ws.column_dimensions["B"].width = 58
ws.conditional_formatting.add("E48:E52",
    FormulaRule(formula=["$E48<>0"], fill=PatternFill("solid", fgColor="FFC7CE"),
                font=Font(name=FN, size=10, bold=True, color="9C0006")))
setup_print(ws)
# ============================================================
# 11b. 指标监测（科技人员占比 / 研发费用占比 / 高新收入占比）
# ============================================================
ws = wb.create_sheet("指标监测")
widths(ws, {"A":16,"B":16,"C":18,"D":18,"E":16,"F":15,"G":17,"H":17,"I":17,"J":24})
title_bar(ws, 10, "一、高新技术企业各项指标监测表（%d年度）" % YEAR,
 "监测口径：①科技人员占当年职工总数比例不低于10%；②研发费用占销售收入比例——销售收入＜5000万元的不低于5%，5000万元～2亿元的不低于4%，"
 "超过2亿元的不低于3%（本表按销售收入自动判定适用档次）；③高新技术产品（服务）收入占同期总收入比例不低于60%。"
 "浅黄底单元格需人工录入（数据取自财务报表及人事台账），研发费用自动取自【年度汇总】。")

def sec(row, text, lastc=10):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=lastc)
    c = ws.cell(row, 1, text); c.font = F_H2; c.fill = fill(C_SUB); c.alignment = A_L
    ws.row_dimensions[row].height = 22

sec(4, "表一  年度指标达标测算")
for i, h in enumerate(["指标项目", "", "数值", "标准要求", "是否达标", "说明"], 1):
    pass
ws.merge_cells("A5:B5"); ws.merge_cells("F5:J5")
for col, h in ((1,"指标项目"), (3,"数值"), (4,"标准要求"), (5,"是否达标"), (6,"说明")):
    c = ws.cell(5, col, h); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C
for col in (2,7,8,9,10):
    ws.cell(5, col).fill = fill(C_GRP)
box(ws, 5, 1, 5, 10); ws.row_dimensions[5].height = 26

T1 = [
 ("当年职工总数（人）",              None, "", "", "在职职工月平均人数，与社保缴纳人数、工资表人数一致", "in", '0'),
 ("其中：科技人员数（人）",          None, '=COUNTIF(人员名册!$H$4:$H$53,"是")', "",
  "左侧“标准要求”列为【人员名册】中标记“是否科技人员=是”的统计参考值，实际以人事台账为准", "in", '0'),
 ("科技人员占比",                    '=IF($C$6=0,"",$C$7/$C$6)', "不低于10%",
  '=IF($C$8="","",IF($C$8>=0.1,"√ 达标","✗ 未达标"))', "累计实际工作满183天以上的研发人员及相关技术人员", "pct", None),
 ("当年销售收入（元）",              None, "", "", "取自利润表“营业收入”中的销售（营业）收入，用于研发费用占比测算", "in", FMT_M),
 ("当年研发费用总额（元）",          '=年度汇总!$O$16', "自动取数", "",
  "自动取自【年度汇总】表一合计，须与研发支出辅助账、纳税申报一致", "cal", FMT_M),
 ("适用研发费用比例标准",            '=IF($C$9=0,"",IF($C$9<50000000,0.05,IF($C$9<=200000000,0.04,0.03)))',
  "按销售收入自动分档", "", "＜5000万→5%；5000万～2亿→4%；＞2亿→3%", "pct", None),
 ("研发费用占销售收入比例",          '=IF($C$9=0,"",$C$10/$C$9)', '=IF($C$11="","",$C$11)',
  '=IF(OR($C$12="",$C$11=""),"",IF($C$12>=$C$11,"√ 达标","✗ 未达标"))',
  "近三个会计年度分别测算；企业成立不足三年的按实际经营年限计算", "pct", None),
 ("当年总收入（元）",                None, "", "", "收入总额减去不征税收入后的余额", "in", FMT_M),
 ("高新技术产品（服务）收入（元）",   None, "", "", "对应主要产品（服务）：3C电子用泡棉/复合材料、新能源汽车电池用缓冲阻燃材料等", "in", FMT_M),
 ("高新收入占比",                    '=IF($C$13=0,"",$C$14/$C$13)', "不低于60%",
  '=IF($C$15="","",IF($C$15>=0.6,"√ 达标","✗ 未达标"))', "高新技术产品（服务）收入占同期总收入比例", "pct", None),
]
for i, (lab, val, std, ok, memo, kind, nf) in enumerate(T1):
    r = 6 + i
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=10)
    ws.cell(r,1,lab).alignment = A_L
    if val is not None: ws.cell(r,3,val)
    if std != "": ws.cell(r,4,std)
    if ok != "": ws.cell(r,5,ok)
    ws.cell(r,6,memo).alignment = A_L
    for c in range(1, 11):
        cell = ws.cell(r, c); cell.border = BORDER
        cell.alignment = A_L if c in (1,6) else A_C
        cell.font = F_BODY
        cell.fill = fill(C_INPUT if (c == 3 and kind == "in") else C_CALC)
    if kind == "pct": ws.cell(r,3).number_format = FMT_P
    elif nf: ws.cell(r,3).number_format = nf
    if r == 12: ws.cell(r,4).number_format = FMT_P
    ws.cell(r,5).font = F_BOLD
    ws.row_dimensions[r].height = 22
ws.conditional_formatting.add("E6:E15",
    FormulaRule(formula=['LEFT($E6,1)="✗"'], fill=PatternFill("solid", fgColor="FFC7CE"),
                font=Font(name=FN, size=10, bold=True, color="9C0006")))
ws.conditional_formatting.add("E6:E15",
    FormulaRule(formula=['LEFT($E6,1)="√"'], fill=PatternFill("solid", fgColor="C6EFCE"),
                font=Font(name=FN, size=10, bold=True, color="006100")))

sec(17, "表二  分月监测（收入与研发费用累计进度）")
h2 = ["年月","当月销售收入\n(元)","当月高新技术产品\n(服务)收入(元)","当月总收入\n(元)","当月研发费用\n(元,自动)",
      "累计销售收入\n(元)","累计研发费用\n(元)","累计研发费用\n占销售收入比例","累计高新收入\n占总收入比例","预警提示"]
for i, h in enumerate(h2, 1):
    c = ws.cell(18, i, h); c.font = F_GRP; c.fill = fill(C_GRP); c.alignment = A_C; c.border = BORDER
ws.row_dimensions[18].height = 38
for i, m in enumerate(MONTHS):
    r = 19 + i
    ws.cell(r,1,m)
    ws.cell(r,5,"=月度归集汇总!$K$%d" % sub_rows[i])
    ws.cell(r,6,"=SUM($B$19:B%d)" % r)
    ws.cell(r,7,"=SUM($E$19:E%d)" % r)
    ws.cell(r,8,'=IF($F%d=0,"",$G%d/$F%d)' % (r, r, r))
    ws.cell(r,9,'=IF(SUM($D$19:D%d)=0,"",SUM($C$19:C%d)/SUM($D$19:D%d))' % (r, r, r))
    ws.cell(r,10,'=IF(OR($H%d="",$C$11=""),"",IF($H%d<$C$11,"⚠ 研发费用占比低于标准","√ 正常"))' % (r, r))
    for c in range(1, 11):
        cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C
        cell.font = F_BODY if c in (1,2,3,4) else F_CALC
        cell.fill = fill(C_INPUT if c in (2,3,4) else C_CALC)
        if c in (2,3,4,5,6,7): cell.number_format = FMT_M
        if c in (8,9): cell.number_format = FMT_P
    ws.row_dimensions[r].height = 19
r = 31
ws.cell(r,1,"全年合计")
for c in (2,3,4,5): ws.cell(r,c,"=ROUND(SUM(%s19:%s30),2)" % (gcl(c), gcl(c)))
ws.cell(r,8,'=IF($B%d=0,"",$E%d/$B%d)' % (r, r, r))
ws.cell(r,9,'=IF($D%d=0,"",$C%d/$D%d)' % (r, r, r))
for c in range(1, 11):
    cell = ws.cell(r, c); cell.border = BORDER; cell.alignment = A_C
    cell.font = F_BOLD; cell.fill = fill(C_TOTAL)
    if c in (2,3,4,5): cell.number_format = FMT_M
    if c in (8,9): cell.number_format = FMT_P
ws.row_dimensions[r].height = 22
ws.merge_cells(start_row=33, start_column=1, end_row=33, end_column=10)
c = ws.cell(33,1,"提示：表二“当月研发费用”自动取自【月度归集汇总】各月合计；销售收入、高新技术产品收入、总收入需按月从财务系统取数填列，"
                 "便于全年动态监测研发费用投入强度，避免年末突击。")
c.font = F_NOTE; c.alignment = A_L; c.fill = fill(C_NOTE)
ws.row_dimensions[33].height = 32
ws.conditional_formatting.add("J19:J30",
    FormulaRule(formula=['LEFT($J19,1)="⚠"'], fill=PatternFill("solid", fgColor="FFC7CE"),
                font=Font(name=FN, size=10, bold=True, color="9C0006")))
ws.freeze_panes = "A6"
setup_print(ws)
# ============================================================
# 12. 费用分配说明（打印稿）
# ============================================================
ws = wb.create_sheet("费用分配说明")
widths(ws, {"A":3,"B":13,"C":16,"D":16,"E":16,"F":16,"G":16,"H":16,"I":10})
ws.merge_cells("A1:I1")
c = ws.cell(1,1,"研发人员及仪器设备、无形资产费用分配说明")
c.font = Font(name=FN, size=16, bold=True); c.alignment = A_C
ws.row_dimensions[1].height = 40
ws.merge_cells("A2:I2")
c = ws.cell(2,1,"（%d年度  ·  按月核算、按年汇总）" % YEAR)
c.font = Font(name=FN, size=11); c.alignment = A_C
ws.row_dimensions[2].height = 20

DOC = [
("H","一、企业基本情况与研发活动概况"),
("P","本公司主营泡棉（发泡材料）生产、泡棉与PET基材贴合（复合）以及分切/模切加工，产品主要应用于3C电子与新能源汽车两大板块。"
     "主要生产工序为：原料配制 → 发泡成型 → 熟化 → PET贴合（复合） → 分切/模切 → 检测包装。"),
("P","公司研发活动主要围绕新型泡棉配方开发、无溶剂/水性胶贴合工艺改进、超薄泡棉高精度分切与毛边控制、"
     "动力电池用阻燃缓冲及导热复合材料开发等方向展开。受工艺特点限制，上述研发活动的小试、中试与验证试验"
     "必须在现有发泡线、贴合线、分切线等生产设备上进行，由具备实际操作经验的技术与生产人员共同完成。"),
("H","二、研发人员与仪器设备生产研发混合使用情况说明"),
("P","1. 人员方面：公司研发人员分为“研发专用”与“研发生产共用”两类。研发中心专职配方工程师、工艺工程师、检测工程师属研发专用人员，"
     "其当月全额工时计入研发项目；发泡、贴合、分切等岗位的部分生产技师在承担正常生产任务的同时参与研发试验，以及外聘技术顾问，"
     "属研发生产共用人员，仅按实际投入研发的工时计入研发项目，其余工时计入生产。"),
("P","2. 设备方面：同样区分“研发专用”与“研发生产共用”。研发专用设备记录全额运行工时；连续发泡生产线、自动贴合复合机、"
     "高精度分切机、模切设备等主要生产设备，在安排试制、试验期间用于研发活动、其余时间用于正常生产，属研发生产共用设备，"
     "按其实际研发使用工时占当月总运行工时的比例分摊折旧（租赁）费。"),
("P","3. 无形资产方面：泡棉配方模拟软件、专利/专有技术等无形资产，同时服务于研发试验与生产质量控制，"
     "亦区分“研发专用”与“研发生产共用”，由研发部门按月据实记录各研发项目的使用工时，按使用工时占比分摊其当月摊销额。"),
("P","4. 鉴于人员与设备存在研发生产混合使用情形，公司采用“实际工时法”，以经审批的工时记录作为唯一分配依据，"
     "将人员人工费用、设备折旧（租赁）费用及无形资产摊销，在各研发项目与生产成本之间进行合理分配，"
     "做到有据可依、可追溯、可复核。"),
("H","三、工作使用情况记录制度（原始记录的形成与审批）"),
("P","1. 记录主体与频次：研发人员按月填报《研发人员工时分配表》（签批用见本台账附表1）；设备管理员/工序班长按月填报"
     "《仪器、设备工时分配表》（签批用见本台账附表2）；无形资产管理员按月填报《无形资产工时分配表》（签批用见本台账附表3）。"
     "填报截止时间为次月3日前。"),
("P","2. 记录内容：人员填报当月出勤总工时、投入各研发项目的工时、非研发（生产及其他）工时，并简述当月主要研发工作内容；"
     "设备与无形资产填报当月总运行（使用）工时、各研发项目占用工时及生产占用工时。"),
("P","3. 原始依据：人员工时以考勤记录、项目周报、试验记录为支撑；设备工时以设备运行台账、点检记录、生产报工单、"
     "试制（试验）通知单为支撑，确保填报数据可与生产系统记录相互印证。"),
("P","4. 审批流程：填报人签字 → 项目负责人确认 → 研发部门负责人审核 → 财务复核入账。签批原件按月装订、专人保管。"),
("P","5. 复核要求：财务复核时必须核对“当月总工时 = 非研发工时 + 各研发项目工时合计”，本台账已设置“工时校验”列自动校验，"
     "不平（显示红色）者退回更正后方可进行费用分摊。研发专用人员/设备的非研发工时应为0，即按全额工时记录。"),
("H","四、费用分配计算方法"),
("F","（一）人员人工费用"),
("Q","当月某项目应分摊的人工费 ＝ 该人员当月总薪酬（含工资、社保、公积金、外聘劳务费） ×（该项目研发工时 ÷ 该人员当月总工时）"),
("F","（二）仪器设备折旧费/租赁使用费"),
("Q","当月某项目应分摊的设备费 ＝ 该设备当月折旧额（或租赁费） ×（该项目使用工时 ÷ 该设备当月总运行工时）"),
("F","（三）无形资产摊销费用"),
("Q","当月某项目应分摊的摊销额 ＝ 该无形资产当月摊销额 ×（该项目使用工时 ÷ 该无形资产当月总使用工时）"),
("P","无形资产（配方模拟软件、专利技术、专有技术等）由研发部门按月填报《无形资产工时分配表》，据实记录各研发项目的使用工时；"
     "研发专用无形资产记录全额使用工时，研发生产共用的按实际使用工时记录，未分配部分计入生产成本。"),
("P","（四）上述计算后未分配至研发项目的剩余部分，全额计入当期生产成本或相应期间费用，确保研发与生产各归其位、不重不漏。"),
("P","（五）分摊金额按月计算，四舍五入保留两位小数；年度金额为当年各月分摊额之和。分配方法在年度内保持一贯，"
     "确需变更的须经公司管理层审批并书面说明理由。"),
("H","五、计算举例（以%d年9月为例，数据详见本台账各明细表）" % YEAR),
("P","1. 人员人工费用：配方工程师张工（工号YF001）当月出勤总工时176小时，其中RD01项目100小时、RD02项目60小时、"
     "非研发（生产及其他）16小时；当月总薪酬 15,000＋3,900＋1,800＋1,200 ＝ 21,900元。"),
("Q","RD01应分摊 ＝ 21,900 × 100 ÷ 176 ＝ 12,443.18元；  RD02应分摊 ＝ 21,900 × 60 ÷ 176 ＝ 7,465.91元；"
     "  计入生产成本 ＝ 21,900 － 12,443.18 － 7,465.91 ＝ 1,990.91元。"),
("P","2. 仪器设备折旧费：连续发泡生产线（资产编号SB001）当月总运行工时320小时，其中RD01占用40小时、RD02占用20小时，"
     "其余260小时用于正常生产；该设备当月计提折旧15,000元。"),
("Q","RD01应分摊 ＝ 15,000 × 40 ÷ 320 ＝ 1,875.00元；  RD02应分摊 ＝ 15,000 × 20 ÷ 320 ＝ 937.50元；"
     "  计入生产成本 ＝ 15,000 － 1,875.00 － 937.50 ＝ 12,187.50元。"),
("H","六、核算与归集流程（按月核算、按年汇总）"),
("P","每月末研发人员与设备/无形资产管理员完成《研发人员工时分配表》《仪器、设备工时分配表》《无形资产工时分配表》填报并逐级签批 → 财务于次月5日前据以编制"
     "《研发人员工时费用分配表》《仪器、设备工时费用分配表》《无形资产费用分配表》，按月按项目核算，"
     "登记研发支出辅助账并出具《研发费用月度归集汇总表》 → 年度终了按年汇总生成《研发费用年度汇总表》与《指标监测表》，"
     "并与“研发支出”科目明细账、企业所得税纳税申报及高新技术企业专项审计数据核对一致。"),
("H","七、资料留存"),
("P","本说明所依据的工时填报签批原件、考勤与设备运行记录、工资表及社保公积金缴费凭证、固定资产折旧计提表、"
     "无形资产摊销表、租赁合同及付款凭证、研发项目立项文件与试验记录，连同本台账一并归档留存备查，"
     "保存期限不少于10年。"),
]
r = 4
for kind, text in DOC:
    if kind == "H":
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        c = ws.cell(r, 1, text); c.font = Font(name=FN, size=12, bold=True, color="1F4E79"); c.alignment = A_L
        ws.row_dimensions[r].height = 26
    elif kind == "F":
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
        c = ws.cell(r, 2, text); c.font = F_BOLD; c.alignment = A_L
        ws.row_dimensions[r].height = 20
    elif kind == "Q":
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
        c = ws.cell(r, 2, text)
        c.font = Font(name=FN, size=10, bold=True, color="C00000"); c.alignment = A_LT
        c.fill = fill(C_NOTE); c.border = BORDER
        ws.row_dimensions[r].height = 17 + 15 * (len(text) // 50)
    else:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
        c = ws.cell(r, 2, text); c.font = Font(name=FN, size=10.5); c.alignment = A_LT
        ws.row_dimensions[r].height = 16 + 15 * (len(text) // 52)
    r += 1

r += 2
sig = [("编  制：", "复  核：", "财务负责人："), ("研发负责人：", "企业负责人：", "日  期：      年    月    日")]
for line in sig:
    for k, t in enumerate(line):
        col = 2 + k * 3
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 2)
        c = ws.cell(r, col, t + "________________"); c.font = Font(name=FN, size=11); c.alignment = A_L
    ws.row_dimensions[r].height = 34
    r += 1
ws.merge_cells(start_row=r + 1, start_column=6, end_row=r + 1, end_column=9)
c = ws.cell(r + 1, 6, "（企业公章）"); c.font = Font(name=FN, size=11); c.alignment = A_C
ws.row_dimensions[r + 1].height = 50
setup_print(ws, landscape=False)
ws.page_margins.left = ws.page_margins.right = 0.5

# ============================================================
# 13/14. 附表1、附表2（月度填报签批模板）
# ============================================================
def build_form(name, title, left, right, proj_label, note, sign_line, unit_row):
    w = wb.create_sheet(name)
    pc = len(left) + 1
    lastc = pc + NP + len(right) - 1
    w.merge_cells(start_row=1, start_column=1, end_row=1, end_column=lastc)
    c = w.cell(1, 1, title); c.font = Font(name=FN, size=16, bold=True); c.alignment = A_C
    w.row_dimensions[1].height = 36
    w.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(4, lastc // 2))
    c = w.cell(2, 1, unit_row); c.font = Font(name=FN, size=11); c.alignment = A_L
    w.merge_cells(start_row=2, start_column=lastc - 3, end_row=2, end_column=lastc)
    c = w.cell(2, lastc - 3, "日期（年/月）：__________"); c.font = Font(name=FN, size=11); c.alignment = A_R
    w.row_dimensions[2].height = 24
    for i, lab in enumerate(left, 1):
        w.merge_cells(start_row=3, start_column=i, end_row=5, end_column=i)
        cc = w.cell(3, i, lab); cc.font = F_SUB; cc.fill = fill(C_SUB); cc.alignment = A_C
    w.merge_cells(start_row=3, start_column=pc, end_row=3, end_column=pc + NP - 1)
    cc = w.cell(3, pc, proj_label); cc.font = F_SUB; cc.fill = fill(C_SUB); cc.alignment = A_C
    for k in range(NP):
        col = pc + k
        cc = w.cell(4, col, "=项目清单!$B$%d" % (4 + k)); cc.font = F_SUB; cc.fill = fill(C_SUB); cc.alignment = A_C
        cc = w.cell(5, col, '=IF(%s4="","",IFERROR(VLOOKUP(%s4,项目清单!$B$4:$C$13,2,0)&"",""))' % (gcl(col), gcl(col)))
        cc.font = Font(name=FN, size=8); cc.fill = fill(C_SUB); cc.alignment = A_C
    for j, lab in enumerate(right):
        col = pc + NP + j
        w.merge_cells(start_row=3, start_column=col, end_row=5, end_column=col)
        cc = w.cell(3, col, lab); cc.font = F_SUB; cc.fill = fill(C_SUB); cc.alignment = A_C
    w.row_dimensions[3].height = 26; w.row_dimensions[4].height = 18; w.row_dimensions[5].height = 36
    for r_ in range(6, 21):
        w.cell(r_, 1, r_ - 5).font = F_BODY
        tc = gcl(pc + NP)          # 研发工时合计
        w.cell(r_, pc + NP, "=IF(SUM(%s%d:%s%d)=0,\"\",SUM(%s%d:%s%d))" % (gcl(pc), r_, gcl(pc+NP-1), r_, gcl(pc), r_, gcl(pc+NP-1), r_))
        w.cell(r_, pc + NP + 1, '=IF(OR(%s%d="",$%s%d=0),"",%s%d/$%s%d)'
               % (tc, r_, gcl(len(left) - 1), r_, tc, r_, gcl(len(left) - 1), r_))
        w.row_dimensions[r_].height = 24
    w.cell(21, 1, "合  计").font = F_BOLD
    w.merge_cells(start_row=21, start_column=1, end_row=21, end_column=2)
    for c_ in range(len(left) - 1, pc + NP + 1):
        w.cell(21, c_, "=IF(SUM(%s6:%s20)=0,\"\",SUM(%s6:%s20))" % (gcl(c_), gcl(c_), gcl(c_), gcl(c_)))
    w.row_dimensions[21].height = 24
    box(w, 3, 1, 21, lastc)
    for r_ in range(6, 22):
        for c_ in range(1, lastc + 1):
            cell = w.cell(r_, c_)
            cell.alignment = A_C if c_ < lastc - 1 else A_L
            cell.font = F_BOLD if r_ == 21 else F_BODY
            if r_ == 21: cell.fill = fill(C_TOTAL)
            if 3 <= c_ <= pc + NP: cell.number_format = FMT_H
            if c_ == pc + NP + 1: cell.number_format = FMT_P
    w.merge_cells(start_row=23, start_column=1, end_row=23, end_column=lastc)
    c = w.cell(23, 1, note); c.font = F_NOTE; c.alignment = A_L
    w.row_dimensions[23].height = 34
    step = max(3, lastc // len(sign_line))
    for k, t in enumerate(sign_line):
        col = 1 + k * step
        w.merge_cells(start_row=25, start_column=col, end_row=25, end_column=min(col + step - 1, lastc))
        c = w.cell(25, col, t + "____________"); c.font = Font(name=FN, size=11); c.alignment = A_L
    w.row_dimensions[25].height = 40
    widths(w, {"A": 6})
    for k in range(NP): widths(w, {gcl(pc + k): 9})
    setup_print(w)
    return w, pc, lastc

w, pc, lastc = build_form(
    "附表1-人员工时签批表",
    "《研发人员工时分配表》（月度填报签批用）",
    ["序号","姓名","工号","当月出勤\n总工时","非研发工时\n(生产及其他)"],
    ["研发工时\n合计","分摊系数\n(研发工时/总工时)","本月主要研发工作内容","本人签字"],
    "各研发项目分配工时（小时）",
    "填报要求：1. 逐人如实填报，当月出勤总工时须与考勤记录一致；2. “当月出勤总工时 = 非研发工时 + 各研发项目工时合计”，"
    "务必核对平衡；3. 研发专用人员记录全额工时，研发生产共用人员按实际投入分别记录；"
    "4. 研发工作内容须简明具体（如“RD03 无溶剂贴合线速匹配试验3批次”）；5. 本表经签批后交财务，原件按月装订留存备查。",
    ["填报人：","项目负责人（确认）：","研发部门负责人（审核）：","财务复核："],
    "填报部门：__________")
widths(w, {"B":12,"C":11,"D":12,"E":13})
widths(w, {gcl(pc+NP):11, gcl(pc+NP+1):13, gcl(pc+NP+2):34, gcl(pc+NP+3):12})

w2, pc2, lastc2 = build_form(
    "附表2-设备工时签批表",
    "《仪器、设备工时分配表》（月度填报签批用）",
    ["序号","资产编号","资产名称","所属工序","当月总运行\n(使用)工时","非研发使用工时\n(生产及其他)"],
    ["研发使用工时\n合计","研发使用占比","备注（试制单号/试验编号）","记录人签字"],
    "各研发项目使用工时（小时）",
    "填报要求：1. 按资产逐台填报，“当月总运行(使用)工时”须与设备运行台账、点检记录一致；"
    "2. “总运行工时 = 非研发使用工时 + 各研发项目使用工时合计”，务必核对平衡；3. 备注列填写对应的试制通知单号或试验编号，便于追溯；"
    "4. 研发专用设备记录全额运行工时，研发生产共用设备按实际使用分别记录，无形资产不在本表填报；"
    "5. 本表经签批后交财务，原件按月装订留存备查。",
    ["记录人：","设备管理员：","项目负责人（确认）：","研发部门负责人（审核）：","财务复核："],
    "填报部门：__________")
widths(w2, {"B":11,"C":22,"D":12,"E":13,"F":14})
widths(w2, {gcl(pc2+NP):12, gcl(pc2+NP+1):11, gcl(pc2+NP+2):26, gcl(pc2+NP+3):12})

w3, pc3, lastc3 = build_form(
    "附表3-无形资产工时签批表",
    "《无形资产工时分配表》（月度填报签批用）",
    ["序号","资产编号","无形资产名称","用途/说明","当月总使用\n工时","非研发使用工时\n(生产及其他)"],
    ["研发使用工时\n合计","研发使用占比","备注（对应项目/试验编号）","记录人签字"],
    "各研发项目使用工时（小时）",
    "填报要求：1. 按无形资产（配方模拟软件、专利技术、专有技术等）逐项填报，“当月总使用工时”须与系统使用日志、使用登记簿一致；"
    "2. “总使用工时 = 非研发使用工时 + 各研发项目使用工时合计”，务必核对平衡；3. 研发专用无形资产记录全额使用工时，研发生产共用的按实际使用分别记录；"
    "4. 备注列填写对应项目或试验编号，便于追溯；5. 本表经签批后交财务，原件按月装订留存备查。",
    ["记录人：","无形资产管理员：","项目负责人（确认）：","研发部门负责人（审核）：","财务复核："],
    "填报部门：__________")
widths(w3, {"B":11,"C":24,"D":14,"E":13,"F":14})
widths(w3, {gcl(pc3+NP):12, gcl(pc3+NP+1):11, gcl(pc3+NP+2):26, gcl(pc3+NP+3):12})

# ============================================================
# 定义名称 & 保存
# ============================================================
for nm, ref in [("项目编号列表", "项目清单!$B$4:$B$13"),
                ("人员工号列表", "人员名册!$B$4:$B$53"),
                ("资产编号列表", "资产台账!$B$4:$B$53")]:
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)

ORDER = ["使用说明","指标监测","项目清单","人员名册","资产台账",
         "研发人员工时分配表","研发人员工时费用分配表","仪器设备工时分配表","仪器设备工时费用分配表",
         "无形资产工时分配表","无形资产费用分配表","其他研发费用","月度归集汇总","年度汇总",
         "费用分配说明","附表1-人员工时签批表","附表2-设备工时签批表","附表3-无形资产工时签批表"]
assert sorted(ORDER) == sorted(wb.sheetnames), set(ORDER) ^ set(wb.sheetnames)
wb._sheets = [wb[n] for n in ORDER]
wb["使用说明"].sheet_properties.tabColor = "1F4E79"
wb["指标监测"].sheet_properties.tabColor = "BF8F00"
for s in ("项目清单","人员名册","资产台账"): wb[s].sheet_properties.tabColor = "70AD47"
for s in ("研发人员工时分配表","仪器设备工时分配表","无形资产工时分配表"): wb[s].sheet_properties.tabColor = "ED7D31"
for s in ("研发人员工时费用分配表","仪器设备工时费用分配表","无形资产费用分配表","其他研发费用"): wb[s].sheet_properties.tabColor = "A9550A"
for s in ("月度归集汇总","年度汇总"): wb[s].sheet_properties.tabColor = "C00000"
for s in ("费用分配说明","附表1-人员工时签批表","附表2-设备工时签批表","附表3-无形资产工时签批表"): wb[s].sheet_properties.tabColor = "7030A0"

wb.active = 0
wb.save(OUT)
print("saved:", OUT)
