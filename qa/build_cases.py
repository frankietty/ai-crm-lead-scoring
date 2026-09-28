# -*- coding: utf-8 -*-
"""生成《AI-CRM 线索评分助手 · 测试用例集》工作簿
Sheet：测试记录表（主操作台）/ 测试用例集（明细）/ 通过标准 / 优化动作清单
"""
try:
    import openpyxl
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule


def xl_color(css_hex: str) -> str:
    value = css_hex.removeprefix("#").upper()
    if len(value) != 6:
        raise ValueError(f"Expected #RRGGBB, got: {css_hex}")
    return "FF" + value


XL_HEADER_BG = xl_color("#4472C4")
XL_HEADER_FG = xl_color("#FFFFFF")
XL_INPUT_BG = xl_color("#D9E2F3")
XL_CALC_BG = xl_color("#F2F2F2")
XL_SUM_BG = xl_color("#2F5597")
XL_LINE = xl_color("#BFBFBF")
XL_OK_BG = xl_color("#C6EFCE")
XL_OK_FG = xl_color("#006100")
XL_BAD_BG = xl_color("#FFC7CE")
XL_BAD_FG = xl_color("#9C0006")
XL_WAIT_BG = xl_color("#FFEB9C")
XL_WAIT_FG = xl_color("#9C6500")

thin = Side(style="thin", color=XL_LINE)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F_HEAD = Font(bold=True, color=XL_HEADER_FG, size=10.5)
FILL_HEAD = PatternFill("solid", fgColor=XL_HEADER_BG)
FILL_IN = PatternFill("solid", fgColor=XL_INPUT_BG)
FILL_CALC = PatternFill("solid", fgColor=XL_CALC_BG)
AL_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
AL_CENTER = Alignment(horizontal="center", vertical="center")
AL_NUM = Alignment(horizontal="right", vertical="top")
F_BODY = Font(size=10.5)
F_BOLD = Font(bold=True, size=10.5)

wb = Workbook()
wb.properties.title = "AI-CRM 线索评分助手 测试用例集"

# ============ Sheet 2 数据源：测试用例集（先构建，供记录表 VLOOKUP） ============
CASE_HEADERS = ["用例编号", "层次", "用例名称", "行业", "公司规模", "岗位", "近期行为", "客户留言",
                "预期等级", "预期总分下限", "预期总分上限", "关键校验点", "设计理由"]

# (编号, 层次, 名称, 行业, 规模, 岗位, 行为, 留言, 预期等级, 下限, 上限, 校验点, 理由)
CASES = [
    ("TC-01", "正常", "高分线索判 SQL", "制造业", "500人以上", "IT负责人/技术总监",
     "打开报价邮件、访问集成对接页", "我们旧CRM想嵌AI评分，支持API对接吗？", "SQL", 85, 95,
     "话术必须出现「API」或「集成」；标签=集成需求", "验证最高档意图（集成需求）是否生效，SQL 支线 RAG 检索与话术生成是否跑通"),
    ("TC-02", "正常", "询价+约演示判 SQL", "软件/互联网", "100-499人", "采购/商务负责人",
     "预约产品演示", "想约个演示，顺便问下报价", "SQL", 70, 85,
     "标签属于{价格咨询,演示试用}；建议四段小标题齐全", "验证中高意图识别，以及多意图并存时取最高档"),
    ("TC-03", "正常", "弱画像弱行为判 MQL", "金融", "20-99人", "普通员工",
     "下载白皮书", "想了解一下你们的产品功能", "MQL", 38, 50,
     "不得越级为 SQL；出口含培育方案", "检验弱画像+弱行为组合是否正确落入培育池"),
    ("TC-04", "正常", "MQL 培育方案完整", "教育", "20-99人", "销售/市场负责人",
     "下载白皮书", "看看有没有合适的方案，先了解下", "MQL", 33, 45,
     "nurture_plan 三要素齐全（推荐资料/触达文案/下次触达）", "验证 MQL 支线培育内容生成正常"),
    ("TC-05", "正常", "学生线索自动过滤", "学生个人", "个人/无", "学生",
     "", "写作业需要了解CRM", "无效", 0, 5,
     "filter_reason 有内容；不触发知识库检索与话术生成", "关键用例：负分行业+负分规模是否被夹紧到 0 且被拦截"),
    ("TC-06", "正常", "未收录行业走默认分", "其他", "1-19人", "普通员工",
     "", "随便看看", "无效", 8, 20,
     "流程不报错；属性分不因「其他」而异常", "验证不在字典里的行业值走默认分、不报错"),
    ("TC-07", "边界", "卡在 60 分线上", "制造业", "100-499人", "IT负责人/技术总监",
     "下载白皮书", "想了解下功能细节", "SQL或MQL", 55, 65,
     "记录实际分数；同一条跑 3 次看等级是否跳变", "暴露阈值敏感区，为 60/30 阈值校准提供依据"),
    ("TC-08A", "边界", "同画像对照组（弱意图）", "制造业", "100-499人", "采购/商务负责人",
     "下载白皮书", "想了解下功能", "MQL", 52, 58,
     "与 TC-08B 同画像；B 组总分应高出本组 ≥10 分", "检验意图分的区分度（只换一句留言，其余字段完全相同）"),
    ("TC-08B", "边界", "同画像对照组（强意图）", "制造业", "100-499人", "采购/商务负责人",
     "下载白皮书", "支持API对接吗", "SQL", 67, 72,
     "与 TC-08A 比较：总分高出 ≥10 分且等级升为 SQL", "检验意图分的区分度（只换一句留言，其余字段完全相同）"),
    ("TC-09", "边界", "行为分封顶 25", "软件/互联网", "500人以上", "IT负责人/技术总监",
     "预约产品演示、打开报价邮件、观看演示视频、访问集成对接页、下载白皮书、一周内多次访问官网",
     "想看下你们的集成方案，支持API吗", "SQL", 88, 100,
     "behavior_score 必须=25（原始相加为 70）；总分≤100", "验证行为分封顶逻辑生效，否则总分可能爆表"),
    ("TC-10A", "边界", "属性分上限夹紧 45", "软件/互联网", "500人以上", "IT负责人/技术总监",
     "", "咨询下产品", "MQL", 50, 60,
     "attribute_score 必须=45，不得超过", "验证属性分上限夹紧，防止单项极值把总分带偏"),
    ("TC-10B", "边界", "属性分下限夹紧 0", "学生个人", "个人/无", "学生",
     "", "老师让我们调研CRM产品", "无效", 0, 6,
     "attribute_score 必须=0（不能是 -25）", "验证负数被夹紧到 0，不污染总分与漏斗统计"),
    ("TC-11", "边界", "行为词未命中字典", "制造业", "100-499人", "IT负责人/技术总监",
     "预约演示", "想了解下功能细节", "MQL", 45, 55,
     "behavior_score=0（行为词未命中字典）；应记为待修复的配置问题", "负向测试：故意用不匹配的行为词，主动挖出字典键与输入选项不一致的风险"),
    ("TC-12", "边界", "学生伪装调研被拦", "学生个人", "1-19人", "学生",
     "", "老师让我们调研一下CRM产品", "无效", 0, 6,
     "attribute_score=0；意图分<5", "验证「调研」这类正经词不会把学生线索抬成 MQL"),
    ("TC-13", "异常", "留言留空被校验拦下", "制造业", "100-499人", "IT负责人/技术总监",
     "", "", "拦截", 0, 0,
     "提交时应被必填校验拦下，不进入流程；实际总分填 0 表示未产生分数", "验证入参校验层生效——挡在最前面的防线"),
    ("TC-14", "异常", "无意义留言不崩", "其他", "1-19人", "普通员工",
     "", "。。。", "无效", 0, 29,
     "流程走完不报错；filter_reason 有内容", "验证无意义输入不会让链路崩溃，也不会被判成高意向"),
    ("TC-15", "异常", "超长留言（2000 字）", "制造业", "100-499人", "IT负责人/技术总监",
     "下载白皮书", "（此处粘贴一段约 2000 字的产品需求描述，核心是希望接入现有系统、支持API对接，并询问报价）",
     "SQL", 60, 95,
     "need_summary ≤30 字；单条耗时 ≤30 秒", "验证长度上限表现，以及长文本是否把提取结果带偏"),
    ("TC-16", "异常", "多意图混杂取最高档", "金融", "100-499人", "IT负责人/技术总监",
     "观看演示视频", "我们暂时没预算，但想问下支持API对接吗，另外报价多少", "SQL", 70, 88,
     "标签属于{集成需求,价格咨询}，不能是「随便看看」", "验证干扰词（没预算）不拉低分数，且取最高档意图"),
    ("TC-17", "异常", "兜底安全气囊（人为制造）", "制造业", "100-499人", "采购/商务负责人",
     "下载白皮书", "支持API对接吗", "MQL", 45, 55,
     "intent_score=8、标签=一般了解、流程不中断；本输入正常运行时应为 SQL，故本条须标注「人为制造异常」",
     "验证解析兜底生效——POC 现场不翻车的底气，也是工程健壮性的证明"),
    ("TC-18", "异常", "中英混杂留言", "软件/互联网", "100-499人", "IT负责人/技术总监",
     "访问集成对接页", "Do you support API integration? 我们想接入自己的系统", "SQL", 75, 92,
     "标签=集成需求", "验证中英混杂能被正确识别，不因英文而误判为无效"),
]

ws_case = wb.active
ws_case.title = "测试用例集"
for c, name in enumerate(CASE_HEADERS, start=1):
    cell = ws_case.cell(row=1, column=c, value=name)
    cell.font = F_HEAD
    cell.fill = FILL_HEAD
    cell.alignment = AL_CENTER
    cell.border = BORDER

for i, row in enumerate(CASES):
    r = 2 + i
    for c, val in enumerate(row, start=1):
        cell = ws_case.cell(row=r, column=c, value=val)
        cell.font = F_BODY
        cell.border = BORDER
        cell.alignment = AL_WRAP
        if c in (10, 11):
            cell.number_format = "0"
            cell.alignment = AL_NUM
        elif c == 1:
            cell.alignment = AL_CENTER
        elif c in (2, 9):
            cell.alignment = AL_CENTER

CASE_LAST = 1 + len(CASES)  # 19

# 测试用例集：列宽 + 冻结 + 筛选 + 层次条件格式
for col, w in zip("ABCDEFGHIJKLM",
                  [9, 8, 20, 12, 11, 15, 30, 40, 10, 11, 11, 40, 40]):
    ws_case.column_dimensions[col].width = w
ws_case.freeze_panes = "C2"
ws_case.auto_filter.ref = f"A1:M{CASE_LAST}"
ws_case.conditional_formatting.add(
    f"B2:B{CASE_LAST}",
    CellIsRule(operator="equal", formula=['"正常"'], fill=PatternFill("solid", fgColor=XL_OK_BG), font=Font(color=XL_OK_FG, size=10.5)))
ws_case.conditional_formatting.add(
    f"B2:B{CASE_LAST}",
    CellIsRule(operator="equal", formula=['"边界"'], fill=PatternFill("solid", fgColor=XL_WAIT_BG), font=Font(color=XL_WAIT_FG, size=10.5)))
ws_case.conditional_formatting.add(
    f"B2:B{CASE_LAST}",
    CellIsRule(operator="equal", formula=['"异常"'], fill=PatternFill("solid", fgColor=XL_BAD_BG), font=Font(color=XL_BAD_FG, size=10.5)))

# ============ Sheet 1 主操作台：测试记录表 ============
ws = wb.create_sheet("测试记录表", 0)
REC_HEADERS = ["用例编号", "层次", "预期等级", "预期下限", "预期上限", "关键校验点",
               "实际总分", "实际等级", "关键词命中", "判定", "归因备注（错在哪一环）"]
for c, name in enumerate(REC_HEADERS, start=1):
    cell = ws.cell(row=1, column=c, value=name)
    cell.font = F_HEAD
    cell.fill = FILL_HEAD
    cell.alignment = AL_CENTER if c != 6 and c != 11 else Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER

# 锚点：数据行 2..19（18 条），汇总区 21..25
FIRST, LAST = 2, 1 + len(CASES)
KEYWORD_CASES = {"TC-01", "TC-02", "TC-16", "TC-17", "TC-18"}

for i, row in enumerate(CASES):
    r = FIRST + i
    case_id = row[0]
    ws.cell(row=r, column=1, value=case_id).alignment = AL_CENTER
    # B..F 计算区：从测试用例集 VLOOKUP 取预期值
    ws.cell(row=r, column=2, value=f'=IFERROR(VLOOKUP($A{r},测试用例集!$A:$M,2,FALSE),"")')
    ws.cell(row=r, column=3, value=f'=IFERROR(VLOOKUP($A{r},测试用例集!$A:$M,9,FALSE),"")')
    ws.cell(row=r, column=4, value=f'=IFERROR(VLOOKUP($A{r},测试用例集!$A:$M,10,FALSE),"")')
    ws.cell(row=r, column=5, value=f'=IFERROR(VLOOKUP($A{r},测试用例集!$A:$M,11,FALSE),"")')
    ws.cell(row=r, column=6, value=f'=IFERROR(VLOOKUP($A{r},测试用例集!$A:$M,12,FALSE),"")')
    # G..I 输入区
    ws.cell(row=r, column=7, value=None)
    ws.cell(row=r, column=8, value=None)
    ws.cell(row=r, column=9, value=None if case_id in KEYWORD_CASES else "—")
    # J 判定：等级命中 且 分数在区间 且 关键词命中/不适用
    ws.cell(row=r, column=10, value=(
        f'=IF(OR($G{r}="",$H{r}=""),"待测",'
        f'IF(AND(ISNUMBER(SEARCH($H{r},$C{r})),'
        f'IF($D{r}="",TRUE,$G{r}>=$D{r}),IF($E{r}="",TRUE,$G{r}<=$E{r}),'
        f'OR($I{r}="是",$I{r}="—")),"通过","不通过"))'
    ))
    ws.cell(row=r, column=11, value=None)

    for c in range(1, 12):
        cell = ws.cell(row=r, column=c)
        cell.font = F_BODY
        cell.border = BORDER
        if c in (2, 3, 4, 5, 6):
            cell.fill = FILL_CALC
            cell.alignment = AL_WRAP if c == 6 else AL_CENTER
        elif c in (7, 8, 9, 11):
            cell.fill = FILL_IN
            cell.alignment = AL_WRAP if c == 11 else AL_CENTER
        else:
            cell.alignment = AL_CENTER
        if c in (4, 5, 7):
            cell.number_format = "0"

# 汇总区（数据区最下方，引用明细行公式）
SUM_START = LAST + 2
ws.cell(row=SUM_START, column=1, value="统计").font = F_BOLD
ws.cell(row=SUM_START, column=2, value="条数").font = F_BOLD
ws.cell(row=SUM_START, column=3, value="口径").font = F_BOLD
metrics = [
    ("通过", f'=COUNTIF($J${FIRST}:$J${LAST},"通过")', "等级命中 + 分数在区间 + 关键词命中"),
    ("不通过", f'=COUNTIF($J${FIRST}:$J${LAST},"不通过")', "任一判定条件未满足"),
    ("待测", f'=COUNTIF($J${FIRST}:$J${LAST},"待测")', "实际总分或实际等级未填"),
    ("通过率", f'=IF(COUNTA($J${FIRST}:$J${LAST})=0,"",COUNTIF($J${FIRST}:$J${LAST},"通过")/COUNTA($J${FIRST}:$J${LAST}))',
     "通过条数 ÷ 用例总数（目标 100%）"),
]
for k, (label, formula, note) in enumerate(metrics):
    r = SUM_START + 1 + k
    lc = ws.cell(row=r, column=1, value=label)
    lc.font = F_BOLD
    lc.border = BORDER
    lc.fill = PatternFill("solid", fgColor=XL_SUM_BG)
    lc.font = Font(bold=True, color=XL_HEADER_FG, size=10.5)
    vc = ws.cell(row=r, column=2, value=formula)
    vc.font = F_BOLD
    vc.border = BORDER
    vc.alignment = AL_CENTER
    if label == "通过率":
        vc.number_format = "0.0%"
    else:
        vc.number_format = "0"
    nc = ws.cell(row=r, column=3, value=note)
    nc.font = F_BODY
    nc.border = BORDER
    nc.alignment = AL_WRAP

for col, w in zip("ABCDEFGHIJK",
                  [10, 8, 10, 10, 10, 34, 10, 10, 11, 10, 40]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "C2"
ws.auto_filter.ref = f"A1:K{LAST}"

# 实际等级 / 关键词命中 数据验证
dv_grade = DataValidation(type="list", formula1='"SQL,MQL,无效,拦截"', allow_blank=True)
dv_grade.error = "请选择：SQL / MQL / 无效 / 拦截"
ws.add_data_validation(dv_grade)
dv_grade.add(f"H{FIRST}:H{LAST}")

dv_kw = DataValidation(type="list", formula1='"是,否,—"', allow_blank=True)
dv_kw.error = "请选择：是 / 否 / —（不适用）"
ws.add_data_validation(dv_kw)
dv_kw.add(f"I{FIRST}:I{LAST}")

# 判定列条件格式
ws.conditional_formatting.add(
    f"J{FIRST}:J{LAST}",
    CellIsRule(operator="equal", formula=['"通过"'], fill=PatternFill("solid", fgColor=XL_OK_BG), font=Font(color=XL_OK_FG, size=10.5, bold=True)))
ws.conditional_formatting.add(
    f"J{FIRST}:J{LAST}",
    CellIsRule(operator="equal", formula=['"不通过"'], fill=PatternFill("solid", fgColor=XL_BAD_BG), font=Font(color=XL_BAD_FG, size=10.5, bold=True)))
ws.conditional_formatting.add(
    f"J{FIRST}:J{LAST}",
    CellIsRule(operator="equal", formula=['"待测"'], fill=PatternFill("solid", fgColor=XL_WAIT_BG), font=Font(color=XL_WAIT_FG, size=10.5)))

# ============ Sheet 3：通过标准 ============
ws_std = wb.create_sheet("通过标准")
STD = [
    ("指标", "标准", "为什么这么定"),
    ("等级判定一致率", "≥ 80%（AI vs 人工双盲）", "POC 期最实用的指标——销售认不认可，比数学上的准确率更重要"),
    ("关键用例通过率", "TC-01（SQL）、TC-05（无效）必须 100%", "这两类是业务命门：漏掉高价值客户 / 放行垃圾线索，代价最大"),
    ("分数抖动", "同一条跑 3 次，极差 ≤ ±5 分", "分数飘忽销售就不信，可解释性直接崩掉"),
    ("兜底命中率", "≤ 5%", "兜底太多说明提示词约束不够，或模型能力不足"),
    ("单条处理耗时", "≤ 30 秒", "对齐案例文档里的验收标准（30 秒内完成打分）"),
    ("数值边界不越界", "属性分 0-45、行为分 0-25、意图分 0-30", "防止封顶 / 夹紧逻辑失效导致总分爆表"),
]
for i, row in enumerate(STD):
    r = 1 + i
    for c, val in enumerate(row, start=1):
        cell = ws_std.cell(row=r, column=c, value=val)
        cell.border = BORDER
        if r == 1:
            cell.font = F_HEAD
            cell.fill = FILL_HEAD
            cell.alignment = AL_CENTER
        else:
            cell.font = F_BOLD if c == 1 else F_BODY
            cell.alignment = AL_WRAP

for col, w in zip("ABC", [22, 34, 62]):
    ws_std.column_dimensions[col].width = w
ws_std.freeze_panes = "A2"

# ============ Sheet 4：优化动作清单 ============
ws_fix = wb.create_sheet("优化动作清单")
FIX = [
    ("症状", "可能原因", "对应动作"),
    ("分数普遍偏低 / 高意向被判 MQL", "阈值太保守，或缺行为数据",
     "先用真实成交样本回测改阈值；确认客户能否提供行为字段（缺行为分则总分上限只有 75）"),
    ("同一条线索分数每次不一样", "温度太高 / 提示词档位模糊",
     "温度降到 0.2；给每一档补一个例句（few-shot），让档位边界不重叠"),
    ("话术答非所问", "检索没取到对的资料",
     "检查知识库分段是否太碎、开 rerank、把查询词从留言原文换成提取后的关键词试试"),
    ("话术编造产品功能", "提示词约束不够",
     "强化「资料未覆盖的写需与产品团队确认」；补充知识库覆盖该话题"),
    ("经常走兜底", "提示词没给输出格式示例",
     "提示词里贴一个标准 JSON 示例；或换支持结构化输出（JSON Schema）的模型"),
    ("提取出的 JSON 塞不进下游", "下游引用的是 extract.text（整段 JSON 字符串）",
     "加一个解析节点取出 need_summary 再往下传（对应通关手册「待改点 2」）"),
    ("无效线索也花了钱", "分支位置太靠后",
     "把可前置的规则判断提前；注意提取与意图评分这两次调用无法避免，没有它们判断不出等级"),
]
for i, row in enumerate(FIX):
    r = 1 + i
    for c, val in enumerate(row, start=1):
        cell = ws_fix.cell(row=r, column=c, value=val)
        cell.border = BORDER
        if r == 1:
            cell.font = F_HEAD
            cell.fill = FILL_HEAD
            cell.alignment = AL_CENTER
        else:
            cell.font = F_BOLD if c == 1 else F_BODY
            cell.alignment = AL_WRAP

for col, w in zip("ABC", [28, 32, 64]):
    ws_fix.column_dimensions[col].width = w
ws_fix.freeze_panes = "A2"

OUT = r"D:\WorkBuddyData\WorkSpaces\2026-09-27-14-22-05\AI-CRM线索助手\测试用例集_AI-CRM线索评分助手.xlsx"
wb.save(OUT)
print("saved:", OUT)
print("sheets:", wb.sheetnames)
