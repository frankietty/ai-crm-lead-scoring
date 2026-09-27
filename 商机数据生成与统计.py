# -*- coding: utf-8 -*-
"""
商机数据生成与统计 —— AI CRM 线索助手·数据分析交付物
1) 按评分模型（属性分45 + 行为分25 + 意图分30，>=60 SQL / >=30 MQL）生成 300 条模拟商机
2) 输出商机数据集 xlsx
3) 统计高意向客户特征，输出 stats json 供分析报告使用
"""
import openpyxl
import random
import json
from collections import Counter, defaultdict

random.seed(42)
N = 300
COLD_P = 0.48  # 非学生线索中“冷线索”（仅泛泛了解）占比，模拟真实漏斗
OUT_DIR = r"D:\WorkBuddyData\WorkSpaces\2026-09-27-14-22-05\AI-CRM线索助手"

# ---------- 字典 ----------
INDUSTRIES = ["制造业", "金融", "软件/互联网", "医疗", "零售电商", "专业服务", "教育", "政府/事业单位", "其他", "学生个人"]
INDUSTRY_W = [22, 14, 18, 10, 12, 8, 8, 4, 6, 8]
INDUSTRY_SCORE = {"制造业": 20, "金融": 20, "软件/互联网": 18, "医疗": 16, "零售电商": 12,
                  "专业服务": 10, "教育": 8, "政府/事业单位": 8, "其他": 6, "学生个人": 0}

SIZES = ["1-50人", "51-200人", "201-1000人", "1000人以上"]
SIZE_W = [42, 28, 20, 10]
SIZE_SCORE = {"1-50人": 5, "51-200人": 8, "201-1000人": 11, "1000人以上": 12}

ROLES = ["IT负责人/技术总监", "采购/商务负责人", "销售/市场负责人", "普通员工", "学生"]
ROLE_W = [20, 14, 16, 42, 8]
ROLE_SCORE = {"IT负责人/技术总监": 13, "采购/商务负责人": 12, "销售/市场负责人": 10, "普通员工": 4, "学生": 0}

CHANNELS = ["官网表单", "白皮书下载", "直播/活动", "社交媒体", "老客户推荐", "销售自拓"]
CHANNEL_W = [26, 18, 14, 22, 10, 10]

BEHAVIORS = ["预约演示", "试用注册", "浏览价格页", "下载白皮书", "观看演示视频", "官网多次访问"]
BEHAVIOR_SCORE = {"预约演示": 10, "试用注册": 8, "浏览价格页": 6, "下载白皮书": 5, "观看演示视频": 4, "官网多次访问": 2}

# 意图类型 -> (关键词组, 意图分基准, 消息模板)
INTENT_TYPES = {
    "集成对接": {"kw": ["集成", "对接", "API", "iframe", "嵌入"], "score": 24,
              "tpl": ["我们的旧系统想把AI评分嵌进去，支持API对接吗？", "能对接我们的{sys}系统吗？想约个演示",
                      "有没有开放接口，我们想集成到自己的CRM里"]},
    "询价预算": {"kw": ["价格", "报价", "预算", "费用", "多少钱"], "score": 20,
              "tpl": ["报价大概什么范围？我们有预算，{time}启动", "想问下价格，{n}人规模用哪个版本合适",
                      "预算有限，费用能不能按年付"]},
    "约演示": {"kw": ["演示", "demo", "试用"], "score": 18,
             "tpl": ["想约个演示，看看评分准不准", "下周能安排线上demo吗？我们把业务方也叫上",
                     "试用版在哪里注册？想先内部试试"]},
    "案例咨询": {"kw": ["案例", "同行", "参考"], "score": 10,
              "tpl": ["想多了解一些同行业案例", "有没有我们这个行业的落地案例可以参考"],
              },
    "功能了解": {"kw": ["功能", "了解", "看看"], "score": 6,
              "tpl": ["先了解一下功能，后面推荐给团队一起看看", "你们这个产品主要功能是什么？发点资料吧"]},
    "学生作业": {"kw": ["学生", "作业", "课程"], "score": 3,
               "tpl": ["学生，写作业需要了解一下CRM是什么概念", "课程作业需要调研一个SaaS产品"]},
}

SYS_WORDS = ["HIS", "ERP", "OA", "用友", "金蝶"]
TIME_WORDS = ["下个月", "月底", "季度末"]
SIZE_N = {"1-50人": "30", "51-200人": "100", "201-1000人": "300", "1000人以上": "800"}

SURNAME = "王李张刘陈杨赵黄周吴徐孙马朱胡郭何高林郑谢"
GIVEN = ["经理", "总监", "总", "工", "主管", "女士", "老师", "先生"]

# ---------- 生成 ----------
rows = []
for i in range(1, N + 1):
    industry = random.choices(INDUSTRIES, weights=INDUSTRY_W)[0]
    if industry == "学生个人":
        role, size = "学生", "1-50人"
    else:
        role = random.choices(ROLES[:4], weights=[22, 16, 18, 36])[0]
        size = random.choices(SIZES, weights=SIZE_W)[0]
    channel = random.choices(CHANNELS, weights=CHANNEL_W)[0]

    # 冷/热双机制：冷线索只做泛了解动作，热线索按职级与行业加权产生强行为
    senior = {"IT负责人/技术总监": 3, "采购/商务负责人": 3, "销售/市场负责人": 2, "普通员工": 1, "学生": 0}[role]
    hot_ind = industry in ("制造业", "金融", "软件/互联网", "医疗")
    cold = (role == "学生") or (random.random() < COLD_P)
    acts = []
    if cold:
        if random.random() < 0.22:
            acts.append("下载白皮书")
        if random.random() < 0.18:
            acts.append("观看演示视频")
        if random.random() < 0.28:
            acts.append("官网多次访问")
        if channel == "白皮书下载" and "下载白皮书" not in acts:
            acts.append("下载白皮书")
    else:
        if random.random() < 0.06 + 0.11 * senior + (0.06 if hot_ind else 0):
            acts.append("预约演示")
        if random.random() < 0.05 + 0.09 * senior + (0.05 if hot_ind else 0):
            acts.append("试用注册")
        if random.random() < 0.04 + 0.08 * senior + (0.04 if hot_ind else 0):
            acts.append("浏览价格页")
        if random.random() < 0.30:
            acts.append("下载白皮书")
        if random.random() < 0.30:
            acts.append("观看演示视频")
        if random.random() < 0.35:
            acts.append("官网多次访问")
        if channel == "白皮书下载" and "下载白皮书" not in acts:
            acts.append("下载白皮书")

    # 意图：与职级/行为相关，冷线索只泛了解
    if role == "学生":
        itype = "学生作业"
    elif cold:
        itype = random.choices(["功能了解", "案例咨询"], weights=[70, 30])[0]
    elif role in ("IT负责人/技术总监", "采购/商务负责人") and senior >= 3:
        itype = random.choices(["集成对接", "询价预算", "约演示", "案例咨询", "功能了解"],
                               weights=[30, 26, 22, 12, 10])[0]
    elif senior == 2:
        itype = random.choices(["询价预算", "约演示", "案例咨询", "功能了解", "集成对接"],
                               weights=[24, 24, 20, 22, 10])[0]
    else:
        itype = random.choices(["功能了解", "案例咨询", "约演示", "询价预算"], weights=[46, 26, 16, 12])[0]

    intent_score = min(30, max(0, INTENT_TYPES[itype]["score"] + random.randint(-4, 4)))
    if itype == "学生作业":
        intent_score = random.randint(0, 5)

    behav_score = min(25, sum(BEHAVIOR_SCORE[a] for a in acts))
    attr_score = INDUSTRY_SCORE[industry] + SIZE_SCORE[size] + ROLE_SCORE[role]
    total = attr_score + behav_score + intent_score
    grade = "SQL" if total >= 60 else ("MQL" if total >= 30 else "无效")

    name = random.choice(SURNAME) + random.choice(GIVEN)
    if industry == "学生个人":
        company = "个人"
    else:
        city = random.choice(["苏州", "深圳", "上海", "广州", "杭州", "武汉", "北京", "成都", "南京", "西安"])
        suffix = {"制造业": "智造设备", "金融": "金控集团", "软件/互联网": "信息科技", "医疗": "医疗设备",
                  "零售电商": "电商科技", "专业服务": "咨询公司", "教育": "培训机构", "政府/事业单位": "政务中心",
                  "其他": "实业公司"}[industry]
        company = city + suffix
    msg = random.choice(INTENT_TYPES[itype]["tpl"]).format(
        sys=random.choice(SYS_WORDS), time=random.choice(TIME_WORDS), n=SIZE_N[size])

    rows.append({
        "线索ID": "L%04d" % i, "客户名称": name, "公司": company, "行业": industry,
        "公司规模": size, "岗位": role, "来源渠道": channel,
        "行为轨迹": "+".join(acts) if acts else "仅访问1次",
        "客户留言": msg, "属性分": attr_score, "行为分": behav_score,
        "意图分": intent_score, "总分": total, "意向等级": grade,
    })

# ---------- 写 xlsx ----------
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "商机数据"
headers = list(rows[0].keys())
ws.append(headers)
for r in rows:
    ws.append([r[h] for h in headers])
xlsx_path = OUT_DIR + r"\商机数据集_模拟300条.xlsx"
wb.save(xlsx_path)

# ---------- 统计 ----------
stats = {"N": N}
g_sql = [r for r in rows if r["意向等级"] == "SQL"]
g_mql = [r for r in rows if r["意向等级"] == "MQL"]
g_bad = [r for r in rows if r["意向等级"] == "无效"]
stats["grade_dist"] = {"SQL": len(g_sql), "MQL": len(g_mql), "无效": len(g_bad)}
stats["avg_total"] = round(sum(r["总分"] for r in rows) / N, 1)

def rate_by(key, high_only=False):
    """各取值 -> (样本数, 高意向率[SQL+MQL占比], SQL率, 平均总分)"""
    bucket = defaultdict(list)
    for r in rows:
        bucket[r[key]].append(r)
    out = {}
    for k, rs in bucket.items():
        hi = [r for r in rs if r["意向等级"] in ("SQL", "MQL")]
        sql = [r for r in rs if r["意向等级"] == "SQL"]
        out[k] = {"n": len(rs), "hi_rate": round(100 * len(hi) / len(rs)),
                  "sql_rate": round(100 * len(sql) / len(rs)),
                  "avg": round(sum(r["总分"] for r in rs) / len(rs), 1)}
    return out

stats["by_industry"] = rate_by("行业")
stats["by_role"] = rate_by("岗位")
stats["by_size"] = rate_by("公司规模")
stats["by_channel"] = rate_by("来源渠道")

# 行为渗透率：高意向组(SQL+MQL) vs 无效组
def behav_pen(group):
    n = len(group)
    return {b: round(100 * sum(1 for r in group if b in r["行为轨迹"]) / n) for b in BEHAVIORS}
stats["behav_high"] = behav_pen(g_sql + g_mql)
stats["behav_bad"] = behav_pen(g_bad)

# 意图类型分布：按关键词归类留言
def intent_of(r):
    msg = r["客户留言"]
    if r["岗位"] == "学生":
        return "学生作业"
    for t, d in INTENT_TYPES.items():
        if any(k in msg for k in d["kw"]):
            return t
    return "功能了解"
cnt = Counter(intent_of(r) for r in rows)
stats["intent_dist"] = dict(cnt)
# 各意图类型的平均总分
intent_avg = defaultdict(list)
for r in rows:
    intent_avg[intent_of(r)].append(r["总分"])
stats["intent_avg"] = {k: round(sum(v) / len(v), 1) for k, v in intent_avg.items()}

# SQL 客户画像速览
stats["sql_industry_top"] = Counter(r["行业"] for r in g_sql).most_common(3)
stats["sql_role_top"] = Counter(r["岗位"] for r in g_sql).most_common(2)
stats["sql_size_top"] = Counter(r["公司规模"] for r in g_sql).most_common(2)
stats["sql_channel_top"] = Counter(r["来源渠道"] for r in g_sql).most_common(2)

with open(r"D:\WorkBuddyData\WorkSpaces\2026-09-27-14-22-05\_stats_out.json", "w", encoding="utf-8") as f:
    json.dump(stats, f, ensure_ascii=False, indent=1)
print("OK", len(rows))
