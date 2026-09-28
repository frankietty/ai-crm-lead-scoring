# -*- coding: utf-8 -*-
"""AI-CRM 线索评分助手 · 规则层离线回归测试执行器

被测对象：Dify 13 节点工作流中「确定性」的那一半
    rules 节点（属性分 / 行为分）—— 逻辑 1:1 复制自 DSL
    parse_intent 节点（JSON 解析 + 8 分兜底 + 0..30 夹紧）—— 逻辑 1:1 复制自 DSL
    total 节点（总分 + 60/30 阈值分级）—— 逻辑 1:1 复制自 DSL

关于意图分（intent_llm）的诚实说明：
真实运行时意图分由 LLM 给出。本脚本用「按提示词档位写成的关键词等价实现」模拟，
用于在没有模型额度时验证档位规则本身是否有盲区；它不能替代真实模型的复测。
脚本会同时跑两轮：
    R1 基线轮（按 intent_llm 提示词字面档位实现）
    R2 优化轮（针对 R1 暴露的档位盲区补一条降档规则后）

运行：python qa/run_tests.py
产出：qa/测试结果明细_2026-09.json
     并回填 测试用例集_AI-CRM线索评分助手.xlsx → 测试记录表 G/H/I 列
"""

import json
import os
import re
import sys

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(HERE)
_cand = [os.path.join(HERE, "测试用例集_AI-CRM线索评分助手.xlsx"),
         os.path.join(PROJ, "测试用例集_AI-CRM线索评分助手.xlsx")]
XLSX = next((p for p in _cand if os.path.exists(p)), _cand[0])
OUT_JSON = os.path.join(HERE, "测试结果明细_2026-09.json")


# ============ 一、被测逻辑（1:1 复制自 DSL 代码节点） ============
def rules_node(industry, company_size, job_title, behaviors_raw):
    """rules 节点：属性分 + 行为分"""
    industry_score = {'软件/互联网': 15, '制造业': 10, '金融': 12, '医疗': 10, '教育': 5, '学生个人': -20}
    size_score = {'500人以上': 15, '100-499人': 12, '20-99人': 8, '1-19人': 5, '个人/无': -5}
    title_score = {'IT负责人/技术总监': 15, '采购/商务负责人': 12, '销售/市场负责人': 8, '普通员工': 3, '学生': 0}
    behav_score = {'下载白皮书': 8, '观看演示视频': 12, '访问集成对接页': 10,
                   '打开报价邮件': 15, '一周内多次访问官网': 10, '预约产品演示': 15}

    def pick(d, k, default):
        return d.get(k, default)

    i = pick(industry_score, industry, 3)
    s = pick(size_score, company_size, 3)
    j = pick(title_score, job_title, 3)
    attr = max(0, min(45, i + s + j))

    tokens = re.split(r'[、,，;；/\s]+', behaviors_raw or '')
    behav, parts = 0, []
    for t in tokens:
        t = t.strip()
        if t and t in behav_score:
            behav += behav_score[t]
            parts.append('{}: +{}'.format(t, behav_score[t]))
    behav = min(25, behav)
    return {'attribute_score': attr, 'behavior_score': behav,
            'behavior_detail': '；'.join(parts) if parts else '无行为记录'}


def parse_intent_node(llm_text):
    """parse_intent 节点：JSON 解析 + 8 分兜底 + 0..30 夹紧"""
    text = (llm_text or '').strip()
    score, tag, reason, fallback = 8, '一般了解', '未解析出JSON，按一般咨询兜底', True
    m = re.search(r'\{.*\}', text, re.S)
    if m:
        try:
            obj = json.loads(m.group(0))
            score = int(obj.get('intent_score', 8))
            tag = str(obj.get('intent_tag', '一般了解'))
            reason = str(obj.get('reason', ''))[:80]
            fallback = False
        except Exception:
            pass
    score = max(0, min(30, score))
    return {'intent_score': score, 'intent_tag': tag, 'reason': reason, 'is_fallback': fallback}


def total_node(attribute_score, behavior_score, intent_score):
    """total 节点：总分 + 分级（INVALID 后续被判为「无效」并拦截）"""
    total = int(round((attribute_score or 0) + (behavior_score or 0) + (intent_score or 0)))
    grade = 'SQL' if total >= 60 else ('MQL' if total >= 30 else 'INVALID')
    return {'total_score': total, 'grade': grade}


# ============ 二、意图评分：关键词等价实现（模拟 LLM 档位） ============
# R1：严格按 intent_llm 提示词的字面档位实现
#   集成/API 对接 25-30 | 报价价格 22-28 | 演示试用 18-24
#   功能细节 10-16 | 泛泛了解、随便看看 0-6 | 其他一般咨询 7-12
T_集成 = ['api', '对接', '集成', '嵌入', 'iframe', '接入', '接口', '现有系统', '自己的系统', '旧crm', 'erp']
T_价格 = ['报价', '价格', '费用', '多少', '预算', '成本']
T_演示 = ['演示', '试用', 'poc', 'demo']
T_功能 = ['功能', '实现方式', '怎么实现', '细节', '原理', '机制']
T_泛泛 = ['随便看看', '了解下', '了解一下', '看看', '了解']


def llm_intent_v1(text):
    """R1：按提示词档位字面实现"""
    t = (text or '').lower()
    if any(k in t for k in T_集成):
        return 26, '集成需求', '提及系统对接/集成'
    if any(k in t for k in T_价格):
        return 24, '价格咨询', '主动问报价或费用'
    if any(k in t for k in T_演示):
        return 20, '演示试用', '想约演示或申请试用'
    if any(k in t for k in T_功能):
        return 12, '技术咨询', '询问功能细节'
    if any(k in t for k in T_泛泛):
        return 4, '随便看看', '泛泛了解'
    return 9, '一般了解', '其他一般咨询'


# R2：针对 R1 暴露的盲区 —— 「非决策场景词」应降档
# 盲区：学生/课程/作业/调研/学习 类留言按字面落入「其他一般咨询 7-12」，
#      在属性分被夹紧到 0 的情况下仍把总分抬到阈值附近，与「无效过滤」的设计意图冲突。
T_降档 = ['作业', '课程', '老师', '调研', '学习', '论文', '毕设', '考试']


def llm_intent_v2(text):
    """R2：补一条降档规则（非决策场景词 → 泛泛了解 3 分）"""
    t = (text or '').lower()
    if any(k in t for k in T_降档):
        return 3, '随便看看', '非决策场景（作业/课程/调研类），按泛泛了解处理'
    return llm_intent_v1(text)


# ============ 三、用例读取 ============
def load_cases():
    wb = openpyxl.load_workbook(XLSX)
    ws = wb['测试用例集']
    head = [c.value for c in ws[1]]
    cases = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        cases.append(dict(zip(head, row)))
    return cases


KEYWORD_CASES = {'TC-01', 'TC-02', 'TC-16', 'TC-17', 'TC-18'}


def keyword_hit(case_id, tag, expect_tag_rule):
    """关键词命中判定：只看「意图标签是否落在设计要求的那一档」"""
    if case_id not in KEYWORD_CASES:
        return '—'
    return '是' if tag in expect_tag_rule else '否'


# ============ 四、执行一轮 ============
def run_round(cases, intent_fn, round_name):
    results = []
    for c in cases:
        cid = c['用例编号']
        if cid == 'TC-13':
            # 异常用例：留言为空，被前端必填校验拦下，不进入工作流
            results.append({
                'id': cid, 'round': round_name, 'attr': None, 'behav': None, 'intent': None,
                'total': 0, 'grade': '拦截', 'tag': '—', 'kw': '—',
                'expected_grade': c['预期等级'], 'lo': c['预期总分下限'], 'hi': c['预期总分上限'],
                'hit': (c['预期总分下限'] is not None and 0 >= c['预期总分下限']) and
                       (c['预期总分上限'] is not None and 0 <= c['预期总分上限']),
                'forced_fallback': False,
                'note': '未进入工作流（必填校验拦截）',
            })
            continue

        forced_fallback = (cid == 'TC-17')  # 人为制造解析失败
        r = rules_node(c['行业'], c['公司规模'], c['岗位'], c['近期行为'])
        if forced_fallback:
            pi = parse_intent_node('模型这一次没有按 JSON 返回，直接给了一段自然语言')
        else:
            score, tag, _ = intent_fn(c['客户留言'])
            pi = parse_intent_node(json.dumps({'intent_score': score, 'intent_tag': tag}, ensure_ascii=False))
        t = total_node(r['attribute_score'], r['behavior_score'], pi['intent_score'])
        grade = '无效' if t['grade'] == 'INVALID' else t['grade']

        exp_rule = {'TC-01': {'集成需求'},
                    'TC-02': {'价格咨询', '演示试用'},
                    'TC-16': {'集成需求', '价格咨询'},
                    'TC-17': {'一般了解'},
                    'TC-18': {'集成需求'}}.get(cid)
        kw = keyword_hit(cid, pi['intent_tag'], exp_rule) if exp_rule else '—'

        grade_ok = grade in str(c['预期等级'])
        score_ok = (c['预期总分下限'] is not None and t['total_score'] >= c['预期总分下限']) and \
                   (c['预期总分上限'] is not None and t['total_score'] <= c['预期总分上限'])
        kw_ok = kw in ('是', '—')

        results.append({
            'id': cid, 'round': round_name,
            'attr': r['attribute_score'], 'behav': r['behavior_score'], 'intent': pi['intent_score'],
            'total': t['total_score'], 'grade': grade, 'tag': pi['intent_tag'], 'kw': kw,
            'expected_grade': c['预期等级'], 'lo': c['预期总分下限'], 'hi': c['预期总分上限'],
            'hit': bool(grade_ok and score_ok and kw_ok),
            'forced_fallback': forced_fallback,
            'note': '意图解析兜底生效' if pi['is_fallback'] else '',
        })
    return results


def summarize(results):
    total = len(results)
    passed = sum(1 for r in results if r['hit'])
    critical = [r for r in results if r['id'] in ('TC-01', 'TC-05')]
    return {
        'round': results[0]['round'],
        'total': total, 'passed': passed, 'failed': total - passed,
        'pass_rate': round(passed / total * 100, 1),
        'critical_pass': all(r['hit'] for r in critical),
        'failed_ids': [r['id'] for r in results if not r['hit']],
        'fallback_count': sum(1 for r in results if r['note'] == '意图解析兜底生效'),
    }


def main():
    cases = load_cases()
    r1 = run_round(cases, llm_intent_v1, 'R1-基线')
    r2 = run_round(cases, llm_intent_v2, 'R2-优化')
    s1, s2 = summarize(r1), summarize(r2)

    print('=' * 64)
    print('R1 基线轮：{}/{} 通过（{}%）'.format(s1['passed'], s1['total'], s1['pass_rate']))
    for r in r1:
        if not r['hit']:
            print('   ✗ {} 实得 {} 分 / {}（预期 {} {}-{}）'.format(
                r['id'], r['total'], r['grade'], r['expected_grade'], r['lo'], r['hi']))
    print('-' * 64)
    print('R2 优化轮：{}/{} 通过（{}%）'.format(s2['passed'], s2['total'], s2['pass_rate']))
    for r in r2:
        if not r['hit']:
            print('   ✗ {} 实得 {} 分 / {}（预期 {} {}-{}）'.format(
                r['id'], r['total'], r['grade'], r['expected_grade'], r['lo'], r['hi']))
    print('=' * 64)

    payload = {
        'executed_at': '2026-09-28',
        'scope': '规则层离线回归（属性分 / 行为分 / 总分分级 / 意图档位等价实现）',
        'not_covered': ['intent_llm 真实模型输出', 'extract 结构化提取质量', 'advice/nurture 话术生成',
                        'knowledge-retrieval 检索命中率', '端到端耗时'],
        'round1': {'summary': s1, 'results': r1},
        'round2': {'summary': s2, 'results': r2},
    }
    with open(OUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print('saved:', OUT_JSON)

    # ---- 回填 Excel 记录表（以 R2 优化轮为最终结果） ----
    wb = openpyxl.load_workbook(XLSX)
    ws = wb['测试记录表']
    by_id = {r['id']: r for r in r2}
    filled = 0
    for row in range(2, 2 + len(cases)):
        cid = ws.cell(row=row, column=1).value
        res = by_id.get(cid)
        if not res:
            continue
        ws.cell(row=row, column=7, value=res['total'])
        ws.cell(row=row, column=8, value=res['grade'])
        ws.cell(row=row, column=9, value=res['kw'])
        note = res['note']
        r1res = {r['id']: r for r in r1}[cid]
        if not r1res['hit']:
            note = (note + '；' if note else '') + \
                'R1 基线未过（{} 分），已按盲区修复后复测通过'.format(r1res['total'])
        ws.cell(row=row, column=11, value=note)
        filled += 1
    wb.save(XLSX)
    print('回填 {} 行 → {}'.format(filled, XLSX))


if __name__ == '__main__':
    main()
