#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
    旅游投诉工单半结构化数据 → 结构化Label表格转换工具

    功能:
    1. 读取2023-2025年旅游投诉工单Excel数据
    2. 清洗文本（处理换行符、特殊字符等）
    3. 基于感知公平理论提取多维标签
    4. 生成可统计量化的结构化Label表格
    5. 标注不满意率高的情境和原因

    作者:lyh
    日期:2026-06-10
================================================================================
"""

import pandas as pd
import numpy as np
import re
import jieba
from collections import Counter
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# 一、配置信息
# ============================================================

# 1. 文件配置
INPUT_FILE = r'D:\Users\PC\Desktop\bbdlw\2023-2025旅游投诉工单.xlsx'      # 输入文件路径
OUTPUT_FILE = r'D:\Users\PC\Desktop\bbdlw\旅游投诉工单结构化Label表格.xlsx'  # 输出文件路径
OUTPUT_DIR = '/mnt/agents/output'                  # 输出目录

# 2. jieba分词自定义词典配置
CUSTOM_WORDS = [
    '长隆', '野生动物世界', '欢乐世界', '水上乐园', '熊猫酒店', '长隆酒店',
    '携程', '飞猪', '抖音', '美团', '淘宝', '微信小程序', '网购平台',
    '虚假宣传', '未履行', '退款', '退票', '退订', '退费', '退款难',
    '服务意识', '服务态度', '服务态度差', '态度恶劣', '辱骂游客',
    '调解', '行政执法', '行政调解', '终止调解',
    '售后服务', '售后客服', '客服电话', '投诉电话', '无法接通',
    '入园', '限流', '封场', '排队', '拥堵', '人山人海',
    '演出取消', '演唱会', '跨年', '烟花表演',
    '人身伤害', '安全隐患', '摔伤', '医疗费',
    '退款申请', '全额退款', '部分退款', '拒绝退款', '不予退款',
    '赔偿', '补偿', '门票期票', '优惠券', '差价',
    '感知公平', '结果公平', '程序公平', '互动公平',
    '未联系', '无法联系', '联系不上', '电话不通',
    '工单超时', '处理超时', '办理期限', '响应不及时'
]


# 3. 标签体系 - 投诉类型关键词字典
COMPLAINT_TYPE_KEYWORDS = {
    '退费退款纠纷': ['退费', '退款', '退票', '退订', '退钱', '返还', '退改', '退款难', '拒退', '不予退款', '扣除费用', '扣费', '退一赔三', '退差价', '退票'],
    '服务未履行': ['未履行', '未兑现', '未提供', '未安排', '未告知', '未通知', '未能成行', '服务缺失', '无法入场', '不能入园', '不给进'],
    '虚假宣传': ['虚假宣传', '虚假广告', '夸大宣传', '误导', '欺骗', '欺诈', '诱导消费', '与宣传不符', '与描述不符', '图文不符', '夸大其词', '营销欺诈'],
    '服务态度': ['态度', '辱骂', '呵斥', '推诿', '搪塞', '不理睬', '一问三不知', '态度恶劣', '态度差', '蛮横', '嚣张', '不尊重', '服务意识'],
    '安全问题': ['安全', '摔伤', '跌倒', '受伤', '流血', '急救', '医务', '医疗', '应急', '救援', '隐患', '危险', '人身伤害', '意外'],
    '价格问题': ['价格', '收费', '乱收费', '加价', '额外收费', '不合理收费', '高价', '宰客', '溢价', '价格欺诈', '差价', '消费券', '收费不合理'],
    '产品质量': ['质量', '产品质量', '商品质量', '食品问题', '餐饮质量', '变质', '损坏', '破旧'],
    '设施管理': ['设施', '设备', '设施损坏', '设备故障', '排队等候', '排队长', '拥挤', '限流', '封场', '关闭', '维修'],
    '市场秩序': ['秩序', '黄牛', '插队', '拥堵', '混乱', '管理混乱', '拉客', '回扣', '勾结'],
    '网络预订': ['网购', '平台', '小程序', '订单', '链接', '页面', '系统', '技术支持', '网络', '预订'],
    '签证政策': ['签证', '过境', '入境', '护照', '签注'],
    '其他问题': ['其他', '咨询', '建议', '表扬']
}

# 4. 标签体系 - 投诉原因关键词字典
CAUSE_KEYWORDS = {
    '未提供服务': ['未提供服务', '未能入场', '不能入园', '无法游玩', '未安排', '未兑现', '未履行约定', '不给进', '被拒绝', '无法进场', '入场受限', '限流'],
    '虚假宣传': ['虚假宣传', '夸大', '误导', '图文不符', '与描述不符', '实际不符', '欺骗', '诱导', '营销欺诈'],
    '退款困难': ['退款难', '拒不退款', '不予退款', '退不了', '拖延退款', '退款无果', '未退款', '扣款', '拒绝退款', '只退部分', '扣手续费', '退款申请被拒', '退费', '退款'],
    '态度恶劣': ['态度恶劣', '态度差', '态度不好', '辱骂', '呵斥', '不尊重', '不耐烦', '冷漠', '无视', '不理睬', '一问三不知', '态度蛮横', '嚣张', '不满'],
    '安全隐患': ['安全问题', '安全隐患', '摔伤', '跌倒', '受伤', '流血', '无安全防护', '应急不当', '医务', '急救', '设施不安全'],
    '过度拥挤': ['拥挤', '人太多', '人山人海', '拥堵', '排队长', '排队久', '排队时间', '排长队', '超负荷', '过度拥挤', '人流量过大'],
    '设施问题': ['设施', '设备损坏', '设施老旧', '维修中', '不开放', '关闭', '停用', '设施不足', '设备故障', '栏杆', '无法通行', '拦截'],
    '信息不透明': ['未告知', '未通知', '不知情', '未说明', '无提示', '信息不透明', '未提前告知', '未明确', '含糊其辞', '无任何提醒'],
    '等待时间长': ['等待', '等候', '等待时间长', '迟迟', '很久', '未按时', '迟到', '延迟', '无车'],
    '管理混乱': ['管理混乱', '管理不善', '组织不力', '协调不当', '无人管理', '缺乏管理', '监管缺失'],
    '违约行为': ['违约', '合同', '协议', '承诺', '单方面', '擅自', '强制', '强迫'],
    '价格欺诈': ['价格欺诈', '乱收费', '加价', '不合理收费', '宰客', '价格虚高', '差价', '多收费', '重复收费']
}

# 5. 标签体系 - 处置方式关键词字典
RESOLUTION_KEYWORDS = {
    '调解成功': ['调解成功', '双方达成一致', '协商解决', '双方同意', '和解', '成功调解', '调解结案'],
    '调解不成功': ['调解不成', '调解失败', '无法达成一致', '终止调解', '未能协商', '未达成一致', '调解无果', '协商不成', '调解不成功'],
    '企业已处理': ['已处理', '已解决', '已安排', '已落实', '已完成整改', '已改善', '已退费', '已退款'],
    '已退款': ['已退款', '已退费', '退回费用', '费用已退', '全额退款', '部分退款', '已原路退回'],
    '补偿赔偿': ['补偿', '赔偿', '赠送', '门票期票', '优惠券', '免费', '赔偿损失', '经济补偿', '抚慰金'],
    '解释说明': ['解释', '说明情况', '已告知', '答复如下', '现回复', '说明如下', '不予立案'],
    '转交部门': ['转交', '转派', '移交', '转由', '交由', '已转'],
    '无法联系': ['无法联系', '联系不上', '电话不通', '无人接听', '号码错误', '未接通', '无法接通'],
    '诉求人撤诉': ['撤诉', '撤回', '撤销', '不再追究', '不再投诉', '取消投诉'],
    '引导诉讼': ['建议诉讼', '司法途径', '法律途径', '向法院', '诉讼解决'],
    '立案查处': ['立案', '查处', '行政处罚', '责令改正', '依法处理'],
    '需进一步处理': ['进一步', '跟进', '继续处理', '正在办理', '加急处理']
}

# 6. 标签体系 - 感知公平维度关键词字典（基于开题报告理论框架）
FAIRNESS_KEYWORDS = {
    '结果公平-物质补偿': ['退款', '退费', '赔偿', '补偿', '经济补偿', '退还', '补偿损失', '免费重游', '赠送', '门票期票', '优惠券', '差价补偿'],
    '结果公平-问题解决': ['已解决', '问题已处理', '整改', '改善', '已落实', '已安排', '妥善解决', '彻底解决'],
    '程序公平-处理时效': ['及时处理', '迅速处理', '加急', '第一时间', '立即', '当天', '响应', '办理期限', '处理时间'],
    '程序公平-流程透明': ['调查', '核实', '核查', '协调', '告知', '说明', '回复如下', '程序', '依据'],
    '程序公平-一致性': ['依法依规', '按照', '依据', '统一', '规范', '标准', '程序正当'],
    '互动公平-尊重关怀': ['歉意', '道歉', '抱歉', '诚恳', '理解', '感谢', '重视', '关注', '诚挚', '遗憾', '关心'],
    '互动公平-沟通解释': ['解释', '说明', '答复', '沟通', '协调', '反馈', '告知', '回复', '说明如下', '详细说明'],
    '互动公平-态度友好': ['耐心', '细致', '积极', '主动', '认真', '热情', '友善', '友好', '周到']
}

# 7. 标签体系 - 不满意原因关键词字典
DISSATISFACTION_KEYWORDS = {
    '问题未解决': ['问题仍然未解决', '未解决', '未改善', '未处理', '问题未处理', '仍未解决', '未落实', '没有给予赔偿'],
    '回复与实际不符': ['回复与实际情况不一致', '与实际不符', '情况不一致', '答复不符'],
    '未收到答复': ['没有收到', '未收到', '未答复', '无回应', '没有答复'],
    '操作仍无效': ['按回复指引操作但问题仍未解决', '操作无效', '仍无法解决'],
    '部分改善': ['有改善但未能解决', '部分改善', '未彻底解决', '改善不明显']
}

# 8. 标签体系 - 情境关键词字典
SCENARIO_KEYWORDS = {
    '景区拥挤排队': ['排队', '拥挤', '人多', '拥堵', '限流', '超负荷', '人山人海', '爆满', '排队长', '排队久', '排队时间'],
    '演出活动问题': ['演唱会', '演出', '表演', '节目', '烟花', '跨年', '活动取消', '看不到演出', '封场'],
    '退改纠纷': ['退', '退款', '退票', '退订', '改期', '改签', '取消', '疫情', '无法出行'],
    '人员服务': ['态度', '服务', '工作人员', '员工', '客服', '保安', '导游', '经理', '接待', '态度恶劣'],
    '设施管理': ['设施', '设备', '维修', '关闭', '不开放', '设施损坏', '设备故障', '洗手间', '餐饮', '餐厅'],
    '安全问题': ['安全', '受伤', '摔伤', '流血', '急救', '医务', '医疗', '应急', '救援', '隐患'],
    '价格争议': ['价格', '收费', '费用', '加价', '宰客', '高价', '不合理', '差价', '消费券'],
    '交通问题': ['交通', '停车', '接驳车', '班车', '路况', '堵车', '停车位', '停车难'],
    '动物相关': ['动物', '猛兽', '投喂', '动物园', '野生动物', '表演', '展区'],
    '酒店住宿': ['酒店', '住宿', '房间', '客房', '入住', '退房', '预订', '房价', '设施陈旧']
}

# 9. 标签体系 - 结果不满意关键词字典
DISSATISFACTION_RESULT_KEYWORDS = {
    '退款退费争议': ['不退不换', '全款退', '差价', '手续费', '扣积分', '扣费', '未收到退款', '未退款', '没有退款', '补差价', '退票', '退费', '退钱', '随时退'],
    '政务部门不作为': ['一个多月', '不主动', '不作为', '不处理', '不执法', '不给予警告', '向上级反馈无回应', '处理方式', '官僚主义', '庸政', '形式主义', '懒政', '职责未做到', '超过一个月', '零回复'],
    '赔偿补偿诉求': ['医药费', '精神损失', '经济损失', '赔偿', '道歉'],
    '虚假办结/走过场': ['不得不同意', '回复已终结', '应付', '延期调查还是已办结', '强制同意', '形于流程', '显示给', '未处理就结工单', '未调解声称已调解', '草草结案', '走形式', '造假'],
    '沟通联系缺失': ['何谈解决', '啥也没有', '无人联系', '未接到', '未接到过', '未有工作人员来电', '根本无联系', '根本没有收到', '没有事先和我沟通', '没有任何承办单位', '没有回访告知', '没有收到电话'],
    '霸王条款/消费欺诈': ['低价吸引', '单方面', '强买强卖', '格式条款', '欺诈', '欺骗', '私自签订', '诈骗', '误导', '违法', '霸王条款'],
    '园区管理问题': ['为什么让我排', '充电宝', '卫生', '大客流', '客流', '排队', '插座', '电瓶车', '自带食品', '身高不够', '食品对动物', '黑名单'],
    '企业服务态度差': ['为什么别人可以', '傲慢', '冷漠', '冷漠无情', '店大欺客', '当无事发生', '态度不好', '搪塞', '敷衍', '毫不在意', '爱来不来', '理由模凌两可', '给出的理由', '视而不见', '轻视'],
    '诉求未落实': ['不同意关闭', '事情没有', '未将', '未将我方', '未解决', '没有得到一个', '没有得到解决', '没有解决', '诉求未', '问题未解决'],
    '政务部门偏袒企业': ['不站消费者', '以长隆单方', '偏帮', '包庇', '和稀泥', '站在酒店', '站在长隆', '站在长隆方', '纵容'],
    '异地维权困难': ['不是广东', '只能吃哑巴亏', '外地', '大老远', '机票'],
    '企业敷衍了事': ['仅一个电话', '处理方法非常不满意', '毫无诚意', '破布娃娃', '纪念品', '记录一下'],
    '调查核实不实': ['不和我消费者', '以偏概全', '实事求是', '怎么不和', '断章取义', '避重就轻', '需要我自己'],
    '处理效率低下': ['十多天', '十来天', '拖延', '效率', '效率双不在线', '未及时答复', '浪费时间', '浪费精力', '磨洋工', '近半年'],
    '执法监管缺位': ['为所欲为', '勾结', '外挂使用者未受到', '徇私', '未依法', '未依法查处', '未受到任何处罚', '未责令', '知法犯法'],
    '企业主体责任缺失': ['不负责', '拒不承认', '推卸责任', '管理主体责任' ],
    '个人信息/权益侵害': ['威胁', '恐吓', '拉黑', '拍照', '殴打', '签名', '身份证', '黑恶势力' ],
    '其他': ['物品遗失', '继续投诉', '表示理解' ],
    '收费争议': ['实时变动调整', '暗中涨价', '涨价', '补交费用', '费用合适', '额外付费消费'],
    '推诿扯皮': ['先后叫', '叫我去', '指向', '职责范围', '踢皮球'],
    '调解机制失效': ['最后一天', '未有任何协商', '未诚心', '未达成', '未达成一致'],
    '监管机制失效': ['监管的意义', '监管的意义在哪里', '被投诉方' ],
    '不可抗力损失承担': ['为什么损失要', '住院', '防疫'],
    '行政处罚争议': ['减速带', '危险驾驶', '处罚', '没有警告', '绿化带', '警告牌子', '霸权'],
    '处理口径反复': ['一会说'],
    '告知义务缺失': ['未告知']
}

# 10. 标签体系 - 核心诉求关键词字典
CAUSE_WAY_KEYWORDS = {
    '退款/退费': ['退款', '退费', '退钱', '退一赔三', '退差价', '退票'],
    '赔偿/补偿': ['赔偿', '补偿', '赔付', '损失费'],
    '查处/监管': ['查处', '监管', '介入', '严肃处理'],
    '道歉': ['道歉', '致歉', '赔礼'],
    '换货/更换': ['换货', '更换', '调换'],
    '补足': ['补足', '补差', '补齐'],
    '延长期限': ['延期', '延长期限', '延保'],
    '整改': ['整改', '改正'],
    '发货/物流': ['发货', '物流', '快递'],
    '其他/未明确': ['-']
}

# 11. 标签体系 - 政府处置方案关键词字典
RESOLUTION_WAY_KEYWORDS = {
    '解释致歉': ['致歉', '歉意', '道歉', '诚挚'],
    '退款/退费': ['退款', '退费', '退钱'],
    '门票延期/换票': ['门票延期', '期票', '再入园', '延期'],
    '赔偿/补偿': ['赔偿', '补偿', '赔付'],
    '引导司法途径': ['司法途径', '诉讼', '法院', '民事诉讼'],
    '责令整改	整改': ['整改'],
    '换货/更换': ['换货', '更换'],
    '无法联系诉求人': ['无法联系', '无人接听', '电话无人'],
    '解释说明': ['解释', '说明', '知悉', '沟通'],
    '解释说明': ['解释', '说明', '知悉', '沟通'],
    '诉求人撤诉': ['撤诉'],
    '其他/未明确': ['-']
}



# ============================================================
# 二、数据预处理函数
# ============================================================

def clean_text(text):
    """清洗文本:处理NaN,替换换行符和多余空格"""
    if pd.isna(text):
        return ''
    text = str(text)
    text = text.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
    text = text.replace('\t', ' ')
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def setup_jieba():
    """配置jieba分词自定义词典"""
    for word in CUSTOM_WORDS:
        jieba.add_word(word)


# ============================================================
# 三、标签提取核心函数
# ============================================================

def extract_labels_from_text(text, keyword_dict, mode='multi', exact_match=False):
    """
    从文本中提取标签

    参数:
        text: 输入文本
        keyword_dict: 关键词字典 {label: [keywords]}
        mode: 'multi'返回多个匹配标签, 'single'返回最佳匹配标签
        exact_match: 是否使用精确匹配

    返回:
        匹配的标签列表或单个最佳标签
    """
    if not text or pd.isna(text):
        return [] if mode == 'multi' else '未标注'

    text = str(text)
    matched_labels = {}

    for label, keywords in keyword_dict.items():
        score = 0
        for keyword in keywords:
            if exact_match:
                if keyword == text or keyword in text.split(';'):
                    score += 100
            else:
                if keyword in text:
                    if len(keyword) <= 2:
                        idx = text.find(keyword)
                        prev_ok = (idx == 0) or (text[idx-1] in ',。；,.; \t') \
                                  or not ('\u4e00' <= text[idx-1] <= '\u9fff')
                        next_ok = (idx + len(keyword) >= len(text)) \
                                  or (text[idx+len(keyword)] in ',。；,.; \t') \
                                  or not ('\u4e00' <= text[idx+len(keyword)] <= '\u9fff')
                        if prev_ok or next_ok:
                            score += len(keyword) * 10
                    else:
                        score += len(keyword) * 10

        if score > 0:
            matched_labels[label] = score

    if not matched_labels:
        return [] if mode == 'multi' else '其他'

    if mode == 'single':
        return max(matched_labels, key=matched_labels.get)
    else:
        sorted_labels = sorted(matched_labels.items(), key=lambda x: x[1], reverse=True)
        return [label for label, score in sorted_labels]


def extract_all_labels(row):
    """为每一行提取所有类型的标签"""

    # 1. 投诉类型 - 从诉求标题 + 事项分类提取
    title = str(row.get('诉求标题', ''))
    cat1 = str(row.get('事项分类一级', ''))
    cat2 = str(row.get('事项分类二级', ''))
    cat3 = str(row.get('事项分类三级', ''))

    title_labels = extract_labels_from_text(title, COMPLAINT_TYPE_KEYWORDS, 'multi')
    cat_labels = extract_labels_from_text(f"{cat1};{cat2};{cat3}", COMPLAINT_TYPE_KEYWORDS, 'multi', exact_match=True)

    all_complaint_labels = list(dict.fromkeys(title_labels + cat_labels))
    if not all_complaint_labels:
        if '市场监管' in cat1:
            all_complaint_labels = ['消费纠纷']
        elif '售后服务' in cat2:
            all_complaint_labels = ['服务未履行']
        else:
            all_complaint_labels = ['其他问题']

    # 2. 投诉原因
    text_cause = ' '.join(filter(None, [
        str(row.get('标签组', '')),
        str(row.get('市民诉求', '')),
        str(row.get('补充信息', ''))
    ]))
    cause_labels = extract_labels_from_text(text_cause, CAUSE_KEYWORDS, 'multi')

    # 3. 处置方式
    text_resolution = ' '.join(filter(None, [
        str(row.get('回复内容', '')),
        str(row.get('办理结果', ''))
    ]))
    resolution_labels = extract_labels_from_text(text_resolution, RESOLUTION_KEYWORDS, 'multi')

    # 4. 感知公平维度
    fairness_labels = extract_labels_from_text(text_resolution, FAIRNESS_KEYWORDS, 'multi')

    # 5. 不满意原因
    text_dissatisfaction = ' '.join(filter(None, [
        str(row.get('办案不满意原因', '')),
        str(row.get('回访意见', ''))    #20260623优化,新增对回访意见不满意原因的归纳
        ]))
    dissatisfaction_labels = extract_labels_from_text(text_dissatisfaction, DISSATISFACTION_KEYWORDS, 'multi')

    # 6. 情境
    text_scenario = ' '.join(filter(None, [
        str(row.get('诉求标题', '')),
        str(row.get('标签组', '')),
        str(row.get('市民诉求', '')),
        str(row.get('补充信息', '')),
        str(row.get('回复内容', ''))
    ]))
    scenario_labels = extract_labels_from_text(text_scenario, SCENARIO_KEYWORDS, 'multi')

    # 7. 结果不满意原因
    text_dissatisfaction_result = ' '.join(filter(None, [
        str(row.get('回访意见', ''))
        ]))
    dissatisfaction_reslut_labels = extract_labels_from_text(text_dissatisfaction_result, DISSATISFACTION_RESULT_KEYWORDS, 'multi')

    # 8. 核心诉求原因
    text_cause_way = ' '.join(filter(None, [
        str(row.get('市民诉求', '')),
        str(row.get('补充信息', '')),
        str(row.get('市民原始诉求', ''))
        ]))
    cause_way_labels = extract_labels_from_text(text_cause_way, CAUSE_WAY_KEYWORDS, 'multi')

    # 9. 政府处置方案原因
    text_resolution_way_result = ' '.join(filter(None, [
        str(row.get('回复内容', ''))
        ]))
    resolution_way_labels = extract_labels_from_text(text_resolution_way_result, RESOLUTION_WAY_KEYWORDS, 'multi')

    return {
        'complaint_type_labels': all_complaint_labels,
        'complaint_type_primary': all_complaint_labels[0] if all_complaint_labels else '其他',
        'cause_labels': cause_labels,
        'cause_primary': cause_labels[0] if cause_labels else '其他',
        'resolution_labels': resolution_labels,
        'resolution_primary': resolution_labels[0] if resolution_labels else '其他',
        'fairness_labels': fairness_labels,
        'fairness_primary': fairness_labels[0] if fairness_labels else '未标注',
        'dissatisfaction_labels': dissatisfaction_labels,
        'dissatisfaction_primary': dissatisfaction_labels[0] if dissatisfaction_labels else '',
        'scenario_labels': scenario_labels,
        'scenario_primary': scenario_labels[0] if scenario_labels else '其他',
        'dissatisfaction_reslut_labels': dissatisfaction_reslut_labels,
        'dissatisfaction_result_primary': dissatisfaction_reslut_labels[0] if dissatisfaction_reslut_labels else '',
        'cause_way_labels': cause_way_labels,
        'cause_way_labels_primary': cause_way_labels[0] if cause_way_labels else '',
        'resolution_way_labels': resolution_way_labels,
        'resolution_way_primary': resolution_way_labels[0] if resolution_way_labels else ''
    }


# ============================================================
# 四、主流程
# ============================================================

def main():
    """主处理流程"""

    # 1. 配置jieba
    setup_jieba()

    # 2. 读取数据
    print(f"读取数据: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE)
    print(f"共 {df.shape[0]} 行, {df.shape[1]} 列")

    # 3. 数据清洗
    data = df.copy()
    text_columns = ['诉求标题', '涉事主体', '事发地点', '标签组', '市民诉求', 
                    '补充信息', '回复内容', '办理结果', '事项分类一级', '事项分类二级', 
                    '事项分类三级', '对象分类一级', '对象分类二级', '核查情况', '回访意见', '市民原始诉求']
    for col in text_columns:
        if col in data.columns:
            data[col] = data[col].apply(clean_text)
    data['序号'] = data['序号'].astype(str)

    # 4. 提取标签（分批处理）
    print("开始提取标签...")
    all_labels = []
    for idx in range(len(data)):
        row = data.iloc[idx]
        labels = extract_all_labels(row)
        all_labels.append(labels)
    labels_df = pd.DataFrame(all_labels)

    # 5. 构建输出表格
    output_columns = {
        '年份': '年份', '序号': '序号', '诉求标题': '诉求标题',
        '涉事主体': '涉事主体',
        '受理时间': '受理时间', '处理时间': '处理时间',
        '工单状态': '工单状态', '来源渠道': '来源渠道',
        '核查情况': '核查情况', '回复时间': '回复时间',
        '回复方式': '回复方式', '是否无法联系诉求人': '是否无法联系诉求人',
        '联系时间': '联系时间', '诉求人是否撤诉': '诉求人是否撤诉',
        '回访时间': '回访时间', '话务满意度': '话务满意度',
        '办案满意度': '办案满意度'
    }

    result = pd.DataFrame()
    for new_col, old_col in output_columns.items():
        result[new_col] = data[old_col] if old_col in data.columns else ''

    # 合并标签
    result['投诉类型标签'] = labels_df['complaint_type_labels'].apply(lambda x: ';'.join(x))
    result['投诉类型_主标签'] = labels_df['complaint_type_primary']
    result['投诉原因标签'] = labels_df['cause_labels'].apply(lambda x: ';'.join(x))
    result['投诉原因_主标签'] = labels_df['cause_primary']
    result['处置方式标签'] = labels_df['resolution_labels'].apply(lambda x: ';'.join(x))
    result['处置方式_主标签'] = labels_df['resolution_primary']
    result['感知公平维度标签'] = labels_df['fairness_labels'].apply(lambda x: ';'.join(x))
    result['感知公平_主标签'] = labels_df['fairness_primary']
    result['不满意原因标签'] = labels_df['dissatisfaction_labels'].apply(lambda x: ';'.join(x))
    result['不满意原因_主标签'] = labels_df['dissatisfaction_primary']
    result['情境标签'] = labels_df['scenario_labels'].apply(lambda x: ';'.join(x))
    result['情境_主标签'] = labels_df['scenario_primary']
    result['是否不满意'] = result['办案满意度'].isin(['不满意', '非常不满意'])
    result['结果不满意原因标签'] = labels_df['dissatisfaction_reslut_labels'].apply(lambda x: ';'.join(x))
    result['结果不满意原因_主标签'] = labels_df['dissatisfaction_result_primary']
    result['核心诉求标签'] = labels_df['cause_way_labels'].apply(lambda x: ';'.join(x))
    result['核心诉求_主标签'] = labels_df['cause_way_labels_primary']
    result['政府处置方案标签'] = labels_df['resolution_way_labels'].apply(lambda x: ';'.join(x))
    result['政府处置方案_主标签'] = labels_df['resolution_way_primary']
    
    
    # 6. 保存Excel
    print(f"保存结果到: {OUTPUT_FILE}")
    result.to_excel(OUTPUT_FILE, index=False)
    print("处理完成!")

    return result


if __name__ == '__main__':
    main()