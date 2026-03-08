#!/usr/bin/env python3
"""
北京邮电大学硕士学位论文格式化脚本（完整版）
格式规范来源：模板-理工科（2025年1月15日）.doc 全部39条批注

用法：
  python3 format_thesis.py <document.xml路径> [--apply-body] [--fix-margins]

  --apply-body : 将正文格式应用到所有未分类段落（默认开启）
  --fix-margins: 修正页面边距为模板规范值（默认开启）
"""

import lxml.etree as ET
import re
import sys
import os
import copy

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'
MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'

# ============================================================
# 字号对照表 (半磅, half-point)
# ============================================================
FONT_SIZES = {
    '初号': 84, '小初': 72, '一号': 52, '小一': 48,
    '二号': 44, '小二': 36, '三号': 32, '小三': 30,
    '四号': 28, '小四': 24, '五号': 21, '小五': 18,
}

# ============================================================
# 格式规范定义（来自模板批注 0-38）
# ============================================================
FORMATS = {
    # === 封面部分 ===
    # 批注1: "学位论文" 黑体 32号 加粗 居中
    'cover_thesis_type': {
        'eastAsia': '黑体', 'ascii': '黑体', 'hAnsi': '黑体', 'cs': 'Times New Roman',
        'sz': 64, 'bold': True,
        'jc': 'center', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # 批注2: 封面题目 宋体/TNR 小二号 加粗 两端对齐
    'cover_title': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小二'], 'bold': True,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # 批注3: 学号至学院 宋体/TNR 四号 加粗 两端对齐 首行缩进6.05字符
    'cover_info': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['四号'], 'bold': True,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '240', 'lineRule': 'auto'},
        'indent': {'firstLineChars': '605', 'firstLine': '1694'},
    },
    # 批注4: 封面日期 宋体/TNR 四号 加粗 居中
    'cover_date': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['四号'], 'bold': True,
        'jc': 'center', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # === 声明页 ===
    # 批注6/8: 声明标题 黑体 三号 居中
    'declaration_title': {
        'eastAsia': '黑体', 'ascii': '黑体', 'hAnsi': '黑体', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['三号'], 'bold': False,
        'jc': 'center', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # 批注7/9: 声明内容 宋体/TNR 小四 两端对齐 首行缩进2字符 固定值20磅
    'declaration_body': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小四'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': {'firstLineChars': '200', 'firstLine': '480'},
    },
    # === 摘要 ===
    # 批注10: 摘要标题 黑体 三号 大纲1级 居中 段后1行
    'abstract_title': {
        'eastAsia': '黑体', 'ascii': '黑体', 'hAnsi': '黑体', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['三号'], 'bold': False,
        'jc': 'center', 'outlineLvl': '0',
        'spacing': {'before': '0', 'after': '312', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # 批注11: 中文摘要内容 宋体/TNR 四号 两端对齐 首行缩进2字符 固定值20磅
    'abstract_body': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['四号'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': {'firstLineChars': '200', 'firstLine': '560'},
    },
    # 批注12: 关键词 宋体/TNR 四号 首行缩进2字符 段前1行 固定值20磅
    'keywords': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['四号'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '312', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': {'firstLineChars': '200', 'firstLine': '560'},
    },
    # 批注13: ABSTRACT标题 TNR 三号 加粗 大纲1级 居中 段后2行
    'abstract_en_title': {
        'eastAsia': 'Times New Roman', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['三号'], 'bold': True,
        'jc': 'center', 'outlineLvl': '0',
        'spacing': {'before': '0', 'after': '624', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # 批注14: 英文摘要内容 TNR 四号 首行缩进2字符 固定值20磅
    'abstract_en_body': {
        'eastAsia': 'Times New Roman', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['四号'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': {'firstLineChars': '200', 'firstLine': '560'},
    },
    # 批注15: KEY WORDS TNR 四号 首行缩进2字符 段前2行 固定值20磅
    'keywords_en': {
        'eastAsia': 'Times New Roman', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['四号'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '624', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': {'firstLineChars': '200', 'firstLine': '560'},
    },
    # === 目录 ===
    # 批注16: 目录标题 黑体 三号 居中 段后2行
    'toc_title': {
        'eastAsia': '黑体', 'ascii': '黑体', 'hAnsi': '黑体', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['三号'], 'bold': False,
        'jc': 'center', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '624', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # 批注16: 目录一级 黑体/TNR 小四 两端对齐 固定值20磅
    'toc1': {
        'eastAsia': '黑体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小四'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': None,
    },
    # 批注16: 目录二级 宋体/TNR 小四 两端对齐 左缩进2字符 固定值20磅
    'toc2': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小四'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': {'leftChars': '200', 'left': '440'},
    },
    # 批注16: 目录三级 宋体/TNR 小四 两端对齐 左缩进4字符 固定值20磅
    'toc3': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小四'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': {'leftChars': '400', 'left': '880'},
    },
    # === 符号说明 ===
    # 批注17: 符号说明标题 黑体 三号 居中 段后2行
    'symbol_title': {
        'eastAsia': '黑体', 'ascii': '黑体', 'hAnsi': '黑体', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['三号'], 'bold': False,
        'jc': 'center', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '624', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # 批注17: 符号说明内容 宋体/TNR 小四 固定值20磅
    'symbol_body': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小四'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': None,
    },
    # === 正文标题 ===
    # 批注18: 一级标题 黑体/TNR 三号 大纲1级 居中 段后2行 单倍行距
    'heading1': {
        'eastAsia': '黑体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['三号'], 'bold': False,
        'jc': 'center', 'outlineLvl': '0',
        'spacing': {'before': '0', 'after': '624', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # 批注19: 二级标题 黑体/TNR 四号 大纲2级 左对齐 段前后0.5行 单倍行距
    'heading2': {
        'eastAsia': '黑体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['四号'], 'bold': False,
        'jc': 'left', 'outlineLvl': '1',
        'spacing': {'before': '156', 'after': '156', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # 批注25: 三级标题 黑体/TNR 小四 大纲3级 左对齐 固定值20磅
    'heading3': {
        'eastAsia': '黑体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小四'], 'bold': False,
        'jc': 'left', 'outlineLvl': '2',
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': None,
    },
    # 批注31: 四级标题 宋体/TNR 小四 大纲4级 两端对齐 首行缩进2字符 固定值20磅
    'heading4': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小四'], 'bold': False,
        'jc': 'both', 'outlineLvl': '3',
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': {'firstLineChars': '200', 'firstLine': '480'},
    },
    # === 正文 ===
    # 批注20: 正文 宋体/TNR 小四 两端对齐 首行缩进2字符 固定值20磅
    'body': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小四'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': {'firstLineChars': '200', 'firstLine': '480'},
    },
    # === 脚注 ===
    # 批注21: 脚注 宋体/TNR 小五 左对齐 单倍行距
    'footnote': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小五'], 'bold': False,
        'jc': 'left', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # === 图表 ===
    # 批注22/27: 图题/表题 楷体/TNR 五号 居中 固定值20磅
    'figure_table_caption': {
        'eastAsia': '楷体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['五号'], 'bold': False,
        'jc': 'center', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': None,
    },
    # 批注23/28: 图注/表注 宋体/TNR 五号 两端对齐 首行缩进2字符 固定值20磅
    'figure_table_note': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['五号'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': {'firstLineChars': '200', 'firstLine': '420'},
    },
    # === 公式 ===
    # 批注26: 公式段落 段前段后6磅 1.5倍行距
    'formula': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小四'], 'bold': False,
        'jc': 'center', 'outlineLvl': None,
        'spacing': {'before': '120', 'after': '120', 'line': '360', 'lineRule': 'auto'},
        'indent': None,
    },
    # === 特殊一级标题 ===
    # 批注33/35/37/38: 参考文献/致谢/附录/个人简历 标题 黑体 三号 大纲1级 居中 段后2行
    'special_heading1': {
        'eastAsia': '黑体', 'ascii': '黑体', 'hAnsi': '黑体', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['三号'], 'bold': False,
        'jc': 'center', 'outlineLvl': '0',
        'spacing': {'before': '0', 'after': '624', 'line': '240', 'lineRule': 'auto'},
        'indent': None,
    },
    # === 参考文献内容 ===
    # 批注33: 宋体/TNR 五号 两端对齐 固定值20磅
    'reference_body': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['五号'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': None,
    },
    # === 致谢内容 ===
    # 批注34: 宋体 小四 两端对齐 单倍行距
    'ack_body': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小四'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '240', 'lineRule': 'auto'},
        'indent': {'firstLineChars': '200', 'firstLine': '480'},
    },
    # === 致谢/个人简历内容 ===
    # 批注36/38: 宋体/TNR 小四 两端对齐 固定值20磅
    'special_body': {
        'eastAsia': '宋体', 'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'cs': 'Times New Roman',
        'sz': FONT_SIZES['小四'], 'bold': False,
        'jc': 'both', 'outlineLvl': None,
        'spacing': {'before': '0', 'after': '0', 'line': '400', 'lineRule': 'exact'},
        'indent': {'firstLineChars': '200', 'firstLine': '480'},
    },
}


# ============================================================
# 段落分类器（增强版）
# ============================================================
SPECIAL_HEADINGS = ['参考文献', '致谢', '致 谢', '附录', '附 录',
                    '攻读学位期间发表的学术论文', '攻读学位期间发表的学术论文目录',
                    '个人简历', '作者简介']

DECLARATION_TITLES = ['独创性（或创新性）声明', '独创性声明', '关于学位论文使用授权的声明',
                      '关于论文使用授权的声明', '学位论文使用授权声明']

COVER_INFO_KEYWORDS = ['学    号', '姓    名', '学科专业', '学习方式', '导    师', '学    院',
                       '学号', '姓名', '专业', '导师', '学院']

# State machine for context-aware classification
SECTION_COVER = 'cover'
SECTION_DECLARATION = 'declaration'
SECTION_ABSTRACT_CN = 'abstract_cn'
SECTION_ABSTRACT_EN = 'abstract_en'
SECTION_TOC = 'toc'
SECTION_SYMBOL = 'symbol'
SECTION_BODY = 'body'
SECTION_REFERENCE = 'reference'
SECTION_ACK = 'ack'
SECTION_APPENDIX = 'appendix'
SECTION_PUB = 'pub'
SECTION_RESUME = 'resume'


def normalize_spaces(text):
    return re.sub(r'\s+', '', text)


def classify_paragraph(text, section_ctx='body'):
    """根据段落文本和上下文判断格式类别"""
    text = text.strip()
    if not text:
        return None

    norm = normalize_spaces(text)

    # 声明页标题
    for dt in DECLARATION_TITLES:
        if normalize_spaces(dt) == norm or dt in text:
            return 'declaration_title'

    # 摘要标题
    if norm in ('摘要', '摘要'):
        return 'abstract_title'
    if text == 'ABSTRACT':
        return 'abstract_en_title'

    # 目录标题
    if norm in ('目录',):
        return 'toc_title'

    # 符号说明标题
    if norm in ('符号说明', '主要符号对照表', '缩略语表'):
        return 'symbol_title'

    # 关键词行
    if text.startswith('关键词') or text.startswith('关 键 词'):
        return 'keywords'
    if text.upper().startswith('KEY WORDS') or text.upper().startswith('KEYWORDS'):
        return 'keywords_en'

    # 特殊一级标题
    for kw in SPECIAL_HEADINGS:
        if text == kw or norm == normalize_spaces(kw):
            return 'special_heading1'

    # 章标题：第X章
    if re.match(r'^第[一二三四五六七八九十百]+章', text):
        return 'heading1'

    # 四级标题：X.Y.Z.W
    if re.match(r'^\d+\.\d+\.\d+\.\d+', text):
        return 'heading4'

    # 三级标题：X.Y.Z
    if re.match(r'^\d+\.\d+\.\d+', text):
        return 'heading3'

    # 二级标题：X.Y（可能无空格，如"2.1引言"）
    if re.match(r'^\d+\.\d+(?!\.\d)', text):
        return 'heading2'

    # 图题/表题
    if re.match(r'^(图|表|Fig|Table)\s*\d', text):
        return 'figure_table_caption'

    # 参考文献条目 [1] 或 1. 格式
    if section_ctx == SECTION_REFERENCE:
        if re.match(r'^\[\d+\]', text) or re.match(r'^\d+[\.\s]', text):
            return 'reference_body'
        return 'reference_body'

    # 封面信息
    if section_ctx == SECTION_COVER:
        for kw in COVER_INFO_KEYWORDS:
            if kw in text:
                return 'cover_info'
        if re.match(r'.*\d{4}\s*年', text) or '月' in text and '日' in text:
            return 'cover_date'

    # 声明页内容
    if section_ctx == SECTION_DECLARATION:
        return 'declaration_body'

    # 符号说明内容
    if section_ctx == SECTION_SYMBOL:
        return 'symbol_body'

    # 致谢内容
    if section_ctx == SECTION_ACK:
        return 'special_body'

    # 个人简历/发表论文
    if section_ctx in (SECTION_PUB, SECTION_RESUME, SECTION_APPENDIX):
        return 'special_body'

    return None


def detect_section_context(text):
    """检测段落是否开始一个新的文档区域"""
    norm = normalize_spaces(text.strip())
    t = text.strip()

    for dt in DECLARATION_TITLES:
        if normalize_spaces(dt) == norm:
            return SECTION_DECLARATION
    if norm in ('摘要',):
        return SECTION_ABSTRACT_CN
    if t == 'ABSTRACT':
        return SECTION_ABSTRACT_EN
    if norm == '目录':
        return SECTION_TOC
    if norm in ('符号说明', '主要符号对照表', '缩略语表'):
        return SECTION_SYMBOL
    if re.match(r'^第[一二三四五六七八九十百]+章', t):
        return SECTION_BODY
    if t == '参考文献':
        return SECTION_REFERENCE
    if t in ('致谢', '致 谢'):
        return SECTION_ACK
    if t in ('附录', '附 录'):
        return SECTION_APPENDIX
    if '攻读学位期间' in t or '发表的学术论文' in t:
        return SECTION_PUB
    if t in ('个人简历', '作者简介'):
        return SECTION_RESUME
    return None


# ============================================================
# 删除多余空行
# ============================================================
def is_empty_paragraph(el):
    """判断元素是否为空段落（无文字内容、无图片、无公式、无分节符）"""
    if el.tag != f'{W}p':
        return False
    pPr = el.find(f'{W}pPr')
    if pPr is not None and pPr.find(f'{W}sectPr') is not None:
        return False
    for t in el.findall(f'.//{W}t'):
        if t.text and t.text.strip():
            return False
    if el.findall(f'.//{W}drawing') or el.findall(f'.//{W}pict'):
        return False
    ns_draw = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
    if el.findall(f'.//{ns_draw}*'):
        return False
    if el.findall(f'.//{MC}AlternateContent'):
        return False
    if el.findall(f'.//{W}object'):
        return False
    return True


def is_table(el):
    return el.tag == f'{W}tbl'


def has_image(el):
    if el.tag != f'{W}p':
        return False
    ns_draw = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
    return bool(el.findall(f'.//{W}drawing') or el.findall(f'.//{W}pict') or
                el.findall(f'.//{ns_draw}*') or el.findall(f'.//{MC}AlternateContent'))


def has_formula(el):
    """段落是否包含公式（OLE对象或oMath）"""
    if el.tag != f'{W}p':
        return False
    M_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
    if el.findall(f'.//{M_NS}oMath') or el.findall(f'.//{M_NS}oMathPara'):
        return True
    if el.findall(f'.//{W}object'):
        return True
    return False


def is_figure_caption(el):
    if el.tag != f'{W}p':
        return False
    texts = []
    for t in el.findall(f'.//{W}t'):
        if t.text:
            texts.append(t.text)
    text = ''.join(texts).strip()
    return bool(re.match(r'^(图|表|Fig|Table)\s*\d', text))


def remove_extra_blank_lines(body):
    """删除正文中多余的空行"""
    children = list(body)
    to_remove = []

    for i, el in enumerate(children):
        if not is_empty_paragraph(el):
            continue
        prev_el = children[i - 1] if i > 0 else None
        next_el = children[i + 1] if i < len(children) - 1 else None
        should_remove = False
        if prev_el is not None and is_table(prev_el):
            should_remove = True
        if next_el is not None and is_table(next_el):
            should_remove = True
        if prev_el is not None and has_image(prev_el):
            should_remove = True
        if next_el is not None and has_image(next_el):
            should_remove = True
        if prev_el is not None and is_figure_caption(prev_el):
            should_remove = True
        if next_el is not None and is_figure_caption(next_el):
            should_remove = True
        if prev_el is not None and is_empty_paragraph(prev_el):
            should_remove = True
        if should_remove:
            to_remove.append(el)

    for el in to_remove:
        body.remove(el)
    return len(to_remove)


# ============================================================
# 格式快照（用于变更报告）
# ============================================================
FONT_SIZE_NAME = {v: k for k, v in FONT_SIZES.items()}


def snapshot_paragraph(p):
    """提取段落当前格式属性，返回 dict"""
    info = {}
    pPr = p.find(f'{W}pPr')

    # 对齐
    if pPr is not None:
        jc = pPr.find(f'{W}jc')
        if jc is not None:
            info['jc'] = jc.get(f'{W}val', '')
        sp = pPr.find(f'{W}spacing')
        if sp is not None:
            info['spacing_line'] = sp.get(f'{W}line', '')
            info['spacing_rule'] = sp.get(f'{W}lineRule', '')

    r = p.find(f'{W}r')
    if r is not None:
        rPr = r.find(f'{W}rPr')
        if rPr is not None:
            rFonts = rPr.find(f'{W}rFonts')
            if rFonts is not None:
                info['eastAsia'] = rFonts.get(f'{W}eastAsia', '')
                info['ascii'] = rFonts.get(f'{W}ascii', '')
            sz = rPr.find(f'{W}sz')
            if sz is not None:
                info['sz'] = sz.get(f'{W}val', '')
            info['bold'] = rPr.find(f'{W}b') is not None
    return info


def diff_format(old, fmt):
    """比较旧快照与目标格式，返回变更列表"""
    changes = []
    JC_NAMES = {'left': '左对齐', 'center': '居中', 'both': '两端对齐', 'right': '右对齐'}

    if old.get('eastAsia') and old['eastAsia'] != fmt['eastAsia']:
        changes.append(f"中文字体 {old['eastAsia']}→{fmt['eastAsia']}")
    if old.get('ascii') and old['ascii'] != fmt['ascii']:
        changes.append(f"西文字体 {old['ascii']}→{fmt['ascii']}")
    if old.get('sz'):
        old_sz = int(old['sz'])
        if old_sz != fmt['sz']:
            old_name = FONT_SIZE_NAME.get(old_sz, f'{old_sz}hp')
            new_name = FONT_SIZE_NAME.get(fmt['sz'], f"{fmt['sz']}hp")
            changes.append(f"字号 {old_name}→{new_name}")
    if 'bold' in old and old['bold'] != fmt['bold']:
        changes.append("加粗→取消" if old['bold'] else "取消→加粗")
    if old.get('jc') and old['jc'] != fmt['jc']:
        changes.append(f"对齐 {JC_NAMES.get(old['jc'], old['jc'])}→{JC_NAMES.get(fmt['jc'], fmt['jc'])}")
    return changes


# ============================================================
# 图表编号检查
# ============================================================
def check_figure_table_numbering(root):
    """扫描所有图题/表题，检查分章编号是否连续正确"""
    chapter = 0
    fig_seq = {}   # chapter -> last seq
    table_seq = {} # chapter -> last seq
    errors = []
    all_captions = []

    for p in root.findall(f'.//{W}p'):
        parent = p.getparent()
        if parent is not None and parent.tag == f'{W}tc':
            continue

        texts = []
        for t in p.findall(f'.//{W}t'):
            if t.text:
                texts.append(t.text)
        text = ''.join(texts).strip()
        if not text:
            continue

        # 检测章节
        m_ch = re.match(r'^第[一二三四五六七八九十百]+章', text)
        if m_ch:
            ch_map = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10}
            ch_text = text[1:text.index('章')]
            if ch_text in ch_map:
                chapter = ch_map[ch_text]
            continue

        # 图题: 图X-Y 或 图X.Y
        m_fig = re.match(r'^图\s*(\d+)[-.](\d+)', text)
        if m_fig:
            ch_num, seq_num = int(m_fig.group(1)), int(m_fig.group(2))
            caption_text = text[:min(len(text), 30)]
            all_captions.append(('图', ch_num, seq_num, caption_text))
            if ch_num != chapter:
                errors.append(f"  ✗ 「{caption_text}」章号={ch_num}，但当前在第{chapter}章")
            expected = fig_seq.get(ch_num, 0) + 1
            if seq_num != expected:
                errors.append(f"  ✗ 「{caption_text}」序号={seq_num}，期望={expected}")
            fig_seq[ch_num] = seq_num
            continue

        # 表题: 表X-Y 或 表X.Y
        m_tbl = re.match(r'^表\s*(\d+)[-.](\d+)', text)
        if m_tbl:
            ch_num, seq_num = int(m_tbl.group(1)), int(m_tbl.group(2))
            caption_text = text[:min(len(text), 30)]
            all_captions.append(('表', ch_num, seq_num, caption_text))
            if ch_num != chapter:
                errors.append(f"  ✗ 「{caption_text}」章号={ch_num}，但当前在第{chapter}章")
            expected = table_seq.get(ch_num, 0) + 1
            if seq_num != expected and not text.startswith('续表'):
                errors.append(f"  ✗ 「{caption_text}」序号={seq_num}，期望={expected}")
            table_seq[ch_num] = seq_num
            continue

    return all_captions, errors


# ============================================================
# 格式应用函数
# ============================================================
def apply_format(p, fmt):
    """将格式 fmt 应用到段落元素 p"""
    pPr = p.find(f'{W}pPr')
    if pPr is None:
        pPr = ET.SubElement(p, f'{W}pPr')
        p.insert(0, pPr)

    for tag in ['spacing', 'ind', 'jc', 'outlineLvl', 'widowControl', 'snapToGrid']:
        for el in pPr.findall(f'{W}{tag}'):
            pPr.remove(el)

    rPr_el = pPr.find(f'{W}rPr')
    pos = list(pPr).index(rPr_el) if rPr_el is not None else len(pPr)

    sp = ET.Element(f'{W}spacing')
    for k, v in fmt['spacing'].items():
        sp.set(f'{W}{k}', v)
    pPr.insert(pos, sp)
    pos += 1

    if fmt.get('indent'):
        ind = ET.Element(f'{W}ind')
        for k, v in fmt['indent'].items():
            ind.set(f'{W}{k}', v)
        pPr.insert(pos, ind)
        pos += 1

    jc = ET.Element(f'{W}jc')
    jc.set(f'{W}val', fmt['jc'])
    pPr.insert(pos, jc)
    pos += 1

    if fmt['outlineLvl'] is not None:
        ol = ET.Element(f'{W}outlineLvl')
        ol.set(f'{W}val', fmt['outlineLvl'])
        pPr.insert(pos, ol)

    for r in p.findall(f'{W}r'):
        rPr = r.find(f'{W}rPr')
        if rPr is None:
            rPr = ET.SubElement(r, f'{W}rPr')
            r.insert(0, rPr)

        rFonts = rPr.find(f'{W}rFonts')
        if rFonts is None:
            rFonts = ET.SubElement(rPr, f'{W}rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(f'{W}ascii', fmt['ascii'])
        rFonts.set(f'{W}eastAsia', fmt['eastAsia'])
        rFonts.set(f'{W}hAnsi', fmt['hAnsi'])
        rFonts.set(f'{W}cs', fmt['cs'])
        if f'{W}hint' in rFonts.attrib:
            del rFonts.attrib[f'{W}hint']

        sz = rPr.find(f'{W}sz')
        if sz is None:
            sz = ET.SubElement(rPr, f'{W}sz')
        sz.set(f'{W}val', str(fmt['sz']))
        szCs = rPr.find(f'{W}szCs')
        if szCs is None:
            szCs = ET.SubElement(rPr, f'{W}szCs')
        szCs.set(f'{W}val', str(fmt['sz']))

        b = rPr.find(f'{W}b')
        bCs = rPr.find(f'{W}bCs')
        if fmt['bold']:
            if b is None:
                ET.SubElement(rPr, f'{W}b')
            if bCs is None:
                ET.SubElement(rPr, f'{W}bCs')
        else:
            if b is not None:
                rPr.remove(b)
            if bCs is not None:
                rPr.remove(bCs)

        for k in rPr.findall(f'{W}kern'):
            rPr.remove(k)

    pPr_rPr = pPr.find(f'{W}rPr')
    if pPr_rPr is not None:
        rFonts = pPr_rPr.find(f'{W}rFonts')
        if rFonts is not None:
            rFonts.set(f'{W}ascii', fmt['ascii'])
            rFonts.set(f'{W}eastAsia', fmt['eastAsia'])
            rFonts.set(f'{W}hAnsi', fmt['hAnsi'])
            rFonts.set(f'{W}cs', fmt['cs'])
        for sz in pPr_rPr.findall(f'{W}sz'):
            sz.set(f'{W}val', str(fmt['sz']))
        for szCs in pPr_rPr.findall(f'{W}szCs'):
            szCs.set(f'{W}val', str(fmt['sz']))
        b = pPr_rPr.find(f'{W}b')
        bCs = pPr_rPr.find(f'{W}bCs')
        if fmt['bold']:
            if b is None:
                ET.SubElement(pPr_rPr, f'{W}b')
            if bCs is None:
                ET.SubElement(pPr_rPr, f'{W}bCs')
        else:
            if b is not None:
                pPr_rPr.remove(b)
            if bCs is not None:
                pPr_rPr.remove(bCs)


def apply_reference_hanging_indent(p, text):
    """参考文献条目的悬挂缩进（批注33）"""
    m = re.match(r'^\[(\d+)\]', text)
    if not m:
        m = re.match(r'^(\d+)[\.\s]', text)
    if not m:
        return

    num = int(m.group(1))
    if num < 10:
        hanging = '340'  # 0.6cm ≈ 340 twips
    elif num < 100:
        hanging = '420'  # 0.74cm ≈ 420 twips
    else:
        hanging = '510'  # 0.9cm ≈ 510 twips

    pPr = p.find(f'{W}pPr')
    if pPr is None:
        return

    for ind_el in pPr.findall(f'{W}ind'):
        pPr.remove(ind_el)

    ind = ET.Element(f'{W}ind')
    ind.set(f'{W}left', hanging)
    ind.set(f'{W}hanging', hanging)

    rPr_el = pPr.find(f'{W}rPr')
    pos = list(pPr).index(rPr_el) if rPr_el is not None else len(pPr)
    sp_el = pPr.find(f'{W}spacing')
    if sp_el is not None:
        pos = list(pPr).index(sp_el) + 1
    pPr.insert(pos, ind)


def fix_page_margins(root):
    """修正页面边距为模板规范值（批注0）"""
    # A4: 左3.17cm 右3.17cm 上2.54cm 下2.54cm 页眉1.5cm 页脚1.75cm
    margins = {
        'top': '1440',      # 2.54cm
        'bottom': '1440',   # 2.54cm
        'left': '1797',     # 3.17cm ≈ 1797 twips (1cm=567twips, 3.17*567=1797.39)
        'right': '1797',    # 3.17cm
        'header': '851',    # 1.5cm ≈ 851 twips
        'footer': '992',    # 1.75cm ≈ 992 twips
        'gutter': '0',
    }
    count = 0
    for pgMar in root.findall(f'.//{W}pgMar'):
        for k, v in margins.items():
            pgMar.set(f'{W}{k}', v)
        count += 1
    return count


def is_toc_paragraph(p):
    """检测是否为目录项段落（通过 pStyle 检测 TOC 样式）"""
    pPr = p.find(f'{W}pPr')
    if pPr is None:
        return None
    pStyle = pPr.find(f'{W}pStyle')
    if pStyle is None:
        return None
    val = pStyle.get(f'{W}val', '')
    if val in ('TOC1', 'TOCHeading', '10', 'a5') or val.lower() in ('toc1',):
        return 'toc1'
    if val in ('TOC2', '20') or val.lower() in ('toc2',):
        return 'toc2'
    if val in ('TOC3', '30') or val.lower() in ('toc3',):
        return 'toc3'
    if val.startswith('TOC') or val.startswith('toc') or val.startswith('目录'):
        return 'toc1'
    return None


def format_footnotes(xml_dir):
    """格式化脚注（批注21）"""
    fn_path = os.path.join(xml_dir, 'footnotes.xml')
    if not os.path.exists(fn_path):
        return 0
    tree = ET.parse(fn_path)
    root = tree.getroot()
    count = 0
    fmt = FORMATS['footnote']
    for p in root.findall(f'.//{W}p'):
        pPr = p.find(f'{W}pPr')
        if pPr is not None:
            pStyle = pPr.find(f'{W}pStyle')
            if pStyle is not None and 'footnote' in pStyle.get(f'{W}val', '').lower():
                apply_format(p, fmt)
                count += 1
    if count > 0:
        tree.write(fn_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    return count


# ============================================================
# 公式自动编号（批注26）
# ============================================================
CN_NUM_MAP = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
              '六': 6, '七': 7, '八': 8, '九': 9, '十': 10}

# Page width 11906 twips - left margin 1797 - right margin 1797 = text area 8312 twips
FORMULA_TAB_CENTER = '4156'  # center of text area
FORMULA_TAB_RIGHT = '8306'   # right edge of text area


def _make_run(text, font_size=24):
    """Create a w:r element with TNR font and given text."""
    r = ET.Element(f'{W}r')
    rPr = ET.SubElement(r, f'{W}rPr')
    rFonts = ET.SubElement(rPr, f'{W}rFonts')
    rFonts.set(f'{W}ascii', 'Times New Roman')
    rFonts.set(f'{W}eastAsia', '宋体')
    rFonts.set(f'{W}hAnsi', 'Times New Roman')
    rFonts.set(f'{W}cs', 'Times New Roman')
    sz = ET.SubElement(rPr, f'{W}sz')
    sz.set(f'{W}val', str(font_size))
    szCs = ET.SubElement(rPr, f'{W}szCs')
    szCs.set(f'{W}val', str(font_size))
    t = ET.SubElement(r, f'{W}t')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r


def _make_tab_run():
    """Create a w:r element containing a tab character."""
    r = ET.Element(f'{W}r')
    ET.SubElement(r, f'{W}tab')
    return r


def number_formulas(root):
    """Add per-chapter numbering (X-Y) to all display-math paragraphs.

    Converts m:oMathPara to inline m:oMath and appends a right-aligned
    tab + equation number using the standard Word tab-stop method.
    Returns dict of {chapter_num: count}.
    """
    body = root.find(f'{W}body')
    if body is None:
        return {}

    chapter_num = 0
    formula_seq = 0
    chapter_formula_counts = {}

    for p in body.iter(f'{W}p'):
        # Track chapter headings
        texts = []
        for t_el in p.findall(f'.//{W}t'):
            if t_el.text:
                texts.append(t_el.text)
        text = ''.join(texts).strip()
        ch_match = re.match(r'^第([一二三四五六七八九十]+)章', text)
        if ch_match:
            chapter_num = CN_NUM_MAP.get(ch_match.group(1), 0)
            formula_seq = 0
            continue

        # Find oMathPara elements in this paragraph
        omp_list = p.findall(f'{M}oMathPara')
        if not omp_list or chapter_num == 0:
            continue

        formula_seq += 1
        eq_label = f'({chapter_num}-{formula_seq})'

        # --- Unwrap oMathPara → oMath ---
        for omp in omp_list:
            omath_elements = omp.findall(f'{M}oMath')
            omp_idx = list(p).index(omp)
            p.remove(omp)
            for i, om in enumerate(omath_elements):
                p.insert(omp_idx + i, om)

        # --- Set up paragraph properties ---
        pPr = p.find(f'{W}pPr')
        if pPr is None:
            pPr = ET.SubElement(p, f'{W}pPr')
            p.insert(0, pPr)

        # Add tab stops (center + right)
        for old_tabs in pPr.findall(f'{W}tabs'):
            pPr.remove(old_tabs)
        tabs = ET.Element(f'{W}tabs')
        tab_center = ET.SubElement(tabs, f'{W}tab')
        tab_center.set(f'{W}val', 'center')
        tab_center.set(f'{W}pos', FORMULA_TAB_CENTER)
        tab_right = ET.SubElement(tabs, f'{W}tab')
        tab_right.set(f'{W}val', 'right')
        tab_right.set(f'{W}pos', FORMULA_TAB_RIGHT)
        pPr.insert(0, tabs)

        # Set spacing: 段前段后6磅(120twips), 1.5倍行距(360)
        for sp_el in pPr.findall(f'{W}spacing'):
            pPr.remove(sp_el)
        spacing = ET.Element(f'{W}spacing')
        spacing.set(f'{W}before', '120')
        spacing.set(f'{W}after', '120')
        spacing.set(f'{W}line', '360')
        spacing.set(f'{W}lineRule', 'auto')
        pPr.append(spacing)

        # Set center alignment
        for jc_el in pPr.findall(f'{W}jc'):
            pPr.remove(jc_el)
        jc = ET.Element(f'{W}jc')
        jc.set(f'{W}val', 'center')
        pPr.append(jc)

        # Remove indent (formulas should not be indented)
        for ind_el in pPr.findall(f'{W}ind'):
            pPr.remove(ind_el)

        # --- Check if there are text runs before the formula ---
        # If the paragraph had text + formula, we keep the text runs
        # and just insert a center-tab before the first oMath
        first_omath = p.find(f'{M}oMath')
        if first_omath is None:
            continue

        has_preceding_text = False
        for child in p:
            if child.tag == f'{M}oMath':
                break
            if child.tag == f'{W}r':
                for t_el in child.findall(f'{W}t'):
                    if t_el.text and t_el.text.strip():
                        has_preceding_text = True
                        break

        if not has_preceding_text:
            # Pure formula paragraph: insert center-tab before oMath
            omath_idx = list(p).index(first_omath)
            p.insert(omath_idx, _make_tab_run())

        # --- Append tab + equation number after the formula ---
        p.append(_make_tab_run())
        p.append(_make_run(eq_label))

        chapter_formula_counts[chapter_num] = chapter_formula_counts.get(chapter_num, 0) + 1

    return chapter_formula_counts


# ============================================================
# 主流程
# ============================================================
def main():
    if len(sys.argv) < 2:
        print("用法: python3 format_thesis.py <document.xml路径>")
        sys.exit(1)

    xml_path = sys.argv[1]
    xml_dir = os.path.dirname(xml_path)
    tree = ET.parse(xml_path)
    root = tree.getroot()
    body = root.find(f'{W}body')

    # Step 1: 删除多余空行
    removed_count = 0
    if body is not None:
        removed_count = remove_extra_blank_lines(body)

    # Step 2: 修正页面边距
    margin_count = fix_page_margins(root)

    # Step 3: 格式化脚注
    fn_count = format_footnotes(xml_dir)

    # Step 3.5: 公式自动编号
    formula_numbered = number_formulas(root)

    # Step 4: 格式化段落（上下文感知）+ 变更跟踪
    stats = {}
    unclassified = []
    change_log = []  # [(text_preview, cat, changes)]
    section_ctx = SECTION_COVER

    for p in root.findall(f'.//{W}p'):
        parent = p.getparent()
        if parent is not None and parent.tag == f'{W}tc':
            continue

        texts = []
        for t in p.findall(f'.//{W}t'):
            if t.text:
                texts.append(t.text)
        text = ''.join(texts).strip()

        if text:
            new_ctx = detect_section_context(text)
            if new_ctx:
                section_ctx = new_ctx

        if has_formula(p) and text:
            old = snapshot_paragraph(p)
            apply_format(p, FORMATS['formula'])
            diffs = diff_format(old, FORMATS['formula'])
            if diffs:
                change_log.append((text[:30], '公式', diffs))
            stats['formula'] = stats.get('formula', 0) + 1
            continue

        toc_level = is_toc_paragraph(p)
        if toc_level and toc_level in FORMATS:
            old = snapshot_paragraph(p)
            apply_format(p, FORMATS[toc_level])
            diffs = diff_format(old, FORMATS[toc_level])
            if diffs:
                change_log.append((text[:30] if text else f'[{toc_level}]', toc_level, diffs))
            stats[toc_level] = stats.get(toc_level, 0) + 1
            continue

        if not text:
            continue

        if section_ctx == SECTION_ABSTRACT_CN:
            cat = classify_paragraph(text, section_ctx)
            if cat is None:
                cat = 'abstract_body'
        elif section_ctx == SECTION_ABSTRACT_EN:
            cat = classify_paragraph(text, section_ctx)
            if cat is None:
                cat = 'abstract_en_body'
        else:
            cat = classify_paragraph(text, section_ctx)

        if cat is None and section_ctx in (SECTION_BODY,):
            if not has_image(p):
                cat = 'body'

        if cat and cat in FORMATS:
            old = snapshot_paragraph(p)
            apply_format(p, FORMATS[cat])
            if cat == 'reference_body':
                apply_reference_hanging_indent(p, text)
            diffs = diff_format(old, FORMATS[cat])
            if diffs:
                change_log.append((text[:30], cat, diffs))
            stats[cat] = stats.get(cat, 0) + 1
        else:
            if text and len(text) > 50:
                text = text[:50] + '...'
            if text:
                unclassified.append(text)

    # Step 5: 检查图表编号
    captions, numbering_errors = check_figure_table_numbering(root)

    tree.write(xml_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    # === 输出报告 ===
    print("=" * 60)
    print("北京邮电大学学位论文格式化完成（完整版）")
    print("=" * 60)
    if removed_count > 0:
        print(f"\n已删除 {removed_count} 个多余空行")
    if margin_count > 0:
        print(f"已修正 {margin_count} 个节的页面边距 (A4, 左右3.17cm, 上下2.54cm)")
    if fn_count > 0:
        print(f"已格式化 {fn_count} 条脚注")
    if formula_numbered:
        total_formulas = sum(formula_numbered.values())
        parts = [f"第{ch}章{cnt}个" for ch, cnt in sorted(formula_numbered.items())]
        print(f"已为 {total_formulas} 个公式自动编号 ({', '.join(parts)})")

    print("\n已格式化的段落统计:")
    label_map = {
        'heading1': '一级标题（第X章）', 'heading2': '二级标题（X.Y）',
        'heading3': '三级标题（X.Y.Z）', 'heading4': '四级标题（X.Y.Z.W）',
        'body': '正文段落', 'abstract_title': '中文摘要标题',
        'abstract_body': '中文摘要内容', 'keywords': '中文关键词',
        'abstract_en_title': '英文摘要标题', 'abstract_en_body': '英文摘要内容',
        'keywords_en': '英文关键词', 'special_heading1': '特殊一级标题',
        'reference_body': '参考文献条目', 'figure_table_caption': '图题/表题',
        'figure_table_note': '图注/表注', 'formula': '公式段落',
        'toc_title': '目录标题', 'toc1': '目录一级', 'toc2': '目录二级', 'toc3': '目录三级',
        'cover_thesis_type': '封面论文类型', 'cover_title': '封面题目',
        'cover_info': '封面个人信息', 'cover_date': '封面日期',
        'declaration_title': '声明页标题', 'declaration_body': '声明页内容',
        'symbol_title': '符号说明标题', 'symbol_body': '符号说明内容',
        'ack_body': '致谢内容', 'special_body': '特殊章节内容',
    }
    for cat, count in sorted(stats.items(), key=lambda x: -x[1]):
        label = label_map.get(cat, cat)
        print(f"  {label}: {count} 个")
    total = sum(stats.values())
    print(f"\n  总计: {total} 个段落已格式化")

    if unclassified:
        print(f"\n未分类的段落 ({len(unclassified)} 个):")
        for t in unclassified[:10]:
            print(f"  · {t}")
        if len(unclassified) > 10:
            print(f"  ... 还有 {len(unclassified) - 10} 个")

    # === 写入变更日志文件 ===
    from datetime import datetime
    log_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(xml_path)
    log_path = os.path.join(log_dir, 'format_changes.log')
    with open(log_path, 'w', encoding='utf-8') as log:
        log.write(f"北京邮电大学学位论文格式化变更报告\n")
        log.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        log.write("=" * 60 + "\n\n")

        # 操作摘要
        log.write("【操作摘要】\n")
        if removed_count > 0:
            log.write(f"  删除多余空行: {removed_count} 个\n")
        if margin_count > 0:
            log.write(f"  修正页面边距: {margin_count} 个节\n")
        if fn_count > 0:
            log.write(f"  格式化脚注: {fn_count} 条\n")
        if formula_numbered:
            total_formulas = sum(formula_numbered.values())
            parts = [f"第{ch}章{cnt}个" for ch, cnt in sorted(formula_numbered.items())]
            log.write(f"  公式自动编号: {total_formulas} 个 ({', '.join(parts)})\n")
        log.write(f"  格式化段落: {total} 个\n\n")

        # 段落统计
        log.write("【段落统计】\n")
        for cat, count in sorted(stats.items(), key=lambda x: -x[1]):
            label = label_map.get(cat, cat)
            log.write(f"  {label}: {count} 个\n")
        log.write(f"  总计: {total} 个\n\n")

        # 变更汇总
        log.write("【格式变更汇总】\n")
        if change_log:
            changed_cats = {}
            for text_preview, cat, diffs in change_log:
                for d in diffs:
                    changed_cats.setdefault(d, 0)
                    changed_cats[d] += 1
            for desc, cnt in sorted(changed_cats.items(), key=lambda x: -x[1]):
                log.write(f"  {desc}: {cnt} 处\n")
        else:
            log.write("  无（所有段落格式已符合规范）\n")
        log.write("\n")

        # 变更明细（完整输出，不截断）
        log.write(f"【格式变更明细】（共 {len(change_log)} 个段落有实际修改）\n")
        for text_preview, cat, diffs in change_log:
            cat_label = label_map.get(cat, cat)
            log.write(f"  [{cat_label}] {text_preview}  ← {', '.join(diffs)}\n")
        log.write("\n")

        # 图表编号检查
        fig_count = sum(1 for c in captions if c[0] == '图')
        tbl_count = sum(1 for c in captions if c[0] == '表')
        log.write(f"【图表编号检查】（共 {fig_count} 个图, {tbl_count} 个表）\n")
        for kind, ch, seq, text_preview in captions:
            log.write(f"  {kind}{ch}-{seq}  {text_preview}\n")
        if numbering_errors:
            log.write(f"\n  ⚠ 发现 {len(numbering_errors)} 个编号问题:\n")
            for err in numbering_errors:
                log.write(f"{err}\n")
        else:
            log.write("  ✓ 所有图表编号连续正确\n")

    # 终端摘要
    fig_count = sum(1 for c in captions if c[0] == '图')
    tbl_count = sum(1 for c in captions if c[0] == '表')
    change_count = len(change_log)
    print(f"\n格式变更: {change_count} 个段落有实际修改")
    print(f"图表编号: {fig_count} 个图, {tbl_count} 个表", end="")
    if numbering_errors:
        print(f"  ⚠ {len(numbering_errors)} 个编号问题")
    else:
        print("  ✓ 编号正确")
    print(f"\n详细变更日志 → {log_path}")


if __name__ == '__main__':
    main()
