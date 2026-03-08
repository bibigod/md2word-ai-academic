"""
列表编号函数
"""
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .elements import set_run_font


def _setup_numbering(doc, list_style, body_cfg=None):
    """
    在文档中创建 Word 内置编号定义（abstractNum + num），返回 numId。
    用户可在 Word/WPS 中直接修改编号格式。
    list_style: 'academic' 或 'circle'
    """
    if body_cfg is None:
        body_cfg = {'font_cn': '仿宋_GB2312', 'font_en': 'Times New Roman'}
    # 获取或创建 numbering part
    numbering_part = doc.part.numbering_part
    numbering_elm = numbering_part._element

    # 编号格式定义
    # numFmt: Word 内置编号格式
    #   'decimal' = 1,2,3  'lowerLetter' = a,b,c
    #   'decimalEnclosedCircleChinese' = ①②③ (Word 内置圈号)
    #   'decimalEnclosedParen' = ⑴⑵⑶ (括号数字)
    if list_style == 'academic':
        # 学术格式: 1）→（1）→ ① → a）
        level_defs = [
            {'numFmt': 'decimal', 'lvlText': '%1）'},
            {'numFmt': 'decimal', 'lvlText': '（%2）'},
            {'numFmt': 'decimalEnclosedCircleChinese', 'lvlText': '%3'},
            {'numFmt': 'lowerLetter', 'lvlText': '%4）'},
        ]
    else:  # circle
        # 圈号格式: ①②③ → ⑴⑵⑶ → a. → 1.
        level_defs = [
            {'numFmt': 'decimalEnclosedCircleChinese', 'lvlText': '%1'},
            {'numFmt': 'decimalEnclosedParen', 'lvlText': '%2'},
            {'numFmt': 'lowerLetter', 'lvlText': '%3.'},
            {'numFmt': 'decimal', 'lvlText': '%4.'},
        ]

    # 查找最大已有 abstractNumId
    existing_abstract = numbering_elm.findall(qn('w:abstractNum'))
    max_abstract_id = -1
    for a in existing_abstract:
        aid = int(a.get(qn('w:abstractNumId'), 0))
        if aid > max_abstract_id:
            max_abstract_id = aid
    abstract_num_id = max_abstract_id + 1

    # 创建 abstractNum
    abstract_num = OxmlElement('w:abstractNum')
    abstract_num.set(qn('w:abstractNumId'), str(abstract_num_id))

    # multiLevelType
    multi = OxmlElement('w:multiLevelType')
    multi.set(qn('w:val'), 'hybridMultilevel')
    abstract_num.append(multi)

    # 缩进参数（单位: cm → twips）
    # left = 文本起始位置, hanging = 编号悬挂量（编号在 left-hanging 处）
    cm_to_twips = lambda cm: int(cm * 360000 / 635)
    base_left_cm = 0.74   # 1级文本缩进
    level_step_cm = 0.55  # 每级递增
    hanging_cm = 0.37     # 编号悬挂（编号宽度）

    for lv in range(4):
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(lv))

        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)

        numFmt = OxmlElement('w:numFmt')
        numFmt.set(qn('w:val'), level_defs[lv]['numFmt'])
        lvl.append(numFmt)

        lvlText = OxmlElement('w:lvlText')
        lvlText.set(qn('w:val'), level_defs[lv]['lvlText'])
        lvl.append(lvlText)

        lvlJc = OxmlElement('w:lvlJc')
        lvlJc.set(qn('w:val'), 'left')
        lvl.append(lvlJc)

        # 缩进: left=文本位置, hanging=编号悬挂
        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        left_cm = base_left_cm + lv * level_step_cm
        ind.set(qn('w:left'), str(cm_to_twips(left_cm)))
        ind.set(qn('w:hanging'), str(cm_to_twips(hanging_cm)))
        pPr.append(ind)
        lvl.append(pPr)

        # 编号字体（确保中文编号使用正确字体）
        rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:eastAsia'), body_cfg['font_cn'])
        rFonts.set(qn('w:ascii'), body_cfg.get('font_en', 'Times New Roman'))
        rFonts.set(qn('w:hAnsi'), body_cfg.get('font_en', 'Times New Roman'))
        rPr.append(rFonts)
        lvl.append(rPr)

        abstract_num.append(lvl)

    # 插入 abstractNum（必须在所有 num 元素之前）
    first_num = numbering_elm.find(qn('w:num'))
    if first_num is not None:
        first_num.addprevious(abstract_num)
    else:
        numbering_elm.append(abstract_num)

    # 创建 num 引用
    existing_nums = numbering_elm.findall(qn('w:num'))
    max_num_id = 0
    for n in existing_nums:
        nid = int(n.get(qn('w:numId'), 0))
        if nid > max_num_id:
            max_num_id = nid
    num_id = max_num_id + 1

    num_elm = OxmlElement('w:num')
    num_elm.set(qn('w:numId'), str(num_id))
    abstractNumId_ref = OxmlElement('w:abstractNumId')
    abstractNumId_ref.set(qn('w:val'), str(abstract_num_id))
    num_elm.append(abstractNumId_ref)
    numbering_elm.append(num_elm)

    return num_id


def _new_num_instance(doc, base_num_id):
    """
    创建一个新的 num 实例（引用同一个 abstractNum），
    使每个独立列表块的编号从1重新开始。
    """
    numbering_elm = doc.part.numbering_part._element

    # 找到 base_num_id 对应的 abstractNumId
    abstract_num_id = None
    for num in numbering_elm.findall(qn('w:num')):
        if int(num.get(qn('w:numId'), 0)) == base_num_id:
            ref = num.find(qn('w:abstractNumId'))
            if ref is not None:
                abstract_num_id = ref.get(qn('w:val'))
            break
    if abstract_num_id is None:
        return base_num_id

    # 找最大 numId
    max_num_id = 0
    for num in numbering_elm.findall(qn('w:num')):
        nid = int(num.get(qn('w:numId'), 0))
        if nid > max_num_id:
            max_num_id = nid
    new_num_id = max_num_id + 1

    # 创建新 num，引用同一个 abstractNum，并添加 lvlOverride 重启编号
    num_elm = OxmlElement('w:num')
    num_elm.set(qn('w:numId'), str(new_num_id))
    abs_ref = OxmlElement('w:abstractNumId')
    abs_ref.set(qn('w:val'), abstract_num_id)
    num_elm.append(abs_ref)

    # 每个 level 都从 1 重新开始
    for lv in range(4):
        override = OxmlElement('w:lvlOverride')
        override.set(qn('w:ilvl'), str(lv))
        start_override = OxmlElement('w:startOverride')
        start_override.set(qn('w:val'), '1')
        override.append(start_override)
        num_elm.append(override)

    numbering_elm.append(num_elm)
    return new_num_id


def make_list_para(doc, text, level, cfg, num_id=None):
    """
    创建列表段落，使用 Word 内置编号（用户可在 Word/WPS 中修改编号格式）。
    使用"列表 N级"样式 + numId/ilvl 引用。
    level: 0-based indent level
    num_id: Word numbering numId
    """
    body_cfg = cfg['body']

    style_name = f'列表 {level+1}级'
    p = doc.add_paragraph(style=style_name)

    # 设置 Word 内置编号引用
    if num_id is not None:
        pPr = p._element.get_or_add_pPr()
        numPr = OxmlElement('w:numPr')
        ilvl = OxmlElement('w:ilvl')
        ilvl.set(qn('w:val'), str(level))
        numPr.append(ilvl)
        numId_elm = OxmlElement('w:numId')
        numId_elm.set(qn('w:val'), str(num_id))
        numPr.append(numId_elm)
        pPr.append(numPr)

    run = p.add_run(text)
    set_run_font(run, body_cfg['font_cn'], body_cfg['font_size_pt'],
                 en_font=body_cfg.get('font_en', 'Times New Roman'))
    return p
