"""
文档样式设置
"""
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

from .numbering import _setup_numbering


def setup_document_styles(doc, cfg):
    """
    设置文档的所有样式：Normal、Heading 1-6、公式、表标题、图标题、列表 1-4级。
    创建 Word 内置编号定义。
    返回 list_num_id。
    """
    body_cfg = cfg['body']
    page_cfg = cfg['page']

    # 默认样式（从配置读取）
    style = doc.styles['Normal']
    style.font.size = Pt(body_cfg['font_size_pt'])
    style.paragraph_format.line_spacing = body_cfg['line_spacing']
    style.paragraph_format.first_line_indent = Cm(body_cfg['first_line_indent_cm'])
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    rpr = style.element.get_or_add_rPr()
    rFonts = rpr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), body_cfg['font_cn'])
    rFonts.set(qn('w:ascii'), body_cfg['font_en'])
    rFonts.set(qn('w:hAnsi'), body_cfg['font_en'])

    # 配置 Heading 1-6 内置样式（用户可在 Word 中统一修改）
    for h_level in range(1, 7):
        h_key = f'h{h_level}'
        h_cfg = cfg['headings'].get(h_key, {'font_cn': '黑体', 'size_pt': 12, 'bold': True})
        style_name = f'Heading {h_level}'
        try:
            h_style = doc.styles[style_name]
        except KeyError:
            h_style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        h_style.font.size = Pt(h_cfg['size_pt'])
        h_style.font.bold = h_cfg.get('bold', True)
        h_style.font.color.rgb = RGBColor(0, 0, 0)
        h_style.paragraph_format.first_line_indent = Cm(0)
        h_style.paragraph_format.space_before = Pt(0)
        h_style.paragraph_format.space_after = Pt(0)
        h_style.paragraph_format.line_spacing = body_cfg['line_spacing']
        h_rpr = h_style.element.get_or_add_rPr()
        h_rFonts = h_rpr.get_or_add_rFonts()
        h_rFonts.set(qn('w:eastAsia'), h_cfg['font_cn'])
        h_rFonts.set(qn('w:ascii'), body_cfg.get('font_en', 'Times New Roman'))
        h_rFonts.set(qn('w:hAnsi'), body_cfg.get('font_en', 'Times New Roman'))

    # 配置"公式"样式（制表位：居中 + 右对齐）
    avail_cm = 21.0 - page_cfg['left_margin_cm'] - page_cfg['right_margin_cm']
    center_cm = avail_cm / 2
    right_cm = avail_cm
    try:
        formula_style = doc.styles['公式']
    except KeyError:
        formula_style = doc.styles.add_style('公式', WD_STYLE_TYPE.PARAGRAPH)
    formula_style.base_style = doc.styles['Normal']
    formula_style.font.size = Pt(body_cfg['font_size_pt'])
    formula_style.paragraph_format.first_line_indent = Cm(0)
    formula_style.paragraph_format.space_before = Pt(3)
    formula_style.paragraph_format.space_after = Pt(3)
    formula_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    f_rpr = formula_style.element.get_or_add_rPr()
    f_rFonts = f_rpr.get_or_add_rFonts()
    f_rFonts.set(qn('w:eastAsia'), body_cfg['font_cn'])
    f_rFonts.set(qn('w:ascii'), body_cfg.get('font_en', 'Times New Roman'))
    f_rFonts.set(qn('w:hAnsi'), body_cfg.get('font_en', 'Times New Roman'))
    # 制表位写入样式
    f_pPr = formula_style.element.get_or_add_pPr()
    for old_tabs in f_pPr.findall(qn('w:tabs')):
        f_pPr.remove(old_tabs)
    f_tabs = f_pPr.makeelement(qn('w:tabs'), {})
    f_tab_center = f_tabs.makeelement(qn('w:tab'), {
        qn('w:val'): 'center',
        qn('w:pos'): str(int(center_cm * 360000 / 635)),
    })
    f_tabs.append(f_tab_center)
    f_tab_right = f_tabs.makeelement(qn('w:tab'), {
        qn('w:val'): 'right',
        qn('w:pos'): str(int(right_cm * 360000 / 635)),
    })
    f_tabs.append(f_tab_right)
    f_pPr.append(f_tabs)

    # 配置"表标题"样式
    cap_cfg = cfg.get('caption', {})
    cap_font = cap_cfg.get('font_cn', '黑体')
    cap_size = cap_cfg.get('font_size_pt', 11)
    try:
        tbl_cap_style = doc.styles['表标题']
    except KeyError:
        tbl_cap_style = doc.styles.add_style('表标题', WD_STYLE_TYPE.PARAGRAPH)
    tbl_cap_style.base_style = doc.styles['Normal']
    tbl_cap_style.font.size = Pt(cap_size)
    tbl_cap_style.font.bold = True
    tbl_cap_style.font.color.rgb = RGBColor(0, 0, 0)
    tbl_cap_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tbl_cap_style.paragraph_format.first_line_indent = Cm(0)
    tbl_cap_style.paragraph_format.space_before = Pt(6)
    tbl_cap_style.paragraph_format.space_after = Pt(4)
    tc_rpr = tbl_cap_style.element.get_or_add_rPr()
    tc_rFonts = tc_rpr.get_or_add_rFonts()
    tc_rFonts.set(qn('w:eastAsia'), cap_font)
    tc_rFonts.set(qn('w:ascii'), body_cfg.get('font_en', 'Times New Roman'))
    tc_rFonts.set(qn('w:hAnsi'), body_cfg.get('font_en', 'Times New Roman'))

    # 配置"图标题"样式
    try:
        fig_cap_style = doc.styles['图标题']
    except KeyError:
        fig_cap_style = doc.styles.add_style('图标题', WD_STYLE_TYPE.PARAGRAPH)
    fig_cap_style.base_style = doc.styles['Normal']
    fig_cap_style.font.size = Pt(cap_size)
    fig_cap_style.font.bold = True
    fig_cap_style.font.color.rgb = RGBColor(0, 0, 0)
    fig_cap_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fig_cap_style.paragraph_format.first_line_indent = Cm(0)
    fig_cap_style.paragraph_format.space_before = Pt(4)
    fig_cap_style.paragraph_format.space_after = Pt(6)
    fc_rpr = fig_cap_style.element.get_or_add_rPr()
    fc_rFonts = fc_rpr.get_or_add_rFonts()
    fc_rFonts.set(qn('w:eastAsia'), cap_font)
    fc_rFonts.set(qn('w:ascii'), body_cfg.get('font_en', 'Times New Roman'))
    fc_rFonts.set(qn('w:hAnsi'), body_cfg.get('font_en', 'Times New Roman'))

    # 配置"列表 1级"~"列表 4级"样式（缩进由 numbering 定义控制）
    list_cfg = cfg.get('list', {})
    for lv in range(4):
        lv_name = f'列表 {lv+1}级'
        try:
            lv_style = doc.styles[lv_name]
        except KeyError:
            lv_style = doc.styles.add_style(lv_name, WD_STYLE_TYPE.PARAGRAPH)
        lv_style.base_style = doc.styles['Normal']
        lv_style.font.size = Pt(body_cfg['font_size_pt'])
        lv_style.font.color.rgb = RGBColor(0, 0, 0)
        lv_style.paragraph_format.first_line_indent = Cm(0)
        lv_style.paragraph_format.space_before = Pt(1)
        lv_style.paragraph_format.space_after = Pt(1)
        lv_rpr = lv_style.element.get_or_add_rPr()
        lv_rFonts = lv_rpr.get_or_add_rFonts()
        lv_rFonts.set(qn('w:eastAsia'), body_cfg['font_cn'])
        lv_rFonts.set(qn('w:ascii'), body_cfg.get('font_en', 'Times New Roman'))
        lv_rFonts.set(qn('w:hAnsi'), body_cfg.get('font_en', 'Times New Roman'))

    # 创建 Word 内置编号定义
    list_num_id = _setup_numbering(doc, list_cfg.get('style', 'academic'), body_cfg)

    return list_num_id
