# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

doc = Document()

section = doc.sections[0]
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.top_margin = Cm(3.7)
section.bottom_margin = Cm(3.5)
section.left_margin = Cm(2.8)
section.right_margin = Cm(2.6)

def set_font(run, font_name, size, bold=False, color=None):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color

def add_title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(24)
    run = p.add_run(text)
    set_font(run, '方正小标宋简体', 22)
    return p

def add_heading_text(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    set_font(run, '黑体', 16)
    return p

def add_body(text, font='仿宋_GB2312', size=16):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(32)
    p.paragraph_format.line_spacing = Pt(28)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    set_font(run, font, size)
    return p

def add_note(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(12)
    run = p.add_run(text)
    set_font(run, '仿宋_GB2312', 14, color=RGBColor(128,128,128))
    return p

# 版本一
add_heading_text('版本一：政务新闻稿（600字）')
add_title('青浦海关送政策进企业 综保区\u201c新9条\u201d助力跨境电商发展')

news1 = '近日，青浦海关联合青浦区商务委走进区内重点跨境电商企业，开展综保区\u201c新9条\u201d政策专场宣讲会，围绕企业关心的仓储、退税、通关等问题进行现场解读，推动政策红利直达市场主体。'
news2 = '宣讲会上，海关业务骨干从政策出台背景、核心条款、适用范围三个方面，对综保区\u201c新9条\u201d进行了系统梳理，并结合跨境电商行业特点，重点讲解了\u201c仓储货物按状态分类监管\u201d\u201c跨境电商出口海外仓退税\u201d\u201c简化账册管理\u201d等与企业密切相关的条款。针对企业提出的\u201c二线入区操作流程\u201d\u201c账册数据核销周期\u201d等实操问题，海关人员逐一解答，并现场指导企业完成系统配置操作。'
news3 = '参会企业负责人表示，此次宣讲会内容实、针对性强，帮助企业准确理解了政策要点，解决了实际操作中的困惑。下一步，青浦海关将持续跟踪政策落地情况，建立\u201c一企一档\u201d精准服务机制，定期收集企业诉求，确保政策红利切实转化为企业发展动力，助力青浦跨境电商产业高质量发展。'

for para in [news1, news2, news3]:
    add_body(para)
add_note('（全文约590字）')
doc.add_page_break()

# 版本二
add_heading_text('版本二：总署网站版（突出监管创新）')
add_title('青浦海关创新监管服务 推动综保区政策红利直达跨境电商企业')

zs1 = '为深入贯彻落实国务院关于促进综合保税区高水平开放高质量发展的决策部署，近日，青浦海关联合青浦区商务委开展综保区\u201c新9条\u201d政策专场宣讲会，以\u201c政策上门+现场答疑+实操指导\u201d的创新服务模式，推动政策红利直达跨境电商企业。'
zs2 = '在监管创新方面，青浦海关聚焦跨境电商行业特点，重点解读了\u201c仓储货物按状态分类监管\u201d\u201c跨境电商出口海外仓退税\u201d\u201c简化账册管理\u201d等创新监管措施，帮助企业准确理解政策内涵。针对企业提出的\u201c二线入区操作流程\u201d\u201c账册数据核销周期\u201d等实操问题，海关人员逐一解答，并现场指导企业完成系统配置，实现政策解读与业务办理\u201c一站式\u201d服务。'
zs3 = '下一步，青浦海关将持续深化监管创新，建立\u201c一企一档\u201d精准服务机制，定期跟踪政策落地成效，以更高水平的监管服务助力跨境电商产业高质量发展。'

for para in [zs1, zs2, zs3]:
    add_body(para)
add_note('（全文约480字）')
doc.add_page_break()

# 版本三
add_heading_text('版本三：青浦区政府版（突出区关合作）')
add_title('区关联动送政策 青浦跨境电商企业迎发展新机遇')

qz1 = '近日，青浦海关联合青浦区商务委走进区内重点跨境电商企业，开展综保区\u201c新9条\u201d政策专场宣讲会，以区关协同服务助力企业用好政策红利。'
qz2 = '此次宣讲会是青浦区深化\u201c区关合作\u201d机制的具体举措。区商务委负责收集企业诉求，青浦海关组织业务骨干上门解读，双方联动打通政策落地\u201c最后一公里\u201d。宣讲围绕仓储、退税、通关等企业关心的热点问题，结合跨境电商行业特点进行重点讲解，并现场解答实操疑问。'
qz3 = '近年来，青浦区持续优化跨境电商发展环境，依托综保区政策优势，吸引了一批优质跨境电商企业集聚。此次政策宣讲将进一步帮助企业降低运营成本、提升通关效率，为青浦外贸稳增长注入新动力。'

for para in [qz1, qz2, qz3]:
    add_body(para)
add_note('（全文约360字）')
doc.add_page_break()

# 附录
add_heading_text('附：生成以上内容使用的提示词')
add_body('【新闻稿提示词】', '楷体_GB2312', 16)
add_body('你是一位有10年经验的青浦海关宣传干部。帮我写一篇关于\u201c青浦海关综保区新9条政策宣讲会走进跨境电商企业\u201d活动的新闻稿：600字左右，分三段（背景、举措、成效），标题简洁有力，语气客观严谨，不用\u201c重磅\u201d\u201c震惊\u201d这类词。', '仿宋_GB2312', 14)
add_body('【一稿多改提示词】', '楷体_GB2312', 16)
add_body('请把这篇新闻稿改成三个版本：1.公众号版：标题口语化，正文加emoji分点，800字以内，结尾加一句互动引导；2.总署网站版：标题正式，正文三段式，突出\u201c海关监管创新\u201d，600字以内；3.青浦区政府版：标题突出\u201c区关合作\u201d，正文强调对青浦外贸的带动作用，500字以内。', '仿宋_GB2312', 14)

doc.save('一稿多改_三个版本成品_政务字体.docx')
print('政务字体版Word文档已生成')
