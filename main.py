from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt


def generate_paragraph(doc, font_size=12, alignment=WD_ALIGN_PARAGRAPH.LEFT, content=''):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.alignment = alignment
    run = paragraph.add_run(content)
    run.font.size = Pt(font_size)
    return paragraph, run


def generate_table(doc, row=1, column=1, left_text='', header=None, left_item=None, style='Table Grid'):
    if left_item is None:
        left_item = []
    if header is None:
        header = []
    # 添加表格
    table = doc.add_table(row, column, style=style)
    # 获取左边列表
    left_content = table.cell(0, 0)
    left_content.text = left_text
    # 垂直居中
    left_content.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # 左右居中
    left_content.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # 合同单元格
    left_content.merge(table.cell(row - 1, 0))
    # 生成列头
    header_row = table.rows[0]
    for index, value in enumerate(header):
        item = header_row.cells[index + 1]
        item.text = value
        item.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        item.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # 生成行标题
    for index, value in enumerate(left_item):
        item = table.rows[index + 1].cells[1]
        item.text = value
        item.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        item.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    return table


if __name__ == '__main__':
    document = Document('考核专报模板.docx')
    generate_paragraph(document, 16, WD_ALIGN_PARAGRAPH.CENTER, '1-7月份全市新签约项目总体情况表')
    generate_paragraph(document, 12, WD_ALIGN_PARAGRAPH.RIGHT, '单位：个')
    generate_table(document, 4, 6, '项目\n总体情况', ['', '新签约\n项目', '其中：\n制造业', '其中：\n服务业', '总投资\n（亿元）'],
                   ['项目数', '新引进项目', '再投资项目'])
    generate_table(document, 2, 6, '重大项目', ['项目数', '其中：\n制造业', '其中：\n服务业', '其中：固投\n20-50亿元', '其中：固投\n50亿元以上'])
    document.add_paragraph()
    generate_table(document, 2, 6, '优强项目', ['投资主体', '境内外\n500强\n含子公司', '上市公司', '独角兽或瞪羚企业', '专精特新企业\n（国家级）'])
    document.add_paragraph()
    generate_table(document, 7, 6, '项目来源地', ['', '个数', '个数占比', '总投资\n（亿元）', '投资占比'],
                   ['长三角', '珠三角', '京津冀', '省内', '境外（外资）', '其他'])
    document.save('test.docx')
