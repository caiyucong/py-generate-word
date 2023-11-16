from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

doc = Document('考核专报模板.docx')
paragraph = doc.add_paragraph()
paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = paragraph.add_run('1-7月份全市新签约项目总体情况表')
run.font.size = Pt(16)
run.font.name = '方正黑体简体'
doc.save('test.docx')
