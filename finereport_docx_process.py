import os
import re
import win32com.client

from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Inches, Pt, RGBColor

word = win32com.client.Dispatch('Word.Application')
word.visible = False
word.DisplayAlerts = False

root_path = os.getcwd()


class doc_process:

    def __init__(self, text, table, target):
        self.table_path = table
        self.to_file = target
        self.new_doc = Document()
        self.doc = Document(text)

    def run(self):
        self.extract_picture()
        self.extract_text()
        self.merge_docx()
        self.modify()
        return True

    def extract_picture(self):
        """
        按照图片产生的顺序保存图片的_rels，之后通过该值导出图片并回插
        :return: list：按出现顺序排列好的图片rels
        """
        pictures = []
        for rel in self.doc.part._rels:
            if "image" in self.doc.part._rels[rel].target_ref:
                pictures.append((rel, int(re.sub('[^0-9]', ' ', self.doc.part._rels[rel].target_ref))))
        return [_[0] for _ in sorted(pictures, key=lambda x: x[1])]

    def extract_table(self, table):
        """
        筛选出唯一的cell（python.docx中合并的单元格会被视为多个一样的1×1的cell
        :param table: 需要处理的table
        :return: 唯一的cell的index
        """
        row_cells, column_cells = [], []
        index = []
        width, length = len(table.columns), len(table.rows)
        k = 0
        for row in table.rows:
            for cell in row.cells:
                if cell not in row_cells:
                    index.append([k // width, k % width])
                    row_cells.append(cell)
                k += 1
        k = 0
        for column in table.columns:
            for cell in column.cells:
                if cell not in column_cells:
                    column_cells.append(cell)
                elif [k % length, k // length] in index:
                    index.remove([k % length, k // length])
                k += 1
        return index

    def extract_text(self):
        """
        提取出表格中的文本并且写入到新的文档中
        :return:
        """
        image_index = 0
        pictures = self.extract_picture()
        for table in self.doc.tables:
            index = self.extract_table(table)

            for _ in index:
                if not len(table.rows[_[0]].cells[_[1]].text) == 0:
                    for paragraph in table.rows[_[0]].cells[_[1]].paragraphs:
                        # sub-title
                        if any([True if t in paragraph.text[0] else False for t in
                                ['一', '二', '三', '四', '五', '六', '七', '八', '九']]):
                            para_heading = self.new_doc.add_heading("", level=2)
                            run = para_heading.add_run(paragraph.text)

                        # title
                        elif any([True if t in paragraph.text[-4:] else False for t in ['运行月报', '分析月报']]):
                            para_heading = self.new_doc.add_heading("", level=1)
                            run = para_heading.add_run(paragraph.text)

                        # text
                        else:
                            text = paragraph.text.split('。', 1)
                            para = self.new_doc.add_paragraph()
                            if len(text) > 1:
                                run = para.add_run(text)

                            elif '图' in text[0][:2]:
                                run = para.add_run(text[0])
                                with open(root_path + "\\image.png", "wb") as p:
                                    p.write(self.doc.part._rels[pictures[image_index]].target_part.blob)
                                    image_index += 1
                                pic = self.new_doc.add_picture(root_path + "\\image.png", height=Inches(3))

        self.new_doc.add_page_break()
        self.new_doc.save(root_path + '\\temporary.docx')
        return

    def merge_docx(self):
        output = word.Documents.Add()
        output.Application.Selection.InsertFile(root_path + '\\temporary.docx')
        output.Application.Selection.InsertFile(self.table_path)
        output.SaveAs(self.to_file) 
        word.Quit()
        return
    
    def font_setting(self,run,text_level):
        font_color = RGBColor(250,0,0)
        
        if text_level == '标题':
            run.font.name = u'黑体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
            run.font.color.rgb = font_color
            run.font.size = Pt(20)
            run.bold = False
            run.italic = False
            
        elif text_level == '子标题':
            run.font.name = u'黑体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
            run.font.color.rgb = font_color
            run.font.size = Pt(16)
            run.bold = False
            run.italic = False
        
        elif text_level == '开头':
            run.font.name = u'楷体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), u'楷体')
            run.font.color.rgb = font_color
            run.font.size = Pt(16)
            run.bold = True
            run.italic = False
            
        elif text_level == '正文':
            run.font.name = u'仿宋'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), u'仿宋')
            run.font.color.rgb = font_color
            run.font.size = Pt(16)
            run.bold = False
            run.italic = False
        
        elif text_level == '图标题':
            run.font.name = u'黑体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
            run.font.color.rgb = font_color
            run.font.size = Pt(12)
            run.bold = False
            run.italic = False
        return

    def modify(self):
        
        # 文档
        doc = Document(self.to_file)

        # 标题
        doc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.paragraphs[0].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        for run in doc.paragraphs[0].runs: self.font_setting(run,'标题')


        # 全文1.5倍行距
        for para in doc.paragraphs[1:]:
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

            # 图片所在段落
            if len(para.text) == 0:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                continue

            # 各级小标题
            elif any([True if t in para.text[0] else False for t in ['一', '二', '三', '四', '五', '六', '七', '八', '九']]):
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in para.runs: self.font_setting(run,'子标题')


            # 正文
            else:
                text = para.text.split('。', 1)
                # 调节图的语句
                if '图' in text[0][:2]:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs: self.font_setting(run,'图标题')

                
                # 调节正文文本
                elif len(text) > 1:
                    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    # 统一每段的缩进为两格
                    run = para.runs[0]
                    t = re.sub(r"/\s+/", "", run.text)
                    run.text = r"  " + re.sub(r"\xa0", "", t)
                    
                    text = run.text.split("。",1)
                    run.text = text[0]
                    self.font_setting(run,'开头')
                    
                    run = para.add_run(text[1])
                    self.font_setting(run,'正文')
                else:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.save(self.to_file)
        return