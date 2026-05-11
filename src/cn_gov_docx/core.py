from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH


class CnGovDocx:
    def __init__(self):
        self.doc = Document()

        self._set_a4_layout()
        self._setup_styles()

    def _set_a4_layout(self):
        section = self.doc.sections[0]

        section.page_width = Cm(21)
        section.page_height = Cm(29.7)

        section.top_margin = Cm(3.7)
        section.bottom_margin = Cm(3.5)
        section.left_margin = Cm(2.8)
        section.right_margin = Cm(2.6)

    def _setup_styles(self):
        styles = self.doc.styles

        # Title
        title = styles["Title"]
        title.font.name = "方正小标宋简体"
        title.font.size = Pt(22)
        title.font.color.rgb = None  # Text 1
        title.paragraph_format.first_line_indent = Pt(0)
        pPr = title.element.find(qn("w:pPr"))
        if pPr is not None:
            pBdr = pPr.find(qn("w:pBdr"))
            if pBdr is not None:
                pPr.remove(pBdr)

        # Heading 1
        h1 = styles["Heading 1"]
        h1.font.name = "黑体"
        h1.font.size = Pt(16)
        h1.font.color.rgb = None  # Text 1
        h1.font.bold = False

        h1.paragraph_format.line_spacing = Pt(28)
        h1.paragraph_format.space_before = Pt(0)
        h1.paragraph_format.space_after = Pt(0)
        h1.paragraph_format.keep_with_next = True
        h1.paragraph_format.keep_together = True
        h1.paragraph_format.first_line_indent = Pt(0)

        # Heading 2
        h2 = styles["Heading 2"]
        h2.font.name = "黑体"
        h2.font.size = Pt(16)
        h2.font.color.rgb = None  # Text 1
        h2.font.bold = False

        h2.paragraph_format.line_spacing = Pt(28)
        h2.paragraph_format.space_before = Pt(0)
        h2.paragraph_format.space_after = Pt(0)
        h2.paragraph_format.keep_with_next = True
        h2.paragraph_format.keep_together = True
        h2.paragraph_format.first_line_indent = Pt(0)

        # Heading 2
        h3 = styles["Heading 3"]
        h3.font.name = "黑体"
        h3.font.size = Pt(16)
        h3.font.color.rgb = None  # Text 1
        h3.font.bold = False

        h3.paragraph_format.line_spacing = Pt(28)
        h3.paragraph_format.space_before = Pt(0)
        h3.paragraph_format.space_after = Pt(0)
        h3.paragraph_format.keep_with_next = True
        h3.paragraph_format.keep_together = True
        h3.paragraph_format.first_line_indent = Pt(0)

        # Normal
        normal = styles["Normal"]
        normal.font.name = "仿宋_GB2312"
        normal.font.size = Pt(16)
        normal.font.color.rgb = None

        normal.paragraph_format.line_spacing = Pt(28)
        normal.paragraph_format.first_line_indent = Pt(32)
        normal.paragraph_format.space_before = Pt(0)
        normal.paragraph_format.space_after = Pt(0)
        normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def _apply_font(self, run, font_name: str):
        run.font.name = font_name
        r = run._element
        r.rPr.rFonts.set(qn("w:eastAsia"), font_name)

    def add_title(self, text: str):
        p = self.doc.add_paragraph(style="Title")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        self._apply_font(run, "方正小标宋简体")

    def add_h1(self, text: str):
        p = self.doc.add_paragraph(style="Heading 1")
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        self._apply_font(run, "黑体")

    def add_h2(self, text: str):
        p = self.doc.add_paragraph(style="Heading 2")
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        self._apply_font(run, "黑体")

    def add_h3(self, text: str):
        p = self.doc.add_paragraph(style="Heading 3")
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        self._apply_font(run, "黑体")

    def add_text(
        self, text: str, indent_first_line: bool = True, remove_newlines: bool = True
    ):
        if remove_newlines:
            text = "".join([x.strip() for x in text.splitlines()])
        p = self.doc.add_paragraph(style="Normal")

        if not indent_first_line:
            p.paragraph_format.first_line_indent = Pt(0)

        run = p.add_run(text)
        self._apply_font(run, "仿宋_GB2312")

    def save(self, filepath: str):
        self.doc.save(filepath)
