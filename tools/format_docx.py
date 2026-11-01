"""Титульный лист и разрывы страниц поверх стилей шаблона kisscm."""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


def field(paragraph, instruction):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    text = OxmlElement("w:instrText")
    text.set(qn("xml:space"), "preserve")
    text.text = f" {instruction} "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for item in (begin, text, end):
        run._r.append(item)


def format_document(path: Path, config: dict):
    doc = Document(path)
    # В качестве основы используются размеры страницы, поля, шрифты,
    # интервалы и стили reference-doc из kisscm.
    for name in ("Title", "Subtitle", "Author", "TOC Heading", "Heading 1",
                 "Heading 2", "Heading 3"):
        if name in doc.styles:
            doc.styles[name].font.color.rgb = RGBColor(0, 0, 0)

    # Оглавление остаётся настоящим полем Word, его можно обновить F9.
    first_heading = next(p for p in doc.paragraphs if p.style.name == "Heading 1")
    first_heading.paragraph_format.page_break_before = True
    publication_metadata = True
    for p in doc.paragraphs:
        if p.text.strip() == "Аннотация":
            publication_metadata = False
        if publication_metadata and p.style.name in ("Body Text", "First Paragraph"):
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if p.style.name.startswith("Heading"):
            p.paragraph_format.keep_with_next = True
        if p.style.name == "Bibliography":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.keep_together = True
        if p._p.xpath(".//w:drawing"):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.keep_with_next = True
        if p.style.name == "Image Caption":
            p.paragraph_format.keep_with_next = False
            p.paragraph_format.first_line_indent = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(12)
        if p.text.startswith("Таблица 1"):
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    paragraphs = doc.paragraphs
    for index, p in enumerate(paragraphs):
        if p._p.xpath(".//m:oMathPara"):
            p.paragraph_format.keep_together = True
            if index:
                paragraphs[index-1].paragraph_format.keep_with_next = True

    for table in doc.tables:
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        widths = [Cm(3.6), Cm(4.5), Cm(6.0), Cm(2.6)]
        for column, width in zip(table.columns, widths):
            column.width = width
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border = OxmlElement("w:" + edge)
            for key, value in (("val", "single"), ("sz", "4"), ("color", "D9D9D9")):
                border.set(qn("w:" + key), value)
            borders.append(border)
        properties = table._tbl.tblPr
        existing = properties.find(qn("w:tblBorders"))
        if existing is not None:
            properties.remove(existing)
        properties.append(borders)
        for row_index, row in enumerate(table.rows):
            if row_index == 0:
                repeat = OxmlElement("w:tblHeader")
                row._tr.get_or_add_trPr().append(repeat)
            no_split = OxmlElement("w:cantSplit")
            row._tr.get_or_add_trPr().append(no_split)
            for column_index, cell in enumerate(row.cells):
                cell.width = widths[column_index]
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                cell_properties = cell._tc.get_or_add_tcPr()
                margins = OxmlElement("w:tcMar")
                for edge in ("top", "left", "bottom", "right"):
                    margin = OxmlElement("w:" + edge)
                    margin.set(qn("w:w"), "100")
                    margin.set(qn("w:type"), "dxa")
                    margins.append(margin)
                cell_properties.append(margins)
                shading = OxmlElement("w:shd")
                shading.set(qn("w:fill"), "F2F2F2" if row_index == 0 else "FFFFFF")
                cell_properties.append(shading)
                for p in cell.paragraphs:
                    p.paragraph_format.first_line_indent = Pt(0)
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.15
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if column_index == 0 else WD_ALIGN_PARAGRAPH.CENTER
                    for run in p.runs:
                        run.font.size = Pt(12)
                        run.font.bold = row_index == 0
                        run.font.color.rgb = RGBColor(0, 0, 0)
        following = table._tbl.getnext()
        if following is not None and following.tag == qn("w:p"):
            from docx.text.paragraph import Paragraph
            Paragraph(following, doc._body).paragraph_format.space_before = Pt(6)

    cover = []

    def cover_line(text, style="Normal", before=0, after=0):
        p = doc.add_paragraph(text, style)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.keep_with_next = False
        cover.append(p._p)
        return p

    cover_line("Российский технологический университет «МИРЭА»", before=24, after=70)
    cover_line("Семестровый проект", "Subtitle", after=10)
    cover_line("Документирование и управление версиями программного обеспечения", after=42)
    cover_line(config["title"], "Title", after=44)
    cover_line("Выполнил " + config["author"], after=8)
    cover_line("Группа " + config["group"], after=44)
    cover_line("Москва 2026")
    pagebreak = doc.add_paragraph()
    pagebreak.add_run().add_break(WD_BREAK.PAGE)
    cover.append(pagebreak._p)
    for element in reversed(cover):
        doc._element.body.insert(0, element)

    for section in doc.sections:
        section.different_first_page_header_footer = True
        # Заменяем весь колонтитул, чтобы не дублировать номер из шаблона.
        for element in list(section.footer._element):
            section.footer._element.remove(element)
        footer = section.footer.add_paragraph()
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        field(footer, "PAGE")
        section.first_page_footer.paragraphs[0].clear()
    settings = doc.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")
    doc.core_properties.title = config["title"]
    doc.core_properties.author = config["author"]
    doc.core_properties.subject = "Семестровый проект по документированию"
    doc.core_properties.comments = ""
    doc.save(path)
