"""
Модуль сборки корпоративных инструкций DOCX.
Построен на базе python-docx с соблюдением фирменных стандартов оформления:
- Четкая типографика (шрифты Calibri/Arial, строгая иерархия заголовков)
- Корпоративные врезки (Callout boxes) для примечаний и предупреждений
- Форматированные таблицы с контрастными заголовками
- Иллюстрации по центру с подписями
- Пошаговые алгоритмы действий
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn


class DocxBuilder:
    # Цветовая палитра
    COLOR_PRIMARY = RGBColor(0x1F, 0x4E, 0x79)     # Темно-синий
    COLOR_SECONDARY = RGBColor(0x2E, 0x75, 0xB6)   # Синий акцент
    COLOR_DARK = RGBColor(0x26, 0x26, 0x26)        # Темно-серый для текста
    COLOR_MUTED = RGBColor(0x59, 0x59, 0x59)       # Серый для подписей
    COLOR_WARNING = RGBColor(0xC0, 0x00, 0x00)     # Красный для предупреждений

    HEX_PRIMARY = "1F4E79"
    HEX_SECONDARY = "2E75B6"
    HEX_LIGHT_BG = "F2F5F8"
    HEX_WARNING_BG = "FFF3F3"
    HEX_WARNING_BORDER = "C00000"
    HEX_TIP_BG = "F0F9F4"
    HEX_TIP_BORDER = "28A745"
    HEX_BORDER_GREY = "D9D9D9"

    def __init__(self, template_path="templates/template.docx"):
        if os.path.exists(template_path):
            self.doc = docx.Document(template_path)
        else:
            self.doc = docx.Document()
            self._setup_page()

    def _setup_page(self):
        section = self.doc.sections[0]
        section.page_width = Inches(8.27)   # A4
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    def add_title_header(self, title, subtitle=None, metadata=None):
        """Создает шапку документа с названием, подзаголовком и метаданными."""
        # Направление / Категория (если указана в metadata)
        if metadata and "Направление" in metadata:
            cat_p = self.doc.add_paragraph()
            cat_p.paragraph_format.space_before = Pt(0)
            cat_p.paragraph_format.space_after = Pt(4)
            cat_run = cat_p.add_run(metadata["Направление"].upper())
            cat_run.font.name = "Calibri"
            cat_run.font.size = Pt(9.5)
            cat_run.font.bold = True
            cat_run.font.color.rgb = self.COLOR_SECONDARY

        # Основной заголовок
        title_p = self.doc.add_paragraph()
        title_p.paragraph_format.space_before = Pt(4)
        title_p.paragraph_format.space_after = Pt(6)
        title_run = title_p.add_run(title)
        title_run.font.name = "Calibri"
        title_run.font.size = Pt(20)
        title_run.font.bold = True
        title_run.font.color.rgb = self.COLOR_PRIMARY

        # Подзаголовок
        if subtitle:
            sub_p = self.doc.add_paragraph()
            sub_p.paragraph_format.space_before = Pt(0)
            sub_p.paragraph_format.space_after = Pt(12)
            sub_run = sub_p.add_run(subtitle)
            sub_run.font.name = "Calibri"
            sub_run.font.size = Pt(12)
            sub_run.font.color.rgb = self.COLOR_MUTED

        # Таблица метаданных
        if metadata:
            meta_items = [(k, v) for k, v in metadata.items() if k != "Направление"]
            if meta_items:
                table = self.doc.add_table(rows=len(meta_items), cols=2)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                for i, (k, v) in enumerate(meta_items):
                    row = table.rows[i]
                    # Ячейка ключа
                    cell_k = row.cells[0]
                    cell_k.width = Inches(1.8)
                    p_k = cell_k.paragraphs[0]
                    p_k.paragraph_format.space_after = Pt(2)
                    p_k.paragraph_format.space_before = Pt(2)
                    r_k = p_k.add_run(k + ":")
                    r_k.font.name = "Calibri"
                    r_k.font.size = Pt(9.5)
                    r_k.font.bold = True
                    r_k.font.color.rgb = self.COLOR_MUTED

                    # Ячейка значения
                    cell_v = row.cells[1]
                    cell_v.width = Inches(4.8)
                    p_v = cell_v.paragraphs[0]
                    p_v.paragraph_format.space_after = Pt(2)
                    p_v.paragraph_format.space_before = Pt(2)
                    r_v = p_v.add_run(v)
                    r_v.font.name = "Calibri"
                    r_v.font.size = Pt(9.5)
                    r_v.font.color.rgb = self.COLOR_DARK

                # Разделительная черта после метаданных
                div_p = self.doc.add_paragraph()
                div_p.paragraph_format.space_before = Pt(8)
                div_p.paragraph_format.space_after = Pt(16)
                pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="4" w:color="{self.HEX_BORDER_GREY}"/></w:pBdr>')
                div_p._p.get_or_add_pPr().append(pBdr)

    def add_heading_1(self, text):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Calibri"
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = self.COLOR_PRIMARY
        return p

    def add_heading_2(self, text):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Calibri"
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = self.COLOR_SECONDARY
        return p

    def add_heading_3(self, text):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = self.COLOR_DARK
        return p

    def add_paragraph(self, text, bold_prefix=None):
        p = self.doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(6)
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = "Calibri"
            r_pre.font.size = Pt(11)
            r_pre.font.bold = True
            r_pre.font.color.rgb = self.COLOR_DARK
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(11)
        r.font.color.rgb = self.COLOR_DARK
        return p

    def add_bullet(self, text, bold_prefix=None, level=0):
        p = self.doc.add_paragraph(style='List Bullet' if level == 0 else 'List Bullet 2')
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = "Calibri"
            r_pre.font.size = Pt(11)
            r_pre.font.bold = True
            r_pre.font.color.rgb = self.COLOR_DARK
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(11)
        r.font.color.rgb = self.COLOR_DARK
        return p

    def add_step(self, step_num, title, description=None, sub_steps=None):
        """Форматирует структурированный шаг инструкции."""
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True

        r_badge = p.add_run(f"Шаг {step_num}. ")
        r_badge.font.name = "Calibri"
        r_badge.font.size = Pt(11.5)
        r_badge.font.bold = True
        r_badge.font.color.rgb = self.COLOR_PRIMARY

        r_title = p.add_run(title)
        r_title.font.name = "Calibri"
        r_title.font.size = Pt(11.5)
        r_title.font.bold = True
        r_title.font.color.rgb = self.COLOR_DARK

        if description:
            self.add_paragraph(description)

        if sub_steps:
            for sub in sub_steps:
                if isinstance(sub, tuple):
                    self.add_bullet(sub[1], bold_prefix=sub[0] + " ")
                else:
                    self.add_bullet(sub)

    def add_callout(self, text, title="Примечание", kind="info"):
        """
        Создает корпоративную врезку (Callout box).
        kind: 'info' (синяя), 'warning' (красная), 'tip' (зеленая)
        """
        if kind == "warning":
            bg_hex = self.HEX_WARNING_BG
            border_hex = self.HEX_WARNING_BORDER
            title_color = self.COLOR_WARNING
        elif kind == "tip":
            bg_hex = self.HEX_TIP_BG
            border_hex = self.HEX_TIP_BORDER
            title_color = RGBColor(0x28, 0xA7, 0x45)
        else:  # info
            bg_hex = self.HEX_LIGHT_BG
            border_hex = self.HEX_SECONDARY
            title_color = self.COLOR_SECONDARY

        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        cell = table.cell(0, 0)
        cell.width = Inches(6.77)
        tcPr = cell._tc.get_or_add_tcPr()

        # Заливка фона
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_hex}"/>')
        tcPr.append(shd)

        # Левая граница - акцентная, остальные скрыты
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="none"/>
                <w:left w:val="single" w:sz="28" w:space="0" w:color="{border_hex}"/>
                <w:bottom w:val="none"/>
                <w:right w:val="none"/>
            </w:tcBorders>
        ''')
        tcPr.append(borders)

        # Внутренние отступы (padding)
        mar = parse_xml(f'''
            <w:tcMar {nsdecls("w")}>
                <w:top w:w="140" w:type="dxa"/>
                <w:left w:w="220" w:type="dxa"/>
                <w:bottom w:w="140" w:type="dxa"/>
                <w:right w:w="220" w:type="dxa"/>
            </w:tcMar>
        ''')
        tcPr.append(mar)

        # Контент
        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)

        if title:
            r_title = p.add_run(f"{title}: ")
            r_title.font.name = "Calibri"
            r_title.font.size = Pt(10.5)
            r_title.font.bold = True
            r_title.font.color.rgb = title_color

        r_text = p.add_run(text)
        r_text.font.name = "Calibri"
        r_text.font.size = Pt(10.5)
        r_text.font.color.rgb = self.COLOR_DARK

        # Отступ после врезки
        spacer = self.doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(0)
        spacer.paragraph_format.space_after = Pt(6)

    def add_image(self, image_path, caption=None, width_inches=5.8):
        """Вставляет изображение по центру с аккуратной подписью."""
        if not os.path.exists(image_path):
            print(f"Warning: Image not found: {image_path}")
            return

        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True

        run = p.add_run()
        run.add_picture(image_path, width=Inches(width_inches))

        if caption:
            cap_p = self.doc.add_paragraph()
            cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap_p.paragraph_format.space_before = Pt(2)
            cap_p.paragraph_format.space_after = Pt(10)
            cap_run = cap_p.add_run(caption)
            cap_run.font.name = "Calibri"
            cap_run.font.size = Pt(9.5)
            cap_run.font.italic = True
            cap_run.font.color.rgb = self.COLOR_MUTED

    def add_table(self, headers, rows, col_widths=None):
        """Создает стилизованную корпоративную таблицу."""
        table = self.doc.add_table(rows=len(rows) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        # Шапка таблицы
        hdr_row = table.rows[0]
        hdr_row._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

        for idx, text in enumerate(headers):
            cell = hdr_row.cells[idx]
            tcPr = cell._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{self.HEX_PRIMARY}"/>'))
            tcPr.append(parse_xml(f'''
                <w:tcMar {nsdecls("w")}>
                    <w:top w:w="120" w:type="dxa"/>
                    <w:left w:w="160" w:type="dxa"/>
                    <w:bottom w:w="120" w:type="dxa"/>
                    <w:right w:w="160" w:type="dxa"/>
                </w:tcMar>
            '''))
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.name = "Calibri"
            r.font.size = Pt(10)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        # Строки данных
        for r_idx, row_data in enumerate(rows):
            row = table.rows[r_idx + 1]
            bg_color = self.HEX_LIGHT_BG if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, val in enumerate(row_data):
                cell = row.cells[c_idx]
                tcPr = cell._tc.get_or_add_tcPr()
                if bg_color != "FFFFFF":
                    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>'))
                tcPr.append(parse_xml(f'''
                    <w:tcBorders {nsdecls("w")}>
                        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{self.HEX_BORDER_GREY}"/>
                        <w:top w:val="none"/>
                        <w:left w:val="none"/>
                        <w:right w:val="none"/>
                    </w:tcBorders>
                '''))
                tcPr.append(parse_xml(f'''
                    <w:tcMar {nsdecls("w")}>
                        <w:top w:w="100" w:type="dxa"/>
                        <w:left w:w="160" w:type="dxa"/>
                        <w:bottom w:w="100" w:type="dxa"/>
                        <w:right w:w="160" w:type="dxa"/>
                    </w:tcMar>
                '''))
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.1
                r = p.add_run(str(val))
                r.font.name = "Calibri"
                r.font.size = Pt(9.5)
                r.font.color.rgb = self.COLOR_DARK

        # Ширины колонок
        if col_widths:
            for row in table.rows:
                for c_idx, w in enumerate(col_widths):
                    row.cells[c_idx].width = Inches(w)

        # Отступ после таблицы
        spacer = self.doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(0)
        spacer.paragraph_format.space_after = Pt(8)

    def save(self, output_path):
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        self.doc.save(output_path)
        print(f"Document saved: {output_path}")
