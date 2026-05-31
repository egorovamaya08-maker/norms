import streamlit as st
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re


def check_word_document(file):

    doc = docx.Document(file)
    issues = []

    # ---------- ВСПОМОГАТЕЛЬНЫЕ ----------

    def is_empty_paragraph(p):
        return len(p.text.strip()) == 0

    def is_bold_paragraph(p):
        runs = [r for r in p.runs if r.text.strip()]
        if not runs:
            return False
        return all(r.bold for r in runs if r.bold is not None)

    level1_headings = {
        "ВВЕДЕНИЕ",
        "ЗАКЛЮЧЕНИЕ",
        "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"
    }

    def is_level1_heading(text):
        text = text.strip()
        if text in level1_headings:
            return True
        return bool(re.match(r'^\d+\.\s+[А-ЯЁ]', text))

    def is_subsection(text):
        return bool(re.match(r'^\d+\.\d+(\.\d+)?\s+', text))

    def subsection_name(text):
        return re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()

    # ---------- ПОИСК СОДЕРЖАНИЯ ----------

    content_start = None

    for i, p in enumerate(doc.paragraphs):
        if "СОДЕРЖАНИЕ" in p.text.upper():
            content_start = i
            break

    # ---------- СОСТОЯНИЯ ----------

    BEFORE_CONTENT = 0
    IN_CONTENT = 1
    BODY = 2

    state = BEFORE_CONTENT

    # ---------- ОСНОВНОЙ ПРОХОД ----------

    figure_counter = 0
    prev_para_empty = False

    for idx, p in enumerate(doc.paragraphs):

        if content_start is not None and idx < content_start:
            continue

        text = p.text.strip()
        text_upper = text.upper()

        if not text:
            prev_para_empty = True
            continue

        pf = p.paragraph_format

        # ==================================================
        # 1. ВХОД В СОДЕРЖАНИЕ
        # ==================================================

        if text_upper == "СОДЕРЖАНИЕ":
            state = IN_CONTENT

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Содержание – выровняйте по центру")

            if not is_bold_paragraph(p):
                issues.append("Содержание – сделайте полужирным")

            continue

        # ==================================================
        # 2. ОБРАБОТКА СОДЕРЖАНИЯ
        # ==================================================

        if state == IN_CONTENT:

            # выход: первый реальный заголовок
            if is_level1_heading(text) and text_upper != "СОДЕРЖАНИЕ":
                state = BODY

            # ВНУТРИ СОДЕРЖАНИЯ НЕТ ЗАГОЛОВКОВ
            # только форматирование
            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Содержание – выровняйте элементы содержания")

            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append("Содержание – интервал перед абзацем должен быть 0 пт")

            if pf.space_after and pf.space_after.pt > 0.5:
                issues.append("Содержание – интервал после абзаца должен быть 0 пт")

            continue

        # ==================================================
        # 3. BODY (ОСНОВНОЙ ТЕКСТ)
        # ==================================================

        is_heading = (
            is_level1_heading(text)
            or is_subsection(text)
            or text_upper == "СОДЕРЖАНИЕ"
        )

        # ---------- 3.1 Шрифт ----------
        if not text.startswith(("Рисунок", "Таблица")):
            for run in p.runs:
                if run.font.name and run.font.name != "Times New Roman":
                    issues.append(f"«{text[:30]}…» – шрифт должен быть Times New Roman")
                if run.font.size and run.font.size != Pt(14):
                    issues.append(f"«{text[:30]}…» – размер шрифта 14")

        # ---------- 3.2 Межстрочный ----------
        if not is_heading:
            spacing = pf.line_spacing
            if spacing and abs(spacing - 1.2) > 0.1:
                issues.append(f"«{text[:30]}…» – интервал 1.2")

        # ---------- 3.3 Заголовки ----------
        if is_level1_heading(text):

            if not is_bold_paragraph(p):
                issues.append(f"«{text[:30]}…» – заголовок должен быть полужирным")

            if text != text.upper():
                issues.append(f"«{text[:30]}…» – заголовок должен быть ПРОПИСНЫМ")

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:30]}…» – выровнять по центру")

            if pf.first_line_indent and pf.first_line_indent.cm != 0:
                issues.append(f"«{text[:30]}…» – убрать отступ")

            if text.endswith("."):
                issues.append(f"«{text[:30]}…» – убрать точку")

            next_p = doc.paragraphs[idx + 1] if idx + 1 < len(doc.paragraphs) else None
            if not (next_p and next_p.text.strip() == ""):
                issues.append(f"«{text[:30]}…» – нужна пустая строка после заголовка")

        # ---------- 3.4 Подразделы ----------
        if is_subsection(text):

            name = subsection_name(text)

            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                issues.append(f'Подраздел "{name}" – отступ 1.0 см')

            if not is_bold_paragraph(p):
                issues.append(f'Подраздел "{name}" – должен быть полужирным')

            if prev_para_empty:
                # ищем предыдущий непустой
                j = idx - 1
                while j >= 0 and is_empty_paragraph(doc.paragraphs[j]):
                    j -= 1

                if j >= 0 and not is_level1_heading(doc.paragraphs[j].text):
                    issues.append(f'Подраздел "{name}" – лишняя пустая строка перед заголовком')

        # ---------- 3.5 Основной текст ----------
        if not is_heading and not text.startswith("Рисунок"):

            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                issues.append(f'«{text[:30]}…» – отступ 1.0 см')

        # ---------- 3.6 Рисунки ----------
        if text.startswith("Рисунок"):
            figure_counter += 1

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – выравнивание по центру")

            if figure_counter == 3:
                if idx > 0 and not is_empty_paragraph(doc.paragraphs[idx - 1]):
                    issues.append("Рисунок 3 – нужна пустая строка перед рисунком")

        prev_para_empty = False

    # ---------- СПИСОК ИСТОЧНИКОВ ----------

    lit_start = None

    for i, p in enumerate(doc.paragraphs):
        if "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" in p.text.upper():
            lit_start = i
            break

    if lit_start is not None:

        block = [p for p in doc.paragraphs[lit_start + 1:] if p.text.strip()]

        if block:

            pf = block[0].paragraph_format

            if pf.left_indent and abs(pf.left_indent.cm) > 0.1:
                issues.append("Список источников – отступ слева 0 см")

            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                issues.append("Список источников – отступ 1.0 см")

            if block[0].alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – выравнивание по ширине")

    # ---------- ФИНАЛ ----------

    issues = list(dict.fromkeys(issues))

    return issues or ["Ошибок не найдено"]
