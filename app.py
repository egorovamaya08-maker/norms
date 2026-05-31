import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re


def check_word_document(file):

    doc = docx.Document(file)
    issues = []

    # --------------------------------------------------
    # ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
    # --------------------------------------------------

    def is_empty_paragraph(p):
        return not p.text or not p.text.strip()

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

    # --------------------------------------------------
    # ПОИСК СОДЕРЖАНИЯ
    # --------------------------------------------------

    content_idx = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().upper() == "СОДЕРЖАНИЕ":
            content_idx = i
            break

    if content_idx is None:
        content_idx = 0

    # --------------------------------------------------
    # СОСТОЯНИЯ
    # --------------------------------------------------

    STATE_BODY = 0
    STATE_CONTENTS = 1

    state = STATE_BODY

    # --------------------------------------------------
    # ОСНОВНОЙ ПРОХОД
    # --------------------------------------------------

    figure_counter = 0

    for idx, p in enumerate(doc.paragraphs):

        if idx < content_idx:
            continue

        text = p.text.strip()
        text_upper = text.upper()

        if not text:
            continue

        pf = p.paragraph_format

        # --------------------------------------------------
        # ВХОД В СОДЕРЖАНИЕ
        # --------------------------------------------------

        if text_upper == "СОДЕРЖАНИЕ":
            state = STATE_CONTENTS

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Содержание – выровняйте по центру")

            if text.endswith("."):
                issues.append("Содержание – удалите точку в конце")

            if not is_bold_paragraph(p):
                issues.append("Содержание – сделайте полужирным")

            continue

        # --------------------------------------------------
        # ВЫХОД ИЗ СОДЕРЖАНИЯ (СТРОГО ПО ПЕРВОМУ ЗАГОЛОВКУ)
        # --------------------------------------------------

        if state == STATE_CONTENTS:
            if is_level1_heading(text) and text_upper != "СОДЕРЖАНИЕ":
                state = STATE_BODY

        # --------------------------------------------------
        # СОДЕРЖАНИЕ: ЕДИНСТВЕННЫЕ ДОПУСТИМЫЕ ПРОВЕРКИ
        # --------------------------------------------------

        if state == STATE_CONTENTS:

            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Содержание – выровняйте элементы содержания")

            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append("Содержание – интервал перед абзацем должен быть 0 пт")

            if pf.space_after and pf.space_after.pt > 0.5:
                issues.append("Содержание – интервал после абзаца должен быть 0 пт")

            continue

        # --------------------------------------------------
        # ЗАГОЛОВКИ (ТОЛЬКО BODY)
        # --------------------------------------------------

        if is_level1_heading(text):

            if not is_bold_paragraph(p):
                issues.append(f"Раздел {text} – должен быть полужирным")

            if text != text.upper():
                issues.append(f"Раздел {text} – используйте заглавные буквы")

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Раздел {text} – выровняйте по центру")

        # --------------------------------------------------
        # ПОДРАЗДЕЛЫ
        # --------------------------------------------------

        if is_subsection(text):

            name = subsection_name(text)

            if pf.first_line_indent is None or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                issues.append(f'Подраздел "{name}" – отступ 1.0 см')

            if not is_bold_paragraph(p):
                issues.append(f'Подраздел "{name}" – должен быть полужирным')

            # проверка пустой строки перед подразделом
            if idx > 0:

                prev = doc.paragraphs[idx - 1]

                if is_empty_paragraph(prev):

                    j = idx - 2
                    prev_non_empty = None

                    while j >= 0:
                        if not is_empty_paragraph(doc.paragraphs[j]):
                            prev_non_empty = doc.paragraphs[j]
                            break
                        j -= 1

                    if prev_non_empty and not is_level1_heading(prev_non_empty.text):
                        issues.append(
                            f'Подраздел "{name}" – лишняя пустая строка перед заголовком'
                        )

        # --------------------------------------------------
        # ОСНОВНОЙ ТЕКСТ
        # --------------------------------------------------

        is_heading = is_level1_heading(text) or is_subsection(text)

        if not is_heading:

            if pf.first_line_indent is None or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                issues.append(f'«{text[:40]}...» – отступ 1.0 см')

            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append(f'«{text[:40]}...» – лишний интервал перед абзацем')

        # --------------------------------------------------
        # РИСУНКИ
        # --------------------------------------------------

        if text.startswith("Рисунок"):

            figure_counter += 1

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – выровнять по центру")

            if figure_counter == 3:
                if idx > 0 and not is_empty_paragraph(doc.paragraphs[idx - 1]):
                    issues.append("Рисунок 3 – нужна пустая строка перед рисунком")

    # --------------------------------------------------
    # СПИСОК ИСТОЧНИКОВ (ЕДИНЫЙ БЛОК)
    # --------------------------------------------------

    lit_start = None

    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            lit_start = i
            break

    if lit_start is not None:

        block = [p for p in doc.paragraphs[lit_start + 1:] if p.text.strip()]

        if block:

            pf = block[0].paragraph_format

            if pf.left_indent and abs(pf.left_indent.cm) > 0.1:
                issues.append("Список источников – отступ слева 0 см")

            if pf.first_line_indent is None or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                issues.append("Список источников – отступ первой строки 1.0 см")

            if block[0].alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – выравнивание по ширине")

            if pf.line_spacing and abs(pf.line_spacing - 1.2) > 0.1:
                issues.append("Список источников – межстрочный интервал 1.2")

    # --------------------------------------------------
    # ФИНАЛ
    # --------------------------------------------------

    issues = list(dict.fromkeys(issues))

    if not issues:
        return ["Ошибок не найдено"]

    return issues


# --------------------------------------------------
# STREAMLIT UI
# --------------------------------------------------

st.set_page_config(page_title="Проверка Word", layout="centered")

st.title("Проверка документа Word")

file = st.file_uploader("Загрузите .docx", type=["docx"])

if file:
    result = check_word_document(file)

    st.subheader("Результаты")
    for r in result:
        st.write("•", r)
