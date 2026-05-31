import streamlit as st
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re


# ==================================================
# ОСНОВНАЯ ФУНКЦИЯ ПРОВЕРКИ
# ==================================================

def check_word_document(file):

    doc = docx.Document(file)
    issues = []

    # -----------------------------
    # ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
    # -----------------------------

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

    # -----------------------------
    # ПОИСК СОДЕРЖАНИЯ
    # -----------------------------

    content_start = None

    for i, p in enumerate(doc.paragraphs):
        if "СОДЕРЖАНИЕ" in p.text.upper():
            content_start = i
            break

    BEFORE = 0
    CONTENT = 1
    BODY = 2

    state = BEFORE

    figure_counter = 0
    prev_empty = False

    # -----------------------------
    # ОСНОВНОЙ ПРОХОД
    # -----------------------------

    for idx, p in enumerate(doc.paragraphs):

        if content_start is not None and idx < content_start:
            continue

        text = p.text.strip()

        if not text:
            prev_empty = True
            continue

        text_upper = text.upper()
        pf = p.paragraph_format

        # =====================================
        # 1. ВХОД В СОДЕРЖАНИЕ
        # =====================================

        if text_upper == "СОДЕРЖАНИЕ":
            state = CONTENT

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Содержание – выровнять по центру")

            if not is_bold_paragraph(p):
                issues.append("Содержание – сделать полужирным")

            continue

        # =====================================
        # 2. ОБРАБОТКА СОДЕРЖАНИЯ
        # =====================================

        if state == CONTENT:

            # выход из содержания
            if is_level1_heading(text) and text_upper != "СОДЕРЖАНИЕ":
                state = BODY

            # ВАЖНО: внутри содержания НЕТ заголовков
            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Содержание – выровнять элементы")

            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append("Содержание – убрать интервал перед")

            if pf.space_after and pf.space_after.pt > 0.5:
                issues.append("Содержание – убрать интервал после")

            continue

        # =====================================
        # 3. ЗАГОЛОВКИ РАЗДЕЛОВ
        # =====================================

        if is_level1_heading(text):

            if not is_bold_paragraph(p):
                issues.append(f"«{text[:30]}…» – заголовок должен быть полужирным")

            if text != text.upper():
                issues.append(f"«{text[:30]}…» – прописные буквы")

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:30]}…» – по центру")

            if pf.first_line_indent and pf.first_line_indent.cm != 0:
                issues.append(f"«{text[:30]}…» – убрать отступ")

        # =====================================
        # 4. ПОДРАЗДЕЛЫ
        # =====================================

        if is_subsection(text):

            name = subsection_name(text)

            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                issues.append(f'Подраздел "{name}" – отступ 1.0 см')

            if prev_empty:
                issues.append(f'Подраздел "{name}" – лишняя пустая строка')

        # =====================================
        # 5. ОСНОВНОЙ ТЕКСТ
        # =====================================

        is_heading = is_level1_heading(text) or is_subsection(text)

        if not is_heading and not text.startswith("Рисунок"):

            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                issues.append(f'«{text[:30]}…» – отступ 1.0 см')

        # =====================================
        # 6. РИСУНКИ
        # =====================================

        if text.startswith("Рисунок"):

            figure_counter += 1

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – по центру")

            if figure_counter == 3:
                if idx > 0 and not is_empty_paragraph(doc.paragraphs[idx - 1]):
                    issues.append("Рисунок 3 – пустая строка перед")

        prev_empty = False

    # =====================================
    # 7. СПИСОК ИСТОЧНИКОВ
    # =====================================

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
                issues.append("Список источников – отступ 0 см")

            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                issues.append("Список источников – отступ 1.0 см")

            if block[0].alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – по ширине")

    return issues


# ==================================================
# STREAMLIT ИНТЕРФЕЙС (ВАЖНО: СНАРУЖИ ФУНКЦИИ)
# ==================================================

st.set_page_config(page_title="Нормоконтроль", layout="centered")

st.title("📊 Проверка Word-документа")

uploaded_file = st.file_uploader("Загрузите .docx файл", type=["docx"])

if uploaded_file is not None:

    with st.spinner("Проверка..."):
        result = check_word_document(uploaded_file)

    st.subheader("Результаты")

    if result:
        for r in result:
            st.write("•", r)
    else:
        st.success("Ошибок не найдено")
