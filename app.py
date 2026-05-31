import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []

    # ====================== ВСПОМОГАТЕЛЬНЫЕ ======================
    def is_empty(p):
        return not p.text.strip()

    def is_bold(p):
        runs = [r for r in p.runs if r.text.strip()]
        return bool(runs) and all(r.bold for r in runs)

    def clean_text(p):
        return p.text.strip()

    # ====================== ПОИСК ГРАНИЦ (улучшенный) ======================
    content_start = None
    content_end = None
    intro_idx = None
    conclusion_idx = None
    lit_idx = None

    for i, p in enumerate(doc.paragraphs):
        txt = clean_text(p).upper()
        clean_lower = clean_text(p).lower()

        if not content_start and any(x in txt for x in ["СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"]):
            content_start = i
        elif not intro_idx and txt == "ВВЕДЕНИЕ":
            intro_idx = i
            if content_start is not None and content_end is None:
                content_end = i - 1
        elif not conclusion_idx and txt == "ЗАКЛЮЧЕНИЕ":
            conclusion_idx = i
        elif not lit_idx and "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" in txt:
            lit_idx = i

    if content_start is None:
        issues.append("Не найдено заголовка 'СОДЕРЖАНИЕ' или 'ОГЛАВЛЕНИЕ'")
    if intro_idx is None:
        issues.append("Не найдено заголовка 'ВВЕДЕНИЕ'")

    # ====================== ПОЛЯ ======================
    margins_ok = True
    for section in doc.sections:
        left = round(section.left_margin.pt * 25.4 / 72)
        right = round(section.right_margin.pt * 25.4 / 72)
        top = round(section.top_margin.pt * 25.4 / 72)
        bottom = round(section.bottom_margin.pt * 25.4 / 72)
        if any(x != 20 for x in (left, right, top, bottom)):
            margins_ok = False
            break
    if not margins_ok:
        issues.append("Все поля страниц должны быть 20 мм")

    # ====================== ОСНОВНОЙ ПРОХОД ======================
    figure_counter = 0
    prev_empty = False
    in_main_content = False

    for idx, p in enumerate(doc.paragraphs):
        text = clean_text(p)
        if not text:
            prev_empty = True
            continue

        # Пропускаем всё до конца содержания
        if content_start is not None and idx <= (content_end or content_start + 30):
            if idx == content_start:  # только сам заголовок содержания
                if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                    issues.append("СОДЕРЖАНИЕ / ОГЛАВЛЕНИЕ – выровняйте по центру")
                if not is_bold(p):
                    issues.append("СОДЕРЖАНИЕ / ОГЛАВЛЕНИЕ – сделайте полужирным")
                if text.endswith("."):
                    issues.append("СОДЕРЖАНИЕ / ОГЛАВЛЕНИЕ – удалите точку")
            continue

        # Начинаем основной текст
        if intro_idx and idx >= intro_idx:
            in_main_content = True

        pf = p.paragraph_format

        # ==================== ЗАГОЛОВКИ РАЗДЕЛОВ ====================
        if text.upper() in ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ"] or re.match(r'^\d+\.', text):
            short = text[:60]

            # Новая страница (кроме Введения)
            if text.upper() != "ВВЕДЕНИЕ" and idx > 0 and not is_empty(doc.paragraphs[idx-1]):
                issues.append(f"«{short}…» – раздел должен начинаться с новой страницы")

            if pf.first_line_indent and abs(pf.first_line_indent.cm) > 0.1:
                issues.append(f"«{short}…» – уберите абзацный отступ")

            if not is_bold(p):
                issues.append(f"«{short}…» – заголовок должен быть полужирным")

            if text != text.upper():
                issues.append(f"«{short}…» – заголовок должен быть ПРОПИСНЫМИ буквами")

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{short}…» – выровняйте по центру")

            if text.endswith("."):
                issues.append(f"«{short}…» – удалите точку в конце")

            if idx + 1 < len(doc.paragraphs) and not is_empty(doc.paragraphs[idx+1]):
                issues.append(f"«{short}…» – после заголовка должна быть пустая строка")

        # ==================== ПОДРАЗДЕЛЫ ====================
        elif re.match(r'^\d+\.\d+', text) and in_main_content:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s*', '', text).strip()
            short = sub_name[:50]

            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.2:
                issues.append(f"Подраздел «{short}…» – абзацный отступ 1,0 см")

            if not is_bold(p):
                issues.append(f"Подраздел «{short}…» – должен быть полужирным")

            # Первая заглавная, остальные строчные
            if sub_name and (sub_name[0].islower() or any(c.isupper() for c in sub_name[3:])):
                issues.append(f"Подраздел «{short}…» – первая буква прописная, остальные строчные")

            if text.endswith("."):
                issues.append(f"Подраздел «{short}…» – удалите точку")

            if prev_empty:
                issues.append(f"Подраздел «{short}…» – уберите пустую строку перед подразделом")

        # ==================== ОСНОВНОЙ ТЕКСТ ====================
        elif in_main_content and not text.startswith(("Рисунок", "Таблица")):
            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.2:
                issues.append(f"Абзац «{text[:50]}…» – отступ первой строки 1,0 см")

            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append(f"Абзац «{text[:50]}…» – интервал перед 0 пт")

        # ==================== РИСУНКИ ====================
        if text.startswith("Рисунок"):
            figure_counter += 1
            m = re.match(r'^Рисунок\s+\d+\s*[–-]\s*(.+)$', text)
            title = m.group(1).strip() if m else ""

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – подпись по центру")

            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – удалите точку")

            if title and title[0].islower():
                issues.append(f"Рисунок {figure_counter} – название с большой буквы")

            # Для рисунка 3 — обе пустые строки
            if figure_counter == 3:
                if idx > 0 and not is_empty(doc.paragraphs[idx-1]):
                    issues.append("Рисунок 3 – добавьте пустую строку перед")
                if idx + 1 < len(doc.paragraphs) and not is_empty(doc.paragraphs[idx+1]):
                    issues.append("Рисунок 3 – добавьте пустую строку после")

        prev_empty = False

    # ====================== СПИСОК ИСТОЧНИКОВ ======================
    if lit_idx is not None:
        first_source = None
        for i in range(lit_idx + 1, len(doc.paragraphs)):
            if clean_text(doc.paragraphs[i]):
                first_source = doc.paragraphs[i]
                break

        if first_source:
            p = first_source
            pf = p.paragraph_format

            if pf.left_indent and abs(pf.left_indent.cm) > 0.1:
                issues.append("Список источников – отступ слева 0 см")

            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.2:
                issues.append("Список источников – отступ первой строки 1,0 см")

            if pf.line_spacing_rule != WD_LINE_SPACING.MULTIPLE or abs((pf.line_spacing or 1) - 1.2) > 0.1:
                issues.append("Список источников – междустрочный интервал 1,2 (множитель)")

            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – выравнивание по ширине")

    issues = list(dict.fromkeys(issues))  # убираем дубли

    if not issues:
        return ["✅ Ошибок не найдено."]
    return issues


# ====================== STREAMLIT ======================
st.set_page_config(page_title="Нормоконтроль ВКР", layout="centered")
st.title("📘 Проверка оформления ВКР")
uploaded_file = st.file_uploader("Загрузите .docx файл", type=["docx"])

if uploaded_file:
    with st.spinner("Проверка..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты:")
    for r in results:
        st.write(f"• {r}")
