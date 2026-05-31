import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []

    # ====================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ======================
    def is_empty_paragraph(p):
        return not p.text.strip()

    def is_bold_paragraph(p):
        runs = [r for r in p.runs if r.text.strip()]
        if not runs:
            return False
        return all(r.bold for r in runs)

    def get_text_clean(p):
        return p.text.strip()

    # ====================== ОПРЕДЕЛЕНИЕ ГРАНИЦ ======================
    content_start = None
    content_end = None
    intro_start = None
    lit_start = None

    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip().upper()
        if txt == "СОДЕРЖАНИЕ" and content_start is None:
            content_start = i
        elif txt == "ВВЕДЕНИЕ" and intro_start is None:
            intro_start = i
            if content_start is not None and content_end is None:
                content_end = i - 1
        elif txt == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            lit_start = i
            break

    if content_start is None:
        issues.append("Не найдено заголовка 'СОДЕРЖАНИЕ'")
    if intro_start is None:
        issues.append("Не найдено заголовка 'ВВЕДЕНИЕ'")

    # ====================== ПОЛЯ СТРАНИЦ ======================
    for section in doc.sections:
        left_mm = section.left_margin.pt * 25.4 / 72
        right_mm = section.right_margin.pt * 25.4 / 72
        top_mm = section.top_margin.pt * 25.4 / 72
        bottom_mm = section.bottom_margin.pt * 25.4 / 72

        if any(abs(x - 20) > 0.6 for x in (left_mm, right_mm, top_mm, bottom_mm)):
            issues.append("Поля страниц – установите все поля по 20 мм (±0.5 мм)")
            break

    # ====================== ОСНОВНОЙ ПРОХОД ======================
    figure_counter = 0
    prev_was_empty = False
    in_main_text = False

    for idx, p in enumerate(doc.paragraphs):
        text = get_text_clean(p)
        if not text:
            prev_was_empty = True
            continue

        # Пропускаем титульную часть и содержание
        if content_start is not None and idx <= content_end:
            if idx == content_start:  # только проверяем сам заголовок "СОДЕРЖАНИЕ"
                if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                    issues.append("СОДЕРЖАНИЕ – выровняйте по центру")
                if not is_bold_paragraph(p):
                    issues.append("СОДЕРЖАНИЕ – сделайте полужирным")
                if text.endswith("."):
                    issues.append("СОДЕРЖАНИЕ – удалите точку в конце")
            continue

        # Начало основного текста
        if idx >= intro_start:
            in_main_text = True

        pf = p.paragraph_format

        # ---------- ЗАГОЛОВКИ РАЗДЕЛОВ (уровень 1) ----------
        if text == "ВВЕДЕНИЕ" or text == "ЗАКЛЮЧЕНИЕ" or re.match(r'^\d+\.\s', text):
            title_short = text[:50]

            if text != "ВВЕДЕНИЕ":
                # Проверка новой страницы (упрощённо)
                if idx > 0 and not is_empty_paragraph(doc.paragraphs[idx-1]):
                    issues.append(f"«{title_short}…» – раздел должен начинаться с новой страницы")

            if pf.first_line_indent and abs(pf.first_line_indent.cm) > 0.1:
                issues.append(f"«{title_short}…» – уберите абзацный отступ")

            if not is_bold_paragraph(p):
                issues.append(f"«{title_short}…» – заголовок должен быть полужирным")

            if text != text.upper():
                issues.append(f"«{title_short}…» – заголовок должен быть ПРОПИСНЫМИ буквами")

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{title_short}…» – выровняйте заголовок по центру")

            if text.endswith("."):
                issues.append(f"«{title_short}…» – удалите точку в конце")

            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx+1]):
                issues.append(f"«{title_short}…» – после заголовка должна быть пустая строка")

        # ---------- ПОДРАЗДЕЛЫ ----------
        elif re.match(r'^\d+\.\d+', text):
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s*', '', text).strip()
            title_short = sub_name[:50]

            # Абзацный отступ 1,0 см
            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.25:
                issues.append(f"Подраздел «{title_short}…» – абзацный отступ должен быть 1,0 см")

            if not is_bold_paragraph(p):
                issues.append(f"Подраздел «{title_short}…» – должен быть полужирным")

            # Первая буква заглавная, остальные — строчные (кроме аббревиатур)
            if sub_name and (sub_name[0].islower() or 
                           (len(sub_name) > 3 and any(c.isupper() for c in sub_name[2:]) and not re.search(r'[А-ЯЁ]{2,}', sub_name))):
                issues.append(f"Подраздел «{title_short}…» – первая буква прописная, остальные строчные")

            if text.endswith("."):
                issues.append(f"Подраздел «{title_short}…» – удалите точку")

            if prev_was_empty:
                issues.append(f"Подраздел «{title_short}…» – уберите пустую строку перед подразделом")

        # ---------- ОСНОВНОЙ ТЕКСТ ----------
        elif in_main_text and not text.startswith("Рисунок") and not text.startswith("Таблица"):
            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.25:
                issues.append(f"Абзац «{text[:45]}…» – установите отступ первой строки 1,0 см")

            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append(f"Абзац «{text[:45]}…» – интервал перед должен быть 0 пт")

        # ---------- РИСУНКИ ----------
        if text.startswith("Рисунок"):
            figure_counter += 1
            m = re.match(r'^Рисунок\s+\d+\s*[–-]\s*(.+)$', text)
            title = m.group(1).strip() if m else ""

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – подпись должна быть по центру")

            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – удалите точку в конце")

            if title and title[0].islower():
                issues.append(f"Рисунок {figure_counter} – название должно начинаться с заглавной буквы")

            # Пустые строки до и после
            if idx > 0 and not is_empty_paragraph(doc.paragraphs[idx-1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку перед подписью")
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx+1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку после подписи")

        prev_was_empty = False

    # ====================== ТАБЛИЦЫ (только в основном тексте) ======================
    for t_idx, table in enumerate(doc.tables, start=1):
        # Простая эвристика — таблицы после ВВЕДЕНИЯ
        if intro_start and list(doc.element.body).index(table._element) < list(doc.element.body).index(doc.paragraphs[intro_start]._element):
            continue

        # Поиск подписи над таблицей
        caption = None
        for i in range(len(doc.paragraphs)-1, -1, -1):
            if doc.paragraphs[i].text.strip().startswith("Таблица"):
                caption = doc.paragraphs[i].text.strip()
                break

        if caption:
            if not re.match(r'^Таблица\s+\d+\s+--\s+', caption):
                issues.append(f"Таблица {t_idx} – формат подписи: «Таблица N -- Название»")
            if caption.rstrip().endswith("."):
                issues.append(f"Таблица {t_idx} – удалите точку в конце названия")
        else:
            issues.append(f"Таблица {t_idx} – отсутствует подпись над таблицей")

    # ====================== СПИСОК ИСТОЧНИКОВ ======================
    if lit_start is not None:
        # Первый источник после заголовка
        first_source_idx = None
        for i in range(lit_start + 1, len(doc.paragraphs)):
            if doc.paragraphs[i].text.strip():
                first_source_idx = i
                break

        if first_source_idx is not None:
            p = doc.paragraphs[first_source_idx]
            pf = p.paragraph_format

            if pf.left_indent and abs(pf.left_indent.cm) > 0.1:
                issues.append("Список источников – отступ слева должен быть 0")

            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.2:
                issues.append("Список источников – отступ первой строки 1,0 см")

            # Множитель 1,2
            if pf.line_spacing_rule != WD_LINE_SPACING.MULTIPLE or abs(pf.line_spacing - 1.2) > 0.05:
                issues.append("Список источников – междустрочный интервал 1,2 (множитель)")

            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – выравнивание по ширине")

    # Убираем дубликаты
    issues = list(dict.fromkeys(issues))

    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует требованиям."]
    
    return issues


# ====================== STREAMLIT ======================
st.set_page_config(page_title="Нормоконтроль ВКР", layout="centered")
st.title("📘 Проверка оформления ВКР / дипломной работы")
st.write("Загрузите .docx файл")

uploaded_file = st.file_uploader("Выберите файл", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем документ..."):
        results = check_word_document(uploaded_file)
    
    st.subheader("Результаты проверки:")
    for issue in results:
        st.write(f"• {issue}")
