import streamlit as st
import docx
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # 1. Проверка полей (допуск 0.2 мм на погрешность Word)
    for i, section in enumerate(doc.sections, start=1):
        margins = {
            'левое': section.left_margin,
            'правое': section.right_margin,
            'верхнее': section.top_margin,
            'нижнее': section.bottom_margin
        }
        for name, val in margins.items():
            if val and abs(val.mm - 20) > 0.2:
                issues.append(f"Раздел {i} - поле '{name}' должно быть 20 мм (текущее: {round(val.mm, 1)})")

    # Вспомогательные функции для безопасного чтения форматирования
    def get_spacing(p):
        ls = p.paragraph_format.line_spacing
        if ls is None: return None
        return ls.pt if isinstance(ls, Pt) else float(ls)
    
    def get_indent(p, attr):
        val = getattr(p.paragraph_format, attr, None)
        return val.cm if val else None

    all_paras = doc.paragraphs
    in_references = False

    for idx, p in enumerate(all_paras):
        text = p.text.strip()
        fmt = p.paragraph_format

        # --- СОДЕРЖАНИЕ ---
        if "СОДЕРЖАНИЕ" in text.upper():
            if text.endswith("."):
                issues.append("Содержание - удалите точку в конце")
            if fmt.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Содержание - выровняйте слово СОДЕРЖАНИЕ по центру")
            continue

        # --- ЗАГОЛОВКИ И НОВЫЕ СТРАНИЦЫ ---
        is_main_section = bool(re.match(r'^\d+\.\s', text)) or text.upper() in ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ"]
        if is_main_section:
            if not fmt.page_break_before:
                issues.append(f"'{text[:40]}...' - должен начинаться с новой страницы (установите 'С новой страницы')")
            
            # Отступы для заголовков и Введения/Заключения
            if get_indent(p, 'first_line_indent') and abs(get_indent(p, 'first_line_indent')) > 0.05:
                issues.append(f"'{text[:40]}...' - уберите отступ первой строки (0 см)")
            if get_indent(p, 'left_indent') and abs(get_indent(p, 'left_indent')) > 0.05:
                issues.append(f"'{text[:40]}...' - отступ слева должен быть 0 см")
            
            sp = get_spacing(p)
            if sp and abs(sp - 1.2) > 0.05:
                issues.append(f"'{text[:40]}...' - междустрочный интервал должен быть 1,2")

        # --- ПУСТЫЕ СТРОКИ ПЕРЕД ПОДРАЗДЕЛАМИ (1.2, 2.1 и т.д.) ---
        if re.match(r'^\d+\.\d+\s', text):
            if idx > 0 and not all_paras[idx-1].text.strip():
                issues.append(f"Подраздел '{text[:40]}...' - удалите пустую строку перед ним")

        # --- РИСУНКИ ---
        if text.startswith("Рисунок"):
            fig_match = re.match(r"Рисунок\s*(\d+)\s*[-–—]?\s*(.*)", text)
            if fig_match:
                fig_num = fig_match.group(1)
                caption = fig_match.group(2).strip()

                if caption.endswith("."):
                    issues.append(f"Рисунок {fig_num} - удалите точку в конце названия")
                if fmt.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                    issues.append(f"Рисунок {fig_num} - выровняйте подпись по центру")
                if fig_num == "5" and caption and caption[0].islower():
                    issues.append("Рисунок 5 - название должно начинаться с заглавной буквы")
                
                # Пустая строка ДО и ПОСЛЕ Рисунок 3
                if fig_num == "3":
                    if idx > 0 and all_paras[idx-1].text.strip():
                        issues.append("Рисунок 3 - добавьте пустую строку ДО рисунка")
                    if idx < len(all_paras)-1 and all_paras[idx+1].text.strip():
                        issues.append("Рисунок 3 - добавьте пустую строку ПОСЛЕ рисунка")

        # --- СПИСОК ИСТОЧНИКОВ ---
        if "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" in text.upper():
            in_references = True
            continue

        if in_references and text and not text.startswith("Рисунок") and not text.startswith("Таблица"):
            if get_indent(p, 'left_indent') and abs(get_indent(p, 'left_indent')) > 0.05:
                issues.append("Список источников - отступ слева должен быть 0 см")
            fi = get_indent(p, 'first_line_indent')
            if not fi or abs(fi - 1.0) > 0.05:
                issues.append("Список источников - отступ первой строки должен быть 1 см")
            sp = get_spacing(p)
            if sp and abs(sp - 1.2) > 0.05:
                issues.append("Список источников - междустрочный интервал должен быть 1,2")
            if fmt.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников - выравнивание должно быть по ширине")

    # --- ТАБЛИЦЫ (только фактические в документе) ---
    for t_idx, table in enumerate(doc.tables, start=1):
        bold_found = False
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if run.bold:
                            bold_found = True
                            break
        if bold_found:
            issues.append(f"Таблица {t_idx} - уберите полужирное начертание")

    # --- РАЗМЕР РИСУНКОВ (проверка выхода за поля) ---
    max_width_cm = 17.0 # 21 (A4) - 2 (левое) - 2 (правое)
    for shape in doc.inline_shapes:
        if shape.type == docx.enum.section.WD_INLINE_SHAPE.PICTURE:
            if shape.width.cm > max_width_cm:
                issues.append(f"Рисунок (встроенное изображение) - уменьшите ширину до {max_width_cm} см, чтобы не выходил за поля (текущая: {round(shape.width.cm, 1)} см)")

    # Очистка дубликатов
    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

# ─────────────────────────────────────────────────────────────
# Streamlit интерфейс
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Нормоконтроль Word", layout="centered")
st.title("📊 Автоматическая проверка документов по ГОСТ/ВУЗ")
st.write("Загрузите `.docx` файл. Скрипт проверит поля, отступы, разрывы страниц, размеры рисунков, таблицы и оформление списка источников.")

uploaded_file = st.file_uploader("Перетащите файл или нажмите для выбора", type=["docx"])

if uploaded_file:
    with st.spinner("🔍 Анализирую структуру, интервалы и разметку..."):
        results = check_word_document(uploaded_file)
    
    st.subheader("Результаты проверки:")
    for res in results:
        st.warning(res) if res.startswith("✅") else st.error(res)
