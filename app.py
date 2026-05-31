import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
import re

def mm_to_emu(mm):
    return Cm(mm/10).emu

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # ---------- 1. ПОЛЯ СТРАНИЦ ----------
    for i, section in enumerate(doc.sections, start=1):
        left_mm = section.left_margin.pt * 25.4 / 72 if section.left_margin else 0
        right_mm = section.right_margin.pt * 25.4 / 72 if section.right_margin else 0
        top_mm = section.top_margin.pt * 25.4 / 72 if section.top_margin else 0
        bottom_mm = section.bottom_margin.pt * 25.4 / 72 if section.bottom_margin else 0
        
        if (abs(left_mm - 20) > 0.5 or abs(right_mm - 20) > 0.5 or 
            abs(top_mm - 20) > 0.5 or abs(bottom_mm - 20) > 0.5):
            issues.append(f"Раздел {i} – установите поля: левое 20, правое 20, верх 20, низ 20 мм (сейчас {left_mm:.1f}, {right_mm:.1f}, {top_mm:.1f}, {bottom_mm:.1f})")

    # ---------- 2. ПРОВЕРКА АБЗАЦЕВ ----------
    in_special_block = False      # внутри содержания или списка литературы
    figure_count = 0
    prev_para_empty = False
    prev_para_was_figure = False
    # Для проверки пустых строк перед подразделами
    level1_headings_pattern = re.compile(r'^\d+\.\s+[А-Я]')  # "1. ТЕКСТ", "2. МЕТОДИКА"
    subsection_pattern = re.compile(r'^\d+\.\d+\s+')        # "1.1 Эволюция"
    
    # Получим список всех элементов документа (абзацы + таблицы) для проверки подписей таблиц
    all_elements = []
    for child in doc.element.body:
        if child.tag.endswith('p'):
            all_elements.append(('paragraph', child))
        elif child.tag.endswith('tbl'):
            all_elements.append(('table', child))
    
    # Функция, чтобы проверить, идёт ли за абзацем таблица
    def next_element_is_table(idx, elements):
        if idx + 1 < len(elements):
            return elements[idx+1][0] == 'table'
        return False
    
    for idx, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            prev_para_empty = True
            continue
        
        p_format = p.paragraph_format
        style_name = p.style.name if p.style else ""
        
        # ------ 2.1 Пропускаем проверки для специальных блоков ------
        if "СОДЕРЖАНИЕ" in text.upper():
            in_special_block = True
        if in_special_block and ("СПИСОК" in text.upper() or "ЗАКЛЮЧЕНИЕ" in text.upper()):
            in_special_block = False
        
        # ------ 2.2 Пустые строки перед подразделами (только для 1.1, 1.2, 2.1 и т.д., не для основных разделов) ------
        if subsection_pattern.match(text) and prev_para_empty:
            issues.append(f"{text[:40]}… – уберите пустую строку перед подразделом")
        
        # ------ 2.3 Абзацный отступ (1,0 см только для основного текста) ------
        indent = p_format.first_line_indent
        ignore_indent = (text.isupper() or 
                         text.startswith(("Рисунок", "Таблица")) or 
                         in_special_block or
                         "СОДЕРЖАНИЕ" in text.upper() or
                         "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" in text.upper())
        if indent and not ignore_indent:
            indent_cm = indent.cm
            if abs(indent_cm - 1.0) > 0.1:
                issues.append(f"«{text[:30]}…» – установите абзацный отступ 1,0 см (сейчас {indent_cm:.1f} см)")
        
        # ------ 2.4 Междустрочный интервал 1,2 ------
        spacing = p_format.line_spacing
        if spacing and not ignore_indent and not in_special_block:
            if p_format.line_spacing_rule == 3:  # MULTIPLE
                if abs(spacing - 1.2) > 0.05:
                    issues.append(f"«{text[:30]}…» – замените интервал на 1,2 (сейчас {spacing:.1f})")
            else:
                issues.append(f"«{text[:30]}…» – установите множитель междустрочного интервала 1,2")
        
        # ------ 2.5 Шрифт и начертание ------
        for run in p.runs:
            if run.font.name and run.font.name != "Times New Roman":
                issues.append(f"«{text[:30]}…» – смените шрифт на Times New Roman")
            if run.font.size and run.font.size != Pt(14):
                issues.append(f"«{text[:30]}…» – установите размер шрифта 14")
            if run.underline:
                issues.append(f"«{text[:30]}…» – удалите подчеркивания")
        
        # ------ 2.6 СОДЕРЖАНИЕ ------
        if "СОДЕРЖАНИЕ" in text.upper():
            if text.endswith("."):
                issues.append("Содержание – удалите точку в конце")
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
        
        # ------ 2.7 Рисунки ------
        if text.startswith("Рисунок"):
            figure_count += 1
            # Удаление точки в конце названия
            if text.rstrip().endswith("."):
                issues.append(f"Рисунок {figure_count} – удалите точку в конце названия")
            # Выравнивание по центру
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_count} – выровняйте подпись по центру")
            # Регистр первого слова после номера
            parts = text.split(maxsplit=2)
            if len(parts) >= 3:
                first_word = parts[2]
                if first_word and first_word[0].islower():
                    issues.append(f"Рисунок {figure_count} – название должно начинаться с большой буквы")
            # Проверка пустых строк до и после (только для рисунка 3)
            if figure_count == 3:
                # Найти предыдущий и следующий абзацы
                prev_para = doc.paragraphs[idx-1] if idx > 0 else None
                next_para = doc.paragraphs[idx+1] if idx+1 < len(doc.paragraphs) else None
                if prev_para and prev_para.text.strip() != "":
                    issues.append("Рисунок 3 – добавьте пустую строку перед рисунком")
                if next_para and next_para.text.strip() != "":
                    issues.append("Рисунок 3 – добавьте пустую строку после рисунка")
        
        # ------ 2.8 Таблицы (проверка подписи над таблицей) ------
        # Ищем абзац, который начинается с "Таблица" и не содержит предлогов (чтобы исключить "в таблице")
        if text.startswith("Таблица") and not any(prep in text.lower() for prep in [" в ", " из ", "на "]):
            # Проверяем, идёт ли после этого абзаца таблица
            # Находим индекс текущего абзаца в all_elements
            try:
                elem_idx = next(i for i, (typ, _) in enumerate(all_elements) if typ == 'paragraph' and all_elements[i][1] is p._element)
                if next_element_is_table(elem_idx, all_elements):
                    # Это реальная подпись таблицы
                    match = re.search(r'Таблица\s+(\d+)', text)
                    if match:
                        tbl_num = int(match.group(1))
                        if tbl_num > len(doc.tables):
                            issues.append(f"Таблица {tbl_num} – номер превышает реальное количество таблиц ({len(doc.tables)})")
                        if text.rstrip().endswith("."):
                            issues.append(f"Таблица {tbl_num} – удалите точку в конце названия")
            except StopIteration:
                pass  # не нашли элемент – игнорируем
        
        # Запоминаем состояние
        prev_para_empty = False
        prev_para_was_figure = text.startswith("Рисунок")
    
    # ---------- 3. ПРОВЕРКА ТАБЛИЦ (содержимое на полужирное) ----------
    for t_idx, table in enumerate(doc.tables, start=1):
        bold_inside = False
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if run.bold:
                            bold_inside = True
        if bold_inside:
            issues.append(f"Таблица {t_idx} – уберите полужирное начертание")
    
    # ---------- 4. ПРОВЕРКА СПИСКА ИСТОЧНИКОВ ----------
    lit_start = None
    for idx, p in enumerate(doc.paragraphs):
        if "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" in p.text.upper():
            lit_start = idx
            break
    if lit_start is not None:
        # Проверяем первый абзац после заголовка (это должен быть первый источник)
        for idx in range(lit_start+1, len(doc.paragraphs)):
            p = doc.paragraphs[idx]
            if not p.text.strip():
                continue
            # Отступ первой строки
            indent = p.paragraph_format.first_line_indent
            if not indent:
                issues.append("Список источников – установите отступ первой строки 1,0 см")
            else:
                if abs(indent.cm - 1.0) > 0.1:
                    issues.append("Список источников – установите отступ первой строки 1,0 см")
            # Отступ слева
            left_indent = p.paragraph_format.left_indent
            if left_indent and left_indent.cm != 0:
                issues.append("Список источников – отступ слева должен быть 0 см")
            # Межстрочный интервал 1,2
            spacing = p.paragraph_format.line_spacing
            if spacing and p.paragraph_format.line_spacing_rule == 3:
                if abs(spacing - 1.2) > 0.05:
                    issues.append("Список источников – междустрочный интервал должен быть 1,2")
            # Выравнивание по ширине
            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – выравнивание по ширине")
            break   # проверяем только первый абзац списка, чтобы не дублировать
    
    # Удаляем дубликаты
    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

# ---------- ИНТЕРФЕЙС STREAMLIT ----------
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write("Загрузите ваш документ в формате .docx для проверки по чек-листу (поля, интервалы, отступы, таблицы, рисунки, список литературы).")

uploaded_file = st.file_uploader("Перетащите файл сюда или нажмите для выбора", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем структуру, интервалы и разметку..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for res in results:
        st.write(f"• {res}")
