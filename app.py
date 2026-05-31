import streamlit as st
import docx
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
import re

def mm_to_emu(mm):
    return Mm(mm).emu

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # ---------- 1. ПОЛЯ СТРАНИЦ ----------
    for i, section in enumerate(doc.sections, start=1):
        # Преобразуем EMU в миллиметры (1 pt = 1/72 дюйма = 25.4/72 мм)
        left_mm = section.left_margin.pt * 25.4 / 72 if section.left_margin else 0
        right_mm = section.right_margin.pt * 25.4 / 72 if section.right_margin else 0
        top_mm = section.top_margin.pt * 25.4 / 72 if section.top_margin else 0
        bottom_mm = section.bottom_margin.pt * 25.4 / 72 if section.bottom_margin else 0
        
        if (abs(left_mm - 20) > 0.5 or abs(right_mm - 20) > 0.5 or 
            abs(top_mm - 20) > 0.5 or abs(bottom_mm - 20) > 0.5):
            issues.append(f"Раздел {i} – установите поля: левое 20, правое 20, верх 20, низ 20 мм (сейчас {left_mm:.1f}, {right_mm:.1f}, {top_mm:.1f}, {bottom_mm:.1f})")

    # ---------- 2. ПРОВЕРКА АБЗАЦЕВ ----------
    # Флаги для отслеживания структуры
    in_special_block = False      # внутри содержания или списка литературы
    figure_count = 0
    table_count = 0
    prev_para_empty = False
    prev_para_was_figure = False
    
    # Список заголовков разделов, которые должны начинаться с новой страницы
    level1_headings = ["ВВЕДЕНИЕ", "1.", "2.", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]
    
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
        
        # ------ 2.2 Проверка начала раздела с новой страницы ------
        if any(text.startswith(h) for h in level1_headings):
            # Получаем номер страницы абзаца (сложно, лучше проверить разрыв перед)
            # Проверим, есть ли явный разрыв страницы перед этим абзацем
            has_page_break = False
            if p.runs:
                for run in p.runs:
                    if run._element.xml.find('w:br') != -1 and 'page' in run._element.xml:
                        has_page_break = True
            if not has_page_break and idx > 0:
                issues.append(f"{text[:30]}… – раздел должен начинаться с новой страницы")
        
        # ------ 2.3 Пустые строки перед подразделами (1.1, 1.2 и т.д.) ------
        if re.match(r'^\d+\.\d+\s', text) and prev_para_empty:
            issues.append(f"{text[:30]}… – уберите пустую строку перед подразделом")
        
        # ------ 2.4 Абзацный отступ (1,0 см только для основного текста) ------
        indent = p_format.first_line_indent
        # Игнорируем заголовки, подписи рисунков, таблицы, содержание
        ignore_indent = (text.isupper() or 
                         text.startswith(("Рисунок", "Таблица")) or 
                         in_special_block or
                         "СОДЕРЖАНИЕ" in text.upper())
        if indent and not ignore_indent:
            indent_cm = indent.cm
            if abs(indent_cm - 1.0) > 0.1:
                issues.append(f"«{text[:30]}…» – установите абзацный отступ 1,0 см (сейчас {indent_cm:.1f} см)")
        
        # ------ 2.5 Междустрочный интервал 1,2 ------
        spacing = p_format.line_spacing
        if spacing and not ignore_indent and not in_special_block:
            # Если line_spacing_rule = MULTIPLE, то spacing = множитель
            if p_format.line_spacing_rule == 3:  # WD_LINE_SPACING.MULTIPLE = 3
                if abs(spacing - 1.2) > 0.05:
                    issues.append(f"«{text[:30]}…» – замените интервал на 1,2 (сейчас {spacing:.1f})")
            else:
                # Если задано в пунктах, проверяем соответствие 14pt → 1,2*14 = 16.8pt?
                # Лучше требовать явный множитель 1,2
                issues.append(f"«{text[:30]}…» – установите множитель междустрочного интервала 1,2")
        
        # ------ 2.6 Шрифт и начертание ------
        for run in p.runs:
            if run.font.name and run.font.name != "Times New Roman":
                issues.append(f"«{text[:30]}…» – смените шрифт на Times New Roman")
            if run.font.size and run.font.size != Pt(14):
                issues.append(f"«{text[:30]}…» – установите размер шрифта 14")
            if run.underline:
                issues.append(f"«{text[:30]}…» – удалите подчеркивания")
        
        # ------ 2.7 СОДЕРЖАНИЕ ------
        if "СОДЕРЖАНИЕ" in text.upper():
            if text.endswith("."):
                issues.append("Содержание – удалите точку в конце")
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
        
        # ------ 2.8 Рисунки ------
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
            # Проверка пустых строк до и после (только для рисунка 3 по запросу)
            if figure_count == 3:
                # Ищем предыдущий и следующий абзацы
                if idx > 0 and doc.paragraphs[idx-1].text.strip() != "":
                    issues.append("Рисунок 3 – добавьте пустую строку перед рисунком")
                if idx < len(doc.paragraphs)-1 and doc.paragraphs[idx+1].text.strip() != "":
                    issues.append("Рисунок 3 – добавьте пустую строку после рисунка")
        
        # ------ 2.9 Таблицы (проверка подписи над таблицей) ------
        if text.startswith("Таблица") and not any(t in text for t in ["в таблице", "из таблицы"]):
            # Извлекаем номер таблицы
            match = re.search(r'Таблица\s+(\d+)', text)
            if match:
                tbl_num = int(match.group(1))
                if tbl_num > len(doc.tables):
                    issues.append(f"Таблица {tbl_num} – не соответствует реальному количеству таблиц в документе")
                if text.rstrip().endswith("."):
                    issues.append(f"Таблица {tbl_num} – удалите точку в конце названия")
        
        # Запоминаем состояние для следующей итерации
        prev_para_empty = False
        prev_para_was_figure = text.startswith("Рисунок")
    
    # ---------- 3. ПРОВЕРКА ТАБЛИЦ (содержимое) ----------
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
    # Находим абзац, содержащий "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"
    lit_start = None
    for idx, p in enumerate(doc.paragraphs):
        if "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" in p.text.upper():
            lit_start = idx
            break
    if lit_start is not None:
        for idx in range(lit_start+1, len(doc.paragraphs)):
            p = doc.paragraphs[idx]
            if not p.text.strip():
                continue
            # Каждый источник должен иметь отступ первой строки 1 см
            indent = p.paragraph_format.first_line_indent
            if indent:
                if abs(indent.cm - 1.0) > 0.1:
                    issues.append("Список источников – установите отступ первой строки 1,0 см")
            else:
                issues.append("Список источников – установите отступ первой строки 1,0 см")
            # Отступ слева 0 см
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
            break   # проверяем только первый абзац списка, иначе дубли
        
    # ---------- 5. ЗАКЛЮЧЕНИЕ НА НОВОЙ СТРАНИЦЕ ----------
    for idx, p in enumerate(doc.paragraphs):
        if "ЗАКЛЮЧЕНИЕ" in p.text.upper():
            has_page_break = False
            if p.runs:
                for run in p.runs:
                    if run._element.xml.find('w:br') != -1 and 'page' in run._element.xml:
                        has_page_break = True
            if not has_page_break and idx > 0:
                issues.append("Заключение – должно начинаться с новой страницы")
            break
    
    # Удаляем дубликаты
    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

# ---------- ИНТЕРФЕЙС STREAMLIT ----------
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write("Загрузите ваш документ в формате .docx для проверки по чек-листу (поля, интервалы, отступы, заголовки, таблицы, рисунки, список литературы).")

uploaded_file = st.file_uploader("Перетащите файл сюда или нажмите для выбора", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем структуру, интервалы и разметку..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for res in results:
        st.write(f"• {res}")
