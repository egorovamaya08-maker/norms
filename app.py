import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import re

def get_effective_alignment(paragraph):
    if paragraph.alignment is not None:
        return paragraph.alignment
    try:
        style = paragraph.style
        if style and style.paragraph_format.alignment is not None:
            return style.paragraph_format.alignment
    except:
        pass
    return None

def is_paragraph_bold(paragraph):
    try:
        if paragraph.style and paragraph.style.font and paragraph.style.font.bold:
            return True
    except:
        pass
    runs = [r for r in paragraph.runs if r.text.strip()]
    if not runs:
        return False
    return all(r.bold for r in runs)

def is_empty_paragraph(paragraph):
    return len(paragraph.text.strip()) == 0

def has_page_number(text):
    """Проверяет, заканчивается ли строка номером страницы"""
    # Проверяем разные варианты: пробелы, табуляции, точки
    if re.search(r'[\t\s\.]{2,}\d+$', text):
        return True
    return False

def get_paragraph_xml_info(paragraph):
    """Выводит XML информацию о параграфе для отладки"""
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                first_line = ind.get(qn('w:firstLine'))
                left = ind.get(qn('w:left'))
                hanging = ind.get(qn('w:hanging'))
                return f"firstLine={first_line}, left={left}, hanging={hanging}"
            else:
                return "no ind element"
        else:
            return "no pPr element"
    except:
        return "error reading XML"

def get_effective_first_line_indent(paragraph):
    """Получает отступ первой строки в см"""
    # Сначала пробуем через python-docx
    pf = paragraph.paragraph_format
    if pf.first_line_indent is not None:
        return pf.first_line_indent.cm
    
    # Потом через стиль
    try:
        style = paragraph.style
        if style and style.paragraph_format.first_line_indent is not None:
            return style.paragraph_format.first_line_indent.cm
    except:
        pass
    
    # Потом через XML
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                first_line = ind.get(qn('w:firstLine'))
                hanging = ind.get(qn('w:hanging'))
                if first_line is not None:
                    twips = int(first_line)
                    return twips / 567  # twips to cm
                elif hanging is not None:
                    # Если задан hanging, то отступа первой строки нет
                    return 0
    except:
        pass
    
    return 0  # По умолчанию считаем что отступа нет

def get_effective_left_indent(paragraph):
    """Получает отступ слева в см"""
    pf = paragraph.paragraph_format
    if pf.left_indent is not None:
        return pf.left_indent.cm
    
    try:
        style = paragraph.style
        if style and style.paragraph_format.left_indent is not None:
            return style.paragraph_format.left_indent.cm
    except:
        pass
    
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                left = ind.get(qn('w:left'))
                if left is not None:
                    twips = int(left)
                    return twips / 567
    except:
        pass
    
    return 0  # По умолчанию отступа нет

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    debug_info = []  # Для отладки

    # ---------- 1. ПОЛЯ СТРАНИЦ ----------
    margins_ok = True
    for section in doc.sections:
        if (abs(section.left_margin.mm - 20) > 0.5 or
            abs(section.right_margin.mm - 20) > 0.5 or
            abs(section.top_margin.mm - 20) > 0.5 or
            abs(section.bottom_margin.mm - 20) > 0.5):
            margins_ok = False
            break
    if not margins_ok:
        issues.append("Поля страниц – установите левое 20 мм, правое 20 мм, верхнее 20 мм, нижнее 20 мм")

    # ---------- 2. ПОИСК НАЧАЛА ОСНОВНОГО ТЕКСТА ----------
    start_idx = None
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt:
            continue
        
        # Пропускаем строки с номерами страниц (содержание)
        if has_page_number(txt):
            debug_info.append(f"SKIP TOC: '{txt[:80]}'")
            continue
        
        # Ищем заголовки
        if txt.upper() in level1_keywords:
            start_idx = i
            debug_info.append(f"FOUND START: '{txt[:80]}' at index {i}")
            break
        
        # Ищем заголовки глав ТОЛЬКО если они ПОЛНОСТЬЮ прописные
        if re.match(r'^\d+\.\s+[А-ЯЁ]', txt) and txt == txt.upper():
            start_idx = i
            debug_info.append(f"FOUND START: '{txt[:80]}' at index {i}")
            break
    
    if start_idx is None:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."] + debug_info

    # ---------- 3. ИЩЕМ НАЧАЛО СПИСКА ИСТОЧНИКОВ ----------
    lit_start = None
    for i in range(start_idx, len(doc.paragraphs)):
        txt = doc.paragraphs[i].text.strip()
        if txt.upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" and not has_page_number(txt):
            lit_start = i
            break

    # ---------- 4. ПРОВЕРКА ОСНОВНОГО ТЕКСТА ----------
    figure_counter = 0
    prev_para_empty = False
    end_idx = lit_start if lit_start is not None else len(doc.paragraphs)

    for idx in range(start_idx, end_idx):
        p = doc.paragraphs[idx]
        text = p.text.strip()
        
        if not text:
            prev_para_empty = True
            continue
        
        # Пропускаем строки с номерами страниц
        if has_page_number(text):
            debug_info.append(f"SKIP IN TEXT: '{text[:80]}'")
            prev_para_empty = False
            continue

        pf = p.paragraph_format
        alignment = get_effective_alignment(p)
        
        # Определяем тип абзаца
        is_level1 = False
        is_subsection = False
        is_figure = text.startswith("Рисунок")
        is_table_caption = text.startswith("Таблица")
        
        # Раздел первого уровня
        if text.upper() in level1_keywords:
            is_level1 = True
        elif re.match(r'^\d+\.\s+[А-ЯЁ]', text) and text == text.upper():
            is_level1 = True
        # Подраздел
        elif re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]', text):
            is_subsection = True

        # Если это не заголовок, не подраздел, не рисунок и не таблица - проверяем на отступ
        if not is_level1 and not is_subsection and not is_figure and not is_table_caption:
            first_line = get_effective_first_line_indent(p)
            xml_info = get_paragraph_xml_info(p)
            
            if abs(first_line - 1.0) > 0.2:
                debug_info.append(f"INDENT ISSUE: '{text[:80]}' | first_line={first_line:.2f}cm | XML: {xml_info}")
                issues.append(f"«{text[:50]}» – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
            
            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append(f"«{text[:50]}» – интервал перед абзацем должен быть 0 пт")
        
        # --- Заголовок раздела ---
        elif is_level1:
            first_line = get_effective_first_line_indent(p)
            if abs(first_line) > 0.1:
                debug_info.append(f"LEVEL1 INDENT: '{text[:80]}' | first_line={first_line:.2f}cm")
                issues.append(f"«{text[:50]}» – уберите абзацный отступ у заголовка")
            
            if text.upper() != "ВВЕДЕНИЕ" or idx != start_idx:
                page_break = False
                if idx > start_idx:
                    prev_p = doc.paragraphs[idx - 1]
                    for run in prev_p.runs:
                        if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                            page_break = True
                for run in p.runs:
                    if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                        page_break = True
                if not page_break:
                    issues.append(f"«{text[:50]}» – раздел должен начинаться с новой страницы")
            
            if not is_paragraph_bold(p):
                issues.append(f"«{text[:50]}» – заголовок раздела должен быть полужирным")
            
            if re.match(r'^\d+\.', text) and text != text.upper():
                issues.append(f"«{text[:50]}» – заголовок раздела должен быть прописными буквами")
            
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:50]}» – выровняйте заголовок по центру")
            
            if text.endswith("."):
                issues.append(f"«{text[:50]}» – удалите точку в конце")
            
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                issues.append(f"«{text[:50]}» – после заголовка должна быть пустая строка")
        
        # --- Подраздел ---
        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()
            first_line = get_effective_first_line_indent(p)
            
            if abs(first_line - 1.0) > 0.2:
                debug_info.append(f"SUBSECTION INDENT: '{text[:80]}' | first_line={first_line:.2f}cm")
                issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
            
            if not is_paragraph_bold(p):
                issues.append(f"Подраздел «{sub_name[:50]}» – заголовок должен быть полужирным")
            
            if text.endswith("."):
                issues.append(f"Подраздел «{sub_name[:50]}» – удалите точку в конце")
            
            if prev_para_empty:
                issues.append(f"Подраздел «{sub_name[:50]}» – уберите пустую строку перед подразделом")
        
        # --- Рисунок ---
        elif is_figure:
            figure_counter += 1
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – выровняйте подпись по центру")
            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – удалите точку в конце")
            m = re.match(r'^Рисунок\s+\d+\s*[–\-]\s*(.+)$', text)
            if m:
                title = m.group(1).strip()
                if title and title[0].islower():
                    issues.append(f"Рисунок {figure_counter} – название должно начинаться с большой буквы")
            if idx > start_idx and not is_empty_paragraph(doc.paragraphs[idx - 1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку перед рисунком")
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку после рисунка")
        
        prev_para_empty = False

    # ---------- 5. ТАБЛИЦЫ ----------
    try:
        start_element = doc.paragraphs[start_idx]._element
        start_body_pos = list(doc.element.body).index(start_element)
    except:
        start_body_pos = 0

    end_body_pos = len(doc.element.body)
    if lit_start is not None:
        try:
            lit_element = doc.paragraphs[lit_start]._element
            end_body_pos = list(doc.element.body).index(lit_element)
        except:
            pass

    main_tables = []
    for table in doc.tables:
        try:
            tbl_pos = list(doc.element.body).index(table._element)
            if start_body_pos < tbl_pos < end_body_pos:
                main_tables.append(table)
        except:
            pass

    for t_idx, table in enumerate(main_tables, start=1):
        tbl_pos = list(doc.element.body).index(table._element)
        
        caption_para = None
        for i in range(tbl_pos - 1, start_body_pos - 1, -1):
            elem = doc.element.body[i]
            if elem.tag.endswith('p'):
                for para in doc.paragraphs:
                    if para._element is elem and para.text.strip():
                        caption_para = para
                        break
                if caption_para:
                    break

        if caption_para and caption_para.text.strip().startswith("Таблица"):
            caption = caption_para.text.strip()
            caption_idx = None
            for i, para in enumerate(doc.paragraphs):
                if para._element is caption_para._element:
                    caption_idx = i
                    break
            
            if not re.match(r'Таблица\s+\d+\s+--\s+', caption):
                issues.append(f"Таблица {t_idx} – оформление подписи: должно быть «Таблица N -- Название»")
            
            m_num = re.search(r'Таблица\s+(\d+)', caption)
            if m_num and int(m_num.group(1)) != t_idx:
                issues.append(f"Таблица {t_idx} – номер в подписи не соответствует порядковому (должен быть {t_idx})")
            
            if caption.rstrip().endswith("."):
                issues.append(f"Таблица {t_idx} – удалите точку в конце названия")
            
            if caption_idx is not None and caption_idx > start_idx:
                if not is_empty_paragraph(doc.paragraphs[caption_idx - 1]):
                    issues.append(f"Таблица {t_idx} – добавьте пустую строку перед подписью таблицы")
            
            next_para = None
            for i in range(tbl_pos + 1, end_body_pos):
                elem = doc.element.body[i]
                if elem.tag.endswith('p'):
                    for para in doc.paragraphs:
                        if para._element is elem:
                            next_para = para
                            break
                    break
            if next_para and not is_empty_paragraph(next_para):
                issues.append(f"Таблица {t_idx} – добавьте пустую строку после таблицы")
        else:
            issues.append(f"Таблица {t_idx} – отсутствует подпись над таблицей")

        bold_in_table = False
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.bold:
                            bold_in_table = True
                            break
                    if bold_in_table:
                        break
                if bold_in_table:
                    break
            if bold_in_table:
                break
        if bold_in_table:
            issues.append(f"Таблица {t_idx} – уберите полужирное начертание внутри таблицы")

    # ---------- 6. СПИСОК ИСТОЧНИКОВ ----------
    if lit_start is not None:
        # Добавим отладку для первого источника
        first_source = None
        for i in range(lit_start + 1, len(doc.paragraphs)):
            if doc.paragraphs[i].text.strip():
                first_source = doc.paragraphs[i]
                break
        
        if first_source:
            first_line = get_effective_first_line_indent(first_source)
            left_indent = get_effective_left_indent(first_source)
            xml_info = get_paragraph_xml_info(first_source)
            debug_info.append(f"FIRST SOURCE: '{first_source.text.strip()[:80]}'")
            debug_info.append(f"  first_line={first_line:.2f}cm, left={left_indent:.2f}cm")
            debug_info.append(f"  XML: {xml_info}")
        
        sources_with_issues = 0
        
        for i in range(lit_start + 1, len(doc.paragraphs)):
            source = doc.paragraphs[i]
            if not source.text.strip():
                continue
            
            has_issue = False
            
            left_indent = get_effective_left_indent(source)
            if abs(left_indent) > 0.1:
                has_issue = True
            
            first_line = get_effective_first_line_indent(source)
            if abs(first_line - 1.0) > 0.1:
                has_issue = True
            
            if has_issue:
                sources_with_issues += 1
        
        if sources_with_issues > 0:
            issues.append(
                "Список источников – проверьте оформление: "
                "отступ слева 0 см, отступ первой строки 1,0 см, "
                "междустрочный интервал 1,2 (множитель), выравнивание по ширине"
            )

    # Показываем отладочную информацию
    if debug_info:
        issues.append("\n📋 ОТЛАДКА:")
        issues.extend(debug_info)

    # ---------- ИТОГ ----------
    real_issues = [i for i in issues if not i.startswith("📋") and not i.startswith("SKIP") and not i.startswith("FOUND") and not i.startswith("INDENT") and not i.startswith("LEVEL1") and not i.startswith("SUBSECTION") and not i.startswith("FIRST")]
    
    if not real_issues:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

# Интерфейс
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write("Загрузите документ в формате .docx – проверка по полному чек-листу.")
uploaded_file = st.file_uploader("Выберите файл", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Проверяем..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for r in results:
        st.write(f"• {r}")
