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
    return bool(re.search(r'[\t\s\.]{2,}\d+$', text))

def is_all_caps(text):
    clean_text = re.sub(r'[\d\s\.,;:!?\-–—()«»""''«»]', '', text)
    if not clean_text:
        return False
    return clean_text == clean_text.upper()

def is_section_header(text):
    if re.match(r'^\d+\.\s+[А-ЯЁ]', text) and is_all_caps(text):
        return True
    if re.match(r'^(?:ГЛАВА|РАЗДЕЛ)\s+\d+', text, re.IGNORECASE):
        return True
    if text.upper().strip() in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}:
        return True
    return False

def normalize_title(text):
    text = re.sub(r'^(?:ГЛАВА|РАЗДЕЛ)\s+\d+[\.\s]*', '', text, flags=re.IGNORECASE)
    text = re.sub(r'^\d+(?:\.\d+)*[\s\.]+', '', text)
    return text.strip().upper()

def extract_toc_entries(doc, start_idx):
    toc_entries = []
    for i in range(start_idx):
        txt = doc.paragraphs[i].text.strip()
        if not txt:
            continue
        if has_page_number(txt):
            clean = re.sub(r'[\t\s\.]{2,}\d+$', '', txt).strip()
            if clean and len(clean) > 5:
                toc_entries.append(clean)
    return toc_entries

def get_list_marker_info(paragraph, doc):
    text = paragraph.text.strip()
    if not text:
        return False, "", True
    
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                numId_elem = numPr.find(qn('w:numId'))
                if numId_elem is not None:
                    numId = numId_elem.get(qn('w:val'))
                    
                    numbering_part = doc.part.numbering_part
                    if numbering_part is not None:
                        numbering_xml = numbering_part._element
                        
                        abstractNumId = None
                        for num in numbering_xml.findall(qn('w:num')):
                            if num.get(qn('w:numId')) == numId:
                                aid_elem = num.find(qn('w:abstractNumId'))
                                if aid_elem is not None:
                                    abstractNumId = aid_elem.get(qn('w:val'))
                                break
                        
                        if abstractNumId:
                            for abstractNum in numbering_xml.findall(qn('w:abstractNum')):
                                if abstractNum.get(qn('w:abstractNumId')) == abstractNumId:
                                    for lvl in abstractNum.findall(qn('w:lvl')):
                                        numFmt = lvl.find(qn('w:numFmt'))
                                        if numFmt is not None:
                                            fmt = numFmt.get(qn('w:val'))
                                            if fmt == 'bullet':
                                                return True, "круглый маркер (•)", False
                                            elif fmt == 'decimal':
                                                return True, "нумерованный (цифры)", True
                                            elif fmt in ['lowerLetter', 'upperLetter']:
                                                return True, "буквенный", True
                                            else:
                                                return True, f"формат '{fmt}'", True
    except:
        pass
    
    if re.match(r'^\d+\)', text):
        return True, "нумерованный", True
    if re.match(r'^[а-яё]\)', text):
        return True, "буквенный", True
    if re.match(r'^[a-z]\)', text):
        return True, "буквенный", True
    if re.match(r'^[\-–—]', text):
        return True, "тире", True
    
    if text and ord(text[0]) in [8226, 8227, 9679, 9702]:
        return True, "круглый маркер (•)", False
    
    return False, "", True

def get_effective_first_line_indent(paragraph):
    pf = paragraph.paragraph_format
    if pf.first_line_indent is not None:
        return pf.first_line_indent.cm
    
    try:
        style = paragraph.style
        if style and style.paragraph_format.first_line_indent is not None:
            return style.paragraph_format.first_line_indent.cm
    except:
        pass
    
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                first_line = ind.get(qn('w:firstLine'))
                hanging = ind.get(qn('w:hanging'))
                if first_line is not None:
                    twips = int(first_line)
                    return twips / 567
                elif hanging is not None:
                    return 0
    except:
        pass
    
    return 0

def get_effective_left_indent(paragraph):
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
    
    return 0

def is_table_continuation(text):
    return bool(re.match(r'^(?:Продолжение|Окончание)\s+таблицы?\s*\d', text))

def is_real_table(table, doc):
    """
    Проверяет, является ли таблица "реальной" (с данными), а не служебной.
    Служебные таблицы обычно пустые или с одной ячейкой для форматирования.
    """
    # Если таблица пустая — не реальная
    if len(table.rows) == 0:
        return False
    
    # Собираем весь текст из таблицы
    total_text = ""
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                total_text += para.text
    
    total_text = total_text.strip()
    
    # Если в таблице нет текста — служебная
    if not total_text:
        return False
    
    # Если таблица состоит из одной ячейки с коротким текстом — возможно служебная
    if len(table.rows) == 1 and len(table.columns) == 1:
        if len(total_text) < 20:
            return False
    
    # Если таблица содержит текст вроде "Продолжение таблицы" в ячейках — это часть другой таблицы
    if re.search(r'Продолжение\s+таблицы', total_text):
        return False
    if re.search(r'Окончание\s+таблицы', total_text):
        return False
    
    return True

def check_word_document(file):
    doc = docx.Document(file)
    auto_issues = []      # Автоматические ошибки
    manual_checks = []    # Для ручной проверки

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
        auto_issues.append("Поля страниц – установите левое 20 мм, правое 20 мм, верхнее 20 мм, нижнее 20 мм")

    # ---------- 2. ПОИСК НАЧАЛА ОСНОВНОГО ТЕКСТА ----------
    start_idx = None
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt:
            continue
        
        if has_page_number(txt):
            continue
        
        if re.match(r'^\d+\.\s+[А-Яа-я]', txt) and not is_all_caps(txt):
            continue
        
        if is_section_header(txt):
            start_idx = i
            break
    
    if start_idx is None:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."], []

    toc_entries = extract_toc_entries(doc, start_idx)

    # ---------- 3. ГРАНИЦЫ СПИСКА ИСТОЧНИКОВ ----------
    lit_start = None
    for i in range(start_idx, len(doc.paragraphs)):
        txt = doc.paragraphs[i].text.strip()
        if txt.upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" and not has_page_number(txt):
            lit_start = i
            break
    
    lit_end = len(doc.paragraphs)
    if lit_start is not None:
        appendix_keywords = ["ПРИЛОЖЕНИЕ", "ПРИЛОЖЕНИЯ", "APPENDIX"]
        for i in range(lit_start + 1, len(doc.paragraphs)):
            txt = doc.paragraphs[i].text.strip().upper()
            if any(txt.startswith(kw) for kw in appendix_keywords):
                lit_end = i
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
        
        if has_page_number(text):
            prev_para_empty = False
            continue
        
        if re.match(r'^\d+\.\s+[А-Яа-я]', text) and not is_all_caps(text):
            prev_para_empty = False
            continue
        
        # Проверяем списки
        is_list, marker_type, marker_valid = get_list_marker_info(p, doc)
        if is_list:
            if not marker_valid:
                auto_issues.append(f"«{text[:50]}» – замените круглый маркер (•) на тире, букву или цифру")
            prev_para_empty = False
            continue
        
        # Продолжение/окончание таблицы
        if is_table_continuation(text):
            first_line = get_effective_first_line_indent(p)
            if abs(first_line) > 0.1:
                auto_issues.append(f"«{text[:50]}» – уберите абзацный отступ (должен быть 0 см)")
            prev_para_empty = False
            continue

        pf = p.paragraph_format
        alignment = get_effective_alignment(p)
        
        is_level1 = is_section_header(text)
        is_subsection = bool(re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]', text)) and not is_level1
        is_figure = text.startswith("Рисунок")
        is_table_caption = text.startswith("Таблица")
        
        normalized = normalize_title(text)
        in_toc = any(normalize_title(e) == normalized for e in toc_entries) if toc_entries else False
        
        if in_toc and not is_level1 and not is_subsection and len(text) > 20:
            is_subsection = True

        # --- Заголовок раздела ---
        if is_level1:
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
                    auto_issues.append(f"«{text[:50]}» – раздел должен начинаться с новой страницы")
            
            first_line = get_effective_first_line_indent(p)
            if abs(first_line) > 0.1:
                auto_issues.append(f"«{text[:50]}» – уберите абзацный отступ у заголовка")
            
            if not is_paragraph_bold(p):
                auto_issues.append(f"«{text[:50]}» – заголовок раздела должен быть полужирным")
            
            if re.match(r'^\d+\.', text) and not is_all_caps(text):
                auto_issues.append(f"«{text[:50]}» – заголовок раздела должен быть прописными буквами")
            
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                auto_issues.append(f"«{text[:50]}» – выровняйте заголовок по центру")
            
            if text.endswith("."):
                auto_issues.append(f"«{text[:50]}» – удалите точку в конце")
            
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                auto_issues.append(f"«{text[:50]}» – после заголовка должна быть пустая строка")
        
        # --- Подраздел ---
        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()
            
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                auto_issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
            
            if not is_paragraph_bold(p):
                auto_issues.append(f"Подраздел «{sub_name[:50]}» – заголовок должен быть полужирным")
            
            if text.endswith("."):
                auto_issues.append(f"Подраздел «{sub_name[:50]}» – удалите точку в конце")
            
            if prev_para_empty:
                auto_issues.append(f"Подраздел «{sub_name[:50]}» – уберите пустую строку перед подразделом")
        
        # --- Рисунок ---
        elif is_figure:
            figure_counter += 1
            fig_match = re.match(r'^(Рисунок\s+\d+(?:\.\d+)?)', text)
            fig_number = fig_match.group(1) if fig_match else f"Рисунок {figure_counter}"
            
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                auto_issues.append(f"{fig_number} – выровняйте подпись по центру")
            
            if text.endswith(".") and not re.search(r'\([^)]*\)\.$', text):
                auto_issues.append(f"{fig_number} – удалите точку в конце")
            
            m = re.match(r'^Рисунок\s+\d+(?:\.\d+)?\s*[–\-]\s*(.+)$', text)
            if m:
                title = m.group(1).strip()
                if title and title[0].islower():
                    auto_issues.append(f"{fig_number} – название должно начинаться с большой буквы")
            
            # Пустые строки — для ручной проверки
            if idx > start_idx and not is_empty_paragraph(doc.paragraphs[idx - 1]):
                manual_checks.append(f"{fig_number} – проверьте наличие пустой строки перед рисунком")
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                manual_checks.append(f"{fig_number} – проверьте наличие пустой строки после рисунка")
        
        # --- Подпись таблицы ---
        elif is_table_caption:
            pass
        
        # --- Обычный текст ---
        else:
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                auto_issues.append(f"«{text[:50]}» – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
            if pf.space_before and pf.space_before.pt > 0.5:
                auto_issues.append(f"«{text[:50]}» – интервал перед абзацем должен быть 0 пт")
        
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

    # Собираем только реальные таблицы
    main_tables = []
    for table in doc.tables:
        try:
            tbl_pos = list(doc.element.body).index(table._element)
            if start_body_pos < tbl_pos < end_body_pos:
                # Проверяем, что таблица реальная (не служебная)
                if is_real_table(table, doc):
                    main_tables.append((tbl_pos, table))
        except:
            pass

    for t_idx, (tbl_pos, table) in enumerate(main_tables, start=1):
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
            
            tbl_num_match = re.match(r'Таблица\s+([\d.]+)', caption)
            tbl_num = tbl_num_match.group(1) if tbl_num_match else str(t_idx)
            
            # Тире в подписи
            if '—' not in caption and '–' not in caption:
                if '--' in caption or ' - ' in caption:
                    auto_issues.append(f"Таблица {tbl_num} – замените дефис на тире (—) в подписи")
            
            # Формат подписи
            if not re.match(r'Таблица\s+[\d.]+\s+[–—]\s+', caption):
                auto_issues.append(f"Таблица {tbl_num} – должно быть «Таблица {tbl_num} — Название»")
            
            # Без точки
            if caption.rstrip().endswith("."):
                auto_issues.append(f"Таблица {tbl_num} – удалите точку в конце названия")
            
            # Пустая строка перед подписью — ручная проверка
            if caption_idx is not None and caption_idx > start_idx:
                if not is_empty_paragraph(doc.paragraphs[caption_idx - 1]):
                    manual_checks.append(f"Таблица {tbl_num} – проверьте наличие пустой строки перед подписью таблицы")
            
            # Пустая строка после таблицы — ручная проверка
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
                manual_checks.append(f"Таблица {tbl_num} – проверьте наличие пустой строки после таблицы")
        else:
            auto_issues.append(f"Таблица {t_idx} – отсутствует подпись над таблицей")

        # Полужирный внутри таблицы
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
            tbl_num = tbl_num_match.group(1) if caption_para and tbl_num_match else str(t_idx)
            auto_issues.append(f"Таблица {tbl_num} – уберите полужирное начертание внутри таблицы")

        # Перенос таблицы — ручная проверка
        rows = len(table.rows)
        if rows > 2:
            tbl_num = tbl_num_match.group(1) if caption_para and tbl_num_match else str(t_idx)
            next_table_pos = end_body_pos
            for next_tbl_pos, _ in main_tables[t_idx:]:
                if next_tbl_pos > tbl_pos:
                    next_table_pos = next_tbl_pos
                    break
            
            markers = []
            for i in range(tbl_pos, min(next_table_pos + 1, len(doc.paragraphs))):
                txt = doc.paragraphs[i].text.strip()
                if txt and re.search(r'(?:Продолжение|Окончание)\s+таблицы?\s*' + re.escape(tbl_num), txt):
                    markers.append((i, txt[:100]))
            
            if not markers:
                manual_checks.append(f"Таблица {tbl_num} – проверьте наличие «Продолжение таблицы {tbl_num}» / «Окончание таблицы {tbl_num}» при переносе на следующую страницу")

    # ---------- 6. СПИСОК ИСТОЧНИКОВ ----------
    if lit_start is not None:
        sources_with_issues = 0
        
        for i in range(lit_start + 1, lit_end):
            source = doc.paragraphs[i]
            txt = source.text.strip()
            
            if not txt:
                continue
            
            if has_page_number(txt):
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
            auto_issues.append(
                "Список источников – проверьте оформление: "
                "отступ слева 0 см, отступ первой строки 1,0 см, "
                "междустрочный интервал 1,2 (множитель), выравнивание по ширине"
            )

    # ---------- ИТОГ ----------
    auto_issues = list(dict.fromkeys(auto_issues))
    manual_checks = list(dict.fromkeys(manual_checks))
    
    result = []
    if not auto_issues and not manual_checks:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    
    if auto_issues:
        result.extend(auto_issues)
    
    if manual_checks:
        result.append("\n📋 Для ручной проверки проверяющего:")
        result.extend(manual_checks)
    
    return result

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
        if r.startswith("📋"):
            st.markdown(f"**{r}**")
        else:
            st.write(f"• {r}")
