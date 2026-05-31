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
    """Проверяет, заканчивается ли строка номером страницы (содержание)"""
    return bool(re.search(r'[\t\s\.]{2,}\d+$', text))

def is_all_caps(text):
    """Проверяет, состоит ли текст из заглавных букв (игнорируя цифры и символы)"""
    clean_text = re.sub(r'[\d\s\.,;:!?\-–—()«»""''«»]', '', text)
    if not clean_text:
        return False
    return clean_text == clean_text.upper()

def is_section_header(text):
    """Определяет, является ли строка заголовком раздела первого уровня"""
    # "1. НАЗВАНИЕ" (номер, точка, пробел, заглавные буквы)
    if re.match(r'^\d+\.\s+[А-ЯЁ]', text) and is_all_caps(text):
        return True
    # "ГЛАВА 1" или "РАЗДЕЛ 1"
    if re.match(r'^(?:ГЛАВА|РАЗДЕЛ)\s+\d+', text, re.IGNORECASE):
        return True
    # Ключевые слова
    if text.upper().strip() in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}:
        return True
    return False

def normalize_title(text):
    """Нормализует заголовок для сравнения с содержанием"""
    text = re.sub(r'^(?:ГЛАВА|РАЗДЕЛ)\s+\d+[\.\s]*', '', text, flags=re.IGNORECASE)
    text = re.sub(r'^\d+(?:\.\d+)*[\s\.]+', '', text)
    return text.strip().upper()

def extract_toc_entries(doc, start_idx):
    """Извлекает все пункты содержания (до start_idx)"""
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

def is_list_item(paragraph, doc):
    """Определяет, является ли параграф элементом списка"""
    text = paragraph.text.strip()
    if not text:
        return False
    
    # Проверяем XML numPr (нумерованный список Word)
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                return True
    except:
        pass
    
    # Текстовые маркеры: "1)", "а)", "-", "•" и т.д.
    if re.match(r'^\d+\)', text):
        return True
    if re.match(r'^[а-яё]\)', text):
        return True
    if re.match(r'^[a-z]\)', text):
        return True
    if re.match(r'^[\-–—]', text):
        return True
    if re.match(r'^\*\s', text):
        return True
    
    # Спецсимволы-маркеры
    if text and ord(text[0]) in [8226, 8227]:  # • ‣
        return True
    if text and ord(text[0]) in range(0x25A0, 0x25FF):  # Геометрические символы
        return True
    
    return False

def get_effective_first_line_indent(paragraph):
    """Получает отступ первой строки в см"""
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
    
    return 0

def find_table_continuation_markers(doc, start_pos, end_pos, table_num):
    """Ищет маркеры продолжения/окончания таблицы"""
    markers = []
    for i in range(start_pos, min(end_pos + 1, len(doc.paragraphs))):
        txt = doc.paragraphs[i].text.strip()
        if txt and re.search(r'(?:Продолжение|Окончание)\s+таблицы?\s*' + str(table_num), txt):
            markers.append((i, txt[:100]))
    return markers

def check_word_document(file):
    doc = docx.Document(file)
    issues = []

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
        
        # Пропускаем содержание
        if has_page_number(txt):
            continue
        
        # Пропускаем нумерованные пункты НЕ капсом (содержание)
        if re.match(r'^\d+\.\s+[А-Яа-я]', txt) and not is_all_caps(txt):
            continue
        
        # Ищем заголовок раздела
        if is_section_header(txt):
            start_idx = i
            break
    
    if start_idx is None:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]

    # Извлекаем пункты содержания для сравнения
    toc_entries = extract_toc_entries(doc, start_idx)

    # ---------- 3. ИЩЕМ ГРАНИЦЫ СПИСКА ИСТОЧНИКОВ ----------
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
        
        # Пропускаем содержание
        if has_page_number(text):
            prev_para_empty = False
            continue
        
        # Пропускаем нумерованные пункты НЕ капсом
        if re.match(r'^\d+\.\s+[А-Яа-я]', text) and not is_all_caps(text):
            prev_para_empty = False
            continue
        
        # Пропускаем элементы списков
        if is_list_item(p, doc):
            prev_para_empty = False
            continue

        pf = p.paragraph_format
        alignment = get_effective_alignment(p)
        
        # Определяем тип абзаца
        is_level1 = is_section_header(text)
        is_subsection = bool(re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]', text)) and not is_level1
        is_figure = text.startswith("Рисунок")
        is_table_caption = text.startswith("Таблица")
        
        # Проверяем, есть ли заголовок в содержании
        normalized = normalize_title(text)
        in_toc = any(normalize_title(e) == normalized for e in toc_entries) if toc_entries else False
        
        # Если это заголовок из содержания, но не раздел — считаем подразделом
        if in_toc and not is_level1 and not is_subsection and len(text) > 20:
            is_subsection = True

        # --- Заголовок раздела ---
        if is_level1:
            # Новая страница (кроме ВВЕДЕНИЯ, если оно первое)
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
            
            # Отступ 0
            first_line = get_effective_first_line_indent(p)
            if abs(first_line) > 0.1:
                issues.append(f"«{text[:50]}» – уберите абзацный отступ у заголовка")
            
            # Полужирный
            if not is_paragraph_bold(p):
                issues.append(f"«{text[:50]}» – заголовок раздела должен быть полужирным")
            
            # Прописные (для номерных разделов)
            if re.match(r'^\d+\.', text) and not is_all_caps(text):
                issues.append(f"«{text[:50]}» – заголовок раздела должен быть прописными буквами")
            
            # Центр
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:50]}» – выровняйте заголовок по центру")
            
            # Без точки
            if text.endswith("."):
                issues.append(f"«{text[:50]}» – удалите точку в конце")
            
            # Пустая строка после
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                issues.append(f"«{text[:50]}» – после заголовка должна быть пустая строка")
        
        # --- Подраздел ---
        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()
            
            # Отступ 1 см
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
            
            # Полужирный
            if not is_paragraph_bold(p):
                issues.append(f"Подраздел «{sub_name[:50]}» – заголовок должен быть полужирным")
            
            # Без точки в конце
            if text.endswith("."):
                issues.append(f"Подраздел «{sub_name[:50]}» – удалите точку в конце")
            
            # Нет пустой строки перед
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
        
        # --- Подпись таблицы (пропускаем, проверим отдельно) ---
        elif is_table_caption:
            pass
        
        # --- Обычный текст ---
        else:
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                issues.append(f"«{text[:50]}» – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append(f"«{text[:50]}» – интервал перед абзацем должен быть 0 пт")
        
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
                main_tables.append((tbl_pos, table))
        except:
            pass

    for t_idx, (tbl_pos, table) in enumerate(main_tables, start=1):
        # Подпись таблицы
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
            
            # Тире в подписи
            if '—' not in caption and '–' not in caption:
                if '--' in caption or ' - ' in caption:
                    issues.append(f"Таблица {t_idx} – замените дефис на тире (—) в подписи")
            
            # Формат "Таблица N -- Название"
            if not re.match(r'Таблица\s+\d+\s+[–—]\s+', caption):
                issues.append(f"Таблица {t_idx} – оформление подписи: должно быть «Таблица N — Название»")
            
            # Номер совпадает
            m_num = re.search(r'Таблица\s+(\d+)', caption)
            if m_num and int(m_num.group(1)) != t_idx:
                issues.append(f"Таблица {t_idx} – номер в подписи не соответствует порядковому (должен быть {t_idx})")
            
            # Без точки
            if caption.rstrip().endswith("."):
                issues.append(f"Таблица {t_idx} – удалите точку в конце названия")
            
            # Пустая строка перед подписью
            if caption_idx is not None and caption_idx > start_idx:
                if not is_empty_paragraph(doc.paragraphs[caption_idx - 1]):
                    issues.append(f"Таблица {t_idx} – добавьте пустую строку перед подписью таблицы")
            
            # Пустая строка после таблицы
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
            issues.append(f"Таблица {t_idx} – уберите полужирное начертание внутри таблицы")

        # Проверка на перенос таблицы
        rows = len(table.rows)
        if rows > 2:
            next_table_pos = end_body_pos
            for next_tbl_pos, _ in main_tables[t_idx:]:
                if next_tbl_pos > tbl_pos:
                    next_table_pos = next_tbl_pos
                    break
            
            markers = find_table_continuation_markers(doc, tbl_pos, next_table_pos, t_idx)
            if not markers:
                issues.append(f"Таблица {t_idx} – проверьте наличие «Продолжение таблицы {t_idx}» / «Окончание таблицы {t_idx}» при переносе")

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
            issues.append(
                "Список источников – проверьте оформление: "
                "отступ слева 0 см, отступ первой строки 1,0 см, "
                "междустрочный интервал 1,2 (множитель), выравнивание по ширине"
            )

    # ---------- ИТОГ ----------
    issues = list(dict.fromkeys(issues))
    if not issues:
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
