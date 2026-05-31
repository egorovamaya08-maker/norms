import streamlit as st
import docx
from docx.oxml.ns import qn
from lxml import etree
import re

def has_page_number(text):
    if re.search(r'[\t\s\.]{2,}\d+$', text):
        return True
    return False

def get_paragraph_xml_info(paragraph):
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

def is_all_caps(text):
    clean_text = re.sub(r'[\d\s\.,;:!?\-–—()«»""''\-\+–—]', '', text)
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

def get_list_marker_info(paragraph, doc):
    """
    Определяет тип маркера для элемента списка.
    Возвращает (marker_type, is_valid)
    marker_type: "тире", "круглый маркер", "нумерованный", "буквенный", "не определен"
    is_valid: True если маркер допустимый (тире), False если недопустимый (круглый)
    """
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                numId_elem = numPr.find(qn('w:numId'))
                if numId_elem is not None:
                    numId = numId_elem.get(qn('w:val'))
                    
                    # Ищем определение списка в document.xml
                    # Ищем numId в <w:num> элементах
                    numbering_part = doc.part.numbering_part
                    if numbering_part:
                        numbering_xml = numbering_part._element
                        
                        # Ищем num с этим numId
                        for num in numbering_xml.findall(qn('w:num')):
                            if num.get(qn('w:numId')) == numId:
                                # Ищем abstractNumId
                                abstractNumId_elem = num.find(qn('w:abstractNumId'))
                                if abstractNumId_elem is not None:
                                    abstractNumId = abstractNumId_elem.get(qn('w:val'))
                                    
                                    # Ищем abstractNum с этим abstractNumId
                                    for abstractNum in numbering_xml.findall(qn('w:abstractNum')):
                                        if abstractNum.get(qn('w:abstractNumId')) == abstractNumId:
                                            # Ищем lvl (уровень 0)
                                            for lvl in abstractNum.findall(qn('w:lvl')):
                                                if lvl.get(qn('w:ilvl')) == '0':
                                                    # Ищем numFmt (формат номера)
                                                    numFmt = lvl.find(qn('w:numFmt'))
                                                    if numFmt is not None:
                                                        fmt = numFmt.get(qn('w:val'))
                                                        
                                                        if fmt == 'bullet':
                                                            # Круглый маркер - недопустимо
                                                            return "круглый маркер (•)", False
                                                        elif fmt == 'decimal':
                                                            return "нумерованный", True
                                                        elif fmt == 'lowerLetter':
                                                            return "буквенный", True
                                                        elif fmt == 'upperLetter':
                                                            return "буквенный", True
                                                        else:
                                                            return f"другой формат ({fmt})", True
                                                    
                                                    # Ищем lvlText (текст маркера)
                                                    lvlText = lvl.find(qn('w:lvlText'))
                                                    if lvlText is not None:
                                                        text = lvlText.get(qn('w:val'))
                                                        if text:
                                                            # Проверяем, содержит ли текст символы маркеров
                                                            if '•' in text or '\u2022' in text:
                                                                return "круглый маркер (•)", False
                                                            if text in ['–', '—', '-']:
                                                                return "тире", True
                                                    
                                                    break
                                        
                                        break
                                break
    except Exception as e:
        pass
    
    return "не определен (проверьте вручную)", True

def is_list_item(text, paragraph, doc):
    """Проверяет, является ли параграф элементом списка"""
    # Проверяем XML numPr
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                marker_type, is_valid = get_list_marker_info(paragraph, doc)
                return True, marker_type, is_valid
    except:
        pass
    
    # Исключаем заголовки
    if is_section_header(text):
        return False, "", True
    
    # Текстовые маркеры
    if re.match(r'^\d+\)', text):
        return True, "нумерованный 1)", True
    if re.match(r'^[а-яё]\)', text):
        return True, "буквенный а)", True
    if re.match(r'^[a-z]\)', text):
        return True, "буквенный a)", True
    if re.match(r'^[\-–—]', text):
        return True, "тире", True
    if re.match(r'^\*\s', text):
        return True, "звездочка *", False
    
    # Спецсимволы-маркеры
    if text:
        first_char = text[0]
        if ord(first_char) in [8226, 8227]:  # • ‣
            return True, f"круглый маркер (U+{ord(first_char):04X})", False
        if ord(first_char) in range(0x25A0, 0x25FF):
            return True, f"маркер-символ (U+{ord(first_char):04X})", False
    
    return False, "", True

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

def find_table_continuation_markers(doc, start_pos, end_pos, table_num):
    markers_found = []
    for i in range(start_pos, min(end_pos + 1, len(doc.paragraphs))):
        txt = doc.paragraphs[i].text.strip()
        if txt:
            if re.search(r'(?:Продолжение|Окончание)\s+таблицы?\s*' + str(table_num), txt):
                markers_found.append((i, txt[:100]))
    return markers_found

def test_document(file):
    doc = docx.Document(file)
    
    st.header("🔍 Анализ документа")
    
    # ============================================
    # ТЕСТ 1: Заголовки разделов
    # ============================================
    st.subheader("📋 Тест 1: Заголовки разделов")
    
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    start_idx = None
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt or has_page_number(txt):
            continue
        if txt.upper() in level1_keywords or (re.match(r'^\d+\.\s+[А-ЯЁ]', txt) and is_all_caps(txt)):
            start_idx = i
            st.write(f"✅ Строка {i}: Начало текста: '{txt[:100]}'")
            break
    
    if start_idx is None:
        st.warning("⚠️ Не найдено начало основного текста")
        return
    
    toc_entries = extract_toc_entries(doc, start_idx)
    
    st.write("**Заголовки в тексте:**")
    section_headers = []
    issues_found = 0
    
    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        
        if not txt or has_page_number(txt):
            continue
        
        # Пропускаем списки
        is_list, _, _ = is_list_item(txt, p, doc)
        if is_list:
            continue
        
        first_line = get_effective_first_line_indent(p)
        normalized = normalize_title(txt)
        in_toc = any(normalize_title(e) == normalized for e in toc_entries)
        
        is_header = False
        header_type = ""
        
        if is_section_header(txt):
            is_header = True
            header_type = "раздел"
        elif in_toc and len(txt) > 20:
            is_header = True
            header_type = "из содержания"
        elif is_all_caps(txt) and len(txt) > 25:
            is_header = True
            header_type = "капсом"
        
        if is_header and len(section_headers) < 30:
            st.write(f"**Строка {i}** ({header_type}): '{txt[:80]}'")
            st.write(f"  Отступ: {first_line:.3f} см {'✅' if abs(first_line) < 0.1 else '❌ должен быть 0 см'}")
            if abs(first_line) > 0.1:
                issues_found += 1
            section_headers.append((i, txt[:80], first_line))
            st.write("---")
    
    if issues_found == 0:
        st.success("✅ Все заголовки без отступа")
    else:
        st.error(f"❌ {issues_found} заголовков с ошибочным отступом")
    
    # ============================================
    # ТЕСТ 2: Списки
    # ============================================
    st.subheader("📝 Тест 2: Элементы списков")
    st.write("**Требования:**")
    st.write("• Маркер должен быть ТИРЕ (–), а не круглый (•)")
    st.write("• Отступ первой строки должен быть 1.0 см")
    
    list_items_found = 0
    list_items_correct = 0
    list_items_wrong = 0
    wrong_marker_items = []
    
    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        
        if not txt or has_page_number(txt):
            continue
        
        is_list, marker_type, marker_valid = is_list_item(txt, p, doc)
        
        if is_list:
            first_line = get_effective_first_line_indent(p)
            xml_info = get_paragraph_xml_info(p)
            
            list_items_found += 1
            
            indent_ok = abs(first_line - 1.0) < 0.15
            
            if list_items_found <= 25:
                st.write(f"**Элемент {list_items_found}** строка {i}:")
                st.write(f"  '{txt[:70]}...'")
                st.write(f"  Тип маркера: {marker_type} {'✅' if marker_valid else '❌ НЕДОПУСТИМ (нужно тире)'}")
                st.write(f"  Отступ: {first_line:.3f} см {'✅' if indent_ok else '❌ нужен 1.0 см'}")
                st.write(f"  XML: {xml_info}")
                
                if not marker_valid or not indent_ok:
                    if not marker_valid:
                        st.write(f"  🔴 Ошибка: замените круглый маркер на тире (–)")
                    if not indent_ok:
                        st.write(f"  🔴 Ошибка: установите отступ 1.0 см (сейчас {first_line:.1f} см)")
                else:
                    st.write(f"  🟢 Всё правильно")
                
                st.write("---")
            
            if marker_valid and indent_ok:
                list_items_correct += 1
            else:
                list_items_wrong += 1
                if not marker_valid:
                    wrong_marker_items.append((i, txt[:50], marker_type))
        
        if list_items_found >= 100:
            break
    
    if list_items_found == 0:
        st.write("ℹ️ Списки не найдены")
    else:
        st.write(f"📊 Всего: {list_items_found}")
        st.write(f"  ✅ Правильно: {list_items_correct}")
        st.write(f"  ❌ С ошибками: {list_items_wrong}")
        
        if wrong_marker_items:
            st.write("---")
            st.write(f"**❌ Элементы с недопустимыми круглыми маркерами (нужно заменить на тире):**")
            for idx, text, marker in wrong_marker_items:
                st.write(f"  • Строка {idx}: [{marker}] '{text}...'")
    
    # ============================================
    # ТЕСТ 3: Таблицы
    # ============================================
    st.subheader("📊 Тест 3: Таблицы")
    st.write("Проверка необходимости «Продолжение таблицы» / «Окончание таблицы»")
    st.write("⚠️ Автоматически невозможно определить границы страниц. Проверяем все таблицы.")
    
    try:
        start_element = doc.paragraphs[start_idx]._element
        start_body_pos = list(doc.element.body).index(start_element)
    except:
        start_body_pos = 0
    
    lit_start = None
    for i in range(start_idx, len(doc.paragraphs)):
        if doc.paragraphs[i].text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            lit_start = i
            break
    
    end_body_pos = len(doc.element.body)
    if lit_start:
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
    
    st.write(f"Таблиц в тексте: {len(main_tables)}")
    
    tables_need_check = 0
    
    for t_idx, (tbl_pos, table) in enumerate(main_tables, start=1):
        st.write(f"---")
        st.write(f"**Таблица {t_idx}** (строка {tbl_pos})")
        
        rows = len(table.rows)
        cols = len(table.columns)
        st.write(f"  Размер: {rows} строк × {cols} столбцов")
        
        # Подпись
        caption = None
        for i in range(tbl_pos - 1, start_body_pos - 1, -1):
            elem = doc.element.body[i]
            if elem.tag.endswith('p'):
                for para in doc.paragraphs:
                    if para._element is elem and para.text.strip():
                        caption = para
                        break
                if caption:
                    break
        
        if caption:
            cap_text = caption.text.strip()
            st.write(f"  Подпись: '{cap_text[:120]}'")
            
            if '—' in cap_text or '–' in cap_text:
                st.write(f"  ✅ Тире правильное")
            elif '--' in cap_text or ' - ' in cap_text:
                st.write(f"  ⚠️ Нужно заменить на тире (—)")
            
            if cap_text.rstrip().endswith("."):
                st.write(f"  ⚠️ Убрать точку в конце")
        
        # Проверка на перенос (для таблиц больше 2 строк)
        if rows > 2:
            next_table_pos = end_body_pos
            for next_tbl_pos, _ in main_tables[t_idx:]:
                if next_tbl_pos > tbl_pos:
                    next_table_pos = next_tbl_pos
                    break
            
            markers = find_table_continuation_markers(doc, tbl_pos, next_table_pos, t_idx)
            
            if not markers:
                st.write(f"  🔴 **Нет «Продолжение таблицы {t_idx}» / «Окончание таблицы {t_idx}»**")
                st.write(f"  💡 Проверьте вручную. Если таблица на нескольких страницах, добавьте:")
                st.write(f"     • «Продолжение таблицы {t_idx}» на следующей странице")
                st.write(f"     • «Окончание таблицы {t_idx}» на последней странице")
                tables_need_check += 1
            else:
                st.write(f"  ✅ Маркеры найдены:")
                for m_pos, m_text in markers:
                    st.write(f"    • Строка {m_pos}: '{m_text}'")
        else:
            st.write(f"  ✅ Маленькая таблица, перенос маловероятен")
    
    if tables_need_check > 0:
        st.warning(f"⚠️ {tables_need_check} таблиц требуют ручной проверки")
    else:
        st.success("✅ Все таблицы в порядке")
    
    # ============================================
    # ИТОГИ
    # ============================================
    st.header("📊 Итоги")
    st.write(f"• Заголовков: {len(section_headers)} (ошибок отступа: {issues_found})")
    st.write(f"• Списков: {list_items_found} (правильно: {list_items_correct}, ошибок: {list_items_wrong})")
    if wrong_marker_items:
        st.write(f"  • Из них с недопустимым маркером: {len(wrong_marker_items)}")
    st.write(f"• Таблиц: {len(main_tables)} (нужна проверка переноса: {tables_need_check})")

# Интерфейс
st.set_page_config(page_title="Проверка документа", layout="wide")
st.title("🧪 Проверка: заголовки, списки, таблицы")
st.write("1. Заголовки — без отступа")
st.write("2. Списки — маркер ТИРЕ, отступ 1.0 см")
st.write("3. Таблицы — Продолжение/Окончание")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализ..."):
        test_document(uploaded_file)
