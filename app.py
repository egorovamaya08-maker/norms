import streamlit as st
import docx
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import re
from lxml import etree

def has_page_number(text):
    """Проверяет, заканчивается ли строка номером страницы"""
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

def get_paragraph_text_with_bullets(paragraph):
    """
    Получает полный текст параграфа, включая маркеры списков.
    Маркеры могут быть в отдельных run'ах или в поле numPr.
    """
    # Проверяем наличие маркера через XML
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                # Это элемент нумерованного или маркированного списка
                # Получаем текст без маркера (сам маркер Word добавляет автоматически)
                return paragraph.text.strip(), True
    except:
        pass
    
    return paragraph.text.strip(), False

def is_list_marker(text):
    """Проверяет, является ли строка элементом списка с маркером"""
    # Стандартные маркеры
    list_patterns = [
        r'^\d+\)',           # 1) 2) 3)
        r'^\d+\.\s',         # 1. 2. 3. (если не заголовок)
        r'^[а-яё]\)',        # а) б) в)
        r'^[a-z]\)',         # a) b) c)
        r'^[\-–—]',          # - – —
    ]
    
    # Специальные символы маркеров (включая все возможные варианты •)
    bullet_chars = [
        '•', '●', '○', '◦',      # разные виды точек
        '▪', '▫', '■', '□',      # квадраты
        '▸', '▹', '►', '▻',      # треугольники
        '◆', '◇', '○', '◎',      # ромбы и круги
        '\uf0b7', '\uf0a7',       # Unicode bullet из разных наборов
        '\u2022', '\u2023',       # Bullet, Triangular bullet
        '\u25E6', '\u2043',       # White bullet, Hyphen bullet
    ]
    
    for pattern in list_patterns:
        if re.match(pattern, text):
            return True
    
    # Проверяем специальные символы
    if text and len(text) > 0:
        first_char = text[0]
        for bullet in bullet_chars:
            if first_char == bullet:
                return True
    
    # Проверяем * (звездочка с пробелом)
    if re.match(r'^\*\s', text):
        return True
    
    # Проверяем Unicode категории (разные виды маркеров)
    if text and ord(text[0]) in range(0x25A0, 0x25FF):  # Geometric Shapes
        return True
    if text and ord(text[0]) in range(0x2600, 0x26FF):  # Miscellaneous Symbols
        return True
    
    return False

def is_list_paragraph(paragraph):
    """
    Проверяет, является ли параграф частью списка через XML (numPr)
    """
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                return True
    except:
        pass
    return False

def is_all_caps(text):
    """Проверяет, написан ли текст заглавными буквами"""
    clean_text = re.sub(r'[\d\s\.,;:!?\-–—()«»""''\-\+–—]', '', text)
    if not clean_text:
        return False
    return clean_text == clean_text.upper()

def normalize_title(text):
    """Нормализует заголовок для сравнения: убирает нумерацию, лишние пробелы"""
    # Убираем нумерацию в начале
    text = re.sub(r'^(?:ГЛАВА|РАЗДЕЛ)\s+\d+[\.\s]*', '', text, flags=re.IGNORECASE)
    text = re.sub(r'^\d+(?:\.\d+)*[\s\.]+', '', text)
    return text.strip().upper()

def extract_toc_entries(doc, start_idx):
    """Извлекает заголовки из содержания (до start_idx)"""
    toc_entries = []
    for i in range(start_idx):
        txt = doc.paragraphs[i].text.strip()
        if not txt:
            continue
        if has_page_number(txt):
            # Это строка содержания
            # Убираем номер страницы в конце
            clean = re.sub(r'[\t\s\.]{2,}\d+$', '', txt).strip()
            if clean and len(clean) > 5:
                toc_entries.append(clean)
    return toc_entries

def check_page_breaks_in_range(doc, start_pos, end_pos):
    """Проверяет наличие разрывов страниц в диапазоне"""
    for i in range(start_pos, min(end_pos + 1, len(doc.paragraphs))):
        para = doc.paragraphs[i]
        # Проверяем lastRenderedPageBreak (автоматический разрыв)
        try:
            rPr = para._element.find(qn('w:rPr'))
            if rPr is not None:
                br = rPr.find(qn('w:lastRenderedPageBreak'))
                if br is not None:
                    return True
        except:
            pass
        
        # Проверяем явные разрывы страниц
        for run in para.runs:
            try:
                br_list = run._element.findall(qn('w:br'))
                for br in br_list:
                    if br.get(qn('w:type')) == 'page':
                        return True
            except:
                pass
        
        # Проверяем pageBreakBefore в свойствах параграфа
        try:
            pPr = para._element.find(qn('w:pPr'))
            if pPr is not None:
                page_break = pPr.find(qn('w:pageBreakBefore'))
                if page_break is not None:
                    return True
        except:
            pass
    
    return False

def table_spans_multiple_pages(doc, table, tbl_pos, next_table_pos):
    """
    Определяет, переносится ли таблица на несколько страниц.
    """
    # Проверка 1: Количество строк (если > 20, скорее всего переносится)
    if len(table.rows) > 20:
        return True
    
    # Проверка 2: Разрывы страниц внутри таблицы
    try:
        # Получаем XML таблицы
        tbl = table._tbl
        
        # Ищем все параграфы внутри таблицы
        all_paras = tbl.findall('.//' + qn('w:p'))
        
        for para_elem in all_paras:
            # Проверяем lastRenderedPageBreak
            for r_elem in para_elem.findall(qn('w:r')):
                for br in r_elem.findall(qn('w:br')):
                    if br.get(qn('w:type')) == 'page':
                        return True
                # Проверяем lastRenderedPageBreak
                rPr = r_elem.find(qn('w:rPr'))
                if rPr is not None:
                    if rPr.find(qn('w:lastRenderedPageBreak')) is not None:
                        return True
            
            # Проверяем pageBreakBefore
            pPr = para_elem.find(qn('w:pPr'))
            if pPr is not None:
                if pPr.find(qn('w:pageBreakBefore')) is not None:
                    return True
    except Exception as e:
        st.write(f"    ⚠️ Ошибка при анализе XML таблицы: {e}")
    
    # Проверка 3: Ищем строки с номерами страниц между началом и концом таблицы
    for i in range(tbl_pos, min(next_table_pos, len(doc.paragraphs))):
        if has_page_number(doc.paragraphs[i].text.strip()):
            return True
    
    return False

def find_table_continuation_markers(doc, start_pos, end_pos, table_num):
    """Ищет маркеры продолжения/окончания таблицы"""
    markers_found = []
    for i in range(start_pos, min(end_pos + 1, len(doc.paragraphs))):
        para = doc.paragraphs[i]
        txt = para.text.strip()
        if txt:
            if re.search(r'(?:Продолжение|Окончание)\s+таблицы?\s*' + str(table_num), txt):
                markers_found.append((i, txt[:100]))
    return markers_found

def test_document(file):
    """Тестовая функция для проверки проблем"""
    doc = docx.Document(file)
    
    st.header("🔍 Анализ документа")
    
    # ============================================
    # ТЕСТ 1: Заголовки разделов
    # ============================================
    st.subheader("📋 Тест 1: Поиск заголовков разделов")
    
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    start_idx = None
    toc_items = []
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt:
            continue
        
        has_page = has_page_number(txt)
        
        if has_page:
            toc_items.append(('page_number', i, txt[:80]))
            continue
        
        if txt.upper() in level1_keywords or (re.match(r'^\d+\.\s+[А-ЯЁ]', txt) and is_all_caps(txt)):
            start_idx = i
            st.write(f"✅ Строка {i}: **Начало основного текста**: '{txt[:100]}'")
            break
    
    if start_idx is None:
        st.warning("⚠️ Не найдено начало основного текста")
        return
    
    st.write(f"📌 Индекс начала основного текста: {start_idx}")
    
    # Извлекаем заголовки из содержания
    toc_entries = extract_toc_entries(doc, start_idx)
    st.write(f"📑 Заголовков в содержании: {len(toc_entries)}")
    if toc_entries:
        st.write("Первые 5 заголовков из содержания:")
        for entry in toc_entries[:5]:
            st.write(f"  • '{entry[:80]}'")
    
    # Анализируем заголовки в тексте
    st.write("---")
    st.write("**Анализ заголовков в тексте:**")
    
    section_headers = []
    issues_found = 0
    
    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        
        if not txt or has_page_number(txt):
            continue
        
        # Пропускаем элементы списков
        if is_list_marker(txt) or is_list_paragraph(p):
            continue
        
        first_line = get_effective_first_line_indent(p)
        xml_info = get_paragraph_xml_info(p)
        normalized = normalize_title(txt)
        
        # Проверка 1: Соответствие содержанию
        is_toc_match = False
        for toc_entry in toc_entries:
            if normalize_title(toc_entry) == normalized:
                is_toc_match = True
                break
        
        # Проверка 2: Заголовок с номером
        has_number = bool(re.match(r'^\d+\.\s+[А-ЯЁ]', txt))
        
        # Проверка 3: Все капсом
        all_caps = is_all_caps(txt)
        
        # Определяем, является ли заголовком
        is_header = False
        header_type = ""
        
        if has_number:
            is_header = True
            header_type = "с номером"
        elif all_caps and len(txt) > 25:
            is_header = True
            header_type = "капсом без номера"
        elif is_toc_match and len(txt) > 20:
            is_header = True
            header_type = "из содержания (не капсом)"
        elif re.match(r'^(?:ГЛАВА|РАЗДЕЛ)\s+\d+', txt, re.IGNORECASE):
            is_header = True
            header_type = "ГЛАВА/РАЗДЕЛ"
        
        if is_header:
            st.write(f"**Строка {i}:** Заголовок ({header_type})")
            st.write(f"  Текст: '{txt[:100]}'")
            st.write(f"  Капсом: {'Да' if all_caps else 'Нет'}")
            st.write(f"  Есть в содержании: {'Да' if is_toc_match else 'Нет'}")
            st.write(f"  Отступ первой строки: {first_line:.3f} см")
            st.write(f"  XML: {xml_info}")
            
            if abs(first_line) > 0.1:
                st.write(f"  🔴 **Ошибка: Заголовок раздела не должен иметь отступ! Требуется 0 см, сейчас {first_line:.1f} см**")
                issues_found += 1
            else:
                st.write(f"  ✅ Отступ отсутствует — правильно")
            
            section_headers.append((header_type, i, txt[:100], first_line))
            st.write("---")
        
        if len(section_headers) >= 40:
            break
    
    if issues_found == 0:
        st.success(f"✅ Все заголовки разделов оформлены правильно (без отступа)")
    else:
        st.error(f"❌ Найдено {issues_found} заголовков с ошибочным отступом")
    
    # ============================================
    # ТЕСТ 2: Проверка списков
    # ============================================
    st.subheader("📝 Тест 2: Проверка всех элементов списков")
    
    list_items_found = 0
    list_items_correct = 0
    list_items_wrong = 0
    list_types = {}
    
    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        
        if not txt or has_page_number(txt):
            continue
        
        # Проверяем и текстовые маркеры, и XML numPr
        has_text_marker = is_list_marker(txt)
        has_xml_list = is_list_paragraph(p)
        
        if has_text_marker or has_xml_list:
            first_line = get_effective_first_line_indent(p)
            left_indent = get_effective_left_indent(p)
            xml_info = get_paragraph_xml_info(p)
            
            list_items_found += 1
            
            # Определяем тип маркера
            if has_xml_list:
                marker_type = "список Word (numPr)"
            elif re.match(r'^\d+\)', txt):
                marker_type = "1)"
            elif re.match(r'^[а-яё]\)', txt):
                marker_type = "а)"
            elif re.match(r'^[a-z]\)', txt):
                marker_type = "a)"
            elif re.match(r'^[\-–—]', txt):
                marker_type = "тире"
            elif txt and ord(txt[0]) in [8226, 8227, 9702, 9642, 9650, 9670, 9679, 8226]:
                marker_type = f"спецсимвол (U+{ord(txt[0]):04X})"
            else:
                marker_type = "другой"
            
            list_types[marker_type] = list_types.get(marker_type, 0) + 1
            
            # Показываем все элементы
            if list_items_found <= 25:
                # Показываем hex первого символа для отладки
                first_char_hex = f"U+{ord(txt[0]):04X}" if txt else "N/A"
                st.write(f"**Элемент списка {list_items_found} [{marker_type}] (строка {i}):**")
                st.write(f"  Текст: '{txt[:80]}...'")
                st.write(f"  Первый символ: '{txt[0]}' ({first_char_hex})")
                st.write(f"  Отступ слева: {left_indent:.3f} см")
                st.write(f"  Отступ первой строки: {first_line:.3f} см")
                st.write(f"  XML: {xml_info}")
                
                if abs(first_line - 1.0) < 0.15:
                    st.write(f"  ✅ Отступ 1.0 см установлен — правильно")
                    list_items_correct += 1
                elif abs(first_line) < 0.1:
                    st.write(f"  ⚠️ Отступ отсутствует (0 см) — нужно установить 1.0 см")
                    list_items_wrong += 1
                else:
                    st.write(f"  ❌ Отступ {first_line:.1f} см — отличается от требуемого 1.0 см")
                    list_items_wrong += 1
                
                st.write("---")
            else:
                if abs(first_line - 1.0) < 0.15:
                    list_items_correct += 1
                else:
                    list_items_wrong += 1
        
        if list_items_found >= 100:
            break
    
    if list_items_found == 0:
        st.write("ℹ️ Списки не найдены")
        # Показываем для отладки несколько параграфов, которые могут быть списками
        st.write("🔍 **Отладка: проверяем параграфы на наличие списков:**")
        count = 0
        for i in range(start_idx, min(start_idx + 50, len(doc.paragraphs))):
            p = doc.paragraphs[i]
            txt = p.text.strip()
            if txt and not has_page_number(txt):
                has_xml = is_list_paragraph(p)
                first_char = txt[0] if txt else ''
                first_char_hex = f"U+{ord(first_char):04X}" if first_char else "N/A"
                if has_xml or ord(first_char) < 128 or ord(first_char) in range(0x2000, 0x3000):
                    st.write(f"  Строка {i}: первый символ '{first_char}' ({first_char_hex}), XML list: {has_xml}, текст: '{txt[:50]}...'")
                    count += 1
                    if count >= 10:
                        break
    else:
        st.write(f"📊 **Всего элементов списков: {list_items_found}**")
        st.write(f"  ✅ С правильным отступом (1.0 см): {list_items_correct}")
        st.write(f"  ❌ Требуют исправления: {list_items_wrong}")
        
        st.write(f"📊 **Типы маркеров:**")
        for marker_type, count in sorted(list_types.items()):
            st.write(f"  • {marker_type}: {count} шт.")
        
        if list_items_wrong == 0:
            st.success("✅ Все элементы списков имеют правильный отступ 1.0 см")
        else:
            st.error(f"❌ {list_items_wrong} элементов списка требуют установки отступа 1.0 см")
    
    # ============================================
    # ТЕСТ 3: Таблицы
    # ============================================
    st.subheader("📊 Тест 3: Анализ таблиц и переносов")
    
    try:
        start_element = doc.paragraphs[start_idx]._element
        start_body_pos = list(doc.element.body).index(start_element)
    except:
        start_body_pos = 0
    
    lit_start = None
    for i in range(start_idx, len(doc.paragraphs)):
        txt = doc.paragraphs[i].text.strip()
        if txt.upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            lit_start = i
            break
    
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
    
    st.write(f"Найдено таблиц в основном тексте: {len(main_tables)}")
    
    for t_idx, (tbl_pos, table) in enumerate(main_tables, start=1):
        st.write(f"---")
        st.write(f"**Таблица {t_idx}** (позиция: {tbl_pos})")
        st.write(f"  Строк: {len(table.rows)}, столбцов: {len(table.columns)}")
        
        next_table_pos = end_body_pos
        for next_tbl_pos, _ in main_tables[t_idx:]:
            if next_tbl_pos > tbl_pos:
                next_table_pos = next_tbl_pos
                break
        
        # Проверяем перенос
        spans_pages = table_spans_multiple_pages(doc, table, tbl_pos, next_table_pos)
        
        st.write(f"  📏 Перенос на несколько страниц: {'✅ ДА' if spans_pages else '❌ Нет (или не обнаружено)'}")
        
        # Детальная отладка
        if not spans_pages:
            st.write(f"  🔍 Отладка: проверяем разрывы страниц в диапазоне {tbl_pos}-{next_table_pos}")
            has_breaks = check_page_breaks_in_range(doc, tbl_pos, next_table_pos)
            st.write(f"    Разрывы страниц в диапазоне: {'Да' if has_breaks else 'Нет'}")
            
            # Проверяем XML таблицы
            try:
                tbl_xml = etree.tostring(table._tbl, encoding='unicode')
                if 'lastRenderedPageBreak' in tbl_xml or 'w:br w:type="page"' in tbl_xml:
                    st.write(f"    ⚠️ Найдены разрывы в XML таблицы, но они не распознаны!")
                    st.code(tbl_xml[:500])
            except:
                pass
        
        # Подпись
        caption_para = None
        for i in range(tbl_pos - 1, start_body_pos - 1, -1):
            elem = doc.element.body[i]
            if elem.tag.endswith('p'):
                for idx, para in enumerate(doc.paragraphs):
                    if para._element is elem and para.text.strip():
                        caption_para = para
                        break
                if caption_para:
                    break
        
        if caption_para:
            caption_text = caption_para.text.strip()
            st.write(f"  📝 Подпись: '{caption_text[:150]}'")
            
            if caption_text.startswith("Таблица"):
                if '—' in caption_text or '–' in caption_text:
                    st.write(f"  ✅ Используется тире")
                elif '--' in caption_text:
                    st.write(f"  ⚠️ Используется двойной дефис")
                elif ' - ' in caption_text:
                    st.write(f"  ⚠️ Используется дефис вместо тире")
                
                if caption_text.rstrip().endswith("."):
                    st.write(f"  ❌ Точка в конце подписи")
                else:
                    st.write(f"  ✅ Нет точки в конце")
        
        # Маркеры переноса
        if spans_pages:
            st.write(f"  🔍 **Проверка маркеров переноса:**")
            markers = find_table_continuation_markers(doc, tbl_pos, next_table_pos, t_idx)
            
            if markers:
                st.write(f"  ✅ Найдены маркеры:")
                for marker_pos, marker_text in markers:
                    st.write(f"    • Строка {marker_pos}: '{marker_text}'")
            else:
                st.write(f"  ❌ **Маркеры переноса не найдены!**")
                st.write(f"  💡 Необходимо добавить:")
                st.write(f"     • 'Продолжение таблицы {t_idx}' на следующих страницах")
                st.write(f"     • 'Окончание таблицы {t_idx}' на последней странице")
        
        # Показываем первые строки
        st.write(f"  📋 Первые 3 строки:")
        for r_idx, row in enumerate(table.rows[:3]):
            cells_text = []
            for cell in row.cells:
                cells_text.append(cell.text.strip()[:30])
            st.write(f"    Строка {r_idx + 1}: {' | '.join(cells_text)}")
    
    # ============================================
    # ИТОГИ
    # ============================================
    st.header("📊 Итоги анализа")
    st.write(f"• Заголовков разделов: {len(section_headers)}")
    st.write(f"• Элементов списков: {list_items_found}")
    st.write(f"• Таблиц: {len(main_tables)}")
    
    tables_needing_continuation = 0
    for t_idx, (tbl_pos, table) in enumerate(main_tables, start=1):
        next_table_pos = end_body_pos
        for next_tbl_pos, _ in main_tables[t_idx:]:
            if next_tbl_pos > tbl_pos:
                next_table_pos = next_tbl_pos
                break
        if table_spans_multiple_pages(doc, table, tbl_pos, next_table_pos):
            markers = find_table_continuation_markers(doc, tbl_pos, next_table_pos, t_idx)
            if not markers:
                tables_needing_continuation += 1
    
    if tables_needing_continuation > 0:
        st.warning(f"⚠️ {tables_needing_continuation} таблиц(ы) требуют добавления 'Продолжение/Окончание таблицы'")

# Интерфейс
st.set_page_config(page_title="Тест проверки документа v5", layout="wide")
st.title("🧪 Тест проверки: заголовки, списки, таблицы")
st.write("Проверяет:")
st.write("1. **Заголовки** — из содержания и капсом, не должны иметь отступ")
st.write("2. **Списки** — все виды маркеров (•, -, 1)), должны иметь отступ 1.0 см")
st.write("3. **Таблицы** — переносы, подписи, маркеры Продолжение/Окончание")

uploaded_file = st.file_uploader("Загрузите документ .docx для теста", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем..."):
        test_document(uploaded_file)
