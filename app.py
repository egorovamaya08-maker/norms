import streamlit as st
import docx
from docx.oxml.ns import qn
import re

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

def is_all_caps(text):
    """Проверяет, написан ли текст заглавными буквами (с учетом кириллицы)"""
    clean_text = re.sub(r'[\d\s\.,;:!?\-–—()«»""''\-\+]', '', text)
    if not clean_text:
        return False
    return clean_text == clean_text.upper()

def is_list_marker(text):
    """Проверяет, является ли строка элементом списка с маркером"""
    # Маркеры: 1) 2) а) б) - • ▪ ○ и т.д.
    list_patterns = [
        r'^\d+\)',           # 1) 2) 3)
        r'^[а-яё]\)',        # а) б) в)
        r'^[a-z]\)',         # a) b) c)
        r'^[\-–—•▪○▸▹►▻◆◇]', # - – — • и другие маркеры
        r'^\*\s',            # * с пробелом
    ]
    for pattern in list_patterns:
        if re.match(pattern, text):
            return True
    return False

def estimate_table_pages(table):
    """Оценивает, на сколько страниц может переноситься таблица"""
    row_count = len(table.rows)
    if row_count <= 20:
        return 1
    elif row_count <= 45:
        return 2
    elif row_count <= 70:
        return 3
    else:
        return 4

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

def check_page_break_in_table(doc, table_element):
    """
    Проверяет, есть ли разрыв страницы внутри таблицы.
    Ищет параграфы внутри таблицы с разрывом страницы.
    """
    try:
        # Проходим по всем строкам и ячейкам таблицы
        for row in table_element.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    # Проверяем наличие разрыва страницы в параграфе
                    for run in para.runs:
                        if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                            return True
                    # Проверяем в свойствах параграфа
                    pPr = para._element.find(qn('w:pPr'))
                    if pPr is not None:
                        for br in pPr.findall(qn('w:br')):
                            if br.get(qn('w:type')) == 'page':
                                return True
    except:
        pass
    return False

def table_spans_multiple_pages(doc, table, tbl_pos, next_table_pos):
    """
    Определяет, переносится ли таблица на несколько страниц.
    Проверяет:
    1. Наличие разрывов страниц внутри таблицы
    2. Количество строк (косвенный признак)
    3. Наличие текста с номерами страниц между частями таблицы
    """
    # Проверка 1: Прямые разрывы страниц
    if check_page_break_in_table(doc, table):
        return True
    
    # Проверка 2: Большое количество строк
    if len(table.rows) > 25:
        return True
    
    # Проверка 3: Ищем строки с номерами страниц между частями таблицы
    # (характерно для таблиц, разорванных автоматически)
    try:
        for i in range(tbl_pos + 1, next_table_pos):
            elem = doc.element.body[i]
            if elem.tag.endswith('p'):
                for para in doc.paragraphs:
                    if para._element is elem and para.text.strip():
                        if has_page_number(para.text.strip()):
                            return True
    except:
        pass
    
    return False

def test_document(file):
    """Тестовая функция для проверки проблем"""
    doc = docx.Document(file)
    
    st.header("🔍 Анализ документа")
    
    # ============================================
    # ТЕСТ 1: Поиск заголовков разделов капсом
    # ============================================
    st.subheader("📋 Тест 1: Поиск заголовков разделов и подразделов")
    
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
        
        # Ищем начало основного текста
        if txt.upper() in level1_keywords or (re.match(r'^\d+\.\s+[А-ЯЁ]', txt) and is_all_caps(txt)):
            start_idx = i
            st.write(f"✅ Строка {i}: **Начало основного текста**: '{txt[:100]}'")
            break
    
    if start_idx is None:
        st.warning("⚠️ Не найдено начало основного текста")
        return
    
    st.write(f"📌 Индекс начала основного текста: {start_idx}")
    
    # Ищем все потенциальные заголовки разделов
    st.write("---")
    st.write("**Анализ заголовков в тексте:**")
    
    section_headers = []
    issues_found = 0
    
    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        
        if not txt or has_page_number(txt):
            continue
        
        first_line = get_effective_first_line_indent(p)
        xml_info = get_paragraph_xml_info(p)
        
        # Заголовок с номером (1., 2., 3. и т.д.) - должен быть капсом
        if re.match(r'^\d+\.\s+[А-ЯЁ]', txt):
            is_caps = is_all_caps(txt)
            
            if is_caps:
                st.write(f"**Строка {i}:** Заголовок раздела (с номером, капсом)")
                st.write(f"  Текст: '{txt[:100]}'")
                st.write(f"  Отступ первой строки: {first_line:.3f} см")
                st.write(f"  XML: {xml_info}")
                
                if abs(first_line) > 0.1:
                    st.write(f"  🔴 **Ошибка: Заголовок раздела не должен иметь отступ! Требуется 0 см, сейчас {first_line:.1f} см**")
                    issues_found += 1
                else:
                    st.write(f"  ✅ Отступ отсутствует — правильно")
                
                section_headers.append(('numbered_caps', i, txt[:100], first_line))
        
        # Текст полностью капсом и длинный (может быть заголовок без номера)
        elif is_all_caps(txt) and len(txt) > 25:
            is_list_item = is_list_marker(txt)
            
            if not is_list_item:
                st.write(f"**Строка {i}:** Возможный заголовок раздела (капсом, без номера)")
                st.write(f"  Текст: '{txt[:100]}'")
                st.write(f"  Отступ первой строки: {first_line:.3f} см")
                st.write(f"  XML: {xml_info}")
                
                if abs(first_line) > 0.1:
                    st.write(f"  🔴 **Вероятная ошибка: Заголовок раздела не должен иметь отступ! Требуется 0 см, сейчас {first_line:.1f} см**")
                    issues_found += 1
                else:
                    st.write(f"  ✅ Отступ отсутствует — правильно")
                
                section_headers.append(('caps_no_number', i, txt[:100], first_line))
        
        st.write("---")
        
        if len(section_headers) >= 40:
            break
    
    if issues_found == 0:
        st.success(f"✅ Все заголовки разделов оформлены правильно (без отступа)")
    else:
        st.error(f"❌ Найдено {issues_found} заголовков с ошибочным отступом")
    
    # ============================================
    # ТЕСТ 2: Проверка ВСЕХ видов списков
    # ============================================
    st.subheader("📝 Тест 2: Проверка всех элементов списков")
    st.write("Проверяем маркированные списки (•, -, *) и нумерованные (1), а), и т.д.)")
    
    list_items_found = 0
    list_items_correct = 0
    list_items_wrong = 0
    list_types = {}
    
    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        
        if not txt or has_page_number(txt):
            continue
        
        # Проверяем, является ли параграф элементом списка
        if is_list_marker(txt):
            first_line = get_effective_first_line_indent(p)
            left_indent = get_effective_left_indent(p)
            xml_info = get_paragraph_xml_info(p)
            
            list_items_found += 1
            
            # Определяем тип маркера
            marker_type = "другой"
            if re.match(r'^\d+\)', txt):
                marker_type = "нумерованный 1)"
            elif re.match(r'^[а-яё]\)', txt):
                marker_type = "буквенный а)"
            elif re.match(r'^[a-z]\)', txt):
                marker_type = "буквенный a)"
            elif re.match(r'^[\-–—]', txt):
                marker_type = "тире"
            elif re.match(r'^[•▪○▸▹►▻◆◇]', txt):
                marker_type = "маркер •"
            elif re.match(r'^\*\s', txt):
                marker_type = "звездочка *"
            
            list_types[marker_type] = list_types.get(marker_type, 0) + 1
            
            # Показываем все элементы списка (или первые 25)
            if list_items_found <= 25:
                st.write(f"**Элемент списка {list_items_found} [{marker_type}] (строка {i}):** '{txt[:80]}...'")
                st.write(f"  • Отступ слева: {left_indent:.3f} см")
                st.write(f"  • Отступ первой строки: {first_line:.3f} см")
                st.write(f"  • XML: {xml_info}")
                
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
    else:
        st.write(f"📊 **Всего элементов списков: {list_items_found}**")
        st.write(f"  ✅ С правильным отступом (1.0 см): {list_items_correct}")
        st.write(f"  ❌ Требуют исправления: {list_items_wrong}")
        
        st.write(f"📊 **Типы найденных маркеров:**")
        for marker_type, count in sorted(list_types.items()):
            st.write(f"  • {marker_type}: {count} шт.")
        
        if list_items_wrong == 0:
            st.success("✅ Все элементы списков имеют правильный отступ 1.0 см")
        else:
            st.error(f"❌ {list_items_wrong} элементов списка требуют установки отступа 1.0 см")
    
    # ============================================
    # ТЕСТ 3: Анализ таблиц
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
        
        # Определяем позицию следующей таблицы
        next_table_pos = end_body_pos
        for next_tbl_pos, _ in main_tables[t_idx:]:
            if next_tbl_pos > tbl_pos:
                next_table_pos = next_tbl_pos
                break
        
        # Проверяем, переносится ли таблица на несколько страниц
        spans_pages = table_spans_multiple_pages(doc, table, tbl_pos, next_table_pos)
        
        st.write(f"  📏 Перенос на несколько страниц: {'✅ Да' if spans_pages else '❌ Нет (или не обнаружено)'}")
        
        # Ищем подпись таблицы
        caption_para = None
        caption_idx = None
        for i in range(tbl_pos - 1, start_body_pos - 1, -1):
            elem = doc.element.body[i]
            if elem.tag.endswith('p'):
                for idx, para in enumerate(doc.paragraphs):
                    if para._element is elem and para.text.strip():
                        caption_para = para
                        caption_idx = idx
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
                    st.write(f"  ⚠️ Используется двойной дефис (--), замените на тире (—)")
                elif ' - ' in caption_text:
                    st.write(f"  ⚠️ Используется дефис вместо тире")
                
                if caption_text.rstrip().endswith("."):
                    st.write(f"  ❌ Точка в конце подписи (удалите)")
                else:
                    st.write(f"  ✅ Нет точки в конце")
            else:
                st.write(f"  ⚠️ Подпись не начинается с 'Таблица'")
        
        # Проверка на перенос
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
                st.write(f"     • На второй странице: 'Продолжение таблицы {t_idx}'")
                st.write(f"     • На последней странице: 'Окончание таблицы {t_idx}'")
        else:
            st.write(f"  ✅ Таблица на одной странице, перенос не требуется")
        
        # Показываем первые строки таблицы
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
    st.write(f"• Элементов списков: {list_items_found} (правильных: {list_items_correct}, с ошибками: {list_items_wrong})")
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
st.set_page_config(page_title="Тест проверки документа v4", layout="wide")
st.title("🧪 Тест проверки: заголовки, списки, таблицы")
st.write("Проверяет:")
st.write("1. **Заголовки разделов капсом** — не должны иметь абзацный отступ")
st.write("2. **Все виды списков** (•, -, 1), а)) — должны иметь отступ 1.0 см")
st.write("3. **Таблицы** — проверка переносов, подписей и маркеров Продолжение/Окончание")

uploaded_file = st.file_uploader("Загрузите документ .docx для теста", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем..."):
        test_document(uploaded_file)
