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

def get_font_size(paragraph):
    """Получает размер шрифта параграфа"""
    try:
        for run in paragraph.runs:
            if run.font.size:
                return run.font.size.pt
        # Пробуем через стиль
        if paragraph.style and paragraph.style.font.size:
            return paragraph.style.font.size.pt
    except:
        pass
    return None

def is_all_caps(text):
    """Проверяет, написан ли текст заглавными буквами (с учетом кириллицы)"""
    # Убираем цифры, пробелы, знаки препинания
    clean_text = re.sub(r'[\d\s\.,;:!?\-–—()«»""''\-\+]', '', text)
    if not clean_text:
        return False
    # Проверяем, что все буквы заглавные
    return clean_text == clean_text.upper()

def test_document(file):
    """Тестовая функция для проверки трёх проблем"""
    doc = docx.Document(file)
    
    st.header("🔍 Анализ документа")
    
    # ============================================
    # ТЕСТ 1: Поиск заголовков разделов капсом
    # ============================================
    st.subheader("📋 Тест 1: Поиск заголовков разделов и подразделов")
    
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    start_idx = None
    toc_items = []
    section_headers = []
    
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
    
    # Ищем все потенциальные заголовки разделов после start_idx
    st.write("---")
    st.write("**Анализ заголовков в тексте:**")
    
    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        
        if not txt or has_page_number(txt):
            continue
        
        first_line = get_effective_first_line_indent(p)
        xml_info = get_paragraph_xml_info(p)
        font_size = get_font_size(p)
        
        # Проверка 1: Заголовок с номером (1., 2., 3. и т.д.) - должен быть капсом
        if re.match(r'^\d+\.\s+[А-ЯЁ]', txt):
            is_caps = is_all_caps(txt)
            st.write(f"**Строка {i}:** Заголовок с номером")
            st.write(f"  Текст: '{txt[:100]}'")
            st.write(f"  Капсом: {'✅ Да' if is_caps else '❌ Нет'}")
            st.write(f"  Отступ первой строки: {first_line:.3f} см {'⚠️ Должен быть 0 для заголовка' if abs(first_line) > 0.1 else '✅'}")
            st.write(f"  XML: {xml_info}")
            st.write(f"  Размер шрифта: {font_size}")
            
            if is_caps and abs(first_line) > 0.1:
                st.write(f"  🔴 **Ошибка: Заголовок раздела не должен иметь отступ!**")
            
            section_headers.append(('numbered', i, txt[:100], is_caps, first_line))
        
        # Проверка 2: Текст полностью капсом и длинный (может быть заголовок без номера)
        elif is_all_caps(txt) and len(txt) > 20:
            # Проверяем, не является ли это частью нумерованного списка
            is_list_item = bool(re.match(r'^\d+\)', txt))
            
            st.write(f"**Строка {i}:** Возможный заголовок без номера (все капсом)")
            st.write(f"  Текст: '{txt[:100]}'")
            st.write(f"  Это элемент списка: {'Да' if is_list_item else 'Нет'}")
            st.write(f"  Отступ первой строки: {first_line:.3f} см")
            st.write(f"  XML: {xml_info}")
            st.write(f"  Размер шрифта: {font_size}")
            
            if not is_list_item and abs(first_line) > 0.1:
                st.write(f"  🔴 **Возможная ошибка: Заголовок раздела не должен иметь отступ!**")
            
            section_headers.append(('caps_no_number', i, txt[:100], True, first_line))
        
        st.write("---")
        
        # Ограничим вывод, чтобы не было слишком много
        if len(section_headers) >= 30:
            break
    
    # ============================================
    # ТЕСТ 2: Поиск нумерованных списков
    # ============================================
    st.subheader("📝 Тест 2: Поиск нумерованных списков")
    
    list_patterns = [
        r'^\d+\)',      # 1) 2) 3)
        r'^[а-яё]\)',   # а) б) в)
        r'^[a-z]\)',    # a) b) c)
        r'^[\-–—•]',    # - – — •
    ]
    
    list_items_found = 0
    
    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        
        if not txt or has_page_number(txt):
            continue
        
        is_list = False
        for pattern in list_patterns:
            if re.match(pattern, txt):
                is_list = True
                break
        
        if is_list:
            first_line = get_effective_first_line_indent(p)
            left_indent = get_effective_left_indent(p)
            xml_info = get_paragraph_xml_info(p)
            
            list_items_found += 1
            st.write(f"**Элемент списка {list_items_found} (строка {i}):** '{txt[:80]}...'")
            st.write(f"  • Отступ слева: {left_indent:.3f} см")
            st.write(f"  • Отступ первой строки: {first_line:.3f} см")
            st.write(f"  • XML: {xml_info}")
            
            # Элементы списка не должны требовать отступ 1.0 см
            if abs(first_line - 1.0) < 0.1:
                st.write(f"  ⚠️ Элемент списка имеет отступ 1.0 см (это может быть неправильно для списка)")
            elif abs(first_line) < 0.1:
                st.write(f"  ✅ Отступ первой строки отсутствует (нормально для списка)")
            
            st.write("---")
            
            if list_items_found >= 20:
                break
    
    if list_items_found == 0:
        st.write("ℹ️ Нумерованные списки не найдены")
    
    # ============================================
    # ТЕСТ 3: Анализ таблиц
    # ============================================
    st.subheader("📊 Тест 3: Анализ таблиц")
    
    # Находим таблицы в основном тексте
    try:
        start_element = doc.paragraphs[start_idx]._element
        start_body_pos = list(doc.element.body).index(start_element)
    except:
        start_body_pos = 0
    
    # Ищем конец основного текста (перед списком литературы)
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
        st.write(f"**Таблица {t_idx}** (позиция в документе: {tbl_pos})")
        st.write(f"  Количество строк: {len(table.rows)}")
        st.write(f"  Количество столбцов: {len(table.columns)}")
        
        # Ищем подпись таблицы (ближайший параграф перед таблицей)
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
            st.write(f"  📝 Подпись: '{caption_text[:100]}'")
            
            # Проверяем формат подписи
            if caption_text.startswith("Таблица"):
                # Проверяем тире
                if '–' in caption_text or '—' in caption_text:
                    st.write(f"  ✅ Используется длинное тире")
                elif '--' in caption_text:
                    st.write(f"  ⚠️ Используется двойное дефис (--), рекомендуется заменить на тире")
                elif ' - ' in caption_text:
                    st.write(f"  ⚠️ Используется дефис вместо тире")
                
                # Проверяем точку в конце
                if caption_text.rstrip().endswith("."):
                    st.write(f"  ❌ Точка в конце подписи (нужно удалить)")
                else:
                    st.write(f"  ✅ Нет точки в конце")
            else:
                st.write(f"  ⚠️ Подпись не начинается с 'Таблица'")
        
        # Проверяем, не переносится ли таблица на следующую страницу
        # Анализируем содержимое таблицы
        st.write(f"  📏 Анализ размера таблицы:")
        
        # Смотрим на следующий параграф после таблицы
        next_para = None
        next_para_idx = None
        for i in range(tbl_pos + 1, end_body_pos):
            elem = doc.element.body[i]
            if elem.tag.endswith('p'):
                for idx, para in enumerate(doc.paragraphs):
                    if para._element is elem:
                        next_para = para
                        next_para_idx = idx
                        break
                break
        
        if next_para and next_para.text.strip():
            st.write(f"  Следующий абзац после таблицы: '{next_para.text.strip()[:100]}'")
            
            # Проверяем, есть ли "Продолжение таблицы" или "Окончание таблицы"
            if re.search(r'Продолжение таблицы|Окончание таблицы', next_para.text.strip()):
                st.write(f"  ✅ Есть указание на продолжение/окончание таблицы")
            else:
                # Анализируем, может ли таблица переноситься
                row_count = len(table.rows)
                if row_count > 5:
                    st.write(f"  ⚠️ Таблица содержит {row_count} строк и может переноситься на следующую страницу")
                    st.write(f"  💡 Рекомендация: проверьте, нужна ли надпись 'Продолжение таблицы {t_idx}' или 'Окончание таблицы {t_idx}'")
        
        # Показываем первые несколько строк таблицы
        st.write(f"  📋 Первые 3 строки таблицы:")
        for r_idx, row in enumerate(table.rows[:3]):
            cells_text = []
            for cell in row.cells:
                cells_text.append(cell.text.strip()[:30])
            st.write(f"    Строка {r_idx + 1}: {' | '.join(cells_text)}")
    
    # ============================================
    # ИТОГИ
    # ============================================
    st.header("📊 Итоги анализа")
    st.write(f"• Найдено заголовков разделов: {len(section_headers)}")
    st.write(f"• Найдено элементов списков: {list_items_found}")
    st.write(f"• Найдено таблиц: {len(main_tables)}")

# Интерфейс
st.set_page_config(page_title="Тест проверки документа v2", layout="wide")
st.title("🧪 Тест проверки: заголовки, списки, таблицы")
st.write("Проверяет:")
st.write("1. Заголовки разделов капсом (с номером и без) — не должны иметь отступ")
st.write("2. Нумерованные списки — не должны требовать отступ 1.0 см")
st.write("3. Таблицы — проверка подписей и переносов")

uploaded_file = st.file_uploader("Загрузите документ .docx для теста", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем..."):
        test_document(uploaded_file)
