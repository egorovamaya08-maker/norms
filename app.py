import streamlit as st
import docx
from docx.oxml.ns import qn
import re

def has_page_number(text):
    if re.search(r'[\t\s\.]{2,}\d+$', text):
        return True
    return False

def get_table_depth(table):
    """Определяет, не является ли таблица вложенной в другую таблицу"""
    depth = 0
    element = table._element
    parent = element.getparent()
    
    while parent is not None:
        if parent.tag == qn('w:tbl'):
            depth += 1
        parent = parent.getparent()
    
    return depth

def find_table_continuation_markers(doc, start_pos, end_pos, table_num):
    """Ищет Продолжение/Окончание таблицы и возвращает их позиции"""
    markers_found = []
    for i in range(start_pos, min(end_pos + 1, len(doc.paragraphs))):
        txt = doc.paragraphs[i].text.strip()
        if txt:
            if re.search(r'(?:Продолжение|Окончание)\s+таблицы?\s*' + str(table_num), txt):
                markers_found.append((i, txt[:100]))
    return markers_found

def is_all_caps(text):
    clean_text = re.sub(r'[\d\s\.,;:!?\-–—()«»""''\-\+–—]', '', text)
    if not clean_text:
        return False
    return clean_text == clean_text.upper()

def test_document(file):
    doc = docx.Document(file)
    
    st.header("📊 Проверка таблиц")
    st.write("Проверяем: подписи, тире, точки, переносы, исключаем части таблиц после Продолжение/Окончание")
    
    # Ищем начало основного текста
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    start_idx = None
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt or has_page_number(txt):
            continue
        if txt.upper() in level1_keywords or (re.match(r'^\d+\.\s+[А-ЯЁ]', txt) and is_all_caps(txt)):
            start_idx = i
            st.write(f"✅ Начало текста: строка {i} ('{txt[:80]}')")
            break
    
    if start_idx is None:
        st.warning("⚠️ Не найдено начало основного текста")
        return
    
    # Ищем список литературы
    lit_start = None
    for i in range(start_idx, len(doc.paragraphs)):
        if doc.paragraphs[i].text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            lit_start = i
            break
    
    try:
        start_element = doc.paragraphs[start_idx]._element
        start_body_pos = list(doc.element.body).index(start_element)
    except:
        start_body_pos = 0
    
    end_body_pos = len(doc.element.body)
    if lit_start:
        try:
            lit_element = doc.paragraphs[lit_start]._element
            end_body_pos = list(doc.element.body).index(lit_element)
        except:
            pass
    
    # Собираем все таблицы верхнего уровня в основном тексте
    all_main_tables = []
    for table in doc.tables:
        depth = get_table_depth(table)
        if depth > 0:
            continue
        
        try:
            tbl_pos = list(doc.element.body).index(table._element)
        except:
            continue
        
        if start_body_pos < tbl_pos < end_body_pos:
            all_main_tables.append((tbl_pos, table))
    
    st.write(f"Всего таблиц верхнего уровня в тексте: {len(all_main_tables)}")
    
    # Находим все маркеры Продолжение/Окончание
    all_continuation_positions = set()
    for tbl_pos, table in all_main_tables:
        # Ищем номер таблицы в подписи
        caption = None
        for i in range(tbl_pos - 1, start_body_pos - 1, -1):
            elem = doc.element.body[i]
            if elem.tag.endswith('p'):
                for para in doc.paragraphs:
                    if para._element is elem and para.text.strip():
                        caption = para.text.strip()
                        break
                if caption:
                    break
        
        if caption:
            # Извлекаем номер таблицы из подписи
            match = re.match(r'Таблица\s+([\d.]+)', caption)
            if match:
                table_num = match.group(1)
                # Ищем Продолжение/Окончание для этой таблицы
                next_table_pos = end_body_pos
                for next_tbl_pos, _ in all_main_tables:
                    if next_tbl_pos > tbl_pos:
                        next_table_pos = next_tbl_pos
                        break
                
                markers = find_table_continuation_markers(doc, tbl_pos, next_table_pos, table_num)
                for marker_pos, _ in markers:
                    all_continuation_positions.add(marker_pos)
    
    st.write(f"Найдено маркеров Продолжение/Окончание: {len(all_continuation_positions)}")
    for pos in sorted(all_continuation_positions):
        st.write(f"  • Строка {pos}: '{doc.paragraphs[pos].text.strip()[:80]}'")
    
    # Определяем, какие таблицы находятся после маркеров Продолжение/Окончание
    # Таблица считается продолжением, если между ней и предыдущим маркером нет другой подписи
    continuation_tables = set()
    
    for idx, (tbl_pos, table) in enumerate(all_main_tables):
        # Проверяем, есть ли перед таблицей маркер Продолжение/Окончание
        # и нет ли между ними подписи другой таблицы
        for marker_pos in sorted(all_continuation_positions):
            if marker_pos < tbl_pos:
                # Проверяем, есть ли между маркером и таблицей подпись другой таблицы
                has_caption_between = False
                for i in range(marker_pos + 1, tbl_pos):
                    txt = doc.paragraphs[i].text.strip()
                    if txt and re.match(r'Таблица\s+[\d.]+\s+[–—]', txt):
                        has_caption_between = True
                        break
                
                if not has_caption_between:
                    continuation_tables.add(tbl_pos)
    
    st.write(f"Таблиц-продолжений (после Продолжение/Окончание): {len(continuation_tables)}")
    
    # Оставляем только основные таблицы (не продолжения)
    main_tables = [(pos, tbl) for pos, tbl in all_main_tables if pos not in continuation_tables]
    
    st.write(f"Основных таблиц для проверки: {len(main_tables)}")
    
    # Показываем, какие таблицы пропущены
    if continuation_tables:
        st.write("---")
        st.write("**Пропущенные таблицы (части после Продолжение/Окончание):**")
        for tbl_pos in sorted(continuation_tables):
            st.write(f"  • Таблица на позиции {tbl_pos} - это продолжение/окончание")
    
    # Проверяем основные таблицы
    st.write("---")
    st.write("**Проверка основных таблиц:**")
    
    tables_need_check = 0
    
    for t_idx, (tbl_pos, table) in enumerate(main_tables, start=1):
        st.write(f"---")
        st.write(f"**Таблица {t_idx}** (позиция {tbl_pos})")
        st.write(f"  Размер: {len(table.rows)} строк × {len(table.columns)} столбцов")
        
        # Ищем подпись
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
            st.write(f"  Подпись: '{cap_text[:150]}'")
            
            if cap_text.startswith("Таблица"):
                # Извлекаем номер таблицы
                match = re.match(r'Таблица\s+([\d.]+)', cap_text)
                table_num = match.group(1) if match else str(t_idx)
                
                if '—' in cap_text or '–' in cap_text:
                    st.write(f"  ✅ Тире правильное")
                elif '--' in cap_text:
                    st.write(f"  ⚠️ Используется двойной дефис (--), нужно заменить на тире")
                elif ' - ' in cap_text:
                    st.write(f"  ⚠️ Используется дефис, нужно заменить на тире")
                
                if cap_text.rstrip().endswith("."):
                    st.write(f"  ❌ Точка в конце подписи (удалите)")
                else:
                    st.write(f"  ✅ Нет точки в конце")
                
                # Проверка на перенос (для таблиц больше 2 строк)
                if len(table.rows) > 2:
                    next_table_pos = end_body_pos
                    for next_tbl_pos, _ in main_tables[t_idx:]:
                        if next_tbl_pos > tbl_pos:
                            next_table_pos = next_tbl_pos
                            break
                    
                    markers = find_table_continuation_markers(doc, tbl_pos, next_table_pos, table_num)
                    
                    if not markers:
                        st.write(f"  🔴 **Нет «Продолжение таблицы {table_num}» / «Окончание таблицы {table_num}»**")
                        st.write(f"  💡 Если таблица на нескольких страницах, добавьте:")
                        st.write(f"     • «Продолжение таблицы {table_num}» на следующей странице")
                        st.write(f"     • «Окончание таблицы {table_num}» на последней странице")
                        tables_need_check += 1
                    else:
                        st.write(f"  ✅ Маркеры переноса найдены:")
                        for m_pos, m_text in markers:
                            st.write(f"    • Строка {m_pos}: '{m_text}'")
            else:
                st.write(f"  ⚠️ Подпись не начинается с 'Таблица'")
        else:
            st.write(f"  ⚠️ Подпись не найдена перед таблицей")
        
        # Показываем первые строки
        st.write(f"  📋 Первые 3 строки:")
        for r_idx, row in enumerate(table.rows[:3]):
            cells_text = []
            for cell in row.cells:
                cells_text.append(cell.text.strip()[:30])
            st.write(f"    Строка {r_idx + 1}: {' | '.join(cells_text)}")
    
    st.write("---")
    st.write("**Итог:**")
    st.write(f"• Всего таблиц верхнего уровня: {len(all_main_tables)}")
    st.write(f"• Пропущено (продолжения): {len(continuation_tables)}")
    st.write(f"• Проверено основных таблиц: {len(main_tables)}")
    st.write(f"• Требуют проверки на перенос: {tables_need_check}")


# Интерфейс
st.set_page_config(page_title="Тест таблиц v2", layout="wide")
st.title("📊 Тест проверки таблиц")
st.write("Проверяет:")
st.write("• Исключает таблицы после «Продолжение таблицы» / «Окончание таблицы»")
st.write("• Наличие подписи 'Таблица N – Название'")
st.write("• Правильность тире и отсутствие точки в конце")
st.write("• Наличие маркеров переноса")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализ таблиц..."):
        test_document(uploaded_file)
