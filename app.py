import streamlit as st
import docx
from docx.oxml.ns import qn
import re

def has_page_number(text):
    if re.search(r'[\t\s\.]{2,}\d+$', text):
        return True
    return False

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

def get_table_depth(table):
    depth = 0
    element = table._element
    parent = element.getparent()
    
    while parent is not None:
        if parent.tag == qn('w:tbl'):
            depth += 1
        parent = parent.getparent()
    
    return depth

def diagnose_tables(file):
    doc = docx.Document(file)
    
    st.header("🔍 Диагностика ВСЕХ таблиц в документе")
    
    # Ищем начало основного текста
    start_idx = None
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt or has_page_number(txt):
            continue
        if txt.upper() in level1_keywords or (re.match(r'^\d+\.\s+[А-ЯЁ]', txt) and is_all_caps(txt)):
            start_idx = i
            st.write(f"✅ Начало текста: строка {i}")
            break
    
    if start_idx is None:
        st.warning("⚠️ Не найдено начало текста")
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
    
    st.write(f"Диапазон body: {start_body_pos} - {end_body_pos}")
    st.write(f"Всего таблиц в документе: {len(doc.tables)}")
    
    # Анализируем КАЖДУЮ таблицу
    st.write("---")
    st.write("**Все таблицы:**")
    
    in_range_count = 0
    out_of_range_count = 0
    nested_count = 0
    empty_count = 0
    
    for t_idx, table in enumerate(doc.tables, start=1):
        depth = get_table_depth(table)
        
        try:
            tbl_pos = list(doc.element.body).index(table._element)
            in_body = True
        except:
            tbl_pos = -1
            in_body = False
        
        rows = len(table.rows)
        cols = len(table.columns)
        
        # Собираем текст из первых 3 строк
        preview = []
        for row in table.rows[:3]:
            cell_texts = []
            for cell in row.cells[:3]:
                cell_texts.append(cell.text.strip()[:30])
            preview.append(" | ".join(cell_texts))
        
        in_range = start_body_pos < tbl_pos < end_body_pos
        
        st.write(f"**Таблица {t_idx}:**")
        st.write(f"  • Позиция в body: {tbl_pos} {'✅ в диапазоне' if in_range else '❌ ВНЕ диапазона'}")
        st.write(f"  • Вложенность: {depth} {'⚠️ ВЛОЖЕННАЯ' if depth > 0 else ''}")
        st.write(f"  • Размер: {rows}×{cols}")
        st.write(f"  • Первые строки:")
        for line in preview:
            if line.strip():
                st.write(f"    '{line}'")
        
        # Ищем подпись
        if in_range and depth == 0:
            caption = None
            for i in range(tbl_pos - 1, max(start_body_pos - 1, -1), -1):
                elem = doc.element.body[i]
                if elem.tag.endswith('p'):
                    for para in doc.paragraphs:
                        if para._element is elem and para.text.strip():
                            caption = para.text.strip()
                            break
                    if caption:
                        break
            
            if caption:
                st.write(f"  • Подпись: '{caption[:100]}'")
            else:
                st.write(f"  • ⚠️ Подпись НЕ НАЙДЕНА")
        
        # Считаем статистику
        if in_range:
            in_range_count += 1
        else:
            out_of_range_count += 1
        
        if depth > 0:
            nested_count += 1
        
        if rows == 0:
            empty_count += 1
        
        st.write("---")
    
    st.header("📊 Статистика")
    st.write(f"• Всего таблиц: {len(doc.tables)}")
    st.write(f"• В диапазоне текста: {in_range_count}")
    st.write(f"• ВНЕ диапазона: {out_of_range_count}")
    st.write(f"• Вложенных: {nested_count}")
    st.write(f"• Пустых: {empty_count}")
    
    # Показываем таблицы ВНЕ диапазона
    st.header("⚠️ Таблицы ВНЕ основного текста")
    for t_idx, table in enumerate(doc.tables, start=1):
        try:
            tbl_pos = list(doc.element.body).index(table._element)
            in_range = start_body_pos < tbl_pos < end_body_pos
        except:
            tbl_pos = -1
            in_range = False
        
        if not in_range:
            st.write(f"Таблица {t_idx}: позиция {tbl_pos} (диапазон: {start_body_pos}-{end_body_pos})")
            # Показываем содержимое
            for row in table.rows[:2]:
                for cell in row.cells[:2]:
                    if cell.text.strip():
                        st.write(f"  '{cell.text.strip()[:50]}'")

# Интерфейс
st.set_page_config(page_title="Диагностика таблиц", layout="wide")
st.title("🔍 Диагностика таблиц")
st.write("Выясняем, откуда берутся лишние таблицы")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем..."):
        diagnose_tables(uploaded_file)
