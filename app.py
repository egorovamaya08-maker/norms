import streamlit as st
import docx
from docx.oxml.ns import qn
import re

def has_page_number(text):
    if re.search(r'[\t\s\.]{2,}\d+$', text):
        return True
    return False

def get_table_depth(table):
    """
    Определяет, не является ли таблица вложенной в другую таблицу.
    Возвращает глубину вложенности (0 = верхний уровень)
    """
    depth = 0
    element = table._element
    parent = element.getparent()
    
    while parent is not None:
        if parent.tag == qn('w:tbl'):
            depth += 1
        parent = parent.getparent()
    
    return depth

def find_table_continuation_markers(doc, start_pos, end_pos, table_num):
    """Ищет Продолжение/Окончание таблицы"""
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
    st.write("Проверяем: подписи, тире, точки, переносы, вложенные таблицы")
    
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
    
    # Собираем ВСЕ таблицы и анализируем их
    st.write("---")
    st.write("**Анализ всех таблиц в документе:**")
    
    all_tables = []
    for table in doc.tables:
        depth = get_table_depth(table)
        try:
            tbl_pos = list(doc.element.body).index(table._element)
        except:
            tbl_pos = -1
        
        all_tables.append({
            'table': table,
            'depth': depth,
            'pos': tbl_pos,
            'rows': len(table.rows),
            'cols': len(table.columns)
        })
    
    st.write(f"Всего таблиц в документе: {len(all_tables)}")
    
    # Показываем все таблицы с их статусом
    for i, t in enumerate(all_tables, 1):
        in_main = start_body_pos < t['pos'] < end_body_pos
        
        status = ""
        if t['depth'] > 0:
            status = f"❌ ВЛОЖЕННАЯ (глубина {t['depth']}) - ПРОПУСКАЕМ"
        elif not in_main:
            status = "❌ Вне основного текста - ПРОПУСКАЕМ"
        else:
            status = "✅ Таблица основного текста - ПРОВЕРЯЕМ"
        
        st.write(f"**Таблица {i}:** позиция {t['pos']}, {t['rows']}×{t['cols']}, {status}")
    
    # Теперь проверяем только таблицы основного текста
    st.write("---")
    st.write("**Проверка таблиц основного текста:**")
    
    main_tables = [(t['pos'], t['table']) for t in all_tables 
                   if t['depth'] == 0 and start_body_pos < t['pos'] < end_body_pos]
    
    st.write(f"Таблиц для проверки: {len(main_tables)}")
    
    if len(main_tables) == 0:
        st.warning("⚠️ Нет таблиц основного текста для проверки")
        return
    
    tables_need_check = 0
    
    for t_idx, (tbl_pos, table) in enumerate(main_tables, start=1):
        st.write(f"---")
        st.write(f"**Проверяемая таблица {t_idx}** (позиция {tbl_pos})")
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
                if '—' in cap_text or '–' in cap_text:
                    st.write(f"  ✅ Тире правильное")
                elif '--' in cap_text:
                    st.write(f"  ⚠️ Используется двойной дефис (--), нужно заменить на тире")
                elif ' - ' in cap_text:
                    st.write(f"  ⚠️ Используется дефис, нужно заменить на тире")
                else:
                    st.write(f"  ⚠️ Проверьте наличие тире после номера таблицы")
                
                if cap_text.rstrip().endswith("."):
                    st.write(f"  ❌ Точка в конце подписи (удалите)")
                else:
                    st.write(f"  ✅ Нет точки в конце")
            else:
                st.write(f"  ⚠️ Подпись не начинается с 'Таблица'")
        else:
            st.write(f"  ⚠️ Подпись не найдена перед таблицей")
        
        # Проверка на перенос
        if len(table.rows) > 2:
            next_table_pos = end_body_pos
            for next_tbl_pos, _ in main_tables[t_idx:]:
                if next_tbl_pos > tbl_pos:
                    next_table_pos = next_tbl_pos
                    break
            
            markers = find_table_continuation_markers(doc, tbl_pos, next_table_pos, t_idx)
            
            if not markers:
                st.write(f"  🔴 **Нет «Продолжение таблицы {t_idx}» / «Окончание таблицы {t_idx}»**")
                st.write(f"  💡 Если таблица на нескольких страницах, добавьте:")
                st.write(f"     • «Продолжение таблицы {t_idx}» на следующей странице")
                st.write(f"     • «Окончание таблицы {t_idx}» на последней странице")
                tables_need_check += 1
            else:
                st.write(f"  ✅ Маркеры переноса найдены:")
                for m_pos, m_text in markers:
                    st.write(f"    • Строка {m_pos}: '{m_text}'")
        else:
            st.write(f"  ✅ Таблица маленькая, перенос маловероятен")
        
        # Показываем первые строки
        st.write(f"  📋 Первые 3 строки:")
        for r_idx, row in enumerate(table.rows[:3]):
            cells_text = []
            for cell in row.cells:
                cells_text.append(cell.text.strip()[:30])
            st.write(f"    Строка {r_idx + 1}: {' | '.join(cells_text)}")
    
    st.write("---")
    st.write("**Итог:**")
    st.write(f"• Всего таблиц в документе: {len(all_tables)}")
    st.write(f"• Вложенных (пропущено): {sum(1 for t in all_tables if t['depth'] > 0)}")
    st.write(f"• Вне основного текста (пропущено): {sum(1 for t in all_tables if t['depth'] == 0 and not (start_body_pos < t['pos'] < end_body_pos))}")
    st.write(f"• Проверено: {len(main_tables)}")
    st.write(f"• Требуют проверки на перенос: {tables_need_check}")


# Интерфейс
st.set_page_config(page_title="Тест таблиц", layout="wide")
st.title("📊 Тест проверки таблиц")
st.write("Проверяет:")
st.write("• Не вложена ли таблица в другую таблицу")
st.write("• Наличие подписи 'Таблица N – Название'")
st.write("• Правильность тире и отсутствие точки в конце")
st.write("• Наличие «Продолжение таблицы» / «Окончание таблицы»")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализ таблиц..."):
        test_document(uploaded_file)
