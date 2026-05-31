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

def test_document(file):
    """Тестовая функция для проверки двух проблем"""
    doc = docx.Document(file)
    
    st.header("🔍 Анализ документа")
    
    # ============================================
    # ТЕСТ 1: Поиск пунктов содержания
    # ============================================
    st.subheader("📋 Тест 1: Поиск содержания и основного текста")
    
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    start_idx = None
    toc_items = []
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt:
            continue
        
        # Проверяем на номер страницы (признак содержания)
        has_page = has_page_number(txt)
        
        # Проверяем, является ли пункт нумерованным и НЕ капсом
        is_numbered_not_caps = bool(re.match(r'^\d+\.\s+[А-Яа-я]', txt)) and txt != txt.upper()
        
        if has_page:
            st.write(f"📄 Строка {i}: **Содержание (с номером страницы)**: '{txt[:80]}' → ПРОПУСКАЕМ")
            toc_items.append(('page_number', i, txt[:80]))
            continue
        
        if is_numbered_not_caps:
            st.write(f"📑 Строка {i}: **Содержание (нумерованный пункт не капсом)**: '{txt[:80]}' → ПРОПУСКАЕМ")
            toc_items.append(('numbered_not_caps', i, txt[:80]))
            continue
        
        # Ищем начало основного текста
        if txt.upper() in level1_keywords or (re.match(r'^\d+\.\s+[А-ЯЁ]', txt) and txt == txt.upper()):
            start_idx = i
            st.write(f"✅ Строка {i}: **Начало основного текста**: '{txt[:80]}'")
            break
    
    if start_idx is None:
        st.warning("⚠️ Не найдено начало основного текста")
        return
    
    st.write(f"📌 Индекс начала основного текста: {start_idx}")
    
    # Проверяем пункты до начала основного текста
    st.write("---")
    st.write("**Пункты, которые были правильно проигнорированы:**")
    for item_type, idx, text in toc_items:
        if idx < start_idx:
            st.write(f"✅ Правильно проигнорирован: '{text}'")
    
    # ============================================
    # ТЕСТ 2: Проверка списка источников
    # ============================================
    st.subheader("📚 Тест 2: Проверка списка источников")
    
    # Ищем "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"
    lit_start = None
    for i in range(start_idx, len(doc.paragraphs)):
        txt = doc.paragraphs[i].text.strip()
        if txt.upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" and not has_page_number(txt):
            lit_start = i
            st.write(f"✅ Найден заголовок списка источников на строке {i}")
            break
    
    if lit_start is None:
        st.warning("⚠️ Список источников не найден")
        return
    
    # Ищем конец списка источников (начало приложений)
    lit_end = len(doc.paragraphs)  # По умолчанию - до конца документа
    appendix_keywords = ["ПРИЛОЖЕНИЕ", "ПРИЛОЖЕНИЯ", "APPENDIX"]
    
    for i in range(lit_start + 1, len(doc.paragraphs)):
        txt = doc.paragraphs[i].text.strip().upper()
        # Ищем заголовок приложения
        if any(txt.startswith(kw) for kw in appendix_keywords):
            lit_end = i
            st.write(f"🛑 Найден конец списка источников на строке {i}: '{doc.paragraphs[i].text.strip()[:80]}'")
            break
    
    st.write(f"📌 Проверяем источники со строки {lit_start + 1} до строки {lit_end - 1}")
    
    # Анализируем каждый источник
    st.write("---")
    st.write("**Анализ каждого источника:**")
    
    source_count = 0
    sources_with_issues = 0
    
    for i in range(lit_start + 1, lit_end):  # Используем lit_end вместо len(doc.paragraphs)
        source = doc.paragraphs[i]
        txt = source.text.strip()
        
        if not txt:
            continue
        
        if has_page_number(txt):
            continue
        
        # Дополнительная проверка: если строка начинается с ключевых слов приложений - пропускаем
        if any(txt.upper().startswith(kw) for kw in appendix_keywords):
            st.write(f"⚠️ Пропущен заголовок приложения: '{txt[:80]}'")
            continue
        
        source_count += 1
        has_issue = False
        
        left_indent = get_effective_left_indent(source)
        first_line = get_effective_first_line_indent(source)
        xml_info = get_paragraph_xml_info(source)
        
        st.write(f"**Источник {source_count}:** '{txt[:60]}...'")
        st.write(f"  • Отступ слева: {left_indent:.3f} см {'✅' if abs(left_indent) < 0.1 else '❌ (должен быть 0 см)'}")
        st.write(f"  • Отступ первой строки: {first_line:.3f} см {'✅' if abs(first_line - 1.0) < 0.1 else '❌ (должен быть 1.0 см)'}")
        st.write(f"  • XML: {xml_info}")
        
        # Проверка XML на наличие hanging
        try:
            pPr = source._element.find(qn('w:pPr'))
            if pPr is not None:
                ind = pPr.find(qn('w:ind'))
                if ind is not None:
                    hanging_xml = ind.get(qn('w:hanging'))
                    first_line_xml = ind.get(qn('w:firstLine'))
                    
                    if hanging_xml and not first_line_xml:
                        st.write(f"  ⚠️ Особый случай: hanging={hanging_xml}, firstLine отсутствует")
                        st.write(f"  → Это значит, что отступа первой строки нет (hanging = выступ)")
                        has_issue = True
        except Exception as e:
            st.write(f"  ⚠️ Ошибка при анализе XML: {e}")
        
        if abs(left_indent) > 0.1:
            has_issue = True
        
        if abs(first_line - 1.0) > 0.1:
            has_issue = True
        
        if has_issue:
            sources_with_issues += 1
            st.write(f"  🔴 **Есть проблемы с оформлением**")
        else:
            st.write(f"  🟢 **Оформление правильное**")
        
        st.write("---")
    
    # Показываем, что после списка источников
    if lit_end < len(doc.paragraphs):
        st.write("📄 **Содержимое после списка источников (пропущено при проверке):**")
        for i in range(lit_end, min(lit_end + 5, len(doc.paragraphs))):
            txt = doc.paragraphs[i].text.strip()
            if txt:
                st.write(f"  Строка {i}: '{txt[:80]}...'")
    
    # Итоги
    st.subheader("📊 Итоги проверки")
    st.write(f"• Пунктов содержания проигнорировано: {len(toc_items)}")
    st.write(f"• Источников проверено: {source_count}")
    st.write(f"• Источников с проблемами: {sources_with_issues}")
    
    if sources_with_issues == 0 and source_count > 0:
        st.success("✅ Все источники оформлены правильно!")
    elif sources_with_issues > 0:
        st.error(f"❌ {sources_with_issues} источник(ов) требуют исправления")
    
    # Дополнительная отладка для первого источника
    if source_count > 0:
        st.subheader("🔧 Детальная отладка первого источника")
        first_source = None
        for i in range(lit_start + 1, lit_end):
            if doc.paragraphs[i].text.strip() and not has_page_number(doc.paragraphs[i].text.strip()):
                first_source = doc.paragraphs[i]
                break
        
        if first_source:
            st.write("XML элемента первого источника:")
            st.code(first_source._element.xml[:1000])

# Интерфейс
st.set_page_config(page_title="Тест проверки документа", layout="wide")
st.title("🧪 Тест проверки двух проблем")
st.write("Этот скрипт проверяет только:")
st.write("1. Правильно ли игнорируются пункты содержания (1., 2., если они не капсом)")
st.write("2. Правильно ли проверяются отступы в списке источников (останавливается перед приложениями)")

uploaded_file = st.file_uploader("Загрузите документ .docx для теста", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем..."):
        test_document(uploaded_file)
