import streamlit as st
import docx
from docx.oxml.ns import qn
import re

def has_page_number(text):
    if re.search(r'[\t\s\.]{2,}\d+$', text):
        return True
    return False

def is_paragraph_bold(paragraph):
    """Проверяет, является ли текст параграфа полужирным"""
    # Способ 1: проверяем runs
    runs = [r for r in paragraph.runs if r.text.strip()]
    if runs:
        # Если все runs с текстом жирные - параграф жирный
        if all(r.bold for r in runs):
            return True, "все runs жирные"
        # Если ни один run не жирный - параграф не жирный
        if not any(r.bold for r in runs):
            return False, "ни один run не жирный"
        # Часть runs жирные, часть нет
        bold_runs = sum(1 for r in runs if r.bold)
        total_runs = len(runs)
        return None, f"частично: {bold_runs}/{total_runs} runs жирные"
    
    # Способ 2: проверяем стиль параграфа
    try:
        if paragraph.style and paragraph.style.font and paragraph.style.font.bold:
            return True, "стиль жирный"
    except:
        pass
    
    # Способ 3: проверяем XML напрямую
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            # Проверяем rPr в pPr (стиль параграфа)
            pPr_rPr = pPr.find(qn('w:rPr'))
            if pPr_rPr is not None:
                bold_elem = pPr_rPr.find(qn('w:b'))
                if bold_elem is not None:
                    return True, "XML pPr/rPr жирный"
            
            # Проверяем rPr в каждом run
            for r in paragraph._element.findall(qn('w:r')):
                rPr = r.find(qn('w:rPr'))
                if rPr is not None:
                    bold_elem = rPr.find(qn('w:b'))
                    if bold_elem is not None:
                        # Проверяем, не установлен ли bold в false
                        val = bold_elem.get(qn('w:val'))
                        if val != 'false' and val != '0':
                            return True, "XML run жирный"
    except:
        pass
    
    return False, "не удалось определить"

def get_paragraph_bold_details(paragraph):
    """Детальная информация о жирности"""
    details = []
    
    # Информация о runs
    for i, run in enumerate(paragraph.runs):
        if run.text.strip():
            bold_status = "жирный" if run.bold else "обычный"
            details.append(f"  Run {i}: '{run.text[:30]}...' - {bold_status}")
    
    # Информация из стиля
    try:
        style_name = paragraph.style.name if paragraph.style else "нет стиля"
        style_bold = paragraph.style.font.bold if paragraph.style and paragraph.style.font else "не определено"
        details.append(f"  Стиль: {style_name}, bold в стиле: {style_bold}")
    except:
        details.append(f"  Стиль: ошибка чтения")
    
    # Информация из XML
    try:
        xml_str = etree.tostring(paragraph._element, encoding='unicode')
        # Ищем все упоминания жирности
        if '<w:b/>' in xml_str or '<w:b ' in xml_str:
            details.append(f"  XML: содержит теги жирности")
        else:
            details.append(f"  XML: нет тегов жирности")
    except:
        pass
    
    return details

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

def test_document(file):
    doc = docx.Document(file)
    
    st.header("🔍 Проверка жирности заголовков")
    
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
    
    toc_entries = extract_toc_entries(doc, start_idx)
    
    st.write("---")
    st.write("**Проверка заголовков:**")
    st.write("Требование: заголовки разделов должны быть ПОЛУЖИРНЫМ")
    
    headers_checked = 0
    headers_bold = 0
    headers_not_bold = 0
    headers_partial = 0
    
    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        
        if not txt or has_page_number(txt):
            continue
        
        # Определяем, является ли заголовком
        is_header = False
        header_type = ""
        
        if is_section_header(txt):
            is_header = True
            header_type = "раздел"
        else:
            # Проверяем по содержанию
            normalized = normalize_title(txt)
            in_toc = any(normalize_title(e) == normalized for e in toc_entries)
            if in_toc and len(txt) > 20:
                is_header = True
                header_type = "из содержания"
            elif is_all_caps(txt) and len(txt) > 25:
                is_header = True
                header_type = "капсом"
        
        if is_header:
            headers_checked += 1
            
            is_bold, bold_info = is_paragraph_bold(p)
            
            st.write(f"**Заголовок {headers_checked}** (строка {i}, {header_type}):")
            st.write(f"  '{txt[:100]}'")
            st.write(f"  Жирность: {bold_info}")
            
            if is_bold == True:
                st.write(f"  ✅ Заголовок полужирный - правильно")
                headers_bold += 1
            elif is_bold == False:
                st.write(f"  ❌ Заголовок НЕ полужирный - нужно сделать жирным!")
                headers_not_bold += 1
            elif is_bold is None:
                st.write(f"  ⚠️ Заголовок частично жирный - проверьте!")
                st.write(f"  Детали:")
                details = get_paragraph_bold_details(p)
                for detail in details:
                    st.write(detail)
                headers_partial += 1
            
            # Показываем детали для первых 5 заголовков
            if headers_checked <= 5:
                st.write(f"  Детали форматирования:")
                details = get_paragraph_bold_details(p)
                for detail in details:
                    st.write(detail)
            
            st.write("---")
        
        if headers_checked >= 30:
            break
    
    # Итоги
    st.write("---")
    st.write("**Итоги проверки жирности:**")
    st.write(f"• Проверено заголовков: {headers_checked}")
    st.write(f"• ✅ Правильно (жирные): {headers_bold}")
    st.write(f"• ❌ Не жирные: {headers_not_bold}")
    st.write(f"• ⚠️ Частично жирные: {headers_partial}")
    
    if headers_not_bold == 0 and headers_partial == 0:
        st.success("✅ Все заголовки правильно оформлены (полужирные)")
    elif headers_not_bold > 0:
        st.error(f"❌ {headers_not_bold} заголовков нужно сделать полужирными")
    elif headers_partial > 0:
        st.warning(f"⚠️ {headers_partial} заголовков имеют частичную жирность")


# Интерфейс
st.set_page_config(page_title="Проверка жирности заголовков", layout="wide")
st.title("🔍 Проверка жирности заголовков разделов")
st.write("Проверяет, являются ли заголовки разделов полужирными")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализ..."):
        test_document(uploaded_file)
