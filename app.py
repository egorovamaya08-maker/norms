import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import parse_xml
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

def is_all_caps(text):
    return bool(re.match(r'^[А-ЯЁ\s\-]+$', text)) and len(text) > 3

def find_toc_heading(doc):
    """
    Поиск заголовка СОДЕРЖАНИЕ в TOC (оглавлении)
    """
    # Метод 1: Поиск в стилях
    for i, p in enumerate(doc.paragraphs):
        if p.style and p.style.name:
            style_name = p.style.name.lower()
            # TOC Heading - стандартный стиль для заголовка оглавления в Word
            if 'toc' in style_name or 'content' in style_name or 'оглавл' in style_name:
                # Проверяем текст параграфа
                text_clean = re.sub(r'[^А-Яа-я]', '', p.text.upper())
                if 'СОДЕРЖАНИЕ' in text_clean or 'ОГЛАВЛЕНИЕ' in text_clean:
                    return i
                # Если текст пустой или короткий, возможно это автоматический заголовок
                if len(p.text.strip()) < 20:
                    return i
    
    # Метод 2: Поиск по номеру - содержание обычно на 2-3 странице
    # Проверяем параграфы с 30 по 80 (где обычно находится оглавление)
    for i in range(30, min(80, len(doc.paragraphs))):
        p = doc.paragraphs[i]
        text = p.text.strip()
        # Если параграф содержит цифры и точки - это содержание
        if re.search(r'\d+\.\.+\.+\d+', text) or re.search(r'\d+\s+\.\.\.\s+\d+', text):
            # Ищем параграф-заголовок перед содержанием
            for j in range(max(0, i-5), i):
                prev_text = doc.paragraphs[j].text.strip().upper()
                if 'СОДЕРЖАНИЕ' in prev_text or 'ОГЛАВЛЕНИЕ' in prev_text:
                    return j
                # Если параграф короткий и по центру - возможно это заголовок
                if len(prev_text) < 30 and get_effective_alignment(doc.paragraphs[j]) == WD_ALIGN_PARAGRAPH.CENTER:
                    if prev_text and prev_text[0].isupper():
                        return j
            return i-1 if i > 0 else i
    
    # Метод 3: Ручной поиск в первых 100 параграфах
    for i in range(min(100, len(doc.paragraphs))):
        p = doc.paragraphs[i]
        text = p.text.strip()
        text_upper = text.upper()
        
        # Очищаем от мусора
        cleaned = re.sub(r'[^А-ЯЁ]', '', text_upper)
        
        # Проверяем точное совпадение
        if cleaned == 'СОДЕРЖАНИЕ' or cleaned == 'ОГЛАВЛЕНИЕ':
            return i
        
        # Проверяем, что параграф короткий, по центру и состоит из заглавных
        if len(text) < 30 and get_effective_alignment(p) == WD_ALIGN_PARAGRAPH.CENTER:
            if text_upper == text and len(text) > 3:
                # Это может быть заголовок СОДЕРЖАНИЕ
                if 'СОДЕРЖ' in cleaned or 'ОГЛАВЛ' in cleaned:
                    return i
    
    return None

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

    # ---------- 2. ПОИСК "СОДЕРЖАНИЕ" В TOC ----------
    content_idx = find_toc_heading(doc)
    
    if content_idx is None:
        issues.append("❌ Не найден заголовок «СОДЕРЖАНИЕ» — проверка невозможна.")
        
        # Диагностика: показываем параграфы, которые могут быть содержанием
        issues.append("")
        issues.append("📋 Поиск возможного оглавления (параграфы 30-80):")
        for i in range(30, min(80, len(doc.paragraphs))):
            p = doc.paragraphs[i]
            text = p.text.strip()
            if text and (re.search(r'\d+\.', text) or len(text) < 30):
                issues.append(f"  {i:3d}: {repr(text[:60])}")
                if get_effective_alignment(p) == WD_ALIGN_PARAGRAPH.CENTER:
                    issues.append(f"       ↑ выровнено по центру")
        
        return issues
    
    issues.append(f"✅ Заголовок «СОДЕРЖАНИЕ» найден в параграфе {content_idx}")
    issues.append(f"   Текст: {repr(doc.paragraphs[content_idx].text[:100])}")
    
    # ---------- 3. ПРОВЕРКА ЗАГОЛОВКА "СОДЕРЖАНИЕ" ----------
    p_content = doc.paragraphs[content_idx]
    text = p_content.text.strip()
    
    if text.endswith("."):
        issues.append("Содержание – удалите точку в конце")
    
    alignment = get_effective_alignment(p_content)
    if alignment != WD_ALIGN_PARAGRAPH.CENTER:
        issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
    
    if not is_paragraph_bold(p_content):
        issues.append("Содержание – сделайте заголовок полужирным")
    
    # Проверка пустой строки после
    empty_found = False
    for i in range(content_idx + 1, min(content_idx + 5, len(doc.paragraphs))):
        if is_empty_paragraph(doc.paragraphs[i]):
            empty_found = True
            break
    
    if not empty_found and content_idx + 1 < len(doc.paragraphs):
        issues.append("Содержание – после заголовка должна быть пустая строка")

    # ---------- 4. ПОИСК ПЕРВОГО ЗАГОЛОВКА РАЗДЕЛА ----------
    start_idx = None
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    # Ищем после содержания
    for i in range(content_idx + 1, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        if not txt:
            continue
        
        # Заголовок с номером (1. ТЕКСТ)
        if re.match(r'^\d+\.\s+[А-Я]', txt):
            start_idx = i
            break
        
        # Ключевые слова
        if txt.upper() in level1_keywords:
            start_idx = i
            break
        
        # Заголовок из прописных по центру
        if is_all_caps(txt) and get_effective_alignment(p) == WD_ALIGN_PARAGRAPH.CENTER:
            start_idx = i
            break
    
    if start_idx is None:
        issues.append("❌ Не найден ни один заголовок раздела после содержания.")
        return issues

    issues.append(f"✅ Первый заголовок раздела в параграфе {start_idx}: {doc.paragraphs[start_idx].text[:50]}")
    
    # ---------- 5. ОСТАЛЬНЫЕ ПРОВЕРКИ ----------
    issues = list(dict.fromkeys(issues))
    return issues

# ========== ИНТЕРФЕЙС STREAMLIT ==========
st.set_page_config(page_title="Нормоконтроль документов", layout="wide")

st.title("📊 Автоматическая проверка документов Word")
st.markdown("---")

uploaded_file = st.file_uploader("📂 Выберите файл для проверки", type=["docx"])

if uploaded_file is not None:
    with st.spinner("🔍 Проверяем документ..."):
        try:
            results = check_word_document(uploaded_file)
            
            st.subheader("📋 Результаты проверки:")
            
            for r in results:
                if r.startswith('✅'):
                    st.success(r)
                elif r.startswith('❌'):
                    st.error(r)
                elif r.startswith('📋') or r.startswith('📊'):
                    st.info(r)
                else:
                    st.warning(r)
                    
        except Exception as e:
            st.error(f"❌ Ошибка при проверке документа: {str(e)}")
            st.exception(e)
else:
    st.info("👈 Загрузите документ для начала проверки")
