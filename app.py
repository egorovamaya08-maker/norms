import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
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

def aggressive_find_content(doc):
    """
    АГРЕССИВНЫЙ поиск - проверяем КАЖДЫЙ параграф
    """
    found_paragraphs = []
    
    print(f"🔍 Поиск во всех {len(doc.paragraphs)} параграфах...")
    
    for i, p in enumerate(doc.paragraphs):
        text = p.text
        text_upper = text.upper()
        
        # Проверяем различные варианты
        if 'СОДЕРЖАНИЕ' in text_upper or 'СОДЕРЖАНИЕ' in text:
            found_paragraphs.append((i, text[:200], 'direct'))
            continue
        
        # Проверяем очищенный текст (удаляем всё кроме букв)
        cleaned = re.sub(r'[^А-Яа-яA-Za-z]', '', text_upper)
        if 'СОДЕРЖАНИЕ' in cleaned:
            found_paragraphs.append((i, text[:200], 'cleaned'))
            continue
        
        # Проверяем через runs
        full_run_text = ''.join([run.text for run in p.runs])
        if 'СОДЕРЖАНИЕ' in full_run_text.upper():
            found_paragraphs.append((i, text[:200], 'runs'))
            continue
        
        # Проверяем частичное совпадение (если слово разбито)
        if 'СОДЕРЖ' in cleaned or 'СОДЕРЖ' in text_upper:
            found_paragraphs.append((i, text[:200], 'partial'))
            continue
    
    return found_paragraphs

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

    # ---------- 2. АГРЕССИВНЫЙ ПОИСК "СОДЕРЖАНИЕ" ----------
    found = aggressive_find_content(doc)
    
    # Показываем все параграфы, где есть упоминание "Содержание"
    issues.append(f"📊 Всего параграфов в документе: {len(doc.paragraphs)}")
    issues.append("")
    
    if found:
        issues.append(f"✅ Найдено {len(found)} параграф(ов) с упоминанием 'Содержание':")
        for idx, text, method in found:
            issues.append(f"  • Параграф {idx} (найден методом: {method})")
            issues.append(f"    Текст: {repr(text[:100])}")
        issues.append("")
        
        # Берем первый найденный параграф как заголовок
        content_idx = found[0][0]
        p_content = doc.paragraphs[content_idx]
        
        # Проверяем форматирование заголовка
        if p_content.text.strip().endswith("."):
            issues.append("Содержание – удалите точку в конце")
        
        alignment = get_effective_alignment(p_content)
        if alignment is not None and alignment != WD_ALIGN_PARAGRAPH.CENTER:
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
            
    else:
        issues.append("❌ Не найден заголовок «СОДЕРЖАНИЕ» — проверка невозможна.")
        
        # Показываем все параграфы, где есть буква "С" (для диагностики)
        issues.append("")
        issues.append("📋 Поиск параграфов, содержащих букву 'С' (первые 30):")
        count = 0
        for i, p in enumerate(doc.paragraphs):
            text = p.text.strip()
            if text and 'С' in text.upper():
                issues.append(f"  {i:3d}: {repr(text[:80])}")
                count += 1
                if count >= 30:
                    issues.append("  ... и другие")
                    break
        
        # Показываем все стили параграфов
        issues.append("")
        issues.append("🎨 Стили параграфов (первые 30):")
        for i in range(min(30, len(doc.paragraphs))):
            p = doc.paragraphs[i]
            style_name = p.style.name if p.style else "Без стиля"
            text_preview = p.text[:50] if p.text else "[пусто]"
            issues.append(f"  {i:3d}: {style_name:30} | {repr(text_preview)}")
        
        return issues
    
    # ---------- 3. ПОИСК ПЕРВОГО ЗАГОЛОВКА РАЗДЕЛА ----------
    content_idx = found[0][0]
    start_idx = None
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    issues.append("")
    issues.append("📋 Поиск первого заголовка раздела после содержания:")
    
    for i in range(content_idx + 1, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        if not txt:
            continue
        
        # Показываем проверяемые параграфы
        if i < content_idx + 20:  # Показываем первые 20 после содержания
            issues.append(f"  Проверяем параграф {i}: {repr(txt[:60])}")
        
        if re.match(r'^\d+\.\s+[А-Я]', txt):
            start_idx = i
            issues.append(f"  ✅ Найден заголовок с номером в параграфе {i}: {txt[:60]}")
            break
        
        if txt.upper() in level1_keywords:
            start_idx = i
            issues.append(f"  ✅ Найден ключевой заголовок в параграфе {i}: {txt[:60]}")
            break
        
        if is_all_caps(txt) and get_effective_alignment(p) == WD_ALIGN_PARAGRAPH.CENTER:
            start_idx = i
            issues.append(f"  ✅ Найден заголовок из прописных в параграфе {i}: {txt[:60]}")
            break
    
    if start_idx is None:
        issues.append("❌ Не найден ни один заголовок раздела после содержания.")
        return issues

    # ---------- 4. ОСТАЛЬНЫЕ ПРОВЕРКИ ----------
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
                elif r.startswith('📊') or r.startswith('📋') or r.startswith('🔍') or r.startswith('🎨'):
                    st.info(r)
                else:
                    st.warning(r)
                    
        except Exception as e:
            st.error(f"❌ Ошибка при проверке документа: {str(e)}")
            st.exception(e)
else:
    st.info("👈 Загрузите документ для начала проверки")
    
