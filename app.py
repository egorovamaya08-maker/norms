import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
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
    """Проверяет, состоит ли строка только из прописных букв, пробелов и дефисов"""
    return bool(re.match(r'^[А-ЯЁ\s\-]+$', text)) and len(text) > 3

def find_content_heading_enhanced(doc):
    """
    МАКСИМАЛЬНО РАСШИРЕННЫЙ поиск заголовка «СОДЕРЖАНИЕ»
    Специально для документов с TOC-Heading стилем
    """
    
    # Стратегия 1: Очистка от всех небуквенных символов
    for i, p in enumerate(doc.paragraphs):
        # Удаляем всё, кроме букв
        cleaned = re.sub(r'[^А-Яа-яA-Za-z]', '', p.text.upper())
        
        # Проверяем, содержит ли очищенный текст слово "СОДЕРЖАНИЕ"
        if "СОДЕРЖАНИЕ" in cleaned or "ОГЛАВЛЕНИЕ" in cleaned:
            return i
        
        # Проверяем оригинальный текст
        stripped = p.text.strip().upper()
        if "СОДЕРЖАНИЕ" in stripped or "ОГЛАВЛЕНИЕ" in stripped:
            return i
    
    # Стратегия 2: Поиск по runs
    for i, p in enumerate(doc.paragraphs):
        full_text = ''
        for run in p.runs:
            full_text += run.text
        
        full_cleaned = re.sub(r'[^А-Яа-я]', '', full_text.upper())
        if "СОДЕРЖАНИЕ" in full_cleaned or "ОГЛАВЛЕНИЕ" in full_cleaned:
            return i
    
    # Стратегия 3: Проверка стиля TOC-Heading
    for i, p in enumerate(doc.paragraphs):
        try:
            if p.style and p.style.name and 'TOC' in p.style.name.upper():
                cleaned_text = re.sub(r'[^А-Яа-я]', '', p.text.upper())
                if "СОДЕРЖАНИЕ" in cleaned_text or "ОГЛАВЛЕНИЕ" in cleaned_text:
                    return i
        except:
            pass
    
    # Стратегия 4: Поиск подстроки в первых 30 параграфах
    for i in range(min(30, len(doc.paragraphs))):
        text = doc.paragraphs[i].text
        if 'содержание' in text.lower() or 'оглавление' in text.lower():
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

    # ---------- 2. ПОИСК "СОДЕРЖАНИЕ" (РАСШИРЕННЫЙ) ----------
    content_idx = find_content_heading_enhanced(doc)
    
    if content_idx is None:
        issues.append("❌ Не найден заголовок «СОДЕРЖАНИЕ» — проверка невозможна.")
        
        # Диагностика
        issues.append("\n📋 Диагностика: первые 15 абзацев документа:")
        count = 0
        for i, p in enumerate(doc.paragraphs):
            p_text = p.text.strip()
            if p_text:
                # Показываем также "сырой" вид для отладки
                raw_repr = repr(p.text[:50])
                issues.append(f"  Абзац {i}: «{p_text[:80]}»")
                issues.append(f"       Raw: {raw_repr}")
                count += 1
                if count >= 15:
                    break
        
        return issues
    
    issues.append(f"✅ Заголовок «СОДЕРЖАНИЕ» найден в абзаце {content_idx}")
    
    # ---------- 3. ПРОВЕРКА ЗАГОЛОВКА "СОДЕРЖАНИЕ" ----------
    p_content = doc.paragraphs[content_idx]
    text = p_content.text.strip()
    
    # Очищаем текст от символа # и других спецсимволов
    clean_text = re.sub(r'^[#\s]+', '', text)
    
    if clean_text.endswith("."):
        issues.append("Содержание – удалите точку в конце")
    
    if get_effective_alignment(p_content) != WD_ALIGN_PARAGRAPH.CENTER:
        issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
    
    if not is_paragraph_bold(p_content):
        issues.append("Содержание – сделайте заголовок полужирным")
    
    # Проверяем наличие пустой строки после
    empty_found = False
    for i in range(content_idx + 1, min(content_idx + 5, len(doc.paragraphs))):
        if is_empty_paragraph(doc.paragraphs[i]):
            empty_found = True
            break
        elif doc.paragraphs[i].text.strip():
            break
    
    if not empty_found and content_idx + 1 < len(doc.paragraphs):
        issues.append("Содержание – после заголовка должна быть пустая строка")

    # ---------- 4. ПОИСК ПЕРВОГО ЗАГОЛОВКА РАЗДЕЛА ----------
    start_idx = None
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    for i in range(content_idx + 1, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        if not txt:
            continue
        
        # Очищаем от спецсимволов
        clean_txt = re.sub(r'^[#\s]+', '', txt)
        
        if re.match(r'^\d+\.\s+[А-Я]', clean_txt):
            start_idx = i
            break
        
        if clean_txt.upper() in level1_keywords:
            start_idx = i
            break
        
        if is_all_caps(clean_txt) and get_effective_alignment(p) == WD_ALIGN_PARAGRAPH.CENTER:
            start_idx = i
            break
    
    if start_idx is None:
        issues.append("❌ Не найден ни один заголовок раздела после содержания.")
        return issues

    issues.append(f"✅ Найден первый заголовок раздела в абзаце {start_idx}")

    # ---------- 5. ПРОВЕРКА ОСНОВНОГО ТЕКСТА (первые 50 абзацев) ----------
    figure_counter = 0
    prev_para_empty = False
    subsection_re = re.compile(r'^\d+\.\d+')
    
    max_check = min(start_idx + 50, len(doc.paragraphs))
    
    for idx in range(start_idx, max_check):
        p = doc.paragraphs[idx]
        text = p.text.strip()
        
        if not text:
            prev_para_empty = True
            continue

        pf = p.paragraph_format
        alignment = get_effective_alignment(p)
        
        # Очищаем текст
        clean_text = re.sub(r'^[#\s]+', '', text)
        
        is_level1 = (clean_text.upper() in level1_keywords or
                     re.match(r'^\d+\.\s+[А-Я]', clean_text) or
                     (is_all_caps(clean_text) and alignment == WD_ALIGN_PARAGRAPH.CENTER and len(clean_text) > 5))
        
        is_subsection = bool(subsection_re.match(clean_text)) and not is_level1
        is_figure = clean_text.startswith("Рисунок")
        
        if is_level1:
            if clean_text.upper() != "ВВЕДЕНИЕ":
                page_break = False
                if idx > 0:
                    prev_p = doc.paragraphs[idx - 1]
                    for run in prev_p.runs:
                        if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                            page_break = True
                if not page_break:
                    issues.append(f"«{clean_text[:50]}» – раздел должен начинаться с новой страницы")
            
            if pf.first_line_indent and abs(pf.first_line_indent.cm) > 0.1:
                issues.append(f"«{clean_text[:50]}» – уберите абзацный отступ у заголовка")
            if not is_paragraph_bold(p):
                issues.append(f"«{clean_text[:50]}» – заголовок раздела должен быть полужирным")
            if clean_text != clean_text.upper():
                issues.append(f"«{clean_text[:50]}» – заголовок раздела должен быть прописными буквами")
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{clean_text[:50]}» – выровняйте заголовок по центру")
            if clean_text.endswith("."):
                issues.append(f"«{clean_text[:50]}» – удалите точку в конце")
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                issues.append(f"«{clean_text[:50]}» – после заголовка должна быть пустая строка")
        
        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', clean_text).strip()
            if not pf.first_line_indent:
                issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см")
            elif pf.first_line_indent and abs(pf.first_line_indent.cm - 1.0) > 0.2:
                issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см")
            if not is_paragraph_bold(p):
                issues.append(f"Подраздел «{sub_name[:50]}» – заголовок должен быть полужирным")
            if sub_name and sub_name[0].islower():
                issues.append(f"Подраздел «{sub_name[:50]}» – первая буква должна быть прописной")
            if clean_text.endswith("."):
                issues.append(f"Подраздел «{sub_name[:50]}» – удалите точку в конце")
            if prev_para_empty:
                issues.append(f"Подраздел «{sub_name[:50]}» – уберите пустую строку перед подразделом")
        
        elif is_figure:
            figure_counter += 1
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – выровняйте подпись по центру")
        
        prev_para_empty = False

    # Убираем дубликаты
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
                elif r.startswith('📋'):
                    st.info(r)
                else:
                    st.warning(f"⚠️ {r}")
                    
        except Exception as e:
            st.error(f"❌ Ошибка при проверке документа: {str(e)}")
            st.exception(e)
