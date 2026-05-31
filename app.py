import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import re

def get_effective_alignment(paragraph):
    if paragraph.alignment is not None:
        return paragraph.alignment
    try:
        style = paragraph.style
        if style and hasattr(style, 'paragraph_format') and style.paragraph_format.alignment is not None:
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

def find_content_start(doc):
    """Ищет явный заголовок СОДЕРЖАНИЕ"""
    for i, p in enumerate(doc.paragraphs):
        clean_text = re.sub(r'[^а-яА-Яa-zA-Z]', '', p.text).upper()
        if "СОДЕРЖАНИЕ" in clean_text or "ОГЛАВЛЕНИЕ" in clean_text:
            return i
    return None

def find_main_text_start(doc, content_idx=None):
    """
    Ищет начало основного текста.
    Если content_idx задан, ищем после него. Если нет - с начала документа.
    Приоритет: ВВЕДЕНИЕ -> Раздел 1 -> Заключение
    """
    start_search = content_idx + 1 if content_idx is not None else 0
    
    for i in range(start_search, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        text = p.text.strip()
        if not text:
            continue
            
        # 1. Ищем ВВЕДЕНИЕ
        if text.upper() == "ВВЕДЕНИЕ":
            return i
            
        # 2. Ищем раздел вида "1. ТЕКСТ"
        if re.match(r'^\d+\.\s+[А-ЯЁ]', text):
            return i
            
    # Если ничего не нашли, возвращаем конец поиска (фоллбэк)
    return start_search

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    warnings = []

    # ---------- 1. ПОЛЯ СТРАНИЦ ----------
    margins_ok = True
    for section in doc.sections:
        try:
            if (abs(section.left_margin.mm - 20) > 1.5 or
                abs(section.right_margin.mm - 20) > 1.5 or
                abs(section.top_margin.mm - 20) > 1.5 or
                abs(section.bottom_margin.mm - 20) > 1.5):
                margins_ok = False
                break
        except:
            margins_ok = False
            break
    if not margins_ok:
        issues.append("Поля страниц – установите 20 мм со всех сторон")

    # ---------- 2. ПОИСК СОДЕРЖАНИЯ И НАЧАЛА ТЕКСТА ----------
    content_idx = find_content_start(doc)
    
    if content_idx is None:
        warnings.append("⚠️ Заголовок «СОДЕРЖАНИЕ» не найден. Проверка начата с первого раздела («ВВЕДЕНИЕ» или «1.»).")
        main_start_idx = find_main_text_start(doc, content_idx=None)
    else:
        # Если содержание найдено, проверяем его оформление
        p_content = doc.paragraphs[content_idx]
        if not get_effective_alignment(p_content) == WD_ALIGN_PARAGRAPH.CENTER:
            issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
        if not is_paragraph_bold(p_content):
            issues.append("Содержание – сделайте заголовок полужирным")
        if p_content.text.strip().endswith("."):
            issues.append("Содержание – удалите точку в конце")
        
        main_start_idx = find_main_text_start(doc, content_idx)

    if main_start_idx >= len(doc.paragraphs):
        issues.append("❌ Не удалось найти начало основного текста (Введение или Раздел 1).")
        return issues + warnings

    # ---------- 3. ПРОХОД ПО ОСНОВНОМУ ТЕКСТУ ----------
    figure_counter = 0
    prev_para_empty = False
    
    # Ограничиваем цикл, чтобы не зависнуть на огромных файлах, но для диплома обычно ок
    for idx in range(main_start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[idx]
        text = p.text.strip()
        
        if not text:
            prev_para_empty = True
            continue
            
        pf = p.paragraph_format
        align = get_effective_alignment(p)
        
        # --- Классификация ---
        is_level1 = False
        is_subsection = False
        is_figure = False
        is_ref_header = False
        
        if text.upper() in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}:
            is_level1 = True
            if text.upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
                is_ref_header = True
        elif re.match(r'^\d+\.\s+[А-ЯЁ]', text):
            is_level1 = True
        elif re.match(r'^\d+\.\d+', text):
            is_subsection = True
        elif text.startswith("Рисунок") or text.startswith("рис."):
            is_figure = True

        # --- Проверки ---
        
        if is_level1:
            title_preview = text[:40]
            
            # Отступ 0
            if pf.first_line_indent and pf.first_line_indent.cm > 0.1:
                issues.append(f"«{title_preview}…» – уберите абзацный отступ")
            
            # Жирный
            if not is_paragraph_bold(p):
                issues.append(f"«{title_preview}…» – должен быть полужирным")
                
            # Все заглавные
            if text != text.upper():
                issues.append(f"«{title_preview}…» – должен быть прописными буквами")
                
            # По центру
            if align != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{title_preview}…» – выравнивание по центру")
                
            # Без точки
            if text.endswith("."):
                issues.append(f"«{title_preview}…» – удалите точку в конце")
                
            # Пустая строка после
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx+1]):
                issues.append(f"«{title_preview}…» – после заголовка нужна пустая строка")

        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()[:40]
            
            # Отступ 1.0 см
            indent = pf.first_line_indent.cm if pf.first_line_indent else 0.0
            if abs(indent - 1.0) > 0.15:
                issues.append(f"Подраздел «{sub_name}…» – отступ 1,0 см (сейчас {indent:.1f})")
            
            # Жирный
            if not is_paragraph_bold(p):
                issues.append(f"Подраздел «{sub_name}…» – должен быть полужирным")
                
            # Первая заглавная
            if sub_name and sub_name[0].islower():
                issues.append(f"Подраздел «{sub_name}…» – первая буква прописная")
                
            # Без точки
            if text.endswith("."):
                issues.append(f"Подраздел «{sub_name}…» – удалите точку в конце")
                
            # Нет пустой строки перед
            if prev_para_empty:
                issues.append(f"Подраздел «{sub_name}…» – уберите пустую строку перед подразделом")

        elif is_figure:
            figure_counter += 1
            if align != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – выравнивание по центру")
            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – удалите точку в конце")
            
            # Название с большой буквы
            m = re.search(r'[–-]\s*(.+)', text)
            if m:
                name_part = m.group(1).strip()
                if name_part and name_part[0].islower():
                    issues.append(f"Рисунок {figure_counter} – название с большой буквы")

        elif is_ref_header:
            if align != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Список источников – заголовок по центру")
            if not is_paragraph_bold(p):
                issues.append("Список источников – заголовок полужирный")
        
        else:
            # Обычный текст
            indent = pf.first_line_indent.cm if pf.first_line_indent else 0.0
            if abs(indent - 1.0) > 0.15:
                if len(text) > 20: 
                     issues.append(f"Текст «{text[:30]}…» – отступ 1,0 см")
            
            if pf.space_before and pf.space_before.pt > 1:
                 issues.append(f"Текст «{text[:30]}…» – интервал перед 0 пт")

        prev_para_empty = False

    # Убираем дубликаты
    issues = list(dict.fromkeys(issues))
    
    if not issues and not warnings:
        return ["✅ Ошибок не найдено."]
        
    return warnings + issues

# ========== ИНТЕРФЕЙС STREAMLIT ==========
st.set_page_config(page_title="Нормоконтроль документов", layout="wide")
st.title("📊 Автоматическая проверка документов Word")
st.markdown("---")

uploaded_file = st.file_uploader("📂 Выберите файл .docx", type=["docx"])

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
                elif r.startswith('⚠️'):
                    st.warning(r)
                else:
                    st.info(f"• {r}")
                    
        except Exception as e:
            st.error(f"❌ Критическая ошибка: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
