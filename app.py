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

def get_first_line_indent_cm(paragraph):
    """
    Получает отступ первой строки в см, учитывая стиль абзаца.
    Если отступ задан явно - берет его. Если нет - берет из стиля.
    """
    pf = paragraph.paragraph_format
    
    # 1. Проверяем явный отступ абзаца
    if pf.first_line_indent is not None:
        return pf.first_line_indent.cm
    
    # 2. Если явного нет, проверяем стиль
    try:
        style = paragraph.style
        if style and hasattr(style, 'paragraph_format'):
            spf = style.paragraph_format
            if spf.first_line_indent is not None:
                return spf.first_line_indent.cm
    except:
        pass
        
    # 3. Если нигде нет, значит 0
    return 0.0

def find_content_start(doc):
    """Ищет явный заголовок СОДЕРЖАНИЕ"""
    for i, p in enumerate(doc.paragraphs):
        clean_text = re.sub(r'[^а-яА-Яa-zA-Z]', '', p.text).upper()
        if "СОДЕРЖАНИЕ" in clean_text or "ОГЛАВЛЕНИЕ" in clean_text:
            return i
    return None

def find_main_text_start(doc, content_idx=None):
    start_search = content_idx + 1 if content_idx is not None else 0
    for i in range(start_search, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        text = p.text.strip()
        if not text:
            continue
        if text.upper() == "ВВЕДЕНИЕ":
            return i
        if re.match(r'^\d+\.\s+[А-ЯЁ]', text):
            return i
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
        p_content = doc.paragraphs[content_idx]
        if not get_effective_alignment(p_content) == WD_ALIGN_PARAGRAPH.CENTER:
            issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
        if not is_paragraph_bold(p_content):
            issues.append("Содержание – сделайте заголовок полужирным")
        if p_content.text.strip().endswith("."):
            issues.append("Содержание – удалите точку в конце")
        
        main_start_idx = find_main_text_start(doc, content_idx)

    if main_start_idx >= len(doc.paragraphs):
        issues.append("❌ Не удалось найти начало основного текста.")
        return issues + warnings

    # ---------- 3. ПРОХОД ПО ОСНОВНОМУ ТЕКСТУ ----------
    figure_counter = 0
    prev_para_empty = False
    
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
        is_table_caption = False
        is_ref_header = False
        is_ref_item = False
        
        # Заголовки уровней
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
        elif text.startswith("Таблица"):
            is_table_caption = True
            
        # Элементы списка литературы (начинаются с цифры и точки, например "1. Zadeh...")
        # Важно отличать их от подразделов (1.1) и разделов (1. ТЕКСТ)
        if is_ref_header and re.match(r'^\d+\.\s', text) and not re.match(r'^\d+\.\d+', text):
             is_ref_item = True

        # --- Проверки ---
        
        if is_level1:
            title_preview = text[:40]
            indent = get_first_line_indent_cm(p)
            
            if indent > 0.1:
                issues.append(f"«{title_preview}…» – уберите абзацный отступ")
            if not is_paragraph_bold(p):
                issues.append(f"«{title_preview}…» – должен быть полужирным")
            if text != text.upper():
                issues.append(f"«{title_preview}…» – должен быть прописными буквами")
            if align != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{title_preview}…» – выравнивание по центру")
            if text.endswith("."):
                issues.append(f"«{title_preview}…» – удалите точку в конце")
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx+1]):
                issues.append(f"«{title_preview}…» – после заголовка нужна пустая строка")

        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()[:40]
            indent = get_first_line_indent_cm(p)
            
            if abs(indent - 1.0) > 0.15:
                issues.append(f"Подраздел «{sub_name}…» – отступ 1,0 см (сейчас {indent:.1f})")
            if not is_paragraph_bold(p):
                issues.append(f"Подраздел «{sub_name}…» – должен быть полужирным")
            if sub_name and sub_name[0].islower():
                issues.append(f"Подраздел «{sub_name}…» – первая буква прописная")
            if text.endswith("."):
                issues.append(f"Подраздел «{sub_name}…» – удалите точку в конце")
            if prev_para_empty:
                issues.append(f"Подраздел «{sub_name}…» – уберите пустую строку перед подразделом")

        elif is_figure:
            figure_counter += 1
            if align != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – выравнивание по центру")
            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – удалите точку в конце")
            m = re.search(r'[–-]\s*(.+)', text)
            if m:
                name_part = m.group(1).strip()
                if name_part and name_part[0].islower():
                    issues.append(f"Рисунок {figure_counter} – название с большой буквы")

        elif is_table_caption:
            # Подписи таблиц обычно без отступа и по центру/левому краю (зависит от ГОСТ, часто по левому)
            # Здесь просто пропускаем проверку отступа 1см, так как это не основной текст
            pass

        elif is_ref_header:
            if align != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Список источников – заголовок по центру")
            if not is_paragraph_bold(p):
                issues.append("Список источников – заголовок полужирный")
        
        elif is_ref_item:
            # Проверка первого источника (если нужно)
            # Обычно в списке источников отступ 0 или висячий отступ. 
            # Если требования нет, пропускаем.
            pass

        else:
            # Обычный основной текст
            indent = get_first_line_indent_cm(p)
            
            # Пропускаем короткие строки и подписи, чтобы не шуметь
            if len(text) > 20: 
                 if abs(indent - 1.0) > 0.15:
                     issues.append(f"Текст «{text[:30]}…» – отступ 1,0 см (сейчас {indent:.1f})")
            
            if pf.space_before and pf.space_before.pt > 1:
                 issues.append(f"Текст «{text[:30]}…» – интервал перед 0 пт")

        prev_para_empty = False

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
