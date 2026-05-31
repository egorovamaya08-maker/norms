import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import re

def get_effective_alignment(paragraph):
    """Безопасное получение выравнивания"""
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
    """Проверка на жирность (стиль или прямое форматирование)"""
    try:
        if paragraph.style and paragraph.style.font and paragraph.style.font.bold:
            return True
    except:
        pass
    runs = [r for r in paragraph.runs if r.text.strip()]
    if not runs:
        return False
    # Считаем жирным, если все значимые куски жирные
    return all(r.bold for r in runs)

def is_empty_paragraph(paragraph):
    return len(paragraph.text.strip()) == 0

def find_content_start(doc):
    """
    Усиленный поиск начала содержания.
    Возвращает индекс абзаца со словом СОДЕРЖАНИЕ или ОГЛАВЛЕНИЕ.
    """
    for i, p in enumerate(doc.paragraphs):
        # Очищаем текст от всего, кроме букв, чтобы найти суть
        clean_text = re.sub(r'[^а-яА-Яa-zA-Z]', '', p.text).upper()
        
        if "СОДЕРЖАНИЕ" in clean_text or "ОГЛАВЛЕНИЕ" in clean_text:
            return i
    return None

def find_main_text_start(doc, content_idx):
    """
    Поиск первого абзаца основного текста (Введение или Раздел 1).
    Ищем ПОСЛЕ содержания.
    """
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    # Начинаем поиск сразу после содержания
    start_search = content_idx + 1
    
    for i in range(start_search, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        text = p.text.strip()
        if not text:
            continue
            
        clean_text = re.sub(r'[^а-яА-Яa-zA-Z0-9\s\.\-]', '', text).upper().strip()
        
        # 1. Точное совпадение с ключевыми словами
        if clean_text in level1_keywords:
            return i
            
        # 2. Совпадение с шаблоном раздела "1. ТЕКСТ"
        # Регулярка: начинается с цифры, точка, пробел, русская буква
        if re.match(r'^\d+\.\s+[А-ЯЁ]', text):
            return i
            
    # Если не нашли явно, возвращаем конец содержания + 1 (фоллбэк)
    return start_search

def check_word_document(file):
    doc = docx.Document(file)
    issues = []

    # ---------- 1. ПОЛЯ СТРАНИЦ ----------
    margins_ok = True
    for section in doc.sections:
        # Используем mm для точности, допуск 1мм
        if (abs(section.left_margin.mm - 20) > 1.5 or
            abs(section.right_margin.mm - 20) > 1.5 or
            abs(section.top_margin.mm - 20) > 1.5 or
            abs(section.bottom_margin.mm - 20) > 1.5):
            margins_ok = False
            break
    if not margins_ok:
        issues.append("Поля страниц – установите 20 мм со всех сторон")

    # ---------- 2. ПОИСК СОДЕРЖАНИЯ ----------
    content_idx = find_content_start(doc)
    
    if content_idx is None:
        # Если не нашли, выводим диагностику первых 20 непустых абзацев
        debug_info = []
        count = 0
        for i, p in enumerate(doc.paragraphs):
            if p.text.strip():
                debug_info.append(f"Абзац {i}: «{p.text[:60]}»")
                count += 1
                if count >= 20: break
        
        issues.append("❌ Не найден заголовок «СОДЕРЖАНИЕ» или «ОГЛАВЛЕНИЕ».")
        issues.append("🔍 Возможно, он оформлен как картинка или таблица. Первые абзацы документа:")
        issues.extend(debug_info)
        return issues

    # ---------- 3. ПРОВЕРКА ЗАГОЛОВКА "СОДЕРЖАНИЕ" ----------
    p_content = doc.paragraphs[content_idx]
    text_raw = p_content.text.strip()
    
    # Проверка точки
    if text_raw.endswith("."):
        issues.append("Содержание – удалите точку в конце")
    
    # Проверка выравнивания
    align = get_effective_alignment(p_content)
    if align != WD_ALIGN_PARAGRAPH.CENTER:
        issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
    
    # Проверка жирности
    if not is_paragraph_bold(p_content):
        issues.append("Содержание – сделайте заголовок полужирным")
    
    # Проверка пустой строки после
    has_empty_after = False
    if content_idx + 1 < len(doc.paragraphs):
        if is_empty_paragraph(doc.paragraphs[content_idx + 1]):
            has_empty_after = True
    if not has_empty_after:
        issues.append("Содержание – после заголовка должна быть пустая строка")

    # ---------- 4. ОПРЕДЕЛЕНИЕ ГРАНИЦ ОСНОВНОГО ТЕКСТА ----------
    main_start_idx = find_main_text_start(doc, content_idx)
    
    # ---------- 5. ПРОХОД ПО ОСНОВНОМУ ТЕКСТУ ----------
    # Проверяем всё от main_start_idx до конца документа
    figure_counter = 0
    prev_para_empty = False
    
    # Ограничим проверку разумным пределом, если документ огромный, 
    # но для диплома обычно проверяем всё.
    for idx in range(main_start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[idx]
        text = p.text.strip()
        
        # Пропускаем пустые, но запоминаем состояние
        if not text:
            prev_para_empty = True
            continue
            
        pf = p.paragraph_format
        align = get_effective_alignment(p)
        
        # --- Классификация абзаца ---
        is_level1 = False
        is_subsection = False
        is_figure = False
        is_ref_header = False
        
        # Заголовки уровня 1: ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, 1. ТЕКСТ
        if text.upper() in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}:
            is_level1 = True
            if text.upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
                is_ref_header = True
        elif re.match(r'^\d+\.\s+[А-ЯЁ]', text):
            is_level1 = True
            
        # Подразделы: 1.1, 1.2.1
        elif re.match(r'^\d+\.\d+', text):
            is_subsection = True
            
        # Рисунки
        elif text.startswith("Рисунок") or text.startswith("рис."):
            is_figure = True

        # --- Проверки ---
        
        if is_level1:
            title_preview = text[:40]
            # 1. Новая страница (кроме Введения, если оно первое)
            # Упрощенная проверка: если это не самый первый абзац основного текста
            if idx != main_start_idx: 
                 # Здесь сложная логика проверки разрыва страницы, 
                 # для простоты пока пропустим или добавим базовую проверку
                 pass 
            
            # 2. Отступ 0
            if pf.first_line_indent and pf.first_line_indent.cm > 0.1:
                issues.append(f"«{title_preview}…» – уберите абзацный отступ")
            
            # 3. Жирный
            if not is_paragraph_bold(p):
                issues.append(f"«{title_preview}…» – должен быть полужирным")
                
            # 4. Все заглавные
            if text != text.upper():
                issues.append(f"«{title_preview}…» – должен быть прописными буквами")
                
            # 5. По центру
            if align != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{title_preview}…» – выравнивание по центру")
                
            # 6. Без точки
            if text.endswith("."):
                issues.append(f"«{title_preview}…» – удалите точку в конце")
                
            # 7. Пустая строка после
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx+1]):
                issues.append(f"«{title_preview}…» – после заголовка нужна пустая строка")

        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()[:40]
            # 1. Отступ 1.0 см
            indent = pf.first_line_indent.cm if pf.first_line_indent else 0.0
            if abs(indent - 1.0) > 0.15:
                issues.append(f"Подраздел «{sub_name}…» – отступ 1,0 см (сейчас {indent:.1f})")
            
            # 2. Жирный
            if not is_paragraph_bold(p):
                issues.append(f"Подраздел «{sub_name}…» – должен быть полужирным")
                
            # 3. Первая заглавная, остальные строчные (упрощенно)
            if sub_name and sub_name[0].islower():
                issues.append(f"Подраздел «{sub_name}…» – первая буква прописная")
                
            # 4. Без точки
            if text.endswith("."):
                issues.append(f"Подраздел «{sub_name}…» – удалите точку в конце")
                
            # 5. Нет пустой строки перед
            if prev_para_empty:
                issues.append(f"Подраздел «{sub_name}…» – уберите пустую строку перед подразделом")

        elif is_figure:
            figure_counter += 1
            # 1. По центру
            if align != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – выравнивание по центру")
            
            # 2. Без точки
            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – удалите точку в конце")
                
            # 3. Название с большой буквы (после тире)
            m = re.search(r'[–-]\s*(.+)', text)
            if m:
                name_part = m.group(1).strip()
                if name_part and name_part[0].islower():
                    issues.append(f"Рисунок {figure_counter} – название с большой буквы")

        elif is_ref_header:
            # Проверка заголовка списка источников
            if align != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Список источников – заголовок по центру")
            if not is_paragraph_bold(p):
                issues.append("Список источников – заголовок полужирный")
        
        else:
            # Обычный текст (не заголовок, не рисунок)
            # 1. Отступ 1.0 см
            indent = pf.first_line_indent.cm if pf.first_line_indent else 0.0
            if abs(indent - 1.0) > 0.15:
                # Игнорируем короткие строки, чтобы не шуметь на подписях таблиц/рисунков без ключевого слова
                if len(text) > 20: 
                     issues.append(f"Текст «{text[:30]}…» – отступ 1,0 см")
            
            # 2. Интервал перед 0
            if pf.space_before and pf.space_before.pt > 1:
                 issues.append(f"Текст «{text[:30]}…» – интервал перед 0 пт")

        prev_para_empty = False # Сброс флага, так как текущий абзац не пустой

    # Убираем дубликаты
    issues = list(dict.fromkeys(issues))
    
    if not issues:
        return ["✅ Ошибок не найдено."]
        
    return issues

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
                elif r.startswith('🔍'):
                    st.info(r)
                else:
                    st.warning(f"⚠️ {r}")
                    
        except Exception as e:
            st.error(f"❌ Критическая ошибка: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
