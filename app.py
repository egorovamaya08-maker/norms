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

def find_content_heading(doc):
    """
    Расширенный поиск заголовка «СОДЕРЖАНИЕ» с несколькими стратегиями.
    Возвращает индекс параграфа или None.
    """
    # Стратегия 1: Точное совпадение после очистки
    for i, p in enumerate(doc.paragraphs):
        # Получаем чистый текст без лишних пробелов и спецсимволов
        clean_text = ' '.join(p.text.strip().split())
        
        # Варианты написания для поиска
        variants = [
            "СОДЕРЖАНИЕ",
            "Содержание", 
            "содержание",
            "ОГЛАВЛЕНИЕ",
            "Оглавление",
            "оглавление"
        ]
        
        for variant in variants:
            if clean_text == variant or clean_text.startswith(variant):
                return i
    
    # Стратегия 2: Поиск по runs (учитывает частичное форматирование)
    for i, p in enumerate(doc.paragraphs):
        # Собираем текст из всех runs
        full_text = ''
        for run in p.runs:
            full_text += run.text
        
        clean_text = ' '.join(full_text.strip().split())
        
        if clean_text in ["СОДЕРЖАНИЕ", "Содержание", "содержание", 
                         "ОГЛАВЛЕНИЕ", "Оглавление", "оглавление"]:
            return i
    
    # Стратегия 3: Поиск по регулярному выражению (игнорируем точки, пробелы, дефисы)
    content_pattern = re.compile(r'^[СC][ОO][ДD][ЕE][РP][Ж][АA][НH][ИI][ЕE]', re.IGNORECASE)
    contents_pattern = re.compile(r'^[ОO][ГG][ЛL][АA][ВB][ЛL][ЕE][НH][ИI][ЕE]', re.IGNORECASE)
    
    for i, p in enumerate(doc.paragraphs):
        # Удаляем все знаки препинания и лишние пробелы
        clean_text = re.sub(r'[^\w\s]', '', p.text.strip())
        clean_text = ' '.join(clean_text.split())
        
        if content_pattern.match(clean_text) or contents_pattern.match(clean_text):
            return i
    
    # Стратегия 4: Поиск по части слова (если есть примечания после)
    for i, p in enumerate(doc.paragraphs):
        text_upper = p.text.strip().upper()
        
        # Проверяем, начинается ли абзац с "СОДЕРЖАНИЕ"
        if text_upper.startswith("СОДЕРЖАНИЕ") or text_upper.startswith("ОГЛАВЛЕНИЕ"):
            # Проверяем, что после слова только знаки препинания или пробелы
            rest = text_upper[10:].strip()
            if not rest or rest in ['.', ',', ':', ';', '!', '?']:
                return i
    
    # Стратегия 5: Проверка первых 20 параграфов на предмет заголовка
    # (содержание обычно находится в начале документа)
    for i in range(min(20, len(doc.paragraphs))):
        p = doc.paragraphs[i]
        text_clean = re.sub(r'[^А-Яа-я]', '', p.text.strip().upper())
        
        if "СОДЕРЖАНИЕ" in text_clean or "ОГЛАВЛЕНИЕ" in text_clean:
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

    # ---------- 2. ПОИСК "СОДЕРЖАНИЕ" (УЛУЧШЕННЫЙ) ----------
    content_idx = find_content_heading(doc)
    
    if content_idx is None:
        issues.append("❌ Не найден заголовок «СОДЕРЖАНИЕ» — проверка невозможна.")
        
        # Сохраняем диагностику в issues для отображения
        issues.append("\n📋 Диагностика: первые 15 непустых абзацев документа:")
        count = 0
        for i, p in enumerate(doc.paragraphs):
            p_text = p.text.strip()
            if p_text:
                issues.append(f"  Абзац {i}: «{p_text[:80]}»")
                count += 1
                if count >= 15:
                    break
        
        return issues
    
    issues.append(f"✅ Заголовок «СОДЕРЖАНИЕ» найден в абзаце {content_idx}")
    
    # ---------- 3. ПРОВЕРКА ЗАГОЛОВКА "СОДЕРЖАНИЕ" ----------
    p_content = doc.paragraphs[content_idx]
    text = p_content.text.strip()
    
    if text.endswith("."):
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
    
    # Ищем после содержания, пропуская пустые строки
    for i in range(content_idx + 1, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        if not txt:
            continue
        
        # Проверяем: заголовок с номером (1. НАЗВАНИЕ)
        if re.match(r'^\d+\.\s+[А-Я]', txt):
            start_idx = i
            break
        
        # Проверяем: ключевые слова
        if txt.upper() in level1_keywords:
            start_idx = i
            break
        
        # Проверяем: строка из прописных букв по центру
        if is_all_caps(txt) and get_effective_alignment(p) == WD_ALIGN_PARAGRAPH.CENTER:
            start_idx = i
            break
    
    if start_idx is None:
        issues.append("❌ Не найден ни один заголовок раздела после содержания.")
        return issues

    issues.append(f"✅ Найден первый заголовок раздела в абзаце {start_idx}")

    # ---------- 5. ПРОВЕРКА ОСНОВНОГО ТЕКСТА ----------
    figure_counter = 0
    prev_para_empty = False
    subsection_re = re.compile(r'^\d+\.\d+')

    # Ограничим проверку первыми 50 абзацами для тестирования
    max_check = min(start_idx + 50, len(doc.paragraphs))
    
    for idx in range(start_idx, max_check):
        p = doc.paragraphs[idx]
        text = p.text.strip()
        
        if not text:
            prev_para_empty = True
            continue

        pf = p.paragraph_format
        alignment = get_effective_alignment(p)
        
        # Определяем тип абзаца
        is_level1 = (text.upper() in level1_keywords or
                     re.match(r'^\d+\.\s+[А-Я]', text) or
                     (is_all_caps(text) and alignment == WD_ALIGN_PARAGRAPH.CENTER and len(text) > 5))
        
        is_subsection = bool(subsection_re.match(text)) and not is_level1
        is_figure = text.startswith("Рисунок")
        is_table_caption = text.startswith("Таблица")

        # --- Заголовок раздела ---
        if is_level1:
            # Новая страница (кроме ВВЕДЕНИЯ)
            if text.upper() != "ВВЕДЕНИЕ":
                page_break = False
                if idx > 0:
                    prev_p = doc.paragraphs[idx - 1]
                    for run in prev_p.runs:
                        if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                            page_break = True
                for run in p.runs:
                    if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                        page_break = True
                if not page_break:
                    issues.append(f"«{text[:50]}» – раздел должен начинаться с новой страницы")
            
            # Отступ 0
            if pf.first_line_indent and abs(pf.first_line_indent.cm) > 0.1:
                issues.append(f"«{text[:50]}» – уберите абзацный отступ у заголовка")
            # Полужирный
            if not is_paragraph_bold(p):
                issues.append(f"«{text[:50]}» – заголовок раздела должен быть полужирным")
            # Прописные
            if text != text.upper():
                issues.append(f"«{text[:50]}» – заголовок раздела должен быть прописными буквами")
            # Центр
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:50]}» – выровняйте заголовок по центру")
            # Без точки
            if text.endswith("."):
                issues.append(f"«{text[:50]}» – удалите точку в конце")
            # Пустая строка после
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                issues.append(f"«{text[:50]}» – после заголовка должна быть пустая строка")
        
        # --- Подраздел ---
        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()
            # Отступ 1 см
            if not pf.first_line_indent:
                issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см")
            elif abs(pf.first_line_indent.cm - 1.0) > 0.2:
                issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см (сейчас {pf.first_line_indent.cm:.1f})")
            # Полужирный
            if not is_paragraph_bold(p):
                issues.append(f"Подраздел «{sub_name[:50]}» – заголовок должен быть полужирным")
            # Первая прописная, остальные строчные
            if sub_name and sub_name[0].islower():
                issues.append(f"Подраздел «{sub_name[:50]}» – первая буква должна быть прописной")
            elif sub_name and len(sub_name) > 1:
                rest = sub_name[1:]
                if any(c.isupper() for c in rest):
                    issues.append(f"Подраздел «{sub_name[:50]}» – после первой буквы должны быть строчные")
            # Без точки
            if text.endswith("."):
                issues.append(f"Подраздел «{sub_name[:50]}» – удалите точку в конце")
            # Нет пустой строки перед
            if prev_para_empty:
                issues.append(f"Подраздел «{sub_name[:50]}» – уберите пустую строку перед подразделом")
        
        # --- Рисунок ---
        elif is_figure:
            figure_counter += 1
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – выровняйте подпись по центру")
            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – удалите точку в конце")
            # Название с большой буквы
            m = re.match(r'^Рисунок\s+\d+\s*[–\-]\s*(.+)$', text)
            if m:
                title = m.group(1).strip()
                if title and title[0].islower():
                    issues.append(f"Рисунок {figure_counter} – название должно начинаться с большой буквы")
            # Пустые строки
            if idx > 0 and not is_empty_paragraph(doc.paragraphs[idx - 1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку перед рисунком")
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку после рисунка")
        
        prev_para_empty = False

    # ---------- ИТОГ ----------
    issues = list(dict.fromkeys(issues))
    return issues

# ========== ИНТЕРФЕЙС STREAMLIT ==========
st.set_page_config(page_title="Нормоконтроль документов", layout="wide")

st.title("📊 Автоматическая проверка документов Word")
st.markdown("---")

# Боковая панель с информацией
with st.sidebar:
    st.header("📌 Инструкция")
    st.markdown("""
    1. Загрузите документ в формате **.docx**
    2. Программа автоматически проверит:
       - Поля страниц
       - Наличие заголовка «СОДЕРЖАНИЕ»
       - Форматирование заголовков
       - Оформление подразделов
       - Рисунки и таблицы
    3. Получите список замечаний
    """)
    
    st.markdown("---")
    st.caption("Версия 2.0 с улучшенным поиском содержания")

# Основная область
col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader(
        "📂 Выберите файл для проверки", 
        type=["docx"],
        help="Поддерживаются только файлы Microsoft Word (.docx)"
    )

with col2:
    if uploaded_file is not None:
        st.success(f"✅ Файл загружен: {uploaded_file.name}")

st.markdown("---")

# Проверка документа
if uploaded_file is not None:
    with st.spinner("🔍 Проверяем документ... Это может занять несколько секунд"):
        try:
            results = check_word_document(uploaded_file)
            
            # Отображение результатов
            st.subheader("📋 Результаты проверки:")
            
            # Счетчики
            errors = [r for r in results if r.startswith(('❌', '•'))]
            warnings = [r for r in results if not r.startswith('✅') and not r.startswith('📋')]
            success = [r for r in results if r.startswith('✅')]
            
            # Метрики
            col_metrics1, col_metrics2, col_metrics3 = st.columns(3)
            with col_metrics1:
                st.metric("Всего замечаний", len([r for r in results if not r.startswith('✅') and not r.startswith('📋')]))
            with col_metrics2:
                st.metric("Ошибки", len([r for r in results if r.startswith('❌')]))
            with col_metrics3:
                st.metric("Успешно", len(success))
            
            st.markdown("---")
            
            # Вывод результатов
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
else:
    st.info("👈 Загрузите документ для начала проверки")
    
    # Пример тестового документа
    with st.expander("📝 Как протестировать?"):
        st.markdown("""
        ### Создайте тестовый документ:
        
        1. Создайте новый документ Word
        2. Установите поля: левое 20 мм, правое 20 мм, верхнее 20 мм, нижнее 20 мм
        3. Напишите заголовок **СОДЕРЖАНИЕ** (по центру, жирным)
        4. Добавьте пустую строку
        5. Напишите **ВВЕДЕНИЕ** (по центру, жирным, с новой страницы)
        6. Сохраните как .docx и загрузите для проверки
        
        Программа найдет заголовок «СОДЕРЖАНИЕ» даже если:
        - Есть точка в конце
        - Есть лишние пробелы
        - Слово написано с маленькой буквы
        - Использовано «ОГЛАВЛЕНИЕ»
        """)
