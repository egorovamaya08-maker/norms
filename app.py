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
            "оглаление"
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
        
        if clean_text in ["СОДЕРЖАНИЕ", "Содержание", "содержание"]:
            return i
    
    # Стратегия 3: Поиск по регулярному выражению (игнорируем точки, пробелы, дефисы)
    content_pattern = re.compile(r'^[СC][ОO][ДD][ЕE][РP][Ж][АA][НH][ИI][ЕE]', re.IGNORECASE)
    
    for i, p in enumerate(doc.paragraphs):
        # Удаляем все знаки препинания и лишние пробелы
        clean_text = re.sub(r'[^\w\s]', '', p.text.strip())
        clean_text = ' '.join(clean_text.split())
        
        if content_pattern.match(clean_text):
            return i
    
    # Стратегия 4: Поиск по части слова (если есть примечания после)
    for i, p in enumerate(doc.paragraphs):
        text_upper = p.text.strip().upper()
        
        # Проверяем, начинается ли абзац с "СОДЕРЖАНИЕ"
        if text_upper.startswith("СОДЕРЖАНИЕ"):
            # Проверяем, что после слова только знаки препинания или пробелы
            rest = text_upper[10:].strip()
            if not rest or rest in ['.', ',', ':', ';', '!', '?']:
                return i
    
    # Стратегия 5: Проверка первых 15 параграфов на предмет заголовка
    # (содержание обычно находится в начале документа)
    for i in range(min(15, len(doc.paragraphs))):
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
        
        # Дополнительная диагностика для отладки
        st.warning("Диагностика поиска «СОДЕРЖАНИЕ»:")
        for i in range(min(20, len(doc.paragraphs))):
            p_text = doc.paragraphs[i].text.strip()
            if p_text:
                st.write(f"Абзац {i}: «{p_text[:100]}»")
        
        return issues
    
    st.success(f"✓ Заголовок «СОДЕРЖАНИЕ» найден в абзаце {content_idx}")
    
    # ---------- 3. ПРОВЕРКА ЗАГОЛОВКА "СОДЕРЖАНИЕ" ----------
    p_content = doc.paragraphs[content_idx]
    text = p_content.text.strip()
    
    if text.endswith("."):
        issues.append("Содержание – удалите точку в конце")
    
    if get_effective_alignment(p_content) != WD_ALIGN_PARAGRAPH.CENTER:
        issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
    
    if not is_paragraph_bold(p_content):
        issues.append("Содержание – сделайте заголовок полужирным")
    
    # Проверяем наличие пустой строки после (учитываем возможные пустые параграфы)
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

    # ... остальная часть вашей функции остается без изменений ...
    # (весь код после поиска первого заголовка)
    
    return issues
