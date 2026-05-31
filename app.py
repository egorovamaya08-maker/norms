def get_table_depth(table, doc):
    """
    Определяет, не является ли таблица вложенной в другую таблицу.
    Возвращает глубину вложенности (0 = верхний уровень)
    """
    depth = 0
    element = table._element
    parent = element.getparent()
    
    while parent is not None:
        if parent.tag == qn('w:tbl'):
            depth += 1
        parent = parent.getparent()
    
    return depth

def test_document(file):
    doc = docx.Document(file)
    
    st.header("🔍 Анализ документа")
    
    # ============================================
    # ТЕСТ 1: Заголовки разделов
    # ============================================
    st.subheader("📋 Тест 1: Заголовки разделов")
    
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    start_idx = None
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt or has_page_number(txt):
            continue
        if txt.upper() in level1_keywords or (re.match(r'^\d+\.\s+[А-ЯЁ]', txt) and is_all_caps(txt)):
            start_idx = i
            st.write(f"✅ Строка {i}: Начало текста: '{txt[:100]}'")
            break
    
    if start_idx is None:
        st.warning("⚠️ Не найдено начало основного текста")
        return
    
    toc_entries = extract_toc_entries(doc, start_idx)
    
    st.write("**Заголовки в тексте:**")
    section_headers = []
    issues_found = 0
    
    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        
        if not txt or has_page_number(txt):
            continue
        
        # Пропускаем списки
        is_list, _, _ = is_list_item(txt, p, doc)
        if is_list:
            continue
        
        first_line = get_effective_first_line_indent(p)
        normalized = normalize_title(txt)
        in_toc = any(normalize_title(e) == normalized for e in toc_entries)
        
        is_header = False
        header_type = ""
        
        if is_section_header(txt):
            is_header = True
            header_type = "раздел"
        elif in_toc and len(txt) > 20:
            is_header = True
            header_type = "из содержания"
        elif is_all_caps(txt) and len(txt) > 25:
            is_header = True
            header_type = "капсом"
        
        if is_header and len(section_headers) < 30:
            st.write(f"**Строка {i}** ({header_type}): '{txt[:80]}'")
            st.write(f"  Отступ: {first_line:.3f} см {'✅' if abs(first_line) < 0.1 else '❌ должен быть 0 см'}")
            if abs(first_line) > 0.1:
                issues_found += 1
            section_headers.append((i, txt[:80], first_line))
            st.write("---")
    
    if issues_found == 0:
        st.success("✅ Все заголовки без отступа")
    else:
        st.error(f"❌ {issues_found} заголовков с ошибочным отступом")
    
    # ============================================
    # ТЕСТ 2: Списки
    # ============================================
    st.subheader("📝 Тест 2: Элементы списков")
    st.write("**Требования:**")
    st.write("• Маркер должен быть ТИРЕ (–), а не круглый (•)")
    st.write("• Отступ первой строки должен быть 1.0 см")
    
    list_items_found = 0
    list_items_correct = 0
    list_items_wrong = 0
    wrong_marker_items = []
    
    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        
        if not txt or has_page_number(txt):
            continue
        
        is_list, marker_type, marker_valid = is_list_item(txt, p, doc)
        
        if is_list:
            first_line = get_effective_first_line_indent(p)
            xml_info = get_paragraph_xml_info(p)
            
            list_items_found += 1
            
            indent_ok = abs(first_line - 1.0) < 0.15
            
            if list_items_found <= 25:
                st.write(f"**Элемент {list_items_found}** строка {i}:")
                st.write(f"  '{txt[:70]}...'")
                st.write(f"  Тип маркера: {marker_type} {'✅' if marker_valid else '❌ НЕДОПУСТИМ (нужно тире)'}")
                st.write(f"  Отступ: {first_line:.3f} см {'✅' if indent_ok else '❌ нужен 1.0 см'}")
                st.write(f"  XML: {xml_info}")
                
                if not marker_valid or not indent_ok:
                    if not marker_valid:
                        st.write(f"  🔴 Ошибка: замените круглый маркер на тире (–)")
                    if not indent_ok:
                        st.write(f"  🔴 Ошибка: установите отступ 1.0 см (сейчас {first_line:.1f} см)")
                else:
                    st.write(f"  🟢 Всё правильно")
                
                st.write("---")
            
            if marker_valid and indent_ok:
                list_items_correct += 1
            else:
                list_items_wrong += 1
                if not marker_valid:
                    wrong_marker_items.append((i, txt[:50], marker_type))
        
        if list_items_found >= 100:
            break
    
    if list_items_found == 0:
        st.write("ℹ️ Списки не найдены")
    else:
        st.write(f"📊 Всего: {list_items_found}")
        st.write(f"  ✅ Правильно: {list_items_correct}")
        st.write(f"  ❌ С ошибками: {list_items_wrong}")
        
        if wrong_marker_items:
            st.write("---")
            st.write(f"**❌ Элементы с недопустимыми круглыми маркерами (нужно заменить на тире):**")
            for idx, text, marker in wrong_marker_items:
                st.write(f"  • Строка {idx}: [{marker}] '{text}...'")
    
    # ============================================
    # ТЕСТ 3: Таблицы
    # ============================================
    st.subheader("📊 Тест 3: Таблицы")
    st.write("Проверка необходимости «Продолжение таблицы» / «Окончание таблицы»")
    st.write("⚠️ Автоматически невозможно определить границы страниц. Проверяем все таблицы.")
    
    try:
        start_element = doc.paragraphs[start_idx]._element
        start_body_pos = list(doc.element.body).index(start_element)
    except:
        start_body_pos = 0
    
    lit_start = None
    for i in range(start_idx, len(doc.paragraphs)):
        if doc.paragraphs[i].text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            lit_start = i
            break
    
    end_body_pos = len(doc.element.body)
    if lit_start:
        try:
            lit_element = doc.paragraphs[lit_start]._element
            end_body_pos = list(doc.element.body).index(lit_element)
        except:
            pass
    
    # Собираем ТОЛЬКО таблицы верхнего уровня (не вложенные)
    main_tables = []
    skipped_tables = 0
    
    for table in doc.tables:
        # Проверяем, не вложена ли таблица в другую таблицу
        depth = get_table_depth(table, doc)
        
        if depth > 0:
            skipped_tables += 1
            continue
        
        try:
            tbl_pos = list(doc.element.body).index(table._element)
            if start_body_pos < tbl_pos < end_body_pos:
                main_tables.append((tbl_pos, table))
        except:
            pass
    
    st.write(f"Таблиц в тексте: {len(main_tables)}")
    if skipped_tables > 0:
        st.write(f"Пропущено вложенных таблиц: {skipped_tables}")
    
    tables_need_check = 0
    
    for t_idx, (tbl_pos, table) in enumerate(main_tables, start=1):
        st.write(f"---")
        st.write(f"**Таблица {t_idx}** (позиция {tbl_pos})")
        
        rows = len(table.rows)
        cols = len(table.columns)
        st.write(f"  Размер: {rows} строк × {cols} столбцов")
        
        # Подпись
        caption = None
        for i in range(tbl_pos - 1, start_body_pos - 1, -1):
            elem = doc.element.body[i]
            if elem.tag.endswith('p'):
                for para in doc.paragraphs:
                    if para._element is elem and para.text.strip():
                        caption = para
                        break
                if caption:
                    break
        
        if caption:
            cap_text = caption.text.strip()
            st.write(f"  Подпись: '{cap_text[:120]}'")
            
            if '—' in cap_text or '–' in cap_text:
                st.write(f"  ✅ Тире правильное")
            elif '--' in cap_text or ' - ' in cap_text:
                st.write(f"  ⚠️ Нужно заменить на тире (—)")
            
            if cap_text.rstrip().endswith("."):
                st.write(f"  ⚠️ Убрать точку в конце")
        else:
            st.write(f"  ⚠️ Подпись не найдена")
        
        # Проверка на перенос (для таблиц больше 2 строк)
        if rows > 2:
            next_table_pos = end_body_pos
            for next_tbl_pos, _ in main_tables[t_idx:]:
                if next_tbl_pos > tbl_pos:
                    next_table_pos = next_tbl_pos
                    break
            
            markers = find_table_continuation_markers(doc, tbl_pos, next_table_pos, t_idx)
            
            if not markers:
                st.write(f"  🔴 **Нет «Продолжение таблицы {t_idx}» / «Окончание таблицы {t_idx}»**")
                st.write(f"  💡 Проверьте вручную. Если таблица на нескольких страницах, добавьте:")
                st.write(f"     • «Продолжение таблицы {t_idx}» на следующей странице")
                st.write(f"     • «Окончание таблицы {t_idx}» на последней странице")
                tables_need_check += 1
            else:
                st.write(f"  ✅ Маркеры найдены:")
                for m_pos, m_text in markers:
                    st.write(f"    • Строка {m_pos}: '{m_text}'")
        else:
            st.write(f"  ✅ Маленькая таблица, перенос маловероятен")
    
    if tables_need_check > 0:
        st.warning(f"⚠️ {tables_need_check} таблиц требуют ручной проверки")
    else:
        st.success("✅ Все таблицы в порядке")
    
    # ============================================
    # ИТОГИ
    # ============================================
    st.header("📊 Итоги")
    st.write(f"• Заголовков: {len(section_headers)} (ошибок отступа: {issues_found})")
    st.write(f"• Списков: {list_items_found} (правильно: {list_items_correct}, ошибок: {list_items_wrong})")
    if wrong_marker_items:
        st.write(f"  • Из них с недопустимым маркером: {len(wrong_marker_items)}")
    st.write(f"• Таблиц: {len(main_tables)} (нужна проверка переноса: {tables_need_check})")
