def check_word_document(file):
    doc = docx.Document(file)
    auto_issues = []
    manual_checks = []

    # --- ПОЛЯ ---
    margins_ok = True
    for section in doc.sections:
        if (abs(section.left_margin.mm - 20) > 0.5 or
            abs(section.right_margin.mm - 20) > 0.5 or
            abs(section.top_margin.mm - 20) > 0.5 or
            abs(section.bottom_margin.mm - 20) > 0.5):
            margins_ok = False
            break
    if not margins_ok:
        auto_issues.append("Поля страниц – установите левое 20 мм, правое 20 мм, верхнее 20 мм, нижнее 20 мм")

    # --- ПОИСК НАЧАЛА ОСНОВНОГО ТЕКСТА ---
    start_idx = None
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt:
            continue
        if has_page_number(txt):
            continue
        if is_section_header(txt):
            start_idx = i
            break
    if start_idx is None:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]

    # --- НУМЕРАЦИЯ СТРАНИЦ ---
    try:
        intro_body_idx = None
        body_elements = list(doc.element.body)
        for i, elem in enumerate(body_elements):
            if elem.tag == qn('w:p'):
                full_text = ''.join(t.text or '' for t in elem.findall('.//w:t', qn('w'))).strip().upper()
                if full_text == 'ВВЕДЕНИЕ':
                    intro_body_idx = i
                    break
        if intro_body_idx is not None:
            page_issues = check_page_numbering(file, intro_body_idx)
            auto_issues.extend(page_issues)
    except:
        pass

    toc_entries = extract_toc_entries(doc, start_idx)

    # --- ГРАНИЦЫ СПИСКА ИСТОЧНИКОВ ---
    lit_start = None
    for i in range(start_idx, len(doc.paragraphs)):
        txt = doc.paragraphs[i].text.strip()
        if txt.upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" and not has_page_number(txt):
            lit_start = i
            break
    lit_end = len(doc.paragraphs)
    if lit_start is not None:
        appendix_keywords = ["ПРИЛОЖЕНИЕ", "ПРИЛОЖЕНИЯ", "APPENDIX"]
        for i in range(lit_start + 1, len(doc.paragraphs)):
            txt = doc.paragraphs[i].text.strip().upper()
            if any(txt.startswith(kw) for kw in appendix_keywords):
                lit_end = i
                break

    # --- ПРОВЕРКА ОСНОВНОГО ТЕКСТА ---
    figure_counter = 0
    prev_para_empty = False
    prev_was_formula = False
    end_idx = lit_start if lit_start is not None else len(doc.paragraphs)
    list_errors = []
    indent_issues = []

    figure_numbers_found = []
    table_numbers_found = []

    para_to_body_idx = {}
    body_elems = list(doc.element.body)
    for i, elem in enumerate(body_elems):
        if elem.tag == qn('w:p'):
            for j, p in enumerate(doc.paragraphs):
                if p._element is elem:
                    para_to_body_idx[j] = i
                    break

    table_captions_info = []

    try:
        start_body_pos = para_to_body_idx[start_idx]
    except:
        start_body_pos = 0

    prev_block_type = None          # 'section', 'subsection' или 'text'
    for idx in range(start_idx, end_idx):
        p = doc.paragraphs[idx]
        text = p.text.strip()
        norm_text = normalize_text(text)

        if not text:
            prev_para_empty = True
            if list_errors:
                first_text = list_errors[0][1]
                auto_issues.append(f"Список начиная с «{first_text[:50]}» и далее – замените круглый маркер (•) на тире, букву или цифру")
                list_errors = []
            continue

        if has_page_number(text):
            prev_para_empty = False
            continue

        is_level1 = is_section_header(norm_text)
        is_subsection = False

        if not is_level1:
            if re.match(r'^\d+\.\d+(\.\d+)?\s*[А-Яа-я]', norm_text):
                is_subsection = True
            else:
                normalized = normalize_title(norm_text)
                in_toc = any(normalize_title(e) == normalized for e in toc_entries) if toc_entries else False
                if in_toc and len(text) > 20:
                    is_subsection = True

        # Определяем тип текущего блока (раздел, подраздел, текст)
        if is_level1:
            current_block_type = "section"
        elif is_subsection:
            current_block_type = "subsection"
        else:
            current_block_type = "text"

        # ================== ПРОВЕРКА ОСНОВНОГО ТЕКСТА ==================
        if not is_level1 and not is_subsection:
            # Списки
            is_list, marker_type, marker_valid = get_list_marker_info(p, doc)
            if is_list:
                if not marker_valid:
                    if list_errors and idx == list_errors[-1][0] + 1:
                        list_errors.append((idx, text, marker_type))
                    else:
                        if list_errors:
                            first_text = list_errors[0][1]
                            auto_issues.append(f"Список начиная с «{first_text[:50]}» и далее – замените круглый маркер (•) на тире, букву или цифру")
                            list_errors = []
                        list_errors.append((idx, text, marker_type))
                else:
                    if list_errors:
                        first_text = list_errors[0][1]
                        auto_issues.append(f"Список начиная с «{first_text[:50]}» и далее – замените круглый маркер (•) на тире, букву или цифру")
                        list_errors = []
                prev_para_empty = False
                continue

            if list_errors:
                first_text = list_errors[0][1]
                auto_issues.append(f"Список начиная с «{first_text[:50]}» и далее – замените круглый маркер (•) на тире, букву или цифру")
                list_errors = []

            # Продолжение таблицы
            if is_table_continuation(norm_text):
                first_line = get_effective_first_line_indent(p)
                if abs(first_line) > 0.1:
                    auto_issues.append(f"«{text[:50]}» – уберите абзацный отступ (должен быть 0 см)")
                prev_para_empty = False
                continue

            # Формулы
            if is_formula_where_line(norm_text):
                errors = check_formula_explanation(text, p, prev_was_formula, prev_para_empty)
                auto_issues.extend(errors)
                prev_para_empty = False
                continue

            if is_formula_or_equation(norm_text):
                prev_was_formula = True
                prev_para_empty = False
                continue

            # Рисунки
            is_figure = (norm_text.startswith("Рисунок") or norm_text.startswith("Рис.")) and not norm_text.startswith("Таблица")
            is_table_caption = norm_text.startswith("Таблица")

            if is_figure:
                num_match = re.search(r'(?:Рисунок|Рис\.)\s*:?\s*(\d+(?:\.\d+)?)', norm_text)
                if num_match:
                    fig_num = num_match.group(1)
                    figure_numbers_found.append(float(fig_num))
                else:
                    figure_counter += 1
                    fig_num = str(figure_counter)
                fig_number = f"Рисунок {fig_num}"

                if norm_text.startswith("Рис."):
                    auto_issues.append(f"{fig_number} – измените «Рис.» на «Рисунок»")
                if re.match(r'Рисунок\s*:', norm_text):
                    auto_issues.append(f"{fig_number} – замените двоеточие на тире (формат: «Рисунок N — Название»)")
                if not re.search(r'Рисунок\s+\d+(?:\.\d+)?\s*[–—]', norm_text) and not re.match(r'Рисунок\s*:', norm_text):
                    if re.search(r'Рисунок\s+\d+(?:\.\d+)?\s*[-]', norm_text):
                        auto_issues.append(f"{fig_number} – замените дефис на тире (—)")
                    else:
                        auto_issues.append(f"{fig_number} – должно быть тире после номера")
                if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.CENTER:
                    auto_issues.append(f"{fig_number} – выровняйте подпись по центру")
                if text.endswith(".") and not re.search(r'\([^)]*\)\.$', text):
                    auto_issues.append(f"{fig_number} – удалите точку в конце")
                m = re.match(r'^(?:Рисунок|Рис\.)\s+\d+(?:\.\d+)?\s*[–—]\s*(.+)$', norm_text)
                if m:
                    title = m.group(1).strip()
                    if title and title[0].islower():
                        auto_issues.append(f"{fig_number} – название должно начинаться с большой буквы")

                sizes = get_font_size_pt(p)
                if sizes:
                    if any(abs(s - 14) > 0.5 for s in sizes):
                        auto_issues.append(f"{fig_number} – установите размер шрифта 14 пт (сейчас {', '.join(str(s) for s in sizes)} пт)")

                body_idx = para_to_body_idx.get(idx)
                if body_idx is not None:
                    empty_errors = check_empty_line_before_after(doc, body_idx, start_body_pos, fig_number)
                    auto_issues.extend(empty_errors)
                prev_para_empty = False
                continue

            elif is_table_caption and not is_table_continuation(norm_text):
                tbl_match = re.match(r'Таблица\s+(\d+(?:\.\d+)?)', norm_text)
                if tbl_match:
                    tbl_num = tbl_match.group(1)
                    tbl_num_float = float(tbl_num)
                    try:
                        table_numbers_found.append(int(tbl_num))
                    except:
                        pass
                    key = f"Таблица {tbl_num}"

                    if not re.match(r'Таблица\s+\d+(?:\.\d+)?\s+[–—]\s+\S', norm_text):
                        auto_issues.append(f"{key} – Исправьте название на «Таблица {tbl_num} – Название»")

                    if text.rstrip().endswith("."):
                        auto_issues.append(f"{key} – удалите точку в конце названия")
                    sizes = get_font_size_pt(p)
                    if sizes:
                        if any(abs(s - 14) > 0.5 for s in sizes):
                            auto_issues.append(f"{key} – установите размер шрифта 14 пт (сейчас {', '.join(str(s) for s in sizes)} пт)")

                    table_captions_info.append({
                        'para_idx': idx,
                        'body_idx': para_to_body_idx.get(idx),
                        'number': tbl_num_float,
                        'number_str': tbl_num,
                        'text': text,
                        'key': key
                    })
                prev_para_empty = False
                continue

            # Обычный текст (не заголовок, не список, не подпись)
            key = norm_text[:50]
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                indent_issues.append((key, first_line))
            if p.paragraph_format.space_before and p.paragraph_format.space_before.pt > 0.5:
                auto_issues.append(f"«{key}» – интервал перед абзацем должен быть 0 пт")

        # ================== ПРОВЕРКА ЗАГОЛОВКОВ ==================
        if is_level1:
            key = text[:80]
            if text.upper() != "ВВЕДЕНИЕ":
                body_idx = para_to_body_idx.get(idx)
                if body_idx is not None and not is_on_new_page(doc, body_idx, start_body_pos):
                    auto_issues.append(f"«{key}» – раздел должен начинаться с новой страницы")
            first_line = get_effective_first_line_indent(p)
            if abs(first_line) > 0.1:
                auto_issues.append(f"«{key}» – уберите абзацный отступ у заголовка")
            if not is_paragraph_bold(p):
                auto_issues.append(f"«{key}» – заголовок раздела должен быть полужирным")
            if re.match(r'^\d+\.', norm_text) and not is_all_caps(norm_text):
                auto_issues.append(f"«{key}» – заголовок раздела должен быть прописными буквами")
            if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.CENTER:
                auto_issues.append(f"«{key}» – выровняйте заголовок по центру")
            if text.endswith("."):
                auto_issues.append(f"«{key}» – удалите точку в конце")
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                auto_issues.append(f"«{key}» – после заголовка должна быть пустая строка")

        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s*', '', norm_text).strip()
            key = f"Подраздел «{sub_name[:50]}»"
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                auto_issues.append(f"{key} – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
            if not is_paragraph_bold(p):
                auto_issues.append(f"{key} – заголовок должен быть полужирным")
            if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.JUSTIFY:
                auto_issues.append(f"{key} – выровняйте по ширине")
            if text.endswith("."):
                auto_issues.append(f"{key} – удалите точку в конце")
            # Пустая строка допустима только сразу после раздела
            if prev_para_empty and prev_block_type != "section":
                auto_issues.append(f"{key} – уберите пустую строку перед подразделом")

        # Обновление состояния для следующей итерации
        prev_block_type = current_block_type
        if text:
            prev_para_empty = False
            prev_was_formula = False
        else:
            prev_para_empty = True

    # --- ЗАВЕРШАЮЩИЕ ПРОВЕРКИ ---
    if indent_issues:
        if len(indent_issues) > 2:
            first_key, first_indent = indent_issues[0]
            auto_issues.append(
                f"Основной текст, начиная со строки «{first_key}» – "
                f"установите абзацный отступ 1,0 см (сейчас {first_indent:.1f} см)"
            )
        else:
            for key, fl in indent_issues:
                auto_issues.append(f"«{key}» – установите абзацный отступ 1,0 см (сейчас {fl:.1f} см)")

    if list_errors:
        first_text = list_errors[0][1]
        auto_issues.append(f"Список начиная с «{first_text[:50]}» и далее – замените круглый маркер (•) на тире, букву или цифру")

    if figure_numbers_found:
        int_nums = sorted(set(int(n) for n in figure_numbers_found if n == int(n)))
        if int_nums:
            if int_nums[0] != 1:
                auto_issues.append("Рисунки – нумерация должна начинаться с 1")
            else:
                for expected in range(1, int_nums[-1] + 1):
                    if expected not in int_nums:
                        auto_issues.append(f"Рисунки – пропущен рисунок {expected}")
                        break

    # Проверка нумерации таблиц
    table_seq_issues = []
    if table_numbers_found:
        nums = sorted(set(table_numbers_found))
        if nums[0] != 1:
            table_seq_issues.append("Таблицы – неверная нумерация")
        else:
            expected = list(range(1, nums[-1] + 1))
            if nums != expected:
                table_seq_issues.append("Таблицы – неверная нумерация")

    # --- ТАБЛИЦЫ (поиск подписей) ---
    end_body_pos = len(body_elems)
    if lit_start is not None:
        try:
            lit_element = doc.paragraphs[lit_start]._element
            end_body_pos = body_elems.index(lit_element)
        except:
            pass

    tables_in_range = []
    for table in doc.tables:
        if get_table_depth(table) > 0:
            continue
        try:
            tbl_pos = body_elems.index(table._element)
        except:
            continue
        if not (start_body_pos < tbl_pos < end_body_pos):
            continue
        if len(table.rows) == 0:
            continue
        tables_in_range.append((tbl_pos, table))

    table_issues = []

    for tbl_pos, table in tables_in_range:
        caption_info = None
        best_distance = None
        for cap in table_captions_info:
            if cap['body_idx'] is None:
                continue
            if cap['body_idx'] >= tbl_pos:
                continue
            tables_between = [pos for pos, _ in tables_in_range if cap['body_idx'] < pos < tbl_pos]
            if tables_between:
                continue
            distance = tbl_pos - cap['body_idx']
            if best_distance is None or distance < best_distance:
                best_distance = distance
                caption_info = cap

        if caption_info:
            key = caption_info['key']
            tbl_num = caption_info['number_str']

            if caption_info['body_idx'] is not None:
                empty_errors = check_empty_line_before_after(
                    doc, caption_info['body_idx'], start_body_pos, key
                )
                table_issues.extend(empty_errors)

            bold_in_table = False
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if run.bold:
                                bold_in_table = True
                                break
                        if bold_in_table:
                            break
                    if bold_in_table:
                        break
                if bold_in_table:
                    break
            if bold_in_table:
                table_issues.append(f"{key} – уберите полужирное начертание внутри таблицы")

            if len(table.rows) > 2:
                current_tbl_index = None
                for i, (pos, _) in enumerate(tables_in_range):
                    if pos == tbl_pos:
                        current_tbl_index = i
                        break
                if current_tbl_index is not None and current_tbl_index + 1 < len(tables_in_range):
                    next_tbl_pos = tables_in_range[current_tbl_index + 1][0]
                else:
                    next_tbl_pos = end_body_pos
                markers_found = False
                for i in range(tbl_pos + 1, next_tbl_pos):
                    if i < len(doc.paragraphs):
                        txt = doc.paragraphs[i].text.strip()
                        if txt and re.search(r'(?:Продолжение|Окончание)\s+таблицы?\s*' + re.escape(tbl_num), txt):
                            markers_found = True
                            break
                if not markers_found:
                    manual_checks.append(
                        f"{key} – проверьте наличие «Продолжение таблицы {tbl_num}» / "
                        f"«Окончание таблицы {tbl_num}» при переносе на следующую страницу"
                    )
        else:
            prev_text = ""
            for i in range(tbl_pos - 1, start_body_pos - 1, -1):
                elem = body_elems[i]
                if elem.tag == qn('w:p'):
                    txt = ''.join(node.text or '' for node in elem.iter() if node.tag == qn('w:t')).strip()
                    if txt:
                        prev_text = txt[:50]
                        break
            table_issues.append(
                f"Таблица – название должно быть перед таблицей (расположена после абзаца: «{prev_text}…»)" if prev_text
                else "Таблица – название должно быть перед таблицей"
            )

    auto_issues.extend(table_seq_issues)
    auto_issues.extend(table_issues)

    # --- СПИСОК ИСТОЧНИКОВ ---
    if lit_start is not None:
        sources_with_issues = 0
        for i in range(lit_start + 1, lit_end):
            source = doc.paragraphs[i]
            txt = source.text.strip()
            if not txt or has_page_number(txt):
                continue
            has_issue = False
            left_indent = get_effective_left_indent(source)
            if abs(left_indent) > 0.1:
                has_issue = True
            first_line = get_effective_first_line_indent(source)
            if abs(first_line - 1.0) > 0.1:
                has_issue = True
            if has_issue:
                sources_with_issues += 1
        if sources_with_issues > 0:
            auto_issues.append(
                "Список источников – проверьте оформление: "
                "отступ слева 0 см, отступ первой строки 1,0 см, "
                "междустрочный интервал 1,2 (множитель), выравнивание по ширине"
            )

    all_issues = auto_issues
    if manual_checks:
        all_issues.append("📋 Для проверки человеком:")
        all_issues.extend(manual_checks)
    return group_issues(all_issues)
