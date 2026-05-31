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

    # ---------- 2. ПОИСК НАЧАЛА ОСНОВНОГО ТЕКСТА ----------
    # Ищем ТОЛЬКО явные заголовки разделов: "ВВЕДЕНИЕ" или "1. НАЗВАНИЕ"
    # Прописные строки по центру НЕ считаем заголовками разделов — они могут быть титульными
    start_idx = None
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip().upper()
        if not txt:
            continue
        
        # Только ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ
        if txt in level1_keywords:
            start_idx = i
            break
        
        # Только заголовки с номером: "1. НАЗВАНИЕ", "2. НАЗВАНИЕ"
        if re.match(r'^\d+\.\s+[А-Я]', txt):
            start_idx = i
            break
    
    if start_idx is None:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]

    # ---------- 3. ПРОВЕРКА ОСНОВНОГО ТЕКСТА ----------
    figure_counter = 0
    prev_para_empty = False
    subsection_re = re.compile(r'^\d+\.\d+')

    for idx in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[idx]
        text = p.text.strip()
        
        if not text:
            prev_para_empty = True
            continue

        pf = p.paragraph_format
        alignment = get_effective_alignment(p)
        
        # Определяем тип абзаца
        is_level1 = (text.upper() in level1_keywords or
                     bool(re.match(r'^\d+\.\s+[А-Я]', text)))
        
        is_subsection = bool(subsection_re.match(text)) and not is_level1
        is_figure = text.startswith("Рисунок")
        is_table_caption = text.startswith("Таблица")

        # --- Заголовок раздела ---
        if is_level1:
            # Новая страница (кроме ВВЕДЕНИЯ, если оно первое)
            if text.upper() != "ВВЕДЕНИЕ" or idx != start_idx:
                page_break = False
                if idx > start_idx:
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
            # Извлекаем название (всё после номера)
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()
            # Отступ 1 см
            if not pf.first_line_indent:
                issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см")
            elif abs(pf.first_line_indent.cm - 1.0) > 0.2:
                issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см (сейчас {pf.first_line_indent.cm:.1f})")
            # Полужирный
            if not is_paragraph_bold(p):
                issues.append(f"Подраздел «{sub_name[:50]}» – заголовок должен быть полужирным")
            # Без точки в конце
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
            if idx > start_idx and not is_empty_paragraph(doc.paragraphs[idx - 1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку перед рисунком")
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку после рисунка")
        
        # --- Подпись таблицы (пропускаем, проверим отдельно) ---
        elif is_table_caption:
            pass
        
        # --- Обычный текст ---
        else:
            # Отступ 1 см
            if not pf.first_line_indent:
                issues.append(f"«{text[:50]}» – установите абзацный отступ 1,0 см")
            elif abs(pf.first_line_indent.cm - 1.0) > 0.2:
                issues.append(f"«{text[:50]}» – установите абзацный отступ 1,0 см (сейчас {pf.first_line_indent.cm:.1f})")
            # Интервал перед 0 пт
            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append(f"«{text[:50]}» – интервал перед абзацем должен быть 0 пт")
        
        prev_para_empty = False

    # ---------- 4. ТАБЛИЦЫ В ОСНОВНОМ ТЕКСТЕ ----------
    try:
        start_element = doc.paragraphs[start_idx]._element
        start_body_pos = list(doc.element.body).index(start_element)
    except:
        start_body_pos = 0

    main_tables = []
    for table in doc.tables:
        try:
            if list(doc.element.body).index(table._element) > start_body_pos:
                main_tables.append(table)
        except:
            pass

    for t_idx, table in enumerate(main_tables, start=1):
        tbl_pos = list(doc.element.body).index(table._element)
        
        # Ищем подпись таблицы
        caption_para = None
        for i in range(tbl_pos - 1, -1, -1):
            elem = doc.element.body[i]
            if elem.tag.endswith('p'):
                for para in doc.paragraphs:
                    if para._element is elem and para.text.strip():
                        caption_para = para
                        break
                if caption_para:
                    break

        if caption_para and caption_para.text.strip().startswith("Таблица"):
            caption = caption_para.text.strip()
            caption_idx = None
            for i, para in enumerate(doc.paragraphs):
                if para._element is caption_para._element:
                    caption_idx = i
                    break
            
            # Формат "Таблица N -- Название"
            if not re.match(r'Таблица\s+\d+\s+--\s+', caption):
                issues.append(f"Таблица {t_idx} – оформление подписи: должно быть «Таблица N -- Название»")
            
            # Номер совпадает
            m_num = re.search(r'Таблица\s+(\d+)', caption)
            if m_num and int(m_num.group(1)) != t_idx:
                issues.append(f"Таблица {t_idx} – номер в подписи не соответствует порядковому (должен быть {t_idx})")
            
            # Без точки
            if caption.rstrip().endswith("."):
                issues.append(f"Таблица {t_idx} – удалите точку в конце названия")
            
            # Пустая строка перед подписью
            if caption_idx is not None and caption_idx > start_idx:
                if not is_empty_paragraph(doc.paragraphs[caption_idx - 1]):
                    issues.append(f"Таблица {t_idx} – добавьте пустую строку перед подписью таблицы")
            
            # Пустая строка после таблицы
            next_para = None
            for i in range(tbl_pos + 1, len(doc.element.body)):
                elem = doc.element.body[i]
                if elem.tag.endswith('p'):
                    for para in doc.paragraphs:
                        if para._element is elem:
                            next_para = para
                            break
                    break
            if next_para and not is_empty_paragraph(next_para):
                issues.append(f"Таблица {t_idx} – добавьте пустую строку после таблицы")
        else:
            issues.append(f"Таблица {t_idx} – отсутствует подпись над таблицей")

        # Полужирный внутри таблицы
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
            issues.append(f"Таблица {t_idx} – уберите полужирное начертание внутри таблицы")

    # ---------- 5. СПИСОК ИСТОЧНИКОВ ----------
    lit_start = None
    for i in range(start_idx, len(doc.paragraphs)):
        if doc.paragraphs[i].text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            lit_start = i
            break
    
    if lit_start is not None:
        # Проверяем ВСЕ источники и собираем проблемы
        lit_issues = []
        sources_with_issues = 0
        
        for i in range(lit_start + 1, len(doc.paragraphs)):
            source = doc.paragraphs[i]
            if not source.text.strip():
                continue
            
            pf = source.paragraph_format
            has_issue = False
            
            # Отступ слева 0 см
            if pf.left_indent and abs(pf.left_indent.cm) > 0.1:
                has_issue = True
            
            # Отступ первой строки 1 см
            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                has_issue = True
            
            # Междустрочный интервал 1,2
            if pf.line_spacing_rule == WD_LINE_SPACING.MULTIPLE:
                if abs(pf.line_spacing - 1.2) > 0.05:
                    has_issue = True
            else:
                has_issue = True
            
            # Выравнивание по ширине
            if get_effective_alignment(source) != WD_ALIGN_PARAGRAPH.JUSTIFY:
                has_issue = True
            
            if has_issue:
                sources_with_issues += 1
        
        if sources_with_issues > 0:
            issues.append(
                f"Список источников – проверьте оформление: "
                f"отступ слева 0 см, отступ первой строки 1,0 см, "
                f"междустрочный интервал 1,2 (множитель), выравнивание по ширине"
            )

    # ---------- ИТОГ ----------
    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

# Интерфейс
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write("Загрузите документ в формате .docx – проверка по полному чек-листу.")
uploaded_file = st.file_uploader("Выберите файл", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Проверяем..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for r in results:
        st.write(f"• {r}")
