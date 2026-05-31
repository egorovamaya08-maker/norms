import streamlit as st
import docx
from docx.shared import Pt, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import re

def get_effective_alignment(paragraph):
    if paragraph.alignment is not None:
        return paragraph.alignment
    try:
        style = paragraph.style
        if style:
            pf = style.paragraph_format
            if pf.alignment is not None:
                return pf.alignment
    except Exception:
        pass
    return None

def is_paragraph_bold(paragraph):
    try:
        if paragraph.style and paragraph.style.font and paragraph.style.font.bold:
            return True
    except Exception:
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

    # ---------- 2. ГРАНИЦЫ ДОКУМЕНТА ----------
    # Ищем заголовок "СОДЕРЖАНИЕ" (точное совпадение, потом частичное)
    content_idx = None
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip().upper()
        if txt == "СОДЕРЖАНИЕ":
            content_idx = i
            break
    if content_idx is None:
        # Попробуем найти как часть текста (если есть пробелы или другие символы)
        for i, p in enumerate(doc.paragraphs):
            if "СОДЕРЖАНИЕ" in p.text.strip().upper():
                content_idx = i
                break

    if content_idx is None:
        issues.append("Не найден заголовок «СОДЕРЖАНИЕ» – проверка структуры невозможна.")
        # Без содержания не рискуем проверять основной текст, возвращаем только поля
        return issues if issues else ["✅ Ошибок не найдено."]

    # Ищем первый заголовок раздела после содержания
    start_idx = None
    level1_headings = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    level1_re = re.compile(r'^\d+\.\s+[А-Я]{2,}')
    # Дополнительный шаблон: полностью прописная строка длиной > 5 символов
    all_caps_re = re.compile(r'^[А-ЯЁ\s]+$')

    for i in range(content_idx + 1, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        if not txt:
            continue
        # Заголовок раздела с номером (1. Название)
        if level1_re.match(txt):
            start_idx = i
            break
        # Точные названия разделов
        if txt.upper() in level1_headings:
            start_idx = i
            break
        # Заголовок без номера, весь из прописных букв, выровнен по центру
        if (all_caps_re.match(txt) and len(txt) > 5
                and get_effective_alignment(p) == WD_ALIGN_PARAGRAPH.CENTER):
            start_idx = i
            break

    # Если раздел не найден, основной текст не проверяем
    if start_idx is None:
        start_idx = len(doc.paragraphs)  # всё, что после содержания, пропускаем

    # ---------- 3. ПРОВЕРКА ЗАГОЛОВКА «СОДЕРЖАНИЕ» ----------
    p_content = doc.paragraphs[content_idx]
    text = p_content.text.strip()
    if text.endswith("."):
        issues.append("Содержание – удалите точку в конце")
    if get_effective_alignment(p_content) != WD_ALIGN_PARAGRAPH.CENTER:
        issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
    if not is_paragraph_bold(p_content):
        issues.append("Содержание – сделайте заголовок полужирным")
    if (content_idx + 1 < len(doc.paragraphs)
            and not is_empty_paragraph(doc.paragraphs[content_idx + 1])):
        issues.append("Содержание – после заголовка должна быть пустая строка")

    # ---------- 4. ОСНОВНАЯ ПРОВЕРКА (с start_idx) ----------
    figure_counter = 0
    prev_para_empty = False
    subsection_re = re.compile(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]')

    for idx, p in enumerate(doc.paragraphs):
        if idx < start_idx:
            continue
        text = p.text.strip()
        if not text:
            prev_para_empty = True
            continue

        pf = p.paragraph_format

        # --- 4.1 Заголовки разделов (включая безномерные прописные) ---
        is_level1 = (text.upper() in level1_headings or
                     level1_re.match(text) or
                     (all_caps_re.match(text) and len(text) > 5 and
                      get_effective_alignment(p) == WD_ALIGN_PARAGRAPH.CENTER))
        if is_level1:
            if text.upper() != "ВВЕДЕНИЕ":  # для ВВЕДЕНИЯ не требуем новую страницу
                page_break = False
                if idx > 0:
                    prev_p = doc.paragraphs[idx - 1]
                    for run in prev_p.runs:
                        if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                            page_break = True
                if not page_break:
                    for run in p.runs:
                        if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                            page_break = True
                if not page_break:
                    issues.append(f"«{text[:50]}» – раздел должен начинаться с новой страницы")
            # Отступ первой строки 0
            if pf.first_line_indent and abs(pf.first_line_indent.cm) > 0.1:
                issues.append(f"«{text[:50]}» – уберите абзацный отступ у заголовка")
            if not is_paragraph_bold(p):
                issues.append(f"«{text[:50]}» – заголовок раздела должен быть полужирным")
            if text != text.upper():
                issues.append(f"«{text[:50]}» – заголовок раздела должен быть прописными буквами")
            if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:50]}» – выровняйте заголовок по центру")
            if text.endswith("."):
                issues.append(f"«{text[:50]}» – удалите точку в конце заголовка")
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                issues.append(f"«{text[:50]}» – после заголовка должна быть пустая строка")

        # --- 4.2 Подразделы ---
        elif subsection_re.match(text):
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()
            if not pf.first_line_indent:
                issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см")
            else:
                if abs(pf.first_line_indent.cm - 1.0) > 0.2:
                    issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см (сейчас {pf.first_line_indent.cm:.1f})")
            if not is_paragraph_bold(p):
                issues.append(f"Подраздел «{sub_name[:50]}» – заголовок должен быть полужирным")
            if sub_name and (sub_name[0].islower() or any(c.isupper() for c in sub_name[1:])):
                issues.append(f"Подраздел «{sub_name[:50]}» – первая буква прописная, остальные строчные")
            if text.endswith("."):
                issues.append(f"Подраздел «{sub_name[:50]}» – удалите точку в конце")
            if prev_para_empty:
                issues.append(f"Подраздел «{sub_name[:50]}» – уберите пустую строку перед подразделом")

        # --- 4.3 Обычный текст (не рисунок и не таблица) ---
        elif not text.startswith("Рисунок") and not text.startswith("Таблица"):
            if not pf.first_line_indent:
                issues.append(f"«{text[:50]}» – установите абзацный отступ 1,0 см")
            else:
                if abs(pf.first_line_indent.cm - 1.0) > 0.2:
                    issues.append(f"«{text[:50]}» – установите абзацный отступ 1,0 см (сейчас {pf.first_line_indent.cm:.1f})")
            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append(f"«{text[:50]}» – интервал перед абзацем должен быть 0 пт")

        # --- 4.4 Рисунки ---
        if text.startswith("Рисунок"):
            figure_counter += 1
            if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – выровняйте подпись по центру")
            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – удалите точку в конце")
            m = re.match(r'^Рисунок\s+\d+\s*[–-]\s*(.+)$', text)
            if m:
                title = m.group(1).strip()
                if title and title[0].islower():
                    issues.append(f"Рисунок {figure_counter} – название должно начинаться с большой буквы")
            if idx > 0 and not is_empty_paragraph(doc.paragraphs[idx - 1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку перед рисунком")
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку после рисунка")

        prev_para_empty = False

    # ---------- 5. ТАБЛИЦЫ (только в основном тексте) ----------
    try:
        start_element = doc.paragraphs[start_idx]._element
        start_body_pos = list(doc.element.body).index(start_element)
    except Exception:
        start_body_pos = 0

    main_tables = []
    for table in doc.tables:
        try:
            if list(doc.element.body).index(table._element) > start_body_pos:
                main_tables.append(table)
        except Exception:
            pass

    for t_idx, table in enumerate(main_tables, start=1):
        tbl_pos = list(doc.element.body).index(table._element)
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
            if not re.match(r'Таблица\s+\d+\s+--\s+', caption):
                issues.append(f"Таблица {t_idx} – оформление подписи: должно быть «Таблица N -- Название»")
            m_num = re.search(r'Таблица\s+(\d+)', caption)
            if m_num and int(m_num.group(1)) != t_idx:
                issues.append(f"Таблица {t_idx} – номер в подписи не соответствует порядковому (должен быть {t_idx})")
            if caption.rstrip().endswith("."):
                issues.append(f"Таблица {t_idx} – удалите точку в конце названия")
            cap_idx = None
            for i, para in enumerate(doc.paragraphs):
                if para._element is caption_para._element:
                    cap_idx = i
                    break
            if cap_idx is not None and cap_idx > 0 and not is_empty_paragraph(doc.paragraphs[cap_idx - 1]):
                issues.append(f"Таблица {t_idx} – добавьте пустую строку перед подписью таблицы")
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

    # ---------- 6. СПИСОК ИСТОЧНИКОВ ----------
    lit_start = None
    for i, p in enumerate(doc.paragraphs):
        if i >= start_idx and p.text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            lit_start = i
            break
    if lit_start is not None:
        first_source = None
        for i in range(lit_start + 1, len(doc.paragraphs)):
            if doc.paragraphs[i].text.strip():
                first_source = doc.paragraphs[i]
                break
        if first_source:
            p = first_source
            pf = p.paragraph_format
            if pf.left_indent and abs(pf.left_indent.cm) > 0.1:
                issues.append("Список источников – отступ слева должен быть 0 см")
            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                issues.append("Список источников – установите отступ первой строки 1,0 см")
            if pf.line_spacing_rule == WD_LINE_SPACING.MULTIPLE:
                if abs(pf.line_spacing - 1.2) > 0.05:
                    issues.append("Список источников – междустрочный интервал должен быть 1,2")
            else:
                issues.append("Список источников – установите множитель междустрочного интервала 1,2")
            if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – выровняйте по ширине")

    # Удаляем дубликаты
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
