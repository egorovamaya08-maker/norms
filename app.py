import streamlit as st
import docx
from docx.shared import Pt, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import re

def get_effective_alignment(paragraph):
    """
    Возвращает действующее выравнивание абзаца: сначала явно заданное,
    затем наследуемое от стиля. Если не удалось определить — None.
    """
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
    """Проверяет, является ли весь текст абзаца полужирным."""
    # Проверка через стиль
    try:
        if paragraph.style and paragraph.style.font and paragraph.style.font.bold:
            return True
    except Exception:
        pass
    # Проверка по прогонам (run)
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
        # Используем миллиметры напрямую
        if (abs(section.left_margin.mm - 20) > 0.5 or
            abs(section.right_margin.mm - 20) > 0.5 or
            abs(section.top_margin.mm - 20) > 0.5 or
            abs(section.bottom_margin.mm - 20) > 0.5):
            margins_ok = False
            break
    if not margins_ok:
        issues.append("Поля страниц – установите левое 20 мм, правое 20 мм, верхнее 20 мм, нижнее 20 мм")

    # ---------- 2. ГРАНИЦЫ ОСНОВНОГО ТЕКСТА ----------
    content_idx = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().upper() == "СОДЕРЖАНИЕ":
            content_idx = i
            break

    if content_idx is not None:
        # Ищем первый заголовок раздела после содержания
        start_idx = None
        for i in range(content_idx + 1, len(doc.paragraphs)):
            txt = doc.paragraphs[i].text.strip()
            if (txt.upper() in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
                or re.match(r'^\d+\.\s+[А-Я]{2,}', txt)):
                start_idx = i
                break
        if start_idx is None:
            start_idx = content_idx + 1   # если разделов не найдено, считаем всё после содержания основным текстом
    else:
        # Если содержания нет – начинаем с начала (риск проверки титулов, но без ориентира иначе)
        start_idx = 0

    # ---------- 3. ПРОВЕРКА ЗАГОЛОВКА «СОДЕРЖАНИЕ» ----------
    if content_idx is not None:
        p = doc.paragraphs[content_idx]
        text = p.text.strip()
        if text.endswith("."):
            issues.append("Содержание – удалите точку в конце")
        if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.CENTER:
            issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
        if not is_paragraph_bold(p):
            issues.append("Содержание – сделайте заголовок полужирным")
        if (content_idx + 1 < len(doc.paragraphs)
            and not is_empty_paragraph(doc.paragraphs[content_idx + 1])):
            issues.append("Содержание – после заголовка должна быть пустая строка")

    # ---------- 4. ОСНОВНАЯ ПРОВЕРКА (начиная с start_idx) ----------
    figure_counter = 0
    prev_para_empty = False

    level1_headings = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    level1_re = re.compile(r'^\d+\.\s+[А-Я]{2,}')
    subsection_re = re.compile(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]')

    for idx, p in enumerate(doc.paragraphs):
        if idx < start_idx:
            continue
        text = p.text.strip()
        if not text:
            prev_para_empty = True
            continue

        pf = p.paragraph_format

        # --- 4.1 Заголовки разделов первого уровня ---
        is_level1 = (text in level1_headings) or level1_re.match(text)
        if is_level1:
            # Новая страница (кроме ВВЕДЕНИЯ, если оно идёт сразу после содержания – упростим:
            # требование не предъявляем только для ВВЕДЕНИЯ при любом расположении)
            if text != "ВВЕДЕНИЕ":
                # проверяем наличие разрыва страницы перед этим абзацем
                page_break = False
                # в предыдущем абзаце (если есть)
                if idx > 0:
                    prev_p = doc.paragraphs[idx - 1]
                    for run in prev_p.runs:
                        if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                            page_break = True
                # в самом абзаце
                if not page_break:
                    for run in p.runs:
                        if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                            page_break = True
                if not page_break:
                    issues.append(f"«{text[:50]}» – раздел должен начинаться с новой страницы")

            # Отступ первой строки 0
            if pf.first_line_indent and abs(pf.first_line_indent.cm) > 0.1:
                issues.append(f"«{text[:50]}» – уберите абзацный отступ у заголовка")
            # Полужирный
            if not is_paragraph_bold(p):
                issues.append(f"«{text[:50]}» – заголовок раздела должен быть полужирным")
            # Прописные буквы
            if text != text.upper():
                issues.append(f"«{text[:50]}» – заголовок раздела должен быть прописными буквами")
            # Выравнивание по центру
            if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:50]}» – выровняйте заголовок по центру")
            # Без точки в конце
            if text.endswith("."):
                issues.append(f"«{text[:50]}» – удалите точку в конце заголовка")
            # Пустая строка после
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                issues.append(f"«{text[:50]}» – после заголовка должна быть пустая строка")

        # --- 4.2 Подразделы ---
        elif subsection_re.match(text):
            # Убираем номер, оставляем название
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()
            # Отступ первой строки 1,0 см
            if not pf.first_line_indent:
                issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см")
            else:
                if abs(pf.first_line_indent.cm - 1.0) > 0.2:
                    issues.append(f"Подраздел «{sub_name[:50]}» – установите абзацный отступ 1,0 см (сейчас {pf.first_line_indent.cm:.1f})")
            # Полужирный
            if not is_paragraph_bold(p):
                issues.append(f"Подраздел «{sub_name[:50]}» – заголовок должен быть полужирным")
            # Первая буква прописная, остальные строчные
            if sub_name:
                if sub_name[0].islower() or any(c.isupper() for c in sub_name[1:]):
                    issues.append(f"Подраздел «{sub_name[:50]}» – первая буква прописная, остальные строчные")
            # Без точки в конце
            if text.endswith("."):
                issues.append(f"Подраздел «{sub_name[:50]}» – удалите точку в конце")
            # Пустая строка перед подразделом не допускается
            if prev_para_empty:
                issues.append(f"Подраздел «{sub_name[:50]}» – уберите пустую строку перед подразделом")

        # --- 4.3 Обычный текст (не рисунок) ---
        elif not text.startswith("Рисунок"):
            # Абзацный отступ 1,0 см
            if not pf.first_line_indent:
                issues.append(f"«{text[:50]}» – установите абзацный отступ 1,0 см")
            else:
                if abs(pf.first_line_indent.cm - 1.0) > 0.2:
                    issues.append(f"«{text[:50]}» – установите абзацный отступ 1,0 см (сейчас {pf.first_line_indent.cm:.1f})")
            # Интервал перед абзацем 0 пт
            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append(f"«{text[:50]}» – интервал перед абзацем должен быть 0 пт")

        # --- 4.4 Рисунки ---
        if text.startswith("Рисунок"):
            figure_counter += 1
            # Выравнивание по центру
            if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – выровняйте подпись по центру")
            # Точка в конце
            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – удалите точку в конце")
            # Название с большой буквы
            m = re.match(r'^Рисунок\s+\d+\s*[–-]\s*(.+)$', text)
            if m:
                title = m.group(1).strip()
                if title and title[0].islower():
                    issues.append(f"Рисунок {figure_counter} – название должно начинаться с большой буквы")
            # Пустая строка до рисунка
            if idx > 0 and not is_empty_paragraph(doc.paragraphs[idx - 1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку перед рисунком")
            # Пустая строка после рисунка
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                issues.append(f"Рисунок {figure_counter} – добавьте пустую строку после рисунка")

        prev_para_empty = False

    # ---------- 5. ТАБЛИЦЫ (только в основном тексте) ----------
    # Определяем позицию первого абзаца основного текста в теле документа
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
        # Поиск подписи – ближайший непустой абзац перед таблицей
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
            # Формат «Таблица N -- Название»
            if not re.match(r'Таблица\s+\d+\s+--\s+', caption):
                issues.append(f"Таблица {t_idx} – оформление подписи: должно быть «Таблица N -- Название»")
            # Номер в подписи должен совпадать с порядковым
            m_num = re.search(r'Таблица\s+(\d+)', caption)
            if m_num and int(m_num.group(1)) != t_idx:
                issues.append(f"Таблица {t_idx} – номер в подписи не соответствует порядковому (должен быть {t_idx})")
            # Точка в конце названия
            if caption.rstrip().endswith("."):
                issues.append(f"Таблица {t_idx} – удалите точку в конце названия")
            # Пустая строка перед подписью
            cap_idx = None
            for i, para in enumerate(doc.paragraphs):
                if para._element is caption_para._element:
                    cap_idx = i
                    break
            if cap_idx is not None and cap_idx > 0 and not is_empty_paragraph(doc.paragraphs[cap_idx - 1]):
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

        # Полужирный шрифт внутри таблицы не допускается
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
        if p.text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
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
            # Отступ слева 0 см
            if pf.left_indent and abs(pf.left_indent.cm) > 0.1:
                issues.append("Список источников – отступ слева должен быть 0 см")
            # Отступ первой строки 1 см
            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                issues.append("Список источников – установите отступ первой строки 1,0 см")
            # Междустрочный интервал 1,2 (множитель)
            if pf.line_spacing_rule == WD_LINE_SPACING.MULTIPLE:
                if abs(pf.line_spacing - 1.2) > 0.05:
                    issues.append("Список источников – междустрочный интервал должен быть 1,2")
            else:
                issues.append("Список источников – установите множитель междустрочного интервала 1,2")
            # Выравнивание по ширине
            if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – выровняйте по ширине")

    # ---------- ИТОГ ----------
    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

# Интерфейс Streamlit
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
