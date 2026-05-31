import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # --------------------------------------------------
    # ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
    # --------------------------------------------------
    def is_empty_paragraph(p):
        return len(p.text.strip()) == 0
    
    def is_bold_paragraph(p):
        try:
            if p.style and p.style.font and p.style.font.bold:
                return True
        except:
            pass
        runs = [r for r in p.runs if r.text.strip()]
        if not runs:
            return False
        return all(r.bold is True for r in runs)
    
    level1_headings = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    def is_level1_heading(text):
        text = text.strip()
        if text in level1_headings:
            return True
        return bool(re.match(r'^\d+\.\s+[А-ЯЁ][А-ЯЁ0-9\s\-\(\)"]+$', text))
    
    def is_subsection(text):
        return bool(re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]', text))
    
    def subsection_name(text):
        return re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()
    
    # --------------------------------------------------
    # ПОИСК ГРАНИЦ СОДЕРЖАНИЯ
    # --------------------------------------------------
    content_start = None
    content_end = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().upper() == "СОДЕРЖАНИЕ":
            content_start = i
            # Ищем конец содержания (следующий заголовок раздела)
            for j in range(i+1, len(doc.paragraphs)):
                txt = doc.paragraphs[j].text.strip().upper()
                if txt == "ВВЕДЕНИЕ" or re.match(r'^\d+\.\s+[А-Я]{2,}', txt) or txt in ["ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]:
                    content_end = j - 1
                    break
            if content_end is None:
                content_end = i + 15  # запас
            break
    
    # --------------------------------------------------
    # ПОЛЯ СТРАНИЦ
    # --------------------------------------------------
    margins_ok = True
    for section in doc.sections:
        left = section.left_margin.pt * 25.4 / 72
        right = section.right_margin.pt * 25.4 / 72
        top = section.top_margin.pt * 25.4 / 72
        bottom = section.bottom_margin.pt * 25.4 / 72
        if abs(left-20) > 0.5 or abs(right-20) > 0.5 or abs(top-20) > 0.5 or abs(bottom-20) > 0.5:
            margins_ok = False
            break
    if not margins_ok:
        issues.append("Поля страниц – установите левое 20 мм, правое 20 мм, верхнее 20 мм, нижнее 20 мм")
    
    # --------------------------------------------------
    # ОСНОВНОЙ ПРОХОД ПО АБЗАЦАМ (ПОСЛЕ СОДЕРЖАНИЯ)
    # --------------------------------------------------
    figure_counter = 0
    prev_para_empty = False
    
    for idx, p in enumerate(doc.paragraphs):
        # Пропускаем всё, что относится к содержанию
        if content_start is not None and content_end is not None and content_start <= idx <= content_end:
            continue
        if content_start is not None and idx < content_start:
            continue
        
        text = p.text.strip()
        if not text:
            prev_para_empty = True
            continue
        
        pf = p.paragraph_format
        
        # ---------- СОДЕРЖАНИЕ (только сам заголовок) ----------
        if text.upper() == "СОДЕРЖАНИЕ":
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
            if text.endswith("."):
                issues.append("Содержание – удалите точку в конце")
            if not is_bold_paragraph(p):
                issues.append("Содержание – сделайте заголовок полужирным")
            # Проверяем пустую строку после
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx+1]):
                issues.append("Содержание – после заголовка должна быть пустая строка")
        
        # ---------- ЗАГОЛОВКИ РАЗДЕЛОВ ----------
        elif is_level1_heading(text):
            section_title = text[:40]
            # Новая страница (кроме ВВЕДЕНИЯ, если оно идёт сразу после содержания)
            if text != "ВВЕДЕНИЕ":
                page_break = False
                if idx > 0:
                    prev = doc.paragraphs[idx-1]
                    for run in prev.runs:
                        if 'w:br' in run._element.xml and 'page' in run._element.xml:
                            page_break = True
                if not page_break and p.runs:
                    for run in p.runs:
                        if 'w:br' in run._element.xml and 'page' in run._element.xml:
                            page_break = True
                if not page_break:
                    issues.append(f"«{section_title}…» – раздел должен начинаться с новой страницы")
            # Абзацный отступ 0
            if pf.first_line_indent and abs(pf.first_line_indent.cm) > 0.1:
                issues.append(f"«{section_title}…» – уберите абзацный отступ у заголовка")
            # Полужирный
            if not is_bold_paragraph(p):
                issues.append(f"«{section_title}…» – заголовок раздела должен быть полужирным")
            # Прописные буквы
            if text != text.upper():
                issues.append(f"«{section_title}…» – заголовок раздела должен быть прописными буквами")
            # Выравнивание по центру
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{section_title}…» – выровняйте заголовок по центру")
            # Без точки в конце
            if text.endswith("."):
                issues.append(f"«{section_title}…» – удалите точку в конце заголовка")
            # Пустая строка после
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx+1]):
                issues.append(f"«{section_title}…» – после заголовка должна быть пустая строка")
        
        # ---------- ПОДРАЗДЕЛЫ (только если это не содержание) ----------
        elif is_subsection(text):
            # Это основной текст, а не содержание
            sub_name = subsection_name(text)
            # Отступ 1,0 см
            if not pf.first_line_indent:
                issues.append(f"Подраздел «{sub_name[:40]}…» – установите абзацный отступ 1,0 см")
            else:
                if abs(pf.first_line_indent.cm - 1.0) > 0.2:
                    issues.append(f"Подраздел «{sub_name[:40]}…» – установите абзацный отступ 1,0 см (сейчас {pf.first_line_indent.cm:.1f})")
            # Полужирный
            if not is_bold_paragraph(p):
                issues.append(f"Подраздел «{sub_name[:40]}…» – заголовок должен быть полужирным")
            # Первая буква прописная, остальные строчные
            if sub_name and (sub_name[0].islower() or any(c.isupper() for c in sub_name[1:])):
                issues.append(f"Подраздел «{sub_name[:40]}…» – первая буква прописная, остальные строчные")
            # Без точки в конце
            if text.endswith("."):
                issues.append(f"Подраздел «{sub_name[:40]}…» – удалите точку в конце")
            # Перед подразделом нет пустой строки
            if prev_para_empty:
                issues.append(f"Подраздел «{sub_name[:40]}…» – уберите пустую строку перед подразделом")
        
        # ---------- ОСНОВНОЙ ТЕКСТ (не заголовки и не рисунки) ----------
        elif not text.startswith("Рисунок"):
            # Абзацный отступ 1,0 см
            if not pf.first_line_indent:
                issues.append(f"«{text[:40]}…» – установите абзацный отступ 1,0 см")
            else:
                if abs(pf.first_line_indent.cm - 1.0) > 0.2:
                    issues.append(f"«{text[:40]}…» – установите абзацный отступ 1,0 см (сейчас {pf.first_line_indent.cm:.1f})")
            # Интервал перед абзацем 0 пт
            if pf.space_before and pf.space_before.pt > 0.5:
                issues.append(f"«{text[:40]}…» – интервал перед абзацем должен быть 0 пт")
        
        # ---------- РИСУНКИ ----------
        if text.startswith("Рисунок"):
            figure_counter += 1
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – выровняйте подпись по центру")
            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – удалите точку в конце")
            # Проверка названия с большой буквы
            m = re.match(r'^Рисунок\s+\d+\s*[–-]\s*(.+)$', text)
            if m:
                title = m.group(1).strip()
                if title and title[0].islower():
                    issues.append(f"Рисунок {figure_counter} – название должно начинаться с большой буквы")
            # Для рисунка 3: пустые строки до и после
            if figure_counter == 3:
                if idx > 0 and not is_empty_paragraph(doc.paragraphs[idx-1]):
                    issues.append("Рисунок 3 – добавьте пустую строку перед рисунком")
                if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx+1]):
                    issues.append("Рисунок 3 – добавьте пустую строку после рисунка")
        
        prev_para_empty = False
    
    # --------------------------------------------------
    # ТАБЛИЦЫ (только после содержания)
    # --------------------------------------------------
    # Определяем таблицы, идущие после content_end
    main_tables = []
    if content_end is not None:
        try:
            end_elem = doc.paragraphs[content_end]._element
            end_pos = list(doc.element.body).index(end_elem)
            for table in doc.tables:
                if list(doc.element.body).index(table._element) > end_pos:
                    main_tables.append(table)
        except:
            main_tables = doc.tables
    else:
        main_tables = doc.tables
    
    for t_idx, table in enumerate(main_tables, start=1):
        # Поиск подписи (предыдущий непустой абзац)
        caption_para = None
        tbl_pos = list(doc.element.body).index(table._element)
        for i in range(tbl_pos-1, -1, -1):
            elem = doc.element.body[i]
            if elem.tag.endswith('p'):
                for p in doc.paragraphs:
                    if p._element is elem and p.text.strip():
                        caption_para = p
                        break
                if caption_para:
                    break
        if caption_para and caption_para.text.strip().startswith("Таблица"):
            caption = caption_para.text.strip()
            # Формат "Таблица N -- Название"
            if not re.match(r'Таблица\s+\d+\s+--\s+', caption):
                issues.append(f"Таблица {t_idx} – оформление подписи: должно быть «Таблица N -- Название»")
            # Номер
            m = re.search(r'Таблица\s+(\d+)', caption)
            if m and int(m.group(1)) != t_idx:
                issues.append(f"Таблица {t_idx} – номер в подписи не соответствует порядковому (должен быть {t_idx})")
            # Точка в конце
            if caption.rstrip().endswith("."):
                issues.append(f"Таблица {t_idx} – удалите точку в конце названия")
            # Пустая строка перед подписью
            cap_idx = None
            for i, p in enumerate(doc.paragraphs):
                if p._element is caption_para._element:
                    cap_idx = i
                    break
            if cap_idx is not None and cap_idx > 0:
                if not is_empty_paragraph(doc.paragraphs[cap_idx-1]):
                    issues.append(f"Таблица {t_idx} – добавьте пустую строку перед подписью таблицы")
            # Пустая строка после таблицы
            next_para = None
            for i in range(tbl_pos+1, len(doc.element.body)):
                elem = doc.element.body[i]
                if elem.tag.endswith('p'):
                    for p in doc.paragraphs:
                        if p._element is elem:
                            next_para = p
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
        if bold_in_table:
            issues.append(f"Таблица {t_idx} – уберите полужирное начертание")
    
    # --------------------------------------------------
    # СПИСОК ИСТОЧНИКОВ
    # --------------------------------------------------
    lit_start = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            lit_start = i
            break
    if lit_start is not None:
        # Ищем первый непустой абзац после заголовка (первый источник)
        first_source = None
        for i in range(lit_start+1, len(doc.paragraphs)):
            if doc.paragraphs[i].text.strip():
                first_source = doc.paragraphs[i]
                break
        if first_source:
            p = first_source
            # Отступ слева 0 см
            if p.paragraph_format.left_indent and abs(p.paragraph_format.left_indent.cm) > 0.1:
                issues.append("Список источников – отступ слева должен быть 0 см")
            # Отступ первой строки 1 см
            if not p.paragraph_format.first_line_indent or abs(p.paragraph_format.first_line_indent.cm - 1.0) > 0.1:
                issues.append("Список источников – установите отступ первой строки 1,0 см")
            # Междустрочный интервал 1,2
            if p.paragraph_format.line_spacing_rule == 3:
                if abs(p.paragraph_format.line_spacing - 1.2) > 0.05:
                    issues.append("Список источников – междустрочный интервал должен быть 1,2")
            else:
                issues.append("Список источников – установите множитель междустрочного интервала 1,2")
            # Выравнивание по ширине
            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – выровняйте по ширине")
    
    # Убираем дубликаты
    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

# ---------- STREAMLIT ИНТЕРФЕЙС ----------
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write("Загрузите ваш документ в формате .docx – проверка по полному чек-листу.")

uploaded_file = st.file_uploader("Выберите файл", type=["docx"])
if uploaded_file is not None:
    with st.spinner("Проверяем..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for r in results:
        st.write(f"• {r}")
