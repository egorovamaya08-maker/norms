import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # ---------- 1. ПОЛЯ СТРАНИЦ (общее сообщение) ----------
    margins_ok = True
    for section in doc.sections:
        left = section.left_margin.pt * 25.4 / 72
        right = section.right_margin.pt * 25.4 / 72
        top = section.top_margin.pt * 25.4 / 72
        bottom = section.bottom_margin.pt * 25.4 / 72
        if abs(left-20)>0.5 or abs(right-20)>0.5 or abs(top-20)>0.5 or abs(bottom-20)>0.5:
            margins_ok = False
            break
    if not margins_ok:
        issues.append("Поля страниц – установите левое 20 мм, правое 20 мм, верхнее 20 мм, нижнее 20 мм")
    
    # ---------- 2. ГРАНИЦЫ СОДЕРЖАНИЯ ----------
    content_start = None
    content_end = None
    for i, p in enumerate(doc.paragraphs):
        if "СОДЕРЖАНИЕ" in p.text.upper():
            content_start = i
            # Ищем конец содержания (ближайший заголовок раздела)
            for j in range(i+1, len(doc.paragraphs)):
                txt = doc.paragraphs[j].text.strip()
                if txt.upper() == "ВВЕДЕНИЕ" or re.match(r'^\d+\.\s+[А-Я]', txt):
                    content_end = j
                    break
            if content_end is None:
                content_end = i + 10
            break
    
    # ---------- 3. ТИТУЛЬНЫЕ СТРАНИЦЫ (до содержания) ----------
    if content_start is not None:
        for i in range(min(content_start, 15)):
            p = doc.paragraphs[i]
            if p.text.strip():
                # Интервал 1,2
                if p.paragraph_format.line_spacing_rule == 3:
                    if abs(p.paragraph_format.line_spacing - 1.2) > 0.05:
                        issues.append(f"Титул – междустрочный интервал должен быть 1,2")
                else:
                    issues.append(f"Титул – установите множитель интервала 1,2")
                # Отступ слева 0
                if p.paragraph_format.left_indent and p.paragraph_format.left_indent.cm != 0:
                    issues.append(f"Титул – отступ слева должен быть 0 см")
                # Интервал перед 0
                if p.paragraph_format.space_before and p.paragraph_format.space_before.pt != 0:
                    issues.append(f"Титул – интервал перед абзацем должен быть 0 пт")
    
    # ---------- 4. ОСНОВНАЯ ПРОВЕРКА (после содержания) ----------
    figure_count = 0
    prev_para_empty = False
    
    # Регулярки
    level1_headings = ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]
    level1_num = re.compile(r'^\d+\.\s+[А-Я]{2,}')     # "1. ТЕОРЕТИЧЕСКИЕ..."
    subsection = re.compile(r'^\d+\.\d+\s+[А-Яа-я]')   # "1.1 Эволюция..."
    
    for idx, p in enumerate(doc.paragraphs):
        # Пропускаем содержание
        if content_start is not None and idx <= content_end:
            continue
        
        text = p.text.strip()
        if not text:
            prev_para_empty = True
            continue
        
        # ----- 4.1 Заголовки разделов -----
        is_l1 = (text in level1_headings) or level1_num.match(text)
        if is_l1:
            # Абзацный отступ = 0
            if p.paragraph_format.first_line_indent and p.paragraph_format.first_line_indent.cm != 0:
                issues.append(f"«{text[:40]}…» – уберите абзацный отступ у заголовка")
            # Полужирный
            bold_ok = (p.style and p.style.font.bold) or all(r.bold for r in p.runs)
            if not bold_ok:
                issues.append(f"«{text[:40]}…» – заголовок раздела должен быть полужирным")
            # Прописные
            if text != text.upper():
                issues.append(f"«{text[:40]}…» – заголовок раздела должен быть прописными буквами")
            # По центру
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:40]}…» – выровняйте заголовок по центру")
            # Без точки в конце
            if text.endswith("."):
                issues.append(f"«{text[:40]}…» – удалите точку в конце заголовка")
            # Пустая строка после
            next_para = doc.paragraphs[idx+1] if idx+1 < len(doc.paragraphs) else None
            if not (next_para and next_para.text.strip() == ""):
                issues.append(f"«{text[:40]}…» – после заголовка должна быть пустая строка")
            # Новая страница (кроме ВВЕДЕНИЯ сразу после содержания)
            if text != "ВВЕДЕНИЕ":
                page_break = False
                # Проверяем разрыв в предыдущем абзаце
                if idx > 0:
                    prev = doc.paragraphs[idx-1]
                    if prev.runs:
                        for run in prev.runs:
                            if 'w:br' in run._element.xml and 'page' in run._element.xml:
                                page_break = True
                # Проверяем в самом абзаце
                if not page_break and p.runs:
                    for run in p.runs:
                        if 'w:br' in run._element.xml and 'page' in run._element.xml:
                            page_break = True
                if not page_break:
                    issues.append(f"«{text[:40]}…» – раздел должен начинаться с новой страницы")
        
        # ----- 4.2 Подразделы -----
        if subsection.match(text):
            # Отступ 1,0 см
            if not p.paragraph_format.first_line_indent:
                issues.append(f"«{text[:40]}…» – установите абзацный отступ 1,0 см для подраздела")
            else:
                if abs(p.paragraph_format.first_line_indent.cm - 1.0) > 0.1:
                    issues.append(f"«{text[:40]}…» – установите абзацный отступ 1,0 см (сейчас {p.paragraph_format.first_line_indent.cm:.1f})")
            # Полужирный
            bold_ok = (p.style and p.style.font.bold) or all(r.bold for r in p.runs)
            if not bold_ok:
                issues.append(f"«{text[:40]}…» – заголовок подраздела должен быть полужирным")
            # Первая буква прописная, остальные строчные
            parts = text.split(maxsplit=1)
            if len(parts) > 1:
                title = parts[1]
                if title and (title[0].islower() or any(c.isupper() for c in title[1:])):
                    issues.append(f"«{text[:40]}…» – заголовок подраздела: первая буква прописная, остальные строчные")
            # Без точки
            if text.endswith("."):
                issues.append(f"«{text[:40]}…» – удалите точку в конце заголовка подраздела")
            # Перед подразделом нет пустой строки
            if prev_para_empty:
                issues.append(f"«{text[:40]}…» – уберите пустую строку перед подразделом")
        
        # ----- 4.3 Рисунки -----
        if text.startswith("Рисунок"):
            figure_count += 1
            # Название с большой буквы
            parts = text.split(maxsplit=2)
            if len(parts) >= 3:
                first_word = parts[2]
                if first_word and first_word[0].islower():
                    issues.append(f"Рисунок {figure_count} – название должно начинаться с большой буквы")
            # Для рисунка 3: пустые строки до и после
            if figure_count == 3:
                prev_para = doc.paragraphs[idx-1] if idx > 0 else None
                next_para = doc.paragraphs[idx+1] if idx+1 < len(doc.paragraphs) else None
                if prev_para and prev_para.text.strip() != "":
                    issues.append("Рисунок 3 – добавьте пустую строку перед рисунком")
                if next_para and next_para.text.strip() != "":
                    issues.append("Рисунок 3 – добавьте пустую строку после рисунка")
        
        prev_para_empty = False
    
    # ---------- 5. ТАБЛИЦЫ (только после содержания) ----------
    # Найдём позицию в body, после которой начинаются основные таблицы
    main_tables = []
    if content_end is not None:
        try:
            end_elem = doc.paragraphs[content_end]._element
            end_pos = list(doc.element.body).index(end_elem)
            for table in doc.tables:
                if list(doc.element.body).index(table._element) > end_pos:
                    main_tables.append(table)
        except:
            main_tables = doc.tables  # fallback
    else:
        main_tables = doc.tables
    
    for t_idx, table in enumerate(main_tables, start=1):
        # Ищем подпись (предыдущий непустой абзац)
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
                if doc.paragraphs[cap_idx-1].text.strip() != "":
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
            if next_para and next_para.text.strip() != "":
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
    
    # ---------- 6. СПИСОК ИСТОЧНИКОВ ----------
    lit_start = None
    for i, p in enumerate(doc.paragraphs):
        if "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" in p.text.upper():
            lit_start = i
            break
    if lit_start is not None:
        first_source = None
        for i in range(lit_start+1, len(doc.paragraphs)):
            if doc.paragraphs[i].text.strip():
                first_source = doc.paragraphs[i]
                break
        if first_source:
            p = first_source
            # Отступ слева 0
            if p.paragraph_format.left_indent and p.paragraph_format.left_indent.cm != 0:
                issues.append("Список источников – отступ слева должен быть 0 см")
            # Первая строка 1 см
            if not p.paragraph_format.first_line_indent or abs(p.paragraph_format.first_line_indent.cm - 1.0) > 0.1:
                issues.append("Список источников – установите отступ первой строки 1,0 см")
            # Интервал 1,2
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
