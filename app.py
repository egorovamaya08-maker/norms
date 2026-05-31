import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # ---------- 1. ПОЛЯ СТРАНИЦ (проверяем все секции, но выводим одно сообщение) ----------
    all_fields_ok = True
    for section in doc.sections:
        left = section.left_margin.pt * 25.4 / 72 if section.left_margin else 0
        right = section.right_margin.pt * 25.4 / 72 if section.right_margin else 0
        top = section.top_margin.pt * 25.4 / 72 if section.top_margin else 0
        bottom = section.bottom_margin.pt * 25.4 / 72 if section.bottom_margin else 0
        if abs(left - 20) > 0.5 or abs(right - 20) > 0.5 or abs(top - 20) > 0.5 or abs(bottom - 20) > 0.5:
            all_fields_ok = False
            break
    if not all_fields_ok:
        issues.append("Поля страниц – установите левое 20 мм, правое 20 мм, верхнее 20 мм, нижнее 20 мм")
    
    # ---------- 2. ОПРЕДЕЛЕНИЕ ГРАНИЦ ОСНОВНОГО ТЕКСТА (ИГНОРИРУЕМ СОДЕРЖАНИЕ И ШАПКУ) ----------
    content_start = None
    content_end = None
    for idx, p in enumerate(doc.paragraphs):
        if "СОДЕРЖАНИЕ" in p.text.upper():
            content_start = idx
            # Ищем конец содержания – следующий заголовок раздела (ВВЕДЕНИЕ или 1.)
            for j in range(idx+1, len(doc.paragraphs)):
                txt = doc.paragraphs[j].text.strip()
                if txt.upper() == "ВВЕДЕНИЕ" or re.match(r'^\d+\.\s+[А-Я]{2,}', txt):
                    content_end = j
                    break
            if content_end is None:
                content_end = idx + 15
            break
    
    # ---------- 3. ПРОВЕРКА ОСНОВНОГО ТЕКСТА (ПОСЛЕ СОДЕРЖАНИЯ) ----------
    figure_count = 0
    prev_para_empty = False
    
    # Заголовки разделов (1-го уровня)
    level1_headings = ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]
    level1_num_pattern = re.compile(r'^\d+\.\s+[А-Я]{2,}')  # "1. ТЕКСТ"
    # Подразделы
    subsection_pattern = re.compile(r'^\d+\.\d+\s+[А-Яа-я]')  # "1.1 Эволюция..."
    
    for idx, p in enumerate(doc.paragraphs):
        # Пропускаем содержание и всё до него
        if content_start is not None and idx <= content_end:
            continue
        if content_start is None and idx < 5:
            continue
            
        text = p.text.strip()
        if not text:
            prev_para_empty = True
            continue
        
        p_format = p.paragraph_format
        
        # ------ 3.1 СОДЕРЖАНИЕ (сам заголовок) ------
        if "СОДЕРЖАНИЕ" in text.upper() and len(text) < 20:
            if text.endswith("."):
                issues.append("Содержание – удалите точку в конце")
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
        
        # ------ 3.2 Заголовки разделов ------
        is_level1 = (text in level1_headings) or level1_num_pattern.match(text)
        if is_level1:
            # Выравнивание по центру
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:30]}…» – выровняйте заголовок по центру")
            # Полужирный (проверяем стиль или runs)
            bold_ok = False
            if p.style and p.style.font and p.style.font.bold:
                bold_ok = True
            elif all(run.bold for run in p.runs):
                bold_ok = True
            if not bold_ok:
                issues.append(f"«{text[:30]}…» – заголовок раздела должен быть полужирным")
            # Прописные буквы
            if text != text.upper():
                issues.append(f"«{text[:30]}…» – заголовок раздела должен быть прописными буквами")
            # Без точки в конце
            if text.endswith("."):
                issues.append(f"«{text[:30]}…» – удалите точку в конце заголовка")
            # Пустая строка после
            next_para = doc.paragraphs[idx+1] if idx+1 < len(doc.paragraphs) else None
            if not (next_para and next_para.text.strip() == ""):
                issues.append(f"«{text[:30]}…» – после заголовка должна быть пустая строка")
            # Новая страница перед разделом (кроме ВВЕДЕНИЯ, если оно сразу после содержания)
            if text not in ["ВВЕДЕНИЕ"] or (text == "ВВЕДЕНИЕ" and idx > content_end + 1):
                has_page_break = False
                if idx > 0:
                    prev = doc.paragraphs[idx-1]
                    if prev.runs and any('w:br' in run._element.xml and 'page' in run._element.xml for run in prev.runs):
                        has_page_break = True
                if not has_page_break and p.runs:
                    if any('w:br' in run._element.xml and 'page' in run._element.xml for run in p.runs):
                        has_page_break = True
                if not has_page_break:
                    issues.append(f"«{text[:30]}…» – раздел должен начинаться с новой страницы")
        
        # ------ 3.3 Заголовки подразделов ------
        if subsection_pattern.match(text):
            # Отступ 1,0 см
            if not p_format.first_line_indent:
                issues.append(f"«{text[:30]}…» – установите абзацный отступ 1,0 см для подраздела")
            else:
                indent_cm = p_format.first_line_indent.cm
                if abs(indent_cm - 1.0) > 0.1:
                    issues.append(f"«{text[:30]}…» – установите абзацный отступ 1,0 см для подраздела (сейчас {indent_cm:.1f} см)")
            # Полужирный
            bold_ok = False
            if p.style and p.style.font and p.style.font.bold:
                bold_ok = True
            elif all(run.bold for run in p.runs):
                bold_ok = True
            if not bold_ok:
                issues.append(f"«{text[:30]}…» – заголовок подраздела должен быть полужирным")
            # Первая буква прописная, остальные строчные
            words = text.split(maxsplit=1)
            if len(words) > 1:
                title = words[1]
                if title and (title[0].islower() or any(c.isupper() for c in title[1:])):
                    issues.append(f"«{text[:30]}…» – заголовок подраздела: первая буква прописная, остальные строчные")
            # Без точки в конце
            if text.endswith("."):
                issues.append(f"«{text[:30]}…» – удалите точку в конце заголовка подраздела")
            # Перед подразделом не должно быть пустой строки
            if prev_para_empty:
                issues.append(f"«{text[:30]}…» – уберите пустую строку перед подразделом")
        
        # ------ 3.4 Рисунки ------
        if text.startswith("Рисунок"):
            figure_count += 1
            # Название с большой буквы
            parts = text.split(maxsplit=2)
            if len(parts) >= 3:
                first_word = parts[2]
                if first_word and first_word[0].islower():
                    issues.append(f"Рисунок {figure_count} – название должно начинаться с большой буквы")
            # Для рисунка 3 – пустая строка перед
            if figure_count == 3:
                prev_para = doc.paragraphs[idx-1] if idx > 0 else None
                if prev_para and prev_para.text.strip() != "":
                    issues.append("Рисунок 3 – добавьте пустую строку перед рисунком")
        
        prev_para_empty = False
    
    # ---------- 4. ПРОВЕРКА ТАБЛИЦ (РЕАЛЬНЫХ ТАБЛИЦ WORD) ----------
    # Сначала определим, какие таблицы относятся к основному тексту (после содержания)
    # Получим все элементы body с их позициями
    body_elements = list(doc.element.body)
    # Найдём позицию элемента, соответствующего концу содержания (content_end)
    start_body_index = 0
    if content_end is not None:
        end_elem = doc.paragraphs[content_end]._element
        try:
            start_body_index = body_elements.index(end_elem) + 1
        except:
            start_body_index = 0
    else:
        start_body_index = 0
    
    # Собираем таблицы, идущие после content_end
    main_tables = []
    for table in doc.tables:
        try:
            tbl_index = body_elements.index(table._element)
            if tbl_index >= start_body_index:
                main_tables.append(table)
        except:
            # Если не удалось определить, считаем, что таблица основная
            main_tables.append(table)
    
    # Проверяем каждую основную таблицу
    for t_idx, table in enumerate(main_tables, start=1):
        # Ищем подпись – предыдущий абзац (не пустой) перед таблицей
        caption_para = None
        table_elem = table._element
        table_pos = body_elements.index(table_elem) if table_elem in body_elements else -1
        if table_pos > 0:
            for i in range(table_pos - 1, -1, -1):
                elem = body_elements[i]
                if elem.tag.endswith('p'):
                    # Найдём соответствующий абзац в doc.paragraphs
                    for p in doc.paragraphs:
                        if p._element is elem and p.text.strip():
                            caption_para = p
                            break
                    if caption_para:
                        break
        
        if caption_para and caption_para.text.strip().startswith("Таблица"):
            caption = caption_para.text.strip()
            # Формат подписи
            if not re.match(r'Таблица\s+\d+\s+--\s+', caption):
                issues.append(f"Таблица {t_idx} – оформление подписи: должно быть «Таблица N -- Название»")
            # Номер таблицы
            match = re.search(r'Таблица\s+(\d+)', caption)
            if match:
                num = int(match.group(1))
                if num != t_idx:
                    issues.append(f"Таблица {num} – номер не соответствует реальному порядку (должна быть {t_idx})")
            # Точка в конце
            if caption.rstrip().endswith("."):
                issues.append(f"Таблица {t_idx} – удалите точку в конце названия")
            # Пустая строка перед подписью
            cap_idx = None
            for k, p in enumerate(doc.paragraphs):
                if p._element is caption_para._element:
                    cap_idx = k
                    break
            if cap_idx is not None and cap_idx > 0:
                prev_para = doc.paragraphs[cap_idx-1]
                if prev_para.text.strip() != "":
                    issues.append(f"Таблица {t_idx} – добавьте пустую строку перед подписью таблицы")
            # Пустая строка после таблицы (следующий абзац после таблицы)
            next_para = None
            for i in range(table_pos + 1, len(body_elements)):
                elem = body_elements[i]
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
        
        # Проверка полужирного шрифта внутри таблицы
        bold_found = False
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.bold:
                            bold_found = True
        if bold_found:
            issues.append(f"Таблица {t_idx} – уберите полужирное начертание")
    
    # ---------- 5. ПРОВЕРКА СПИСКА ИСТОЧНИКОВ ----------
    lit_start = None
    for idx, p in enumerate(doc.paragraphs):
        if "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" in p.text.upper():
            lit_start = idx
            break
    if lit_start is not None:
        # Ищем первый непустой абзац после заголовка
        first_source = None
        for idx in range(lit_start+1, len(doc.paragraphs)):
            if doc.paragraphs[idx].text.strip():
                first_source = doc.paragraphs[idx]
                break
        if first_source:
            p = first_source
            # Отступ слева 0 см
            left_indent = p.paragraph_format.left_indent
            if left_indent and left_indent.cm != 0:
                issues.append("Список источников – отступ слева должен быть 0 см")
            # Отступ первой строки 1 см
            first_line = p.paragraph_format.first_line_indent
            if not first_line or abs(first_line.cm - 1.0) > 0.1:
                issues.append("Список источников – установите отступ первой строки 1,0 см")
            # Межстрочный интервал 1,2
            spacing = p.paragraph_format.line_spacing
            if spacing and p.paragraph_format.line_spacing_rule == 3:
                if abs(spacing - 1.2) > 0.05:
                    issues.append("Список источников – междустрочный интервал должен быть 1,2")
            # Выравнивание по ширине
            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – выровняйте по ширине")
    
    # Удаляем дубликаты
    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

# ---------- ИНТЕРФЕЙС STREAMLIT ----------
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write("Загрузите ваш документ в формате .docx для проверки по чек-листу.")

uploaded_file = st.file_uploader("Перетащите файл сюда или нажмите для выбора", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем документ..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for res in results:
        st.write(f"• {res}")
