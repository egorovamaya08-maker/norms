import streamlit as st
import docx
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # ---------- 1. ПОЛЯ СТРАНИЦ ----------
    for i, section in enumerate(doc.sections, start=1):
        left_mm = section.left_margin.pt * 25.4 / 72 if section.left_margin else 0
        right_mm = section.right_margin.pt * 25.4 / 72 if section.right_margin else 0
        top_mm = section.top_margin.pt * 25.4 / 72 if section.top_margin else 0
        bottom_mm = section.bottom_margin.pt * 25.4 / 72 if section.bottom_margin else 0
        
        if (abs(left_mm - 20) > 0.5 or abs(right_mm - 20) > 0.5 or 
            abs(top_mm - 20) > 0.5 or abs(bottom_mm - 20) > 0.5):
            issues.append(f"Раздел {i} – установите поля: левое 20, правое 20, верх 20, низ 20 мм")
    
    # ---------- 2. ОПРЕДЕЛЕНИЕ ГРАНИЦ ОСНОВНОГО ТЕКСТА (ИГНОРИРУЕМ СОДЕРЖАНИЕ) ----------
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
            # Если не нашли, то конец содержания – через 10 абзацев (защита)
            if content_end is None:
                content_end = idx + 10
            break
    
    # ---------- 3. ПРОВЕРКА ОСНОВНОГО ТЕКСТА (ПОСЛЕ СОДЕРЖАНИЯ) ----------
    figure_count = 0
    prev_para_empty = False
    
    # Шаблоны заголовков
    level1_headings = ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]
    level1_num_pattern = re.compile(r'^\d+\.\s+[А-Я]{2,}')  # "1. ТЕКСТ", "2. МЕТОДИКА..."
    subsection_pattern = re.compile(r'^\d+\.\d+\s+[А-Яа-я]')  # "1.1 Эволюция..."
    
    for idx, p in enumerate(doc.paragraphs):
        # Пропускаем всё, что относится к содержанию
        if content_start is not None and idx <= content_end:
            continue
        if content_start is None and idx < 5:  # если нет "СОДЕРЖАНИЕ", пропускаем первые 5 абзацев
            continue
            
        text = p.text.strip()
        if not text:
            prev_para_empty = True
            continue
        
        p_format = p.paragraph_format
        
        # ------ 3.1 Шрифт и размер (для основного текста, кроме подписей рис. и табл.) ------
        if not text.startswith(("Рисунок", "Таблица")):
            for run in p.runs:
                if run.font.name and run.font.name != "Times New Roman":
                    issues.append(f"«{text[:30]}…» – смените шрифт на Times New Roman")
                if run.font.size and run.font.size != Pt(14):
                    issues.append(f"«{text[:30]}…» – установите размер шрифта 14")
                if run.underline:
                    issues.append(f"«{text[:30]}…» – удалите подчеркивания")
        
        # ------ 3.2 Междустрочный интервал 1,2 (кроме заголовков и подписей) ------
        ignore_spacing = (text.isupper() or text.startswith(("Рисунок", "Таблица")) or 
                          level1_num_pattern.match(text) or any(text.startswith(h) for h in level1_headings) or
                          subsection_pattern.match(text))
        if not ignore_spacing:
            spacing = p_format.line_spacing
            if spacing:
                if p_format.line_spacing_rule == 3:  # MULTIPLE
                    if abs(spacing - 1.2) > 0.05:
                        issues.append(f"«{text[:30]}…» – замените интервал на 1,2")
                else:
                    issues.append(f"«{text[:30]}…» – установите множитель междустрочного интервала 1,2")
        
        # ------ 3.3 Абзацный отступ 1,0 см (только для обычного текста) ------
        indent = p_format.first_line_indent
        ignore_indent = (text.isupper() or text.startswith(("Рисунок", "Таблица")) or 
                         level1_num_pattern.match(text) or any(text.startswith(h) for h in level1_headings) or
                         subsection_pattern.match(text))
        if indent and not ignore_indent:
            indent_cm = indent.cm
            if abs(indent_cm - 1.0) > 0.1:
                issues.append(f"«{text[:30]}…» – установите абзацный отступ 1,0 см (сейчас {indent_cm:.1f} см)")
        
        # ------ 3.4 Заголовки разделов ------
        is_level1 = (text in level1_headings) or level1_num_pattern.match(text)
        if is_level1:
            # Полужирный
            bold_ok = False
            if p.style and p.style.font.bold:
                bold_ok = True
            elif all(run.bold for run in p.runs):
                bold_ok = True
            if not bold_ok:
                issues.append(f"«{text[:30]}…» – заголовок раздела должен быть полужирным")
            # Прописные буквы
            if text != text.upper():
                issues.append(f"«{text[:30]}…» – заголовок раздела должен быть прописными буквами")
            # Выравнивание по центру
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:30]}…» – выровняйте заголовок по центру")
            # Без абзацного отступа
            if p_format.first_line_indent and p_format.first_line_indent.cm != 0:
                issues.append(f"«{text[:30]}…» – уберите абзацный отступ у заголовка")
            # Без точки в конце
            if text.endswith("."):
                issues.append(f"«{text[:30]}…» – удалите точку в конце заголовка")
            # Пустая строка после
            next_para = doc.paragraphs[idx+1] if idx+1 < len(doc.paragraphs) else None
            if not (next_para and next_para.text.strip() == ""):
                issues.append(f"«{text[:30]}…» – после заголовка должна быть пустая строка")
            # Новая страница перед разделом (кроме первого раздела после содержания)
            # Если раздел идёт сразу после содержания (ВВЕДЕНИЕ) – не требуем, иначе требуем
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
        
        # ------ 3.5 Заголовки подразделов (1.1, 1.2, ...) ------
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
            if p.style and p.style.font.bold:
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
        
        # ------ 3.6 Рисунки ------
        if text.startswith("Рисунок"):
            figure_count += 1
            # Точка в конце
            if text.rstrip().endswith("."):
                issues.append(f"Рисунок {figure_count} – удалите точку в конце названия")
            # Выравнивание по центру
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_count} – выровняйте подпись по центру")
            # Название с большой буквы
            parts = text.split(maxsplit=2)
            if len(parts) >= 3:
                first_word = parts[2]
                if first_word and first_word[0].islower():
                    issues.append(f"Рисунок {figure_count} – название должно начинаться с большой буквы")
            # Для Рисунка 3 – добавьте пустую строку перед рисунком
            if figure_count == 3:
                prev_para = doc.paragraphs[idx-1] if idx > 0 else None
                if prev_para and prev_para.text.strip() != "":
                    issues.append("Рисунок 3 – добавьте пустую строку перед рисунком")
        
        prev_para_empty = False
    
    # ---------- 4. ПРОВЕРКА ТАБЛИЦ (полностью) ----------
    # Собираем все таблицы и сопоставляем с подписями
    # Игнорируем таблицы, которые находятся до content_end
    # Для этого определим индекс первой таблицы после содержания
    first_table_index = None
    for table in doc.tables:
        # Получим XML-элемент таблицы
        try:
            tbl_elem = table._element
            tbl_pos = list(doc.element.body).index(tbl_elem)
            # Найдём позицию content_end в body
            if content_end is not None:
                end_elem = doc.paragraphs[content_end]._element
                end_pos = list(doc.element.body).index(end_elem)
                if tbl_pos > end_pos:
                    first_table_index = tbl_pos
                    break
        except:
            pass
    
    # Если первая таблица найдена, собираем все последующие
    main_tables = []
    if first_table_index is not None:
        for table in doc.tables:
            try:
                if list(doc.element.body).index(table._element) >= first_table_index:
                    main_tables.append(table)
            except:
                pass
    else:
        # Если не удалось определить, проверяем все таблицы (но это рискованно)
        main_tables = doc.tables
    
    # Проходим по основным таблицам
    for t_idx, table in enumerate(main_tables, start=1):
        # Ищем подпись (предыдущий непустой абзац)
        caption_para = None
        # Найдём позицию таблицы в doc.paragraphs (приблизительно)
        # Проще: перебираем абзацы и ищем таблицу после них
        table_elem = table._element
        table_pos = list(doc.element.body).index(table_elem)
        # Ищем предыдущий абзац
        for i in range(table_pos - 1, -1, -1):
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
            # Проверка формата "Таблица N -- Название"
            if not re.match(r'Таблица\s+\d+\s+--\s+', caption):
                issues.append(f"Таблица {t_idx} – оформление подписи: должно быть «Таблица N -- Название»")
            # Сверка номера
            match = re.search(r'Таблица\s+(\d+)', caption)
            if match:
                num = int(match.group(1))
                if num != t_idx:
                    issues.append(f"Таблица {num} – номер не соответствует реальному порядку (должна быть {t_idx})")
            # Точка в конце
            if caption.rstrip().endswith("."):
                issues.append(f"Таблица {t_idx} – удалите точку в конце названия")
            # Пустая строка перед подписью
            # Найдём абзац перед подписью
            cap_idx = None
            for k, p in enumerate(doc.paragraphs):
                if p._element is caption_para._element:
                    cap_idx = k
                    break
            if cap_idx is not None and cap_idx > 0:
                prev_para = doc.paragraphs[cap_idx-1]
                if prev_para.text.strip() != "":
                    issues.append(f"Таблица {t_idx} – добавьте пустую строку перед подписью таблицы")
            # Пустая строка после таблицы
            # Найдём следующий абзац после таблицы
            next_para = None
            for i in range(table_pos + 1, len(doc.element.body)):
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
            # Нет подписи
            issues.append(f"Таблица {t_idx} – отсутствует подпись над таблицей")
    
    # Проверка полужирного шрифта внутри таблиц (для всех основных таблиц)
    for t_idx, table in enumerate(main_tables, start=1):
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
        # Ищем первый непустой абзац после заголовка (первый источник)
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
st.write("Загрузите ваш документ в формате .docx для проверки по полному чек-листу (поля, интервалы, отступы, заголовки, таблицы, рисунки, список литературы).")

uploaded_file = st.file_uploader("Перетащите файл сюда или нажмите для выбора", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем документ..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for res in results:
        st.write(f"• {res}")
