import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # ---------- 1. ПОЛЯ СТРАНИЦ (общее сообщение) ----------
    page_margins_ok = True
    for section in doc.sections:
        left_mm = section.left_margin.pt * 25.4 / 72
        right_mm = section.right_margin.pt * 25.4 / 72
        top_mm = section.top_margin.pt * 25.4 / 72
        bottom_mm = section.bottom_margin.pt * 25.4 / 72
        if (abs(left_mm - 20) > 0.5 or abs(right_mm - 20) > 0.5 or 
            abs(top_mm - 20) > 0.5 or abs(bottom_mm - 20) > 0.5):
            page_margins_ok = False
            break
    if not page_margins_ok:
        issues.append("Поля страниц – установите левое 20 мм, правое 20 мм, верхнее 20 мм, нижнее 20 мм")
    
    # ---------- 2. ОПРЕДЕЛЕНИЕ ГРАНИЦ ----------
    # Найдём индекс абзаца "СОДЕРЖАНИЕ" и конец содержания
    content_start = None
    content_end = None
    for idx, p in enumerate(doc.paragraphs):
        if "СОДЕРЖАНИЕ" in p.text.upper():
            content_start = idx
            for j in range(idx+1, len(doc.paragraphs)):
                txt = doc.paragraphs[j].text.strip()
                if txt.upper() == "ВВЕДЕНИЕ" or re.match(r'^\d+\.\s+[А-Я]', txt):
                    content_end = j
                    break
            if content_end is None:
                content_end = idx + 10
            break
    
    # ---------- 3. ПРОВЕРКА ОСНОВНОГО ТЕКСТА (после содержания) ----------
    figure_count = 0
    prev_para_empty = False
    
    # Шаблоны заголовков
    level1_headings = ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]
    level1_num_pattern = re.compile(r'^\d+\.\s+[А-Я]{2,}')  # "1. ТЕОРЕТИЧЕСКИЕ..."
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
        
        # ------ 3.1 Проверка отступа у заголовка "ВВЕДЕНИЕ" и других разделов ------
        if text.upper() == "ВВЕДЕНИЕ" or level1_num_pattern.match(text) or text in level1_headings:
            if p_format.first_line_indent and p_format.first_line_indent.cm != 0:
                issues.append(f"«{text[:30]}…» – уберите абзацный отступ у заголовка")
        
        # ------ 3.2 Проверка новой страницы для разделов ------
        if text in level1_headings or level1_num_pattern.match(text):
            # Для ВВЕДЕНИЯ не требуем (оно идёт после содержания)
            if text != "ВВЕДЕНИЕ":
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
        
        # ------ 3.3 Проверка подразделов: пустая строка перед ------
        if subsection_pattern.match(text):
            if prev_para_empty:
                issues.append(f"«{text[:30]}…» – уберите пустую строку перед подразделом")
        
        # ------ 3.4 Рисунки ------
        if text.startswith("Рисунок"):
            figure_count += 1
            # Название с большой буквы (для всех рисунков)
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
    
    # ---------- 4. ПРОВЕРКА ТАБЛИЦ (только основные, после содержания) ----------
    # Собираем все таблицы, которые находятся после content_end
    main_tables = []
    if content_end is not None:
        # Находим позицию последнего абзаца содержания в body
        try:
            end_para_elem = doc.paragraphs[content_end]._element
            end_pos = list(doc.element.body).index(end_para_elem)
            for table in doc.tables:
                tbl_pos = list(doc.element.body).index(table._element)
                if tbl_pos > end_pos:
                    main_tables.append(table)
        except:
            main_tables = doc.tables
    else:
        main_tables = doc.tables
    
    for t_idx, table in enumerate(main_tables, start=1):
        # Ищем подпись (предыдущий непустой абзац)
        caption_para = None
        table_elem = table._element
        table_pos = list(doc.element.body).index(table_elem)
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
    
    # ---------- 6. ПРОВЕРКА ИНТЕРВАЛОВ НА ПЕРВЫХ СТРАНИЦАХ (до содержания) ----------
    # Проверяем первые абзацы до content_start
    if content_start is not None:
        for idx in range(min(content_start, 10)):
            p = doc.paragraphs[idx]
            if p.text.strip():
                # Междустрочный интервал 1,2
                spacing = p.paragraph_format.line_spacing
                if spacing and p.paragraph_format.line_spacing_rule == 3:
                    if abs(spacing - 1.2) > 0.05:
                        issues.append(f"«{p.text[:30]}…» – междустрочный интервал должен быть 1,2 (титульный лист)")
                # Отступ слева 0 см
                left_indent = p.paragraph_format.left_indent
                if left_indent and left_indent.cm != 0:
                    issues.append(f"«{p.text[:30]}…» – отступ слева должен быть 0 см (титульный лист)")
                # Интервал перед 0 пт
                space_before = p.paragraph_format.space_before
                if space_before and space_before.pt != 0:
                    issues.append(f"«{p.text[:30]}…» – интервал перед абзацем должен быть 0 пт (титульный лист)")
    
    # Удаляем дубликаты
    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

# ---------- ИНТЕРФЕЙС STREAMLIT ----------
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write("Загрузите ваш документ в формате .docx для проверки по полному чек-листу.")

uploaded_file = st.file_uploader("Перетащите файл сюда или нажмите для выбора", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем документ..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for res in results:
        st.write(f"• {res}")
