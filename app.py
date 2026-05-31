import streamlit as st
import docx
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
import re

def mm_to_emu(mm):
    return Mm(mm).emu

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
    
    # ---------- 2. ПРОВЕРКА АБЗАЦЕЙ ----------
    in_special_block = False      # внутри содержания или списка литературы
    figure_count = 0
    prev_para_empty = False
    prev_para_was_figure = False
    # Шаблоны для заголовков
    level1_headings = ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]
    # Для разделов 1., 2. и т.д.
    level1_num_pattern = re.compile(r'^\d+\.\s+[А-Я]')
    subsection_pattern = re.compile(r'^\d+\.\d+\s+[А-Яа-я]')
    
    for idx, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            prev_para_empty = True
            continue
        
        p_format = p.paragraph_format
        style_name = p.style.name if p.style else ""
        
        # ------ Пропускаем специальные блоки ------
        if "СОДЕРЖАНИЕ" in text.upper():
            in_special_block = True
        if in_special_block and ("СПИСОК" in text.upper() or "ЗАКЛЮЧЕНИЕ" in text.upper()):
            in_special_block = False
        
        # ------ 2.1 Шрифт и начертание (основной текст) ------
        if not in_special_block:
            for run in p.runs:
                if run.font.name and run.font.name != "Times New Roman":
                    issues.append(f"«{text[:30]}…» – смените шрифт на Times New Roman")
                # Размер 14 пт (для подписей рисунков и таблиц допускается меньше)
                if run.font.size and run.font.size != Pt(14) and not text.startswith(("Рисунок", "Таблица")):
                    issues.append(f"«{text[:30]}…» – установите размер шрифта 14")
                if run.underline:
                    issues.append(f"«{text[:30]}…» – удалите подчеркивания")
        
        # ------ 2.2 Междустрочный интервал 1,2 ------
        spacing = p_format.line_spacing
        if spacing and not in_special_block:
            if p_format.line_spacing_rule == 3:  # MULTIPLE
                if abs(spacing - 1.2) > 0.05:
                    issues.append(f"«{text[:30]}…» – замените интервал на 1,2")
            else:
                issues.append(f"«{text[:30]}…» – установите множитель междустрочного интервала 1,2")
        
        # ------ 2.3 Абзацный отступ 1,0 см (для основного текста) ------
        indent = p_format.first_line_indent
        ignore_indent = (text.isupper() or 
                         text.startswith(("Рисунок", "Таблица")) or 
                         in_special_block or
                         "СОДЕРЖАНИЕ" in text.upper() or
                         level1_num_pattern.match(text) or
                         any(text.startswith(h) for h in level1_headings))
        if indent and not ignore_indent:
            indent_cm = indent.cm
            if abs(indent_cm - 1.0) > 0.1:
                issues.append(f"«{text[:30]}…» – установите абзацный отступ 1,0 см (сейчас {indent_cm:.1f} см)")
        
        # ------ 2.4 СОДЕРЖАНИЕ ------
        if "СОДЕРЖАНИЕ" in text.upper():
            if text.endswith("."):
                issues.append("Содержание – удалите точку в конце")
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
            # Проверка пустой строки после
            next_para = doc.paragraphs[idx+1] if idx+1 < len(doc.paragraphs) else None
            if not (next_para and next_para.text.strip() == ""):
                issues.append("Содержание – после заголовка должна быть пустая строка")
        
        # ------ 2.5 Заголовки разделов (1-го уровня) ------
        is_level1 = (text in level1_headings) or level1_num_pattern.match(text)
        if is_level1:
            # Полужирный
            if not all(run.bold for run in p.runs):
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
            # Разрыв страницы перед разделом (кроме первого раздела после содержания)
            if idx > 0:
                has_page_break = False
                # Ищем разрыв страницы в предыдущем абзаце или в runs текущего
                if idx > 0:
                    prev = doc.paragraphs[idx-1]
                    if prev.runs and any('w:br' in run._element.xml and 'page' in run._element.xml for run in prev.runs):
                        has_page_break = True
                if not has_page_break and p.runs:
                    if any('w:br' in run._element.xml and 'page' in run._element.xml for run in p.runs):
                        has_page_break = True
                if not has_page_break:
                    issues.append(f"«{text[:30]}…» – раздел должен начинаться с новой страницы")
        
        # ------ 2.6 Заголовки подразделов (1.1, 1.2, ...) ------
        if subsection_pattern.match(text):
            # Отступ 1,0 см
            if not p_format.first_line_indent or abs(p_format.first_line_indent.cm - 1.0) > 0.1:
                issues.append(f"«{text[:30]}…» – установите абзацный отступ 1,0 см для подраздела")
            # Полужирный
            if not all(run.bold for run in p.runs):
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
        
        # ------ 2.7 Рисунки ------
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
            # Пустые строки до и после (для всех рисунков)
            prev_para = doc.paragraphs[idx-1] if idx > 0 else None
            next_para = doc.paragraphs[idx+1] if idx+1 < len(doc.paragraphs) else None
            if prev_para and prev_para.text.strip() != "":
                issues.append(f"Рисунок {figure_count} – добавьте пустую строку перед рисунком")
            if next_para and next_para.text.strip() != "":
                issues.append(f"Рисунок {figure_count} – добавьте пустую строку после рисунка")
        
        # ------ 2.8 Таблицы (подписи) – проверяем позже, отдельно ------
        
        prev_para_empty = False
    
    # ---------- 3. ПРОВЕРКА ПОДПИСЕЙ ТАБЛИЦ (по реальным таблицам) ----------
    # Получаем все таблицы и связываем с предыдущим абзацем
    tables = doc.tables
    # Получаем все элементы в теле документа по порядку
    body_elements = list(doc.element.body)
    elements = []
    for elem in body_elements:
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag == 'p':
            elements.append(('paragraph', elem))
        elif tag == 'tbl':
            elements.append(('table', elem))
    
    table_idx = 0
    for i, (typ, elem) in enumerate(elements):
        if typ == 'table':
            table_idx += 1
            # Ищем предыдущий абзац (непустой)
            prev_para = None
            for j in range(i-1, -1, -1):
                if elements[j][0] == 'paragraph':
                    para_elem = elements[j][1]
                    for p in doc.paragraphs:
                        if p._element is para_elem and p.text.strip():
                            prev_para = p
                            break
                    if prev_para:
                        break
            if prev_para and prev_para.text.strip().startswith("Таблица"):
                caption = prev_para.text.strip()
                # Проверка формата: "Таблица X -- Название"
                if not re.match(r'Таблица\s+\d+\s+--\s+', caption):
                    issues.append(f"Таблица {table_idx} – оформление подписи: должно быть «Таблица N -- Название»")
                # Номер таблицы должен соответствовать порядковому
                match = re.search(r'Таблица\s+(\d+)', caption)
                if match:
                    num = int(match.group(1))
                    if num != table_idx:
                        issues.append(f"Таблица {num} – номер не соответствует реальному порядку (должна быть {table_idx})")
                # Точка в конце
                if caption.rstrip().endswith("."):
                    issues.append(f"Таблица {table_idx} – удалите точку в конце названия")
                # Пустые строки до и после подписи
                # Находим индекс этого абзаца в doc.paragraphs
                para_idx = None
                for k, p in enumerate(doc.paragraphs):
                    if p._element is prev_para._element:
                        para_idx = k
                        break
                if para_idx is not None:
                    if para_idx > 0:
                        prev_prev = doc.paragraphs[para_idx-1]
                        if prev_prev.text.strip() != "":
                            issues.append(f"Таблица {table_idx} – добавьте пустую строку перед подписью таблицы")
                    # После подписи идёт таблица, но после таблицы должна быть пустая строка
                    # Найдём следующий после таблицы абзац
                    next_para = None
                    for m in range(i+1, len(elements)):
                        if elements[m][0] == 'paragraph':
                            next_elem = elements[m][1]
                            for p in doc.paragraphs:
                                if p._element is next_elem:
                                    next_para = p
                                    break
                            break
                    if next_para and next_para.text.strip() != "":
                        issues.append(f"Таблица {table_idx} – добавьте пустую строку после таблицы")
            else:
                issues.append(f"Таблица {table_idx} – отсутствует подпись над таблицей")
    
    # ---------- 4. ПРОВЕРКА СОДЕРЖИМОГО ТАБЛИЦ (полужирное) ----------
    for t_idx, table in enumerate(tables, start=1):
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
st.write("Загрузите ваш документ в формате .docx для проверки по полному чек-листу (поля, интервалы, отступы, заголовки, таблицы, рисунки, список литературы).")

uploaded_file = st.file_uploader("Перетащите файл сюда или нажмите для выбора", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем документ..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for res in results:
        st.write(f"• {res}")
