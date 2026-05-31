import streamlit as st
import docx
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # 1. ПРОВЕРКА ПОЛЕЙ СТРАНИЦ
    for i, section in enumerate(doc.sections, start=1):
        left_mm = round(section.left_margin.mm) if section.left_margin else 20
        right_mm = round(section.right_margin.mm) if section.right_margin else 20
        top_mm = round(section.top_margin.mm) if section.top_margin else 20
        bottom_mm = round(section.bottom_margin.mm) if section.bottom_margin else 20
        
        if left_mm != 20 or right_mm != 20 or top_mm != 20 or bottom_mm != 20:
            issues.append(f"Раздел {i} - установите поля: левое 20, правое 20, верх 20, низ 20 мм")

    # Вычисляем доступную ширину для картинок (в системных единицах EMU)
    first_sec = doc.sections[0]
    page_width = first_sec.page_width if first_sec.page_width else Cm(21)
    l_margin = first_sec.left_margin if first_sec.left_margin else Cm(2)
    r_margin = first_sec.right_margin if first_sec.right_margin else Cm(2)
    max_usable_width = page_width - l_margin - r_margin

    # Безопасно собираем элементы в единый список, используя параллельные счётчики
    body_elements = []
    p_idx = 0
    t_idx = 0
    for child in doc.element.body:
        if child.tag.endswith('p'):
            if p_idx < len(doc.paragraphs):
                body_elements.append(('p', doc.paragraphs[p_idx]))
                p_idx += 1
        elif child.tag.endswith('tbl'):
            if t_idx < len(doc.tables):
                body_elements.append(('tbl', doc.tables[t_idx]))
                t_idx += 1

    figure_count = 0
    current_table_id = None
    in_bibliography = False

    # 2. ЛИНЕЙНЫЙ АНАЛИЗ ДОКУМЕНТА
    for idx, (typ, elem) in enumerate(body_elements):
        if typ == 'p':
            text = elem.text.strip()
            if not text:
                continue
            
            p_format = elem.paragraph_format

            # Проверяем наличие картинок/рисунков внутри абзаца
            drawings = elem._p.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}extent')
            
            # Маркеры крупных разделов
            is_main_heading = False
            heading_label = ""
            
            if "ВВЕДЕНИЕ" in text.upper() and len(text) < 15:
                is_main_heading = True
                heading_label = "Введение"
            elif "ЗАКЛЮЧЕНИЕ" in text.upper() and len(text) < 15:
                is_main_heading = True
                heading_label = "Заключение"
            elif "СПИСОК" in text.upper() and ("ИСТОЧНИК" in text.upper() or "ЛИТЕРАТУР" in text.upper()):
                is_main_heading = True
                heading_label = "Список литературы"
                in_bibliography = True
            elif re.match(r'^Раздел\s+\d+', text, re.IGNORECASE) or re.match(r'^\d+\.\s+[А-Я]', text):
                is_main_heading = True
                match_num = re.search(r'\d+', text)
                heading_label = f"Раздел {match_num.group()}" if match_num else "Раздел"

            # Сброс флага списка литературы при переходе к другому разделу
            if is_main_heading and "СПИСОК" not in text.upper():
                in_bibliography = False

            # Проверка крупных разделов (Введение, Разделы, Заключение)
            if is_main_heading:
                has_page_break = False
                has_text_before = any(body_elements[i][0] == 'p' and body_elements[i][1].text.strip() for i in range(idx))
                
                if has_text_before:
                    if p_format.page_break_before:
                        has_page_break = True
                    else:
                        # Проверяем ручные разрывы страниц
                        if any(b.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page' 
                               for r in elem.runs for b in r._r.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br')):
                            has_page_break = True
                        elif idx > 0 and body_elements[idx-1][0] == 'p':
                            prev_p = body_elements[idx-1][1]
                            if any(b.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page' 
                                   for r in prev_p.runs for b in r._r.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br')):
                                has_page_break = True
                    
                    if not has_page_break:
                        issues.append(f"{heading_label} - начните с новой страницы")

                if p_format.first_line_indent and p_format.first_line_indent.cm > 0:
                    issues.append(f"{heading_label} - удалите абзацный отступ")

            # Проверка подразделов (1.1, 1.2, 2.1 и т.д.)
            elif re.match(r'^\d+\.\d+', text):
                sub_num = re.match(r'^\d+\.\d+', text).group()
                sub_label = f"Подраздел {sub_num}"
                
                if idx > 0 and body_elements[idx-1][0] == 'p' and not body_elements[idx-1][1].text.strip():
                    issues.append(f"{sub_label} - уберите пустую строку перед подразделом")

            # Проверка подписей к Рисункам
            elif text.startswith("Рисунок"):
                figure_count += 1
                fig_label = f"Рисунок {figure_count}"
                
                if text.endswith("."):
                    issues.append(f"{fig_label} - удалите точку в конце названия")
                if elem.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                    issues.append(f"{fig_label} - выровняйте подпись по центру")
                
                parts = text.split(maxsplit=2)
                if len(parts) >= 3 and parts[2] and parts[2][0].islower():
                    issues.append(f"{fig_label} - название должно начинаться с большой буквы")

                # Анализ пустых строк вокруг рисунка
                has_prev_empty = False
                if idx > 0 and body_elements[idx-1][0] == 'p' and not body_elements[idx-1][1].text.strip():
                    has_prev_empty = True
                if idx > 1 and body_elements[idx-1][0] == 'p' and body_elements[idx-1][1]._p.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}extent'):
                    if idx > 2 and body_elements[idx-2][0] == 'p' and not body_elements[idx-2][1].text.strip():
                        has_prev_empty = True
                    else:
                        has_prev_empty = False
                        
                if not has_prev_empty:
                    issues.append(f"{fig_label} - добавьте пустую строку перед рисунком")

                has_next_empty = idx + 1 < len(body_elements) and body_elements[idx+1][0] == 'p' and not body_elements[idx+1][1].text.strip()
                if not has_next_empty:
                    issues.append(f"{fig_label} - добавьте пустую строку после рисунка")

            # Анализ ширины графики
            if drawings:
                for ext in drawings:
                    cx = int(ext.attrib.get('cx', 0))
                    if cx > max_usable_width:
                        issues.append(f"Рисунок {figure_count + 1} - уменьшите размер, чтобы он не выходил за границы полей")

            # Анализ заголовков Таблиц
            if text.startswith("Таблица") and not any(w in text.lower() for w in [" в ", " из ", "на "]):
                t_match = re.search(r'\d+', text)
                if t_match:
                    current_table_id = t_match.group()
                    if text.endswith("."):
                        issues.append(f"Таблица {current_table_id} - удалите точку в конце названия")

            # Проверка слова СОДЕРЖАНИЕ
            if "СОДЕРЖАНИЕ" in text.upper() and len(text) < 15:
                if text.endswith("."):
                    issues.append("Содержание - удалите точку в конце")
                if elem.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                    issues.append("Содержание - выровняйте слово СОДЕРЖАНИЕ по центру")

            # Проверка элементов Списка литературы
            if in_bibliography and not is_main_heading:
                indent = p_format.first_line_indent
                if not indent or round(indent.cm, 1) != 1.0:
                    issues.append("Список литературы - установите отступ первой строки 1,0 см")
                if p_format.left_indent and p_format.left_indent.cm != 0:
                    issues.append("Список литературы - отступ слева должен быть 0 см")
                if elem.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                    issues.append("Список литературы - выровняйте по ширине")

        # Проверка свойств самой таблицы (только если она идет сразу за заголовком)
        elif typ == 'tbl':
            if current_table_id:
                bold_inside = False
                for row in elem.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                if run.bold:
                                    bold_inside = True
                                    break
                if bold_inside:
                    issues.append(f"Таблица {current_table_id} - уберите полужирное начертание")
                current_table_id = None 

    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

# ИНТЕРФЕЙС СТРИМЛИТ
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write("Загрузите ваш документ в формате .docx для мгновенного нормоконтроля по чек-листу.")

uploaded_file = st.file_uploader("Перетащите файл сюда или нажмите для выбора", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем структуру, интервалы и разметку..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for res in results:
        st.write(res)
