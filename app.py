import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # ---------- 1. ПОЛЯ СТРАНИЦ (одно сообщение) ----------
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
    
    # ---------- 2. НАХОДИМ ГРАНИЦЫ ОСНОВНОГО ТЕКСТА ----------
    # Индекс абзаца "СОДЕРЖАНИЕ"
    content_idx = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().upper() == "СОДЕРЖАНИЕ":
            content_idx = i
            break
    
    # Индекс первого заголовка раздела после содержания
    start_idx = None
    if content_idx is not None:
        for i in range(content_idx + 1, len(doc.paragraphs)):
            txt = doc.paragraphs[i].text.strip()
            if txt.upper() == "ВВЕДЕНИЕ" or re.match(r'^\d+\.\s+[А-Я]{2,}', txt) or txt.upper() in ["ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]:
                start_idx = i
                break
    else:
        # Если нет содержания, начинаем с первого абзаца
        start_idx = 0
    
    if start_idx is None:
        start_idx = 0
    
    # ---------- 3. ПРОВЕРКА ЗАГОЛОВКА "СОДЕРЖАНИЕ" ----------
    if content_idx is not None:
        p = doc.paragraphs[content_idx]
        text = p.text.strip()
        if text.endswith("."):
            issues.append("Содержание – удалите точку в конце")
        if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
            issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
        if not all(r.bold for r in p.runs if r.text.strip()):
            issues.append("Содержание – сделайте заголовок полужирным")
        if content_idx + 1 < len(doc.paragraphs) and doc.paragraphs[content_idx+1].text.strip() != "":
            issues.append("Содержание – после заголовка должна быть пустая строка")
    
    # ---------- 4. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ----------
    def is_empty(p):
        return len(p.text.strip()) == 0
    
    def is_bold(p):
        try:
            if p.style and p.style.font and p.style.font.bold:
                return True
        except:
            pass
        runs = [r for r in p.runs if r.text.strip()]
        if not runs:
            return False
        return all(r.bold for r in runs)
    
    # ---------- 5. ОСНОВНАЯ ПРОВЕРКА (начиная с start_idx) ----------
    figure_counter = 0
    prev_para_empty = False
    
    level1_headings = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    level1_num = re.compile(r'^\d+\.\s+[А-Я]{2,}')
    subsection = re.compile(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]')
    
    for idx, p in enumerate(doc.paragraphs):
        if idx < start_idx:
            continue
        text = p.text.strip()
        if not text:
            prev_para_empty = True
            continue
        
        pf = p.paragraph_format
        
        # ----- 5.1 ЗАГОЛОВКИ РАЗДЕЛОВ -----
        is_l1 = (text in level1_headings) or level1_num.match(text)
        if is_l1:
            # Новая страница (кроме ВВЕДЕНИЯ, если оно идёт сразу после содержания)
            # Для ВВЕДЕНИЯ проверяем только если оно не первое после содержания? По требованию: раздел 1 должен начинаться с новой страницы, а ВВЕДЕНИЕ может идти на той же странице, что и содержание? Не уверен. Лучше требовать для всех, кроме ВВЕДЕНИЯ, если оно идёт сразу после содержания. Но для простоты: требуем новую страницу для всех разделов, кроме ВВЕДЕНИЯ.
            if text != "ВВЕДЕНИЕ":
                page_break = False
                if idx > 0:
                    prev_para = doc.paragraphs[idx-1]
                    for run in prev_para.runs:
                        if 'w:br' in run._element.xml and 'page' in run._element.xml:
                            page_break = True
                if not page_break and p.runs:
                    for run in p.runs:
                        if 'w:br' in run._element.xml and 'page' in run._element.xml:
                            page_break = True
                if not page_break:
                    issues.append(f"«{text[:40]}…» – раздел должен начинаться с новой страницы")
            # Абзацный отступ 0
            if pf.first_line_indent and abs(pf.first_line_indent.cm) > 0.1:
                issues.append(f"«{text[:40]}…» – уберите абзацный отступ у заголовка")
            # Полужирный
            if not is_bold(p):
                issues.append(f"«{text[:40]}…» – заголовок раздела должен быть полужирным")
            # Прописные буквы
            if text != text.upper():
                issues.append(f"«{text[:40]}…» – заголовок раздела должен быть прописными буквами")
            # Выравнивание по центру
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:40]}…» – выровняйте заголовок по центру")
            # Без точки в конце
            if text.endswith("."):
                issues.append(f"«{text[:40]}…» – удалите точку в конце заголовка")
            # Пустая строка после
            if idx+1 < len(doc.paragraphs) and not is_empty(doc.paragraphs[idx+1]):
                issues.append(f"«{text[:40]}…» – после заголовка должна быть пустая строка")
        
        # ----- 5.2 ПОДРАЗДЕЛЫ (только в основном тексте, так как содержание пропущено) -----
        elif subsection.match(text):
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()
            # Отступ 1,0 см
            if not pf.first_line_indent:
                issues.append(f"Подраздел «{sub_name[:40]}…» – установите абзацный отступ 1,0 см")
            else:
                if abs(pf.first_line_indent.cm - 1.0) > 0.2:
                    issues.append(f"Подраздел «{sub_name[:40]}…» – установите абзацный отступ 1,0 см (сейчас {pf.first_line_indent.cm:.1f})")
            # Полужирный
            if not is_bold(p):
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
        
        # ----- 5.3 ОБЫЧНЫЙ ТЕКСТ (не заголовки и не рисунки) -----
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
        
        # ----- 5.4 РИСУНКИ -----
        if text.startswith("Рисунок"):
            figure_counter += 1
            # Выравнивание по центру
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
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
            # Рисунок 3: пустые строки до и после
            if figure_counter == 3:
                if idx > 0 and not is_empty(doc.paragraphs[idx-1]):
                    issues.append("Рисунок 3 – добавьте пустую строку перед рисунком")
                if idx+1 < len(doc.paragraphs) and not is_empty(doc.paragraphs[idx+1]):
                    issues.append("Рисунок 3 – добавьте пустую строку после рисунка")
        
        prev_para_empty = False
    
    # ---------- 6. ТАБЛИЦЫ (только те, что после start_idx) ----------
    # Находим позицию в body первого абзаца основного текста
    try:
        start_elem = doc.paragraphs[start_idx]._element
        start_body_pos = list(doc.element.body).index(start_elem)
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
            if cap_idx is not None and cap_idx > 0 and not is_empty(doc.paragraphs[cap_idx-1]):
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
            if next_para and not is_empty(next_para):
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
    
    # ---------- 7. СПИСОК ИСТОЧНИКОВ ----------
    lit_start = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            lit_start = i
            break
    if lit_start is not None:
        # Ищем первый непустой абзац после заголовка
        first_source = None
        for i in range(lit_start+1, len(doc.paragraphs)):
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
            # Междустрочный интервал 1,2
            if pf.line_spacing_rule == 3:
                if abs(pf.line_spacing - 1.2) > 0.05:
                    issues.append("Список источников – междустрочный интервал должен быть 1,2")
            else:
                issues.append("Список источников – установите множитель междустрочного интервала 1,2")
            # Выравнивание по ширине
            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – выровняйте по ширине")
    
    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

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
