import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # ---------- 1. ПОЛЯ СТРАНИЦ ----------
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
    
    # ---------- 2. ГРАНИЦЫ СОДЕРЖАНИЯ ----------
    content_start = None
    content_end = None
    for i, p in enumerate(doc.paragraphs):
        if "СОДЕРЖАНИЕ" in p.text.upper():
            content_start = i
            for j in range(i+1, len(doc.paragraphs)):
                txt = doc.paragraphs[j].text.strip().upper()
                if txt == "ВВЕДЕНИЕ" or re.match(r'^\d+\.\s+[А-Я]{2,}', txt) or txt in ["ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]:
                    content_end = j - 1
                    break
            if content_end is None:
                content_end = i + 10
            break
    
    # ---------- 3. ПРОВЕРКА СОДЕРЖАНИЯ ----------
    if content_start is not None:
        p = doc.paragraphs[content_start]
        text = p.text.strip()
        if text.endswith("."):
            issues.append("Содержание – удалите точку в конце")
        if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
            issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
    
    # ---------- 4. ПРОВЕРКА ПОСЛЕ СОДЕРЖАНИЯ ----------
    figure_count = 0
    prev_para_empty = False
    
    level1_words = ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]
    level1_num = re.compile(r'^\d+\.\s+[А-Я]{2,}')
    subsection = re.compile(r'^\d+\.\d+\s+[А-Яа-я]')
    
    for idx, p in enumerate(doc.paragraphs):
        if content_start is not None and idx <= content_end:
            continue
        
        text = p.text.strip()
        if not text:
            prev_para_empty = True
            continue
        
        p_format = p.paragraph_format
        
        # ----- Заголовки разделов -----
        is_level1 = (text in level1_words) or level1_num.match(text)
        if is_level1:
            if text != "ВВЕДЕНИЕ":
                page_break = False
                if idx > 0:
                    prev = doc.paragraphs[idx-1]
                    if prev.runs:
                        for run in prev.runs:
                            if 'w:br' in run._element.xml and 'page' in run._element.xml:
                                page_break = True
                if not page_break and p.runs:
                    for run in p.runs:
                        if 'w:br' in run._element.xml and 'page' in run._element.xml:
                            page_break = True
                if not page_break:
                    issues.append(f"«{text[:40]}…» – раздел должен начинаться с новой страницы")
            if p_format.first_line_indent and abs(p_format.first_line_indent.cm) > 0.1:
                issues.append(f"«{text[:40]}…» – уберите абзацный отступ у заголовка")
            bold_ok = (p.style and p.style.font.bold) or all(r.bold for r in p.runs)
            if not bold_ok:
                issues.append(f"«{text[:40]}…» – заголовок раздела должен быть полужирным")
            if text != text.upper():
                issues.append(f"«{text[:40]}…» – заголовок раздела должен быть прописными буквами")
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"«{text[:40]}…» – выровняйте заголовок по центру")
            if text.endswith("."):
                issues.append(f"«{text[:40]}…» – удалите точку в конце заголовка")
            next_para = doc.paragraphs[idx+1] if idx+1 < len(doc.paragraphs) else None
            if not (next_para and next_para.text.strip() == ""):
                issues.append(f"«{text[:40]}…» – после заголовка должна быть пустая строка")
        
        # ----- Подразделы -----
        if subsection.match(text):
            if not p_format.first_line_indent:
                issues.append(f"«{text[:40]}…» – установите абзацный отступ 1,0 см для подраздела")
            else:
                indent_cm = p_format.first_line_indent.cm
                if abs(indent_cm - 1.0) > 0.2:
                    issues.append(f"«{text[:40]}…» – установите абзацный отступ 1,0 см для подраздела (сейчас {indent_cm:.1f} см)")
            bold_ok = (p.style and p.style.font.bold) or all(r.bold for r in p.runs)
            if not bold_ok:
                issues.append(f"«{text[:40]}…» – заголовок подраздела должен быть полужирным")
            parts = text.split(maxsplit=1)
            if len(parts) > 1:
                title = parts[1]
                if title and (title[0].islower() or any(c.isupper() for c in title[1:])):
                    issues.append(f"«{text[:40]}…» – заголовок подраздела: первая буква прописная, остальные строчные")
            if text.endswith("."):
                issues.append(f"«{text[:40]}…» – удалите точку в конце заголовка подраздела")
            if prev_para_empty:
                issues.append(f"«{text[:40]}…» – уберите пустую строку перед подразделом")
        
        # ----- Рисунки -----
        if text.startswith("Рисунок"):
            figure_count += 1
            parts = text.split(maxsplit=2)
            if len(parts) >= 3:
                first_word = parts[2]
                if first_word and first_word[0].islower():
                    issues.append(f"Рисунок {figure_count} – название должно начинаться с большой буквы")
            if figure_count == 3:
                prev_para = doc.paragraphs[idx-1] if idx > 0 else None
                next_para = doc.paragraphs[idx+1] if idx+1 < len(doc.paragraphs) else None
                if prev_para and prev_para.text.strip() != "":
                    issues.append("Рисунок 3 – добавьте пустую строку перед рисунком")
                if next_para and next_para.text.strip() != "":
                    issues.append("Рисунок 3 – добавьте пустую строку после рисунка")
        
        prev_para_empty = False
    
    # ---------- 5. ТАБЛИЦЫ ПОСЛЕ СОДЕРЖАНИЯ ----------
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
            if not re.match(r'Таблица\s+\d+\s+--\s+', caption):
                issues.append(f"Таблица {t_idx} – оформление подписи: должно быть «Таблица N -- Название»")
            m = re.search(r'Таблица\s+(\d+)', caption)
            if m and int(m.group(1)) != t_idx:
                issues.append(f"Таблица {t_idx} – номер в подписи не соответствует порядковому (должен быть {t_idx})")
            if caption.rstrip().endswith("."):
                issues.append(f"Таблица {t_idx} – удалите точку в конце названия")
            cap_idx = None
            for i, p in enumerate(doc.paragraphs):
                if p._element is caption_para._element:
                    cap_idx = i
                    break
            if cap_idx is not None and cap_idx > 0:
                if doc.paragraphs[cap_idx-1].text.strip() != "":
                    issues.append(f"Таблица {t_idx} – добавьте пустую строку перед подписью таблицы")
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
            if p.paragraph_format.left_indent and p.paragraph_format.left_indent.cm != 0:
                issues.append("Список источников – отступ слева должен быть 0 см")
            if not p.paragraph_format.first_line_indent or abs(p.paragraph_format.first_line_indent.cm - 1.0) > 0.1:
                issues.append("Список источников – установите отступ первой строки 1,0 см")
            if p.paragraph_format.line_spacing_rule == 3:
                if abs(p.paragraph_format.line_spacing - 1.2) > 0.05:
                    issues.append("Список источников – междустрочный интервал должен быть 1,2")
            else:
                issues.append("Список источников – установите множитель междустрочного интервала 1,2")
            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Список источников – выровняйте по ширине")
    
    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

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
