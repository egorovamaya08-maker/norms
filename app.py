import streamlit as st
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # --------------------------------------------------
    # ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
    # --------------------------------------------------
    def is_empty(p):
        return len(p.text.strip()) == 0
    
    def is_bold(p):
        # Проверяем прямое форматирование, затем стиль
        if p.style and p.style.font and p.style.font.bold:
            return True
        runs = [r for r in p.runs if r.text.strip()]
        if not runs:
            return False
        return all(r.bold is True for r in runs)
    
    def is_centered(p):
        # Проверяем выравнивание абзаца и его стиля
        align = p.alignment
        if align == WD_ALIGN_PARAGRAPH.CENTER:
            return True
        if p.style and p.style.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            return True
        return False
    
    def get_indent(pf):
        """Безопасное получение отступа в см"""
        return pf.first_line_indent.cm if pf.first_line_indent is not None else 0.0
    
    # --------------------------------------------------
    # ОПРЕДЕЛЕНИЕ ГРАНИЦ ЗОН ДОКУМЕНТА
    # --------------------------------------------------
    toc_start = None
    toc_end = None
    ref_start = None
    
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if txt.upper() == "СОДЕРЖАНИЕ":
            toc_start = i
            
        # Ищем конец содержания: первый заголовок уровня 1 или список источников
        if toc_start is not None and toc_end is None:
            is_l1 = txt.upper() in level1_keywords or (re.match(r'^\d+\.\s+', txt) and txt == txt.upper())
            if is_l1 or txt.upper().startswith("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"):
                toc_end = i - 1
                
        if txt.upper().startswith("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"):
            ref_start = i

    # Фоллбэки
    if toc_start is not None and toc_end is None:
        toc_end = toc_start
    main_start = (toc_end + 1) if toc_end is not None else 0
    if ref_start is None:
        ref_start = len(doc.paragraphs)

    # --------------------------------------------------
    # ПОЛЯ СТРАНИЦ (глобальная проверка)
    # --------------------------------------------------
    margins_ok = True
    for sec in doc.sections:
        if (abs(sec.left_margin.cm - 2.0) > 0.1 or abs(sec.right_margin.cm - 2.0) > 0.1 or
            abs(sec.top_margin.cm - 2.0) > 0.1 or abs(sec.bottom_margin.cm - 2.0) > 0.1):
            margins_ok = False
            break
    if not margins_ok:
        issues.append("Поля страниц – установите 20 мм со всех сторон")

    # --------------------------------------------------
    # ОСНОВНОЙ ПРОХОД ПО АБЗАЦАМ
    # --------------------------------------------------
    fig_counter = 0
    ref_checked = False
    
    for idx, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if is_empty(p):
            continue
            
        # Пропускаем титульные листы и всё до основного текста
        if idx < main_start:
            continue
            
        pf = p.paragraph_format

        # ---------- ЗАГОЛОВОК СОДЕРЖАНИЯ ----------
        if txt.upper() == "СОДЕРЖАНИЕ":
            if not is_centered(p):
                issues.append("Содержание – выровняйте слово по центру")
            if not is_bold(p):
                issues.append("Содержание – заголовок должен быть полужирным")
            if txt.endswith("."):
                issues.append("Содержание – удалите точку в конце")
            if idx + 1 < len(doc.paragraphs) and not is_empty(doc.paragraphs[idx+1]):
                issues.append("Содержание – после заголовка должна быть пустая строка")
            continue

        # ---------- СПИСОК ИСТОЧНИКОВ ----------
        if idx >= ref_start:
            if idx == ref_start:
                if not is_centered(p): issues.append("Список источников – заголовок по центру")
                if not is_bold(p): issues.append("Список источников – заголовок полужирный")
                continue
            
            if not ref_checked:
                left_ind = pf.left_indent.cm if pf.left_indent else 0.0
                first_ind = get_indent(pf)
                
                if abs(left_ind) > 0.1:
                    issues.append("Список источников – отступ слева должен быть 0 см")
                if abs(first_ind - 1.0) > 0.15:
                    issues.append("Список источников – отступ первой строки 1,0 см")
                if pf.line_spacing_rule == WD_LINE_SPACING.MULTIPLE:
                    if abs(pf.line_spacing - 1.2) > 0.05:
                        issues.append("Список источников – междустрочный интервал должен быть 1,2")
                else:
                    issues.append("Список источников – установите множитель междустрочного интервала 1,2")
                if not (p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY or 
                        (p.style and p.style.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY)):
                    issues.append("Список источников – выровняйте текст по ширине")
                ref_checked = True
            continue

        # ---------- ЗАГОЛОВКИ РАЗДЕЛОВ (УРОВЕНЬ 1) ----------
        is_l1 = (txt.upper() in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ"} or 
                 (re.match(r'^\d+\.\s+[А-ЯЁ]', txt) and txt == txt.upper()))
        if is_l1:
            title = txt[:40]
            # Новая страница (кроме Введения, если сразу после содержания)
            if txt != "ВВЕДЕНИЕ":
                has_break = False
                if idx > 0:
                    for r in doc.paragraphs[idx-1].runs:
                        if 'w:br' in r._element.xml and 'type="page"' in r._element.xml:
                            has_break = True
                            break
                if not has_break:
                    for r in p.runs:
                        if 'w:br' in r._element.xml and 'type="page"' in r._element.xml:
                            has_break = True
                            break
                if not has_break:
                    issues.append(f"«{title}…» – раздел должен начинаться с новой страницы")
            
            if abs(get_indent(pf)) > 0.1:
                issues.append(f"«{title}…» – уберите абзацный отступ")
            if not is_bold(p):
                issues.append(f"«{title}…» – должен быть полужирным")
            if txt != txt.upper():
                issues.append(f"«{title}…» – должен быть прописными буквами")
            if not is_centered(p):
                issues.append(f"«{title}…» – выравнивание по центру")
            if txt.endswith("."):
                issues.append(f"«{title}…» – удалите точку в конце")
            if idx + 1 < len(doc.paragraphs) and not is_empty(doc.paragraphs[idx+1]):
                issues.append(f"«{title}…» – после заголовка должна быть пустая строка")
            continue

        # ---------- ПОДРАЗДЕЛЫ (1.1, 1.2 и т.д.) ----------
        if re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]', txt):
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', txt).strip()
            if abs(get_indent(pf) - 1.0) > 0.15:
                issues.append(f"Подраздел «{sub_name[:30]}…» – установите отступ 1,0 см")
            if not is_bold(p):
                issues.append(f"Подраздел «{sub_name[:30]}…» – должен быть полужирным")
            # Регистр: первая заглавная, остальные строчные
            if sub_name and not (sub_name[0].isupper() and sub_name[1:].lower() == sub_name[1:]):
                issues.append(f"Подраздел «{sub_name[:30]}…» – первая буква прописная, остальные строчные")
            if txt.endswith("."):
                issues.append(f"Подраздел «{sub_name[:30]}…» – удалите точку в конце")
            if idx > 0 and is_empty(doc.paragraphs[idx-1]):
                issues.append(f"Подраздел «{sub_name[:30]}…» – уберите пустую строку перед подразделом")
            continue

        # ---------- РИСУНКИ ----------
        if re.match(r'^[Рр]исунок\s*\d+', txt):
            fig_counter += 1
            if not is_centered(p):
                issues.append(f"Рисунок {fig_counter} – выровняйте подпись по центру")
            if txt.endswith("."):
                issues.append(f"Рисунок {fig_counter} – удалите точку в конце")
            
            m = re.match(r'^[Рр]исунок\s*\d+\s*[–-]\s*(.+)$', txt)
            if m:
                title = m.group(1).strip()
                if title and title[0].islower():
                    issues.append(f"Рисунок {fig_counter} – название должно начинаться с большой буквы")
            
            # Пустые строки до и после (строгое требование для рис. 3, применяем ко всем для единообразия)
            if idx > 0 and not is_empty(doc.paragraphs[idx-1]):
                issues.append(f"Рисунок {fig_counter} – добавьте пустую строку перед рисунком")
            if idx + 1 < len(doc.paragraphs) and not is_empty(doc.paragraphs[idx+1]):
                issues.append(f"Рисунок {fig_counter} – добавьте пустую строку после рисунка")
            continue

        # ---------- ОСНОВНОЙ ТЕКСТ ----------
        if abs(get_indent(pf) - 1.0) > 0.15:
            issues.append(f"«{txt[:30]}…» – установите абзацный отступ 1,0 см")
        if pf.space_before and pf.space_before.pt > 0.5:
            issues.append(f"«{txt[:30]}…» – интервал перед абзацем должен быть 0 пт")

    # --------------------------------------------------
    # ТАБЛИЦЫ (только в основном тексте)
    # --------------------------------------------------
    ref_elem = doc.paragraphs[toc_end]._element if toc_end is not None else None
    main_tables = []
    for table in doc.tables:
        if ref_elem:
            t_parent = table._element.getparent()
            if t_parent is not None and t_parent == ref_elem.getparent():
                children = list(t_parent)
                if children.index(table._element) > children.index(ref_elem):
                    main_tables.append(table)
        else:
            main_tables.append(table)

    for t_idx, table in enumerate(main_tables, start=1):
        t_pos = list(doc.element.body).index(table._element)
        caption = None
        # Ищем подпись, пропуская пустые абзацы
        for i in range(t_pos-1, -1, -1):
            elem = doc.element.body[i]
            if elem.tag.endswith('p'):
                p_obj = next((p for p in doc.paragraphs if p._element is elem), None)
                if p_obj and not is_empty(p_obj):
                    caption = p_obj
                    break
            if elem.tag in ('w:tbl', 'w:sectPr'):
                break # Не ищем за другой таблицей или разрывом раздела

        if caption and re.match(r'^[Тт]аблица\s+\d+', caption.text.strip()):
            cap_txt = caption.text.strip()
            norm = cap_txt.replace('--', '–').replace('-', '–')
            if not re.match(r'[Тт]аблица\s+\d+\s*–\s+', norm):
                issues.append(f"Таблица {t_idx} – оформление: «Таблица N – Название»")
            
            m = re.search(r'[Тт]аблица\s+(\d+)', cap_txt)
            if m and int(m.group(1)) != t_idx:
                issues.append(f"Таблица {t_idx} – номер в подписи не соответствует порядку (должен быть {t_idx})")
            if cap_txt.endswith("."):
                issues.append(f"Таблица {t_idx} – удалите точку в конце названия")
            
            cap_idx = doc.paragraphs.index(caption)
            if cap_idx > 0 and not is_empty(doc.paragraphs[cap_idx-1]):
                issues.append(f"Таблица {t_idx} – добавьте пустую строку перед подписью")
            
            # Пустая строка после таблицы
            after_found = False
            for i in range(t_pos+1, len(doc.element.body)):
                elem = doc.element.body[i]
                if elem.tag.endswith('p'):
                    next_p = next((p for p in doc.paragraphs if p._element is elem), None)
                    if next_p and not is_empty(next_p):
                        issues.append(f"Таблица {t_idx} – добавьте пустую строку после таблицы")
                    after_found = True
                    break
                if elem.tag == 'w:tbl': break
            if not after_found:
                issues.append(f"Таблица {t_idx} – добавьте пустую строку после таблицы")
                
            # Проверка на полужирный внутри ячеек
            has_bold = any(r.bold for row in table.rows for cell in row.cells 
                           for para in cell.paragraphs for r in para.runs if r.text.strip() and r.bold)
            if has_bold:
                issues.append(f"Таблица {t_idx} – уберите полужирное начертание в ячейках")
        else:
            issues.append(f"Таблица {t_idx} – отсутствует подпись над таблицей")

    # --------------------------------------------------
    # ФИНАЛИЗАЦИЯ
    # --------------------------------------------------
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
