import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import re
from collections import defaultdict

def get_effective_alignment(paragraph):
    if paragraph.alignment is not None:
        return paragraph.alignment
    try:
        style = paragraph.style
        if style and style.paragraph_format.alignment is not None:
            return style.paragraph_format.alignment
    except:
        pass
    return None

def is_paragraph_bold(paragraph):
    runs = [r for r in paragraph.runs if r.text.strip()]
    if runs:
        if any(r.bold for r in runs):
            return True
    
    try:
        if paragraph.style and paragraph.style.font and paragraph.style.font.bold:
            return True
    except:
        pass
    
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            pPr_rPr = pPr.find(qn('w:rPr'))
            if pPr_rPr is not None:
                bold_elem = pPr_rPr.find(qn('w:b'))
                if bold_elem is not None:
                    val = bold_elem.get(qn('w:val'))
                    if val != 'false' and val != '0':
                        return True
            
            for r in paragraph._element.findall(qn('w:r')):
                rPr = r.find(qn('w:rPr'))
                if rPr is not None:
                    bold_elem = rPr.find(qn('w:b'))
                    if bold_elem is not None:
                        val = bold_elem.get(qn('w:val'))
                        if val != 'false' and val != '0':
                            return True
    except:
        pass
    
    return False

def is_empty_paragraph(paragraph):
    return len(paragraph.text.strip()) == 0

def has_page_number(text):
    return bool(re.search(r'[\t\s\.]{2,}\d+$', text))

def is_all_caps(text):
    clean_text = re.sub(r'[\d\s\.,;:!?\-–—()«»""''«»]', '', text)
    if not clean_text:
        return False
    return clean_text == clean_text.upper()

def is_section_header(text):
    if re.match(r'^\d+\.\s+[А-ЯЁ]', text) and is_all_caps(text):
        return True
    if re.match(r'^(?:ГЛАВА|РАЗДЕЛ)\s+\d+', text, re.IGNORECASE):
        return True
    if text.upper().strip() in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}:
        return True
    return False

def normalize_title(text):
    text = re.sub(r'^(?:ГЛАВА|РАЗДЕЛ)\s+\d+[\.\s]*', '', text, flags=re.IGNORECASE)
    text = re.sub(r'^\d+(?:\.\d+)*[\s\.]+', '', text)
    return text.strip().upper()

def extract_toc_entries(doc, start_idx):
    toc_entries = []
    for i in range(start_idx):
        txt = doc.paragraphs[i].text.strip()
        if not txt:
            continue
        if has_page_number(txt):
            clean = re.sub(r'[\t\s\.]{2,}\d+$', '', txt).strip()
            if clean and len(clean) > 5:
                toc_entries.append(clean)
    return toc_entries

def is_dash_char(ch):
    """Проверяет, является ли символ тире (включая приватные области Unicode)"""
    code = ord(ch)
    if code in [0x2D, 0x2010, 0x2011, 0x2012, 0x2013, 0x2014, 0x2015]:
        return True
    if 0xE000 <= code <= 0xF8FF:
        return True
    return False

def is_bullet_char(ch):
    """Проверяет, является ли символ круглым маркером"""
    code = ord(ch)
    return code in [0x2022, 0x2023, 0x25E6, 0x25CF, 0x25CB, 0x26AB, 0x2B24]

def get_list_marker_info(paragraph, doc):
    text = paragraph.text.strip()
    if not text:
        return False, "", True
    
    first_char = text[0] if text else ""
    
    if is_dash_char(first_char):
        return True, "тире", True
    
    if re.match(r'^\d+\)\s', text):
        return True, "нумерованный", True
    
    if re.match(r'^[а-яё]\)\s', text):
        return True, "буквенный", True
    if re.match(r'^[a-z]\)\s', text):
        return True, "буквенный", True
    
    if is_bullet_char(first_char):
        return True, "круглый маркер (•)", False
    
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                numId_elem = numPr.find(qn('w:numId'))
                if numId_elem is not None:
                    numId = numId_elem.get(qn('w:val'))
                    numbering_part = doc.part.numbering_part
                    if numbering_part is not None:
                        numbering_xml = numbering_part._element
                        abstractNumId = None
                        for num in numbering_xml.findall(qn('w:num')):
                            if num.get(qn('w:numId')) == numId:
                                aid_elem = num.find(qn('w:abstractNumId'))
                                if aid_elem is not None:
                                    abstractNumId = aid_elem.get(qn('w:val'))
                                break
                        if abstractNumId:
                            for abstractNum in numbering_xml.findall(qn('w:abstractNum')):
                                if abstractNum.get(qn('w:abstractNumId')) == abstractNumId:
                                    for lvl in abstractNum.findall(qn('w:lvl')):
                                        numFmt = lvl.find(qn('w:numFmt'))
                                        lvlText_elem = lvl.find(qn('w:lvlText'))
                                        
                                        if numFmt is not None:
                                            fmt = numFmt.get(qn('w:val'))
                                            if fmt == 'bullet':
                                                if lvlText_elem is not None:
                                                    txt_val = lvlText_elem.get(qn('w:val'))
                                                    if txt_val:
                                                        clean = re.sub(r'%\d+', '', txt_val).strip()
                                                        if clean and is_dash_char(clean[0]):
                                                            return True, "тире", True
                                                return True, "круглый маркер (•)", False
                                            elif fmt == 'decimal':
                                                return True, "нумерованный (цифры)", True
                                            elif fmt in ['lowerLetter', 'upperLetter']:
                                                return True, "буквенный", True
                                            else:
                                                return True, f"формат '{fmt}'", True
    except:
        pass
    
    return False, "", True

def get_effective_first_line_indent(paragraph):
    pf = paragraph.paragraph_format
    if pf.first_line_indent is not None:
        return pf.first_line_indent.cm
    try:
        style = paragraph.style
        if style and style.paragraph_format.first_line_indent is not None:
            return style.paragraph_format.first_line_indent.cm
    except:
        pass
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                first_line = ind.get(qn('w:firstLine'))
                hanging = ind.get(qn('w:hanging'))
                if first_line is not None:
                    return int(first_line) / 567
                elif hanging is not None:
                    return 0
    except:
        pass
    return 0

def get_effective_left_indent(paragraph):
    pf = paragraph.paragraph_format
    if pf.left_indent is not None:
        return pf.left_indent.cm
    try:
        style = paragraph.style
        if style and style.paragraph_format.left_indent is not None:
            return style.paragraph_format.left_indent.cm
    except:
        pass
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                left = ind.get(qn('w:left'))
                if left is not None:
                    return int(left) / 567
    except:
        pass
    return 0

def is_table_continuation(text):
    return bool(re.match(r'^(?:Продолжение|Окончание)\s+таблицы?\s*\d', text))

def get_table_depth(table):
    depth = 0
    element = table._element
    parent = element.getparent()
    while parent is not None:
        if parent.tag == qn('w:tbl'):
            depth += 1
        parent = parent.getparent()
    return depth

def find_nearest_caption(doc, tbl_pos, start_body_pos):
    for i in range(tbl_pos - 1, start_body_pos - 1, -1):
        elem = doc.element.body[i]
        if elem.tag.endswith('p'):
            for para in doc.paragraphs:
                if para._element is elem and para.text.strip():
                    return para.text.strip(), i
    return None, None

def is_formula_or_equation(text):
    if re.search(r'[=≠≤≥±×÷∫∑∏√∞∂∇∈∉⊂⊃∪∩]', text):
        return True
    if re.search(r'[α-ωΑ-Ω]', text):
        return True
    if re.search(r'[₀-₉ₐ-ₜₓᵦ-ᵧ⁰-⁹ⁱ⁻⁺ⁿ]', text):
        return True
    return False

def is_formula_where_line(text):
    return bool(re.match(r'^[Гг]де\s*[:\s]?\s*[А-Яа-яA-Za-z\-–—]', text))

def check_formula_explanation(text, paragraph, prev_was_formula, prev_para_empty):
    errors = []
    
    short_text = text[:60] + "…" if len(text) > 60 else text
    key = f"Пояснение к формуле «{short_text}»"
    
    if text.startswith("Где"):
        errors.append(f"{key} – «где» должно быть с маленькой буквы")
    
    if re.match(r'^[Гг]де\s*:', text):
        errors.append(f"{key} – уберите двоеточие после «где»")
    
    first_line = get_effective_first_line_indent(paragraph)
    if abs(first_line - 1.0) > 0.2:
        errors.append(f"{key} – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
    
    if prev_para_empty and prev_was_formula:
        errors.append(f"{key} – уберите пустую строку перед пояснением (должно идти сразу после формулы)")
    
    explanation_text = re.sub(r'^[Гг]де\s*:?\s*', '', text).strip()
    
    lines = [l.strip() for l in explanation_text.split('\t') if l.strip()]
    if len(lines) > 1:
        for i, line in enumerate(lines):
            if i < len(lines) - 1:
                if not line.endswith(';'):
                    var_match = re.match(r'^([^\s–—]+)\s*[–—]\s*(.+)$', line)
                    var_name = var_match.group(1) if var_match else line[:20]
                    errors.append(f"{key} – после «{var_name}» должна быть точка с запятой (;)")
            else:
                if not line.endswith('.'):
                    var_match = re.match(r'^([^\s–—]+)\s*[–—]\s*(.+)$', line)
                    var_name = var_match.group(1) if var_match else line[:20]
                    errors.append(f"{key} – после «{var_name}» должна быть точка (.)")
    else:
        parts = re.findall(r'([^\s–—]+)\s*[–—]\s*([^;.]+)([;.]?)', explanation_text)
        
        if len(parts) > 1:
            for i, (var, desc, punct) in enumerate(parts):
                if i < len(parts) - 1:
                    if punct != ';':
                        errors.append(f"{key} – после «{var}» должна быть точка с запятой (;)")
                else:
                    if punct != '.':
                        errors.append(f"{key} – после «{var}» должна быть точка (.)")
        elif len(parts) == 1:
            var, desc, punct = parts[0]
            if punct != '.':
                errors.append(f"{key} – после «{var}» должна быть точка (.)")
    
    return errors

def group_issues(issues_list):
    """Группирует ошибки по ключу"""
    auto_issues = []
    manual_issues = []
    manual_section = False
    
    for issue in issues_list:
        if issue.startswith("📋 Для проверки человеком"):
            manual_section = True
            continue
        if manual_section:
            manual_issues.append(issue)
        else:
            auto_issues.append(issue)
    
    grouped = defaultdict(list)
    standalone = []
    
    for issue in auto_issues:
        match = re.match(r'^(?:«([^»]+)»|(Рисунок\s+[\d.]+)|(Таблица\s+[\d.]+)|(Подраздел\s+«([^»]+)»)|(Пояснение к формуле «([^»]+)»)|(Список начиная с «([^»]+)»))\s*[–-]\s*(.+)$', issue)
        
        if match:
            key = None
            if match.group(1):
                key = match.group(1)[:80]
            elif match.group(2):
                key = match.group(2)
            elif match.group(3):
                key = match.group(3)
            elif match.group(5):
                key = f"Подраздел «{match.group(5)[:50]}»"
            elif match.group(6):
                key = f"Пояснение к формуле «{match.group(7)[:60]}»"
            elif match.group(8):
                key = f"Список начиная с «{match.group(9)[:50]}»"
            
            if key:
                message = match.group(10)
                grouped[key].append(message)
            else:
                standalone.append(issue)
        else:
            standalone.append(issue)
    
    result = []
    
    for issue in standalone:
        result.append(issue)
    
    for key, messages in grouped.items():
        if len(messages) == 1:
            if key.startswith("Подраздел «") or key.startswith("Пояснение к формуле «") or key.startswith("Список начиная с «"):
                result.append(f"{key} – {messages[0]}")
            elif key.startswith("Рисунок ") or key.startswith("Таблица "):
                result.append(f"{key} – {messages[0]}")
            else:
                result.append(f"«{key}» – {messages[0]}")
        else:
            combined = "; ".join(messages)
            if key.startswith("Подраздел «") or key.startswith("Пояснение к формуле «") or key.startswith("Список начиная с «"):
                result.append(f"{key} – {combined}")
            elif key.startswith("Рисунок ") or key.startswith("Таблица "):
                result.append(f"{key} – {combined}")
            else:
                result.append(f"«{key}» – {combined}")
    
    if manual_issues:
        result.append("\n📋 Для проверки человеком:")
        result.extend(manual_issues)
    
    return result

def check_word_document(file):
    doc = docx.Document(file)
    auto_issues = []
    manual_checks = []

    # ---------- 1. ПОЛЯ СТРАНИЦ ----------
    margins_ok = True
    for section in doc.sections:
        if (abs(section.left_margin.mm - 20) > 0.5 or
            abs(section.right_margin.mm - 20) > 0.5 or
            abs(section.top_margin.mm - 20) > 0.5 or
            abs(section.bottom_margin.mm - 20) > 0.5):
            margins_ok = False
            break
    if not margins_ok:
        auto_issues.append("Поля страниц – установите левое 20 мм, правое 20 мм, верхнее 20 мм, нижнее 20 мм")

    # ---------- 2. ПОИСК НАЧАЛА ОСНОВНОГО ТЕКСТА ----------
    start_idx = None
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt:
            continue
        if has_page_number(txt):
            continue
        if re.match(r'^\d+\.\s+[А-Яа-я]', txt) and not is_all_caps(txt):
            continue
        if is_section_header(txt):
            start_idx = i
            break
    
    if start_idx is None:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]

    toc_entries = extract_toc_entries(doc, start_idx)

    # ---------- 3. ГРАНИЦЫ СПИСКА ИСТОЧНИКОВ ----------
    lit_start = None
    for i in range(start_idx, len(doc.paragraphs)):
        txt = doc.paragraphs[i].text.strip()
        if txt.upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" and not has_page_number(txt):
            lit_start = i
            break
    
    lit_end = len(doc.paragraphs)
    if lit_start is not None:
        appendix_keywords = ["ПРИЛОЖЕНИЕ", "ПРИЛОЖЕНИЯ", "APPENDIX"]
        for i in range(lit_start + 1, len(doc.paragraphs)):
            txt = doc.paragraphs[i].text.strip().upper()
            if any(txt.startswith(kw) for kw in appendix_keywords):
                lit_end = i
                break

    # ---------- 4. ПРОВЕРКА ОСНОВНОГО ТЕКСТА ----------
    figure_counter = 0
    prev_para_empty = False
    prev_was_formula = False
    end_idx = lit_start if lit_start is not None else len(doc.paragraphs)
    
    list_errors = []
    
    for idx in range(start_idx, end_idx):
        p = doc.paragraphs[idx]
        text = p.text.strip()
        
        if not text:
            prev_para_empty = True
            if list_errors:
                first_text = list_errors[0][1]
                auto_issues.append(f"Список начиная с «{first_text[:50]}» и далее – замените круглый маркер (•) на тире, букву или цифру")
                list_errors = []
            continue
        if has_page_number(text):
            prev_para_empty = False
            continue
        if re.match(r'^\d+\.\s+[А-Яа-я]', text) and not is_all_caps(text):
            prev_para_empty = False
            continue
        
        is_list, marker_type, marker_valid = get_list_marker_info(p, doc)
        if is_list:
            if not marker_valid:
                if list_errors and idx == list_errors[-1][0] + 1:
                    list_errors.append((idx, text, marker_type))
                else:
                    if list_errors:
                        first_text = list_errors[0][1]
                        auto_issues.append(f"Список начиная с «{first_text[:50]}» и далее – замените круглый маркер (•) на тире, букву или цифру")
                        list_errors = []
                    list_errors.append((idx, text, marker_type))
            else:
                if list_errors:
                    first_text = list_errors[0][1]
                    auto_issues.append(f"Список начиная с «{first_text[:50]}» и далее – замените круглый маркер (•) на тире, букву или цифру")
                    list_errors = []
            prev_para_empty = False
            continue
        
        if list_errors:
            first_text = list_errors[0][1]
            auto_issues.append(f"Список начиная с «{first_text[:50]}» и далее – замените круглый маркер (•) на тире, букву или цифру")
            list_errors = []
        
        if is_table_continuation(text):
            first_line = get_effective_first_line_indent(p)
            if abs(first_line) > 0.1:
                auto_issues.append(f"«{text[:50]}» – уберите абзацный отступ (должен быть 0 см)")
            prev_para_empty = False
            continue

        pf = p.paragraph_format
        alignment = get_effective_alignment(p)
        
        # --- ПРОВЕРКА ПОЯСНЕНИЙ К ФОРМУЛАМ ---
        if is_formula_where_line(text):
            errors = check_formula_explanation(text, p, prev_was_formula, prev_para_empty)
            auto_issues.extend(errors)
            prev_para_empty = False
            continue
        
        if is_formula_or_equation(text):
            prev_was_formula = True
            prev_para_empty = False
            continue
        
        # Определяем тип абзаца
        is_level1 = False
        is_subsection = False
        is_figure = text.startswith("Рисунок")
        is_table_caption = text.startswith("Таблица")
        
        if is_section_header(text):
            is_level1 = True
        elif re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]', text):
            is_subsection = True
        else:
            normalized = normalize_title(text)
            in_toc = any(normalize_title(e) == normalized for e in toc_entries) if toc_entries else False
            if in_toc and len(text) > 20:
                is_subsection = True

        # --- Заголовок раздела ---
        if is_level1:
            key = text[:80]
            
            if text.upper() != "ВВЕДЕНИЕ" or idx != start_idx:
                page_break = False
                if idx > start_idx:
                    prev_p = doc.paragraphs[idx - 1]
                    for run in prev_p.runs:
                        if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                            page_break = True
                for run in p.runs:
                    if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                        page_break = True
                if not page_break:
                    auto_issues.append(f"«{key}» – раздел должен начинаться с новой страницы")
            
            first_line = get_effective_first_line_indent(p)
            if abs(first_line) > 0.1:
                auto_issues.append(f"«{key}» – уберите абзацный отступ у заголовка")
            
            if not is_paragraph_bold(p):
                auto_issues.append(f"«{key}» – заголовок раздела должен быть полужирным")
            
            if re.match(r'^\d+\.', text) and not is_all_caps(text):
                auto_issues.append(f"«{key}» – заголовок раздела должен быть прописными буквами")
            
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                auto_issues.append(f"«{key}» – выровняйте заголовок по центру")
            
            if text.endswith("."):
                auto_issues.append(f"«{key}» – удалите точку в конце")
            
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                auto_issues.append(f"«{key}» – после заголовка должна быть пустая строка")
        
        # --- Подраздел ---
        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()
            key = f"Подраздел «{sub_name[:50]}»"
            
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                auto_issues.append(f"{key} – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
            
            if not is_paragraph_bold(p):
                auto_issues.append(f"{key} – заголовок должен быть полужирным")
            
            if text.endswith("."):
                auto_issues.append(f"{key} – удалите точку в конце")
            
            if prev_para_empty:
                auto_issues.append(f"{key} – уберите пустую строку перед подразделом")
        
        # --- Рисунок ---
        elif is_figure:
            figure_counter += 1
            fig_match = re.match(r'^(Рисунок\s+\d+(?:\.\d+)?)', text)
            fig_number = fig_match.group(1) if fig_match else f"Рисунок {figure_counter}"
            
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                auto_issues.append(f"{fig_number} – выровняйте подпись по центру")
            
            if text.endswith(".") and not re.search(r'\([^)]*\)\.$', text):
                auto_issues.append(f"{fig_number} – удалите точку в конце")
            
            m = re.match(r'^Рисунок\s+\d+(?:\.\d+)?\s*[–\-]\s*(.+)$', text)
            if m:
                title = m.group(1).strip()
                if title and title[0].islower():
                    auto_issues.append(f"{fig_number} – название должно начинаться с большой буквы")
            
            # ТОЛЬКО пустая строка ПОСЛЕ рисунка → в ручную проверку
            if idx + 1 < len(doc.paragraphs) and not is_empty_paragraph(doc.paragraphs[idx + 1]):
                manual_checks.append(f"{fig_number} – проверьте наличие пустой строки после рисунка")
        
        # --- Подпись таблицы ---
        elif is_table_caption:
            pass
        
        # --- Обычный текст ---
        else:
            key = text[:50]
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                auto_issues.append(f"«{key}» – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
            if pf.space_before and pf.space_before.pt > 0.5:
                auto_issues.append(f"«{key}» – интервал перед абзацем должен быть 0 пт")
        
        prev_para_empty = False
        prev_was_formula = False
    
    if list_errors:
        first_text = list_errors[0][1]
        auto_issues.append(f"Список начиная с «{first_text[:50]}» и далее – замените круглый маркер (•) на тире, букву или цифру")

    # ---------- 5. ТАБЛИЦЫ ----------
    try:
        start_element = doc.paragraphs[start_idx]._element
        start_body_pos = list(doc.element.body).index(start_element)
    except:
        start_body_pos = 0

    end_body_pos = len(doc.element.body)
    if lit_start is not None:
        try:
            lit_element = doc.paragraphs[lit_start]._element
            end_body_pos = list(doc.element.body).index(lit_element)
        except:
            pass

    tables_in_range = []
    for table in doc.tables:
        if get_table_depth(table) > 0:
            continue
        try:
            tbl_pos = list(doc.element.body).index(table._element)
        except:
            continue
        if not (start_body_pos < tbl_pos < end_body_pos):
            continue
        if len(table.rows) == 0:
            continue
        tables_in_range.append((tbl_pos, table))

    main_tables = []
    
    for tbl_pos, table in tables_in_range:
        caption, cap_pos = find_nearest_caption(doc, tbl_pos, start_body_pos)
        
        if caption and re.match(r'Таблица\s+[\d.]+\s+[–—]', caption):
            main_tables.append((tbl_pos, table, caption, cap_pos))
        elif caption and is_table_continuation(caption):
            continue
        else:
            caption2, cap_pos2 = find_nearest_caption(doc, cap_pos if cap_pos else tbl_pos, start_body_pos)
            if caption2 and is_table_continuation(caption2):
                continue
            main_tables.append((tbl_pos, table, None, None))

    for t_idx, (tbl_pos, table, caption, cap_pos) in enumerate(main_tables, start=1):
        if caption:
            tbl_num_match = re.match(r'Таблица\s+([\d.]+)', caption)
            tbl_num = tbl_num_match.group(1) if tbl_num_match else str(t_idx)
            key = f"Таблица {tbl_num}"
            
            if '—' not in caption and '–' not in caption:
                if '--' in caption or ' - ' in caption:
                    auto_issues.append(f"{key} – замените дефис на тире (—) в подписи")
            
            if not re.match(r'Таблица\s+[\d.]+\s+[–—]\s+', caption):
                auto_issues.append(f"{key} – должно быть «Таблица {tbl_num} — Название»")
            
            if caption.rstrip().endswith("."):
                auto_issues.append(f"{key} – удалите точку в конце названия")
            
            # Пустая строка ПЕРЕД подписью → авто
            if cap_pos is not None and cap_pos > start_idx:
                if not is_empty_paragraph(doc.paragraphs[cap_pos - 1]):
                    auto_issues.append(f"{key} – добавьте пустую строку перед подписью таблицы")
            
            # Пустая строка ПОСЛЕ таблицы → в ручную проверку
            next_para = None
            for i in range(tbl_pos + 1, end_body_pos):
                elem = doc.element.body[i]
                if elem.tag.endswith('p'):
                    for para in doc.paragraphs:
                        if para._element is elem:
                            next_para = para
                            break
                    break
            if next_para and not is_empty_paragraph(next_para):
                manual_checks.append(f"{key} – проверьте наличие пустой строки после таблицы")
        else:
            auto_issues.append(f"Таблица {t_idx} – отсутствует подпись над таблицей")

        bold_in_table = False
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.bold:
                            bold_in_table = True
                            break
                    if bold_in_table:
                        break
                if bold_in_table:
                    break
            if bold_in_table:
                break
        if bold_in_table:
            tbl_num = tbl_num_match.group(1) if caption and tbl_num_match else str(t_idx)
            key = f"Таблица {tbl_num}"
            auto_issues.append(f"{key} – уберите полужирное начертание внутри таблицы")

        # Перенос таблицы → ТОЛЬКО в ручную проверку
        if len(table.rows) > 2:
            tbl_num = tbl_num_match.group(1) if caption and tbl_num_match else str(t_idx)
            key = f"Таблица {tbl_num}"
            next_main_pos = end_body_pos
            for next_pos, _, _, _ in main_tables[t_idx:]:
                if next_pos > tbl_pos:
                    next_main_pos = next_pos
                    break
            
            markers_found = False
            for i in range(tbl_pos + 1, next_main_pos):
                if i < len(doc.paragraphs):
                    txt = doc.paragraphs[i].text.strip()
                    if txt and re.search(r'(?:Продолжение|Окончание)\s+таблицы?\s*' + re.escape(tbl_num), txt):
                        markers_found = True
                        break
            
            if not markers_found:
                manual_checks.append(f"Таблица {tbl_num} – проверьте наличие «Продолжение таблицы {tbl_num}» / «Окончание таблицы {tbl_num}» при переносе на следующую страницу")

    # ---------- 6. СПИСОК ИСТОЧНИКОВ ----------
    if lit_start is not None:
        sources_with_issues = 0
        for i in range(lit_start + 1, lit_end):
            source = doc.paragraphs[i]
            txt = source.text.strip()
            if not txt or has_page_number(txt):
                continue
            
            has_issue = False
            left_indent = get_effective_left_indent(source)
            if abs(left_indent) > 0.1:
                has_issue = True
            first_line = get_effective_first_line_indent(source)
            if abs(first_line - 1.0) > 0.1:
                has_issue = True
            if has_issue:
                sources_with_issues += 1
        
        if sources_with_issues > 0:
            auto_issues.append(
                "Список источников – проверьте оформление: "
                "отступ слева 0 см, отступ первой строки 1,0 см, "
                "междустрочный интервал 1,2 (множитель), выравнивание по ширине"
            )

    # ---------- ГРУППИРУЕМ И ВЫВОДИМ ----------
    all_issues = auto_issues
    if manual_checks:
        all_issues.append("📋 Для проверки человеком:")
        all_issues.extend(manual_checks)
    
    return group_issues(all_issues)

# Интерфейс
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write("Загрузите документ в формате .docx – проверка по полному чек-листу.")
uploaded_file = st.file_uploader("Выберите файл", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Проверяем..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for r in results:
        if r.startswith("📋"):
            st.markdown(f"**{r}**")
        else:
            st.write(f"• {r}")
