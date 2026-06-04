import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import re
from collections import defaultdict, Counter
import zipfile
from lxml import etree
import unicodedata

NSMAP = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# ------------------------------------------------------------
# Нормализация текста
# ------------------------------------------------------------
def normalize_text(text):
    if not text:
        return text
    result = []
    for ch in text:
        cat = unicodedata.category(ch)
        if cat == 'Cf':
            continue
        if ch.isspace():
            result.append(' ')
        else:
            if result and result[-1].isdigit() and ch.isalpha():
                result.append(' ')
            result.append(ch)
    return re.sub(r'\s+', ' ', ''.join(result)).strip()

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
    """Проверяет, является ли параграф полужирным"""
    for run in paragraph.runs:
        if run.bold:
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
                if pPr_rPr.find(qn('w:b')) is not None:
                    return True
        
        for r in paragraph._element.findall(qn('w:r')):
            rPr = r.find(qn('w:rPr'))
            if rPr is not None:
                if rPr.find(qn('w:b')) is not None:
                    return True
    except:
        pass
    
    return False

def is_empty_paragraph(paragraph):
    return len(paragraph.text.strip()) == 0

def has_page_number(text):
    return bool(re.search(r'[\t\s\.]{2,}\d+$', text))

def is_section_header(text):
    """Заголовок раздела первого уровня"""
    cleaned = normalize_text(text)
    if not cleaned:
        return False
    
    upper_cleaned = cleaned.upper()
    
    # Служебные разделы
    if upper_cleaned in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}:
        return True
    
    # Начинается с "ГЛАВА" или "РАЗДЕЛ"
    if re.match(r'^(ГЛАВА|РАЗДЕЛ)\s+\d+', upper_cleaned):
        return True
    
    # Текстовый заголовок ЗАГЛАВНЫМИ БУКВАМИ (длиной более 3 символов)
    only_letters = re.sub(r'[\d\s\.,;:!?\-–—()«»""''«»]', '', cleaned)
    if only_letters and len(only_letters) > 3:
        if only_letters == only_letters.upper():
            return True
    
    return False

def is_subsection_header(text):
    """Подраздел вида 1.1 Текст, 1.1.1 Текст"""
    cleaned = text.strip()
    if not cleaned:
        return False
    if re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-яA-Za-z]', cleaned):
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
    code = ord(ch)
    if code in [0x2D, 0x2010, 0x2011, 0x2012, 0x2013, 0x2014, 0x2015]:
        return True
    if code in [0xF02D]:
        return True
    return False

def is_bullet_char(ch):
    code = ord(ch)
    return code in [0x2022, 0x2023, 0x25E6, 0x25CF, 0x25CB, 0x26AB, 0x2B24]

def is_arrow_char(ch):
    code = ord(ch)
    return (0x2190 <= code <= 0x21FF or
            0x2794 <= code <= 0x27BF or
            0x2B00 <= code <= 0x2BFF)

def get_list_marker_info(paragraph, doc):
    text = paragraph.text.strip()
    if not text:
        return False, "", True
    first_char = text[0] if text else ""
    if is_dash_char(first_char):
        return True, "тире", True
    if re.match(r'^\d+\)\s', text):
        return True, "нумерованный", True
    if re.match(r'^\d+\.\s', text):
        return True, "нумерованный", True
    if re.match(r'^[а-яё]\)\s', text, re.IGNORECASE):
        return True, "буквенный", True
    if re.match(r'^[a-z]\)\s', text, re.IGNORECASE):
        return True, "буквенный", True
    if is_bullet_char(first_char):
        return True, "круглый маркер (•)", False
    if is_arrow_char(first_char):
        return True, f"недопустимый маркер (стрелка {first_char})", False
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
                                                        else:
                                                            return True, "круглый маркер (•)", False
                                                else:
                                                    return True, "круглый маркер (•)", False
                                            elif fmt == 'decimal':
                                                return True, "нумерованный (цифры)", True
                                            elif fmt in ['lowerLetter', 'upperLetter']:
                                                return True, "буквенный", True
                                            else:
                                                return True, f"формат '{fmt}'", True
    except:
        pass
    left_indent = get_effective_left_indent(paragraph)
    first_line = get_effective_first_line_indent(paragraph)
    if left_indent > 0.5 and first_line < 0:
        if not re.match(r'[А-Яа-яёЁA-Za-z0-9]', first_char) and first_char != ' ':
            return True, f"недопустимый маркер ({first_char})", False
    return False, "", True

def get_effective_first_line_indent(paragraph):
    """Получает эффективный отступ первой строки в см
    Проверяет только явно заданный отступ в параграфе, игнорируя стиль
    """
    # Проверяем XML напрямую - только явный отступ в параграфе
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                first_line = ind.get(qn('w:firstLine'))
                if first_line is not None:
                    # Отступ задан явно в параграфе
                    twips = int(first_line)
                    cm = twips * 2.54 / 1440
                    # Если отступ очень маленький (менее 0.05 см), считаем его нулевым
                    if abs(cm) < 0.05:
                        return 0.0
                    return cm
    except:
        pass
    
    # Если в XML нет явного отступа, возвращаем 0
    return 0.0

def is_on_new_page(doc, body_idx, start_body_pos=0, min_empty_paragraphs=10):
    """Проверяет, начинается ли элемент с новой страницы"""
    body_elems = list(doc.element.body)
    if body_idx == start_body_pos:
        return True
    
    # Проверяем сам элемент и элементы перед ним
    for i in range(body_idx, start_body_pos - 1, -1):
        if i < 0:
            break
            
        elem = body_elems[i]
        
        # Проверка разрыва раздела
        if elem.tag == qn('w:sectPr'):
            if i == len(body_elems) - 1:
                continue
            type_el = elem.find(qn('w:type'))
            val = type_el.get(qn('w:val')) if type_el is not None else None
            if val != 'continuous':
                return True
        
        if elem.tag == qn('w:p'):
            # Проверка явного разрыва страницы
            for br in elem.findall('.//w:br', NSMAP):
                if br.get(qn('w:type')) == 'page':
                    return True
            
            # Проверка свойства pageBreakBefore у параграфа
            pPr = elem.find(qn('w:pPr'))
            if pPr is not None:
                if pPr.find(qn('w:pageBreakBefore')) is not None:
                    return True
                
                # Проверка sectPr внутри pPr
                sectPr = pPr.find(qn('w:sectPr'))
                if sectPr is not None:
                    type_el = sectPr.find(qn('w:type'))
                    val = type_el.get(qn('w:val')) if type_el is not None else None
                    if val != 'continuous':
                        return True
            
            # Проверяем, является ли элемент пустым
            if not has_content(elem):
                continue
            
            # Если нашли непустой элемент, проверяем, не является ли он на новой странице
            # из-за большого количества пустых строк
            blank_count = 0
            for j in range(i - 1, start_body_pos - 1, -1):
                prev_elem = body_elems[j]
                if prev_elem.tag == qn('w:p') and not has_content(prev_elem):
                    blank_count += 1
                else:
                    break
            if blank_count >= min_empty_paragraphs:
                return True
            
            # Если это не новый раздел, продолжаем поиск
            if i < body_idx:
                continue
                
        # Если дошли до начала, возвращаем False
        if i == start_body_pos:
            return False
    
    return False

def get_alignment_from_xml(paragraph):
    """Получает выравнивание из XML напрямую"""
    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            jc = pPr.find(qn('w:jc'))
            if jc is not None:
                val = jc.get(qn('w:val'))
                if val == 'center':
                    return WD_ALIGN_PARAGRAPH.CENTER
                elif val == 'right':
                    return WD_ALIGN_PARAGRAPH.RIGHT
                elif val == 'left':
                    return WD_ALIGN_PARAGRAPH.LEFT
                elif val == 'both':
                    return WD_ALIGN_PARAGRAPH.JUSTIFY
    except:
        pass
    
    # Проверяем стиль
    try:
        style = paragraph.style
        if style and style.paragraph_format.alignment is not None:
            return style.paragraph_format.alignment
    except:
        pass
    
    return paragraph.alignment

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
                    return int(left) * 2.54 / 1440
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

def check_page_numbering(file, intro_start_idx):
    issues = []
    try:
        with zipfile.ZipFile(file, 'r') as z:
            doc_xml = etree.fromstring(z.read('word/document.xml'))
            nsmap = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            }
            body = doc_xml.find('.//w:body', nsmap)
            body_elements = list(body)
            if intro_start_idx >= len(body_elements):
                return []
            target_sectPr = None
            for i in range(intro_start_idx, len(body_elements)):
                if body_elements[i].tag == qn('w:sectPr'):
                    target_sectPr = body_elements[i]
                    break
            if target_sectPr is None:
                target_sectPr = body.find('.//w:sectPr', nsmap)
            if target_sectPr is None:
                return []
            footer_refs = target_sectPr.findall('.//w:footerReference', nsmap)
            if not footer_refs:
                return []
            try:
                rels_xml = etree.fromstring(z.read('word/_rels/document.xml.rels'))
                rels_map = {}
                for rel in rels_xml:
                    rel_id = rel.get('Id')
                    target = rel.get('Target')
                    rel_type = rel.get('Type')
                    if 'footer' in rel_type.lower():
                        rels_map[rel_id] = target
            except:
                return []
            for footer_ref in footer_refs:
                ref_id = footer_ref.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                file_name = rels_map.get(ref_id)
                if not file_name:
                    continue
                full_path = 'word/' + file_name
                try:
                    content = z.read(full_path)
                    tree = etree.fromstring(content)
                    paragraphs = []
                    for para in tree.findall('.//w:p', nsmap):
                        texts = []
                        for t in para.findall('.//w:t', nsmap):
                            if t.text:
                                texts.append(t.text)
                        full_text = ''.join(texts)
                        has_page = False
                        for fld in para.findall('.//w:fldChar', nsmap):
                            fld_type = fld.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType')
                            if fld_type == 'begin':
                                has_page = True
                        for instr in para.findall('.//w:instrText', nsmap):
                            if instr.text and 'PAGE' in instr.text:
                                has_page = True
                        jc = para.find('.//w:jc', nsmap)
                        align = jc.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if jc is not None else None
                        font_size = None
                        if has_page:
                            for r in para.findall('.//w:r', nsmap):
                                rPr = r.find('.//w:rPr', nsmap)
                                if rPr is not None:
                                    sz = rPr.find('.//w:sz', nsmap)
                                    if sz is None:
                                        sz = rPr.find('.//w:szCs', nsmap)
                                    if sz is not None:
                                        sz_val = sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                                        if sz_val:
                                            font_size = int(sz_val) / 2
                                            break
                        paragraphs.append({
                            'text': full_text,
                            'has_page': has_page,
                            'alignment': align,
                            'font_size': font_size,
                            'is_empty': not bool(full_text)
                        })
                    has_page_in_footer = any(p['has_page'] for p in paragraphs)
                    if not has_page_in_footer:
                        static_numbers = [p for p in paragraphs if p['text'].isdigit()]
                        if static_numbers:
                            issues.append("Нумерация страниц – используйте автонумерацию (поле PAGE) вместо статичного номера")
                        else:
                            issues.append("Нумерация страниц – отсутствует нумерация в колонтитуле")
                        return issues
                    page_para_errors = []
                    for j, p in enumerate(paragraphs):
                        if p['has_page']:
                            if p['font_size'] is not None and abs(p['font_size'] - 14) > 0.5:
                                page_para_errors.append(f"размер шрифта {p['font_size']} pt (должен быть 14)")
                            if p['alignment'] != 'center':
                                page_para_errors.append("выравнивание должно быть по центру")
                            if j + 1 < len(paragraphs) and paragraphs[j + 1]['is_empty']:
                                page_para_errors.append("уберите пустую строку после номера")
                        elif p['is_empty'] and j > 0 and paragraphs[j-1]['has_page']:
                            page_para_errors.append("уберите пустую строку после номера")
                    if page_para_errors:
                        issues.append(f"Нумерация страниц – {', '.join(page_para_errors)}")
                    return issues
                except:
                    continue
    except Exception:
        pass
    return issues

def get_category_order(issue):
    if issue.startswith("Таблицы –"):
        return (4, 0)
    if issue.startswith("Рисунки –"):
        return (3, 0)
    if issue.startswith("«") and "» –" in issue:
        return (1, 1)
    if issue.startswith("Подраздел «"):
        return (2, 1)
    if issue.startswith("Рисунок "):
        return (3, 1)
    if issue.startswith("Таблица "):
        return (4, 1)
    if issue.startswith("Пояснение к формуле «"):
        return (5, 1)
    if issue.startswith("Список, начиная с «"):
        return (6, 1)
    if issue.startswith("Нумерация страниц"):
        return (7, 1)
    if issue.startswith("Поля страниц"):
        return (8, 1)
    if issue.startswith("Основной текст, начиная со строки «"):
        return (9, 1)
    if issue.startswith("Список источников"):
        return (10, 1)
    return (11, 1)

def group_issues(issues_list):
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
        match = re.match(
            r'^(?:«([^»]+)»|(Рисунок\s+[\d.]+)|(Таблица\s+[\d.]+)|'
            r'(Подраздел\s+«([^»]+)»)|(Пояснение к формуле «([^»]+)»)|'
            r'(Список начиная с «([^»]+)»)|(Нумерация страниц))\s*[–-]\s*(.+)$',
            issue
        )
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
            elif match.group(10):
                key = "Нумерация страниц"
            if key:
                message = match.group(11)
                grouped[key].append(message)
            else:
                standalone.append(issue)
        else:
            standalone.append(issue)
    figure_keys = [k for k in grouped if re.match(r'^Рисунок\s+\d', k)]
    if len(figure_keys) > 3:
        all_fig_messages = []
        for k in figure_keys:
            all_fig_messages.extend(grouped[k])
        msg_counts = Counter(all_fig_messages)
        unique_msgs = list(msg_counts.keys())
        first_seen = {}
        for msg in all_fig_messages:
            if msg not in first_seen:
                first_seen[msg] = len(first_seen)
        unique_msgs.sort(key=lambda m: (-msg_counts[m], first_seen[m]))
        replacements = {
            "удалите точку в конце": "в названии не должно быть точки в конце",
            "уберите точку в конце": "в названии не должно быть точки в конце"
        }
        transformed = [replacements.get(m, m) for m in unique_msgs]
        combined_fig = "Рисунки – " + ", ".join(transformed)
        for k in figure_keys:
            del grouped[k]
        standalone.insert(0, combined_fig)
    caption_msg_pattern = re.compile(r'^Исправьте название на «Таблица [\d.]+ – Название»$')
    caption_msgs_keys = []
    for key in list(grouped.keys()):
        if key.startswith("Таблица ") and re.match(r'^Таблица\s+\d', key):
            msgs = grouped[key]
            new_msgs = []
            for m in msgs:
                if caption_msg_pattern.match(m):
                    caption_msgs_keys.append(key)
                else:
                    new_msgs.append(m)
            if new_msgs:
                grouped[key] = new_msgs
            else:
                del grouped[key]
    if len(caption_msgs_keys) >= 2:
        for k in caption_msgs_keys:
            if k in grouped and not grouped[k]:
                del grouped[k]
        standalone.insert(0, "Таблицы – исправьте название на «Таблица № – Название»")
    else:
        for k in caption_msgs_keys:
            grouped[k].append("Исправьте название на «Таблица " + k.split()[-1] + " – Название»")
    before_msgs = [issue for issue in standalone if issue.startswith("Таблица – название должно быть перед таблицей")]
    if len(before_msgs) >= 2:
        for msg in before_msgs:
            standalone.remove(msg)
        standalone.insert(0, "Таблицы – название должно быть перед таблицей")
    before_keys = [k for k, v in grouped.items() if k == "Таблица" and any(m.startswith("название должно быть перед таблицей") for m in v)]
    if len(before_keys) >= 2:
        for k in before_keys:
            del grouped[k]
        standalone.insert(0, "Таблицы – название должно быть перед таблицей")
    table_common_msgs = [issue for issue in standalone if issue.startswith("Таблицы –")]
    if len(table_common_msgs) > 1:
        parts = [issue[len("Таблицы –"):] for issue in table_common_msgs]
        combined = "Таблицы – " + ", ".join(parts)
        for issue in table_common_msgs:
            standalone.remove(issue)
        standalone.append(combined)
    result = []
    for issue in standalone:
        result.append(issue)
    for key, messages in grouped.items():
        unique_msgs = list(dict.fromkeys(messages))
        if len(unique_msgs) == 1:
            if key.startswith("Подраздел «") or key.startswith("Пояснение к формуле «") or key.startswith("Список начиная с «"):
                result.append(f"{key} – {unique_msgs[0]}")
            elif key.startswith("Рисунок ") or key.startswith("Таблица ") or key == "Нумерация страниц":
                result.append(f"{key} – {unique_msgs[0]}")
            else:
                result.append(f"«{key}» – {unique_msgs[0]}")
        else:
            combined = "; ".join(unique_msgs)
            if key.startswith("Подраздел «") or key.startswith("Пояснение к формуле «") or key.startswith("Список начиная с «"):
                result.append(f"{key} – {combined}")
            elif key.startswith("Рисунок ") or key.startswith("Таблица ") or key == "Нумерация страниц":
                result.append(f"{key} – {combined}")
            else:
                result.append(f"«{key}» – {combined}")
    result.sort(key=lambda x: (get_category_order(x), x))
    cont_pattern = re.compile(r'^Таблица\s+(\d+(?:\.\d+)?)\s+–\s+проверьте\s+наличие\s+«Продолжение таблицы \1»\s+/\s+«Окончание таблицы \1» при переносе на следующую страницу$')
    cont_msgs = [m for m in manual_issues if cont_pattern.match(m)]
    if len(cont_msgs) >= 2:
        for m in cont_msgs:
            manual_issues.remove(m)
        manual_issues.append("Таблицы – проверьте наличие «Продолжение таблицы» / «Окончание таблицы» при переносе на следующую страницу")
    if manual_issues:
        result.append("\n📋 Для проверки человеком:")
        result.extend(manual_issues)
    return result

def get_font_size_pt(paragraph):
    sizes = set()
    for run in paragraph.runs:
        if run.font.size:
            sizes.add(run.font.size.pt)
    return sizes

def check_empty_line_before_after(doc, idx, start_idx, label):
    errors = []
    body_elems = list(doc.element.body)
    if idx > start_idx:
        prev_elem = body_elems[idx - 1]
        is_empty_prev = (
            prev_elem.tag == qn('w:p') and
            not prev_elem.xpath('string(.)').strip()
        )
        if not is_empty_prev:
            errors.append(f"{label} – добавьте пустую строку перед подписью")
    if idx + 1 < len(body_elems):
        next_elem = body_elems[idx + 1]
        is_empty_next = (
            next_elem.tag == qn('w:p') and
            not next_elem.xpath('string(.)').strip()
        )
        if not is_empty_next:
            errors.append(f"{label} – добавьте пустую строку после подписи")
    return errors

def has_content(elem):
    """Проверяет, есть ли у элемента реальное содержимое"""
    texts = [node.text or '' for node in elem.iter() if node.tag == qn('w:t')]
    if any(t.strip() for t in texts):
        return True
    
    if elem.find('.//w:drawing', NSMAP) is not None:
        return True
    
    if elem.find('.//w:object', NSMAP) is not None:
        return True
    
    return False

def is_on_new_page(doc, body_idx, start_body_pos=0, min_empty_paragraphs=10):
    """Проверяет, начинается ли элемент с новой страницы"""
    body_elems = list(doc.element.body)
    if body_idx == start_body_pos:
        return True
    
    blank_count = 0
    for i in range(body_idx - 1, start_body_pos - 1, -1):
        elem = body_elems[i]
        
        # Проверка разрыва раздела
        if elem.tag == qn('w:sectPr'):
            if i == len(body_elems) - 1:
                continue
            type_el = elem.find(qn('w:type'))
            val = type_el.get(qn('w:val')) if type_el is not None else None
            if val != 'continuous':
                return True
        
        if elem.tag == qn('w:p'):
            # Проверка явного разрыва страницы
            for br in elem.findall('.//w:br', NSMAP):
                if br.get(qn('w:type')) == 'page':
                    return True
            
            # Проверка свойства pageBreakBefore у параграфа
            pPr = elem.find(qn('w:pPr'))
            if pPr is not None:
                if pPr.find(qn('w:pageBreakBefore')) is not None:
                    return True
                
                sectPr = pPr.find(qn('w:sectPr'))
                if sectPr is not None:
                    type_el = sectPr.find(qn('w:type'))
                    val = type_el.get(qn('w:val')) if type_el is not None else None
                    if val != 'continuous':
                        return True
            
            if has_content(elem):
                return blank_count >= min_empty_paragraphs
            else:
                blank_count += 1
                continue
        
        if elem.tag == qn('w:tbl'):
            return blank_count >= min_empty_paragraphs
    
    return True

# ------------------------------------------------------------
# Функция анализа подразделов (тестовый модуль)
# ------------------------------------------------------------
def analyze_subsections(doc):
    """Анализирует подразделы и определяет, нужно ли убирать пустую строку"""
    results = []
    
    empty_paragraph_indices = set()
    for idx, para in enumerate(doc.paragraphs):
        if not para.text.strip():
            empty_paragraph_indices.add(idx)
    
    prev_nonempty_was_section = False
    prev_nonempty_text = ""
    
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        
        is_section = is_section_header(text)
        is_sub = is_subsection_header(text)
        
        if is_sub:
            has_empty_before = (idx - 1) in empty_paragraph_indices
            
            is_technical_prev = False
            if idx > 0:
                prev_text = doc.paragraphs[idx - 1].text.strip()
                if re.search(r'(?:Продолжение|Окончание)\s+таблицы', prev_text, re.IGNORECASE):
                    is_technical_prev = True
            
            error_condition = has_empty_before and not prev_nonempty_was_section and not is_technical_prev
            
            results.append({
                "index": idx,
                "text": text[:80],
                "has_empty_before": has_empty_before,
                "prev_was_section": prev_nonempty_was_section,
                "prev_text": prev_nonempty_text[:60] if prev_nonempty_text else "—",
                "is_technical_prev": is_technical_prev,
                "error_should_show": error_condition,
                "error_msg": f"Подраздел «{text[:50]}» – уберите пустую строку перед подразделом" if error_condition else None
            })
        
        if is_section:
            prev_nonempty_was_section = True
        else:
            prev_nonempty_was_section = False
        
        prev_nonempty_text = text
    
    return results

# ------------------------------------------------------------
# Функция анализа заголовков разделов (тестовый модуль)
# ------------------------------------------------------------
def analyze_section_headers(doc):
    """Анализирует заголовки разделов и определяет проблемы с отступами и новой страницей"""
    results = []
    
    # Находим начало основного текста
    start_idx = None
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt:
            continue
        if has_page_number(txt):
            continue
        if is_section_header(txt):
            start_idx = i
            break
    
    if start_idx is None:
        return results
    
    # Получаем соответствие параграфов и body элементов
    para_to_body_idx = {}
    body_elems = list(doc.element.body)
    for i, elem in enumerate(body_elems):
        if elem.tag == qn('w:p'):
            for j, p in enumerate(doc.paragraphs):
                if p._element is elem:
                    para_to_body_idx[j] = i
                    break
    
    try:
        start_body_pos = para_to_body_idx[start_idx]
    except:
        start_body_pos = 0
    
    for idx, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            continue
        if has_page_number(text):
            continue
        
        if is_section_header(text):
            # Получаем фактический отступ первой строки
            first_line = get_effective_first_line_indent(p)
            
            # Получаем выравнивание
            alignment = get_alignment_from_xml(p)
            
            # Проверяем, начинается ли с новой страницы
            body_idx = para_to_body_idx.get(idx)
            starts_new_page = False
            
            if body_idx is not None:
                starts_new_page = is_on_new_page(doc, body_idx, start_body_pos)
                
                # Дополнительная проверка: разрыв страницы в стиле
                if not starts_new_page:
                    try:
                        style = p.style
                        if style and style.paragraph_format.page_break_before:
                            starts_new_page = True
                    except:
                        pass
            
            # Проверяем пустую строку после
            has_empty_after = False
            if idx + 1 < len(doc.paragraphs):
                if is_empty_paragraph(doc.paragraphs[idx + 1]):
                    has_empty_after = True
            
            # Для отступа используем порог 0.1 см (1 мм)
            # Если отступ меньше 0.1 см, считаем что его нет
            first_line_ok = abs(first_line) <= 0.1
            
            results.append({
                "index": idx,
                "text": text[:80],
                "first_line_indent_cm": first_line,
                "first_line_indent_ok": first_line_ok,
                "starts_new_page": starts_new_page,
                "starts_new_page_ok": starts_new_page or text.upper() == "ВВЕДЕНИЕ",
                "has_empty_after": has_empty_after,
                "has_empty_after_ok": has_empty_after,
                "is_bold": is_paragraph_bold(p),
                "alignment": alignment,
                "alignment_ok": alignment == WD_ALIGN_PARAGRAPH.CENTER
            })
    
    return results

# ------------------------------------------------------------
# Главная проверка документа
# ------------------------------------------------------------
def check_word_document(file):
    doc = docx.Document(file)
    auto_issues = []
    manual_checks = []

    # --- ПОЛЯ ---
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

    # --- ПОИСК НАЧАЛА ОСНОВНОГО ТЕКСТА ---
    start_idx = None
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt:
            continue
        if has_page_number(txt):
            continue
        if is_section_header(txt):
            start_idx = i
            break
    if start_idx is None:
        return ["✅ Ошибок не найдено. Документ соответствует чек-листу."]

    # --- НУМЕРАЦИЯ СТРАНИЦ ---
    try:
        intro_body_idx = None
        body_elements = list(doc.element.body)
        for i, elem in enumerate(body_elements):
            if elem.tag == qn('w:p'):
                full_text = ''.join(t.text or '' for t in elem.findall('.//w:t', qn('w'))).strip().upper()
                if full_text == 'ВВЕДЕНИЕ':
                    intro_body_idx = i
                    break
        if intro_body_idx is not None:
            page_issues = check_page_numbering(file, intro_body_idx)
            auto_issues.extend(page_issues)
    except:
        pass

    toc_entries = extract_toc_entries(doc, start_idx)

    # --- ГРАНИЦЫ СПИСКА ИСТОЧНИКОВ ---
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

    # --- ПРОВЕРКА ОСНОВНОГО ТЕКСТА ---
    figure_counter = 0
    prev_para_empty = False
    prev_was_formula = False

    end_idx = lit_start if lit_start is not None else len(doc.paragraphs)
    list_errors = []
    indent_issues = []

    figure_numbers_found = []
    table_numbers_found = []

    para_to_body_idx = {}
    body_elems = list(doc.element.body)
    for i, elem in enumerate(body_elems):
        if elem.tag == qn('w:p'):
            for j, p in enumerate(doc.paragraphs):
                if p._element is elem:
                    para_to_body_idx[j] = i
                    break

    table_captions_info = []

    try:
        start_body_pos = para_to_body_idx[start_idx]
    except:
        start_body_pos = 0

    for idx in range(start_idx, end_idx):
        p = doc.paragraphs[idx]
        text = p.text.strip()
        norm_text = normalize_text(text)

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

        is_level1 = is_section_header(norm_text)
        is_subsection = False

        if not is_level1:
            if is_subsection_header(norm_text):
                is_subsection = True
            else:
                normalized = normalize_title(norm_text)
                in_toc = any(normalize_title(e) == normalized for e in toc_entries) if toc_entries else False
                if in_toc and len(text) > 20:
                    is_subsection = True

        # --- ПРОВЕРКА ДЛЯ ОСНОВНОГО ТЕКСТА (не заголовки) ---
        if not is_level1 and not is_subsection:
            # Проверка списков
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

            if is_table_continuation(norm_text):
                first_line = get_effective_first_line_indent(p)
                if abs(first_line) > 0.1:
                    auto_issues.append(f"«{text[:50]}» – уберите абзацный отступ (должен быть 0 см)")
                prev_para_empty = False
                continue

            if is_formula_where_line(norm_text):
                errors = check_formula_explanation(text, p, prev_was_formula, prev_para_empty)
                auto_issues.extend(errors)
                prev_para_empty = False
                continue

            if is_formula_or_equation(norm_text):
                prev_was_formula = True
                prev_para_empty = False
                continue

            # Проверка рисунков
            is_figure = norm_text.startswith("Рисунок") or norm_text.startswith("Рис.")
            is_table_caption = norm_text.startswith("Таблица")

            if is_figure:
                num_match = re.search(r'(?:Рисунок|Рис\.)\s*:?\s*(\d+(?:\.\d+)?)', norm_text)
                if num_match:
                    fig_num = num_match.group(1)
                    figure_numbers_found.append(float(fig_num))
                else:
                    figure_counter += 1
                    fig_num = str(figure_counter)
                fig_number = f"Рисунок {fig_num}"

                if norm_text.startswith("Рис."):
                    auto_issues.append(f"{fig_number} – измените «Рис.» на «Рисунок»")
                if re.match(r'Рисунок\s*:', norm_text):
                    auto_issues.append(f"{fig_number} – замените двоеточие на тире (формат: «Рисунок N — Название»)")
                if not re.search(r'Рисунок\s+\d+(?:\.\d+)?\s*[–—]', norm_text) and not re.search(r'Рисунок\s+\d+(?:\.\d+)?\s*--', norm_text) and not re.match(r'Рисунок\s*:', norm_text):
                    if re.search(r'Рисунок\s+\d+(?:\.\d+)?\s*[-]', norm_text):
                        auto_issues.append(f"{fig_number} – замените дефис на тире (—)")
                    else:
                        auto_issues.append(f"{fig_number} – должно быть тире после номера")
                if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.CENTER:
                    auto_issues.append(f"{fig_number} – выровняйте подпись по центру")
                if text.endswith(".") and not re.search(r'\([^)]*\)\.$', text):
                    auto_issues.append(f"{fig_number} – удалите точку в конце")
                m = re.match(r'^(?:Рисунок|Рис\.)\s+\d+(?:\.\d+)?\s*[–—]?\s*(.+)$', norm_text)
                if m:
                    title = m.group(1).strip()
                    if title and title[0].islower():
                        auto_issues.append(f"{fig_number} – название должно начинаться с большой буквы")

                sizes = get_font_size_pt(p)
                if sizes:
                    if any(abs(s - 14) > 0.5 for s in sizes):
                        auto_issues.append(f"{fig_number} – установите размер шрифта 14 пт (сейчас {', '.join(str(s) for s in sizes)} пт)")

                body_idx = para_to_body_idx.get(idx)
                if body_idx is not None:
                    empty_errors = check_empty_line_before_after(doc, body_idx, start_body_pos, fig_number)
                    auto_issues.extend(empty_errors)
                prev_para_empty = False
                continue

            elif is_table_caption and not is_table_continuation(norm_text):
                tbl_match = re.match(r'Таблица\s+(\d+(?:\.\d+)?)', norm_text)
                if tbl_match:
                    tbl_num = tbl_match.group(1)
                    tbl_num_float = float(tbl_num)
                    table_numbers_found.append(tbl_num_float)
                    key = f"Таблица {tbl_num}"

                    if not re.match(r'Таблица\s+\d+(?:\.\d+)?\s+[–—]?\s*\S', norm_text) and not re.search(r'Таблица\s+\d+(?:\.\d+)?\s*--', norm_text):
                        auto_issues.append(f"{key} – Исправьте название на «Таблица {tbl_num} – Название»")

                    if text.rstrip().endswith("."):
                        auto_issues.append(f"{key} – удалите точку в конце названия")
                    sizes = get_font_size_pt(p)
                    if sizes:
                        if any(abs(s - 14) > 0.5 for s in sizes):
                            auto_issues.append(f"{key} – установите размер шрифта 14 пт (сейчас {', '.join(str(s) for s in sizes)} пт)")

                    table_captions_info.append({
                        'para_idx': idx,
                        'body_idx': para_to_body_idx.get(idx),
                        'number': tbl_num_float,
                        'number_str': tbl_num,
                        'text': text,
                        'key': key
                    })
                prev_para_empty = False
                continue

            # Основной текст
            key = norm_text[:50]
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                indent_issues.append((key, first_line))
            if p.paragraph_format.space_before and p.paragraph_format.space_before.pt > 0.5:
                auto_issues.append(f"«{key}» – интервал перед абзацем должен быть 0 пт")

        # === Проверка заголовков разделов ===
        if is_level1:
            key = text[:80]
            
            # Для ВВЕДЕНИЯ не проверяем новую страницу
            if text.upper() != "ВВЕДЕНИЕ":
                body_idx = para_to_body_idx.get(idx)
                starts_new_page = False
                if body_idx is not None:
                    starts_new_page = is_on_new_page(doc, body_idx, start_body_pos)
                    
                    # Проверка page_break_before в стиле
                    if not starts_new_page:
                        try:
                            if p.style and p.style.paragraph_format.page_break_before:
                                starts_new_page = True
                        except:
                            pass
                    
                    # Проверка разрыва страницы в предыдущем параграфе
                    if not starts_new_page and idx > 0:
                        prev_para = doc.paragraphs[idx - 1]
                        if hasattr(prev_para, '_element'):
                            pPr = prev_para._element.find(qn('w:pPr'))
                            if pPr is not None:
                                if pPr.find(qn('w:pageBreakBefore')) is not None:
                                    starts_new_page = True
                
                if not starts_new_page:
                    auto_issues.append(f"«{key}» – раздел должен начинаться с новой страницы")
            
            # Проверка отступа первой строки - только явный отступ в параграфе
            # Используем функцию, которая игнорирует стиль
            first_line = get_effective_first_line_indent(p)
            # Если выравнивание по центру, отступ не имеет значения
            alignment = get_effective_alignment(p)
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                if round(abs(first_line), 2) > 0.05:
                    auto_issues.append(f"«{key}» – уберите абзацный отступ у заголовка (сейчас {first_line:.2f} см)")
            
            # Проверка полужирного начертания
            if not is_paragraph_bold(p):
                auto_issues.append(f"«{key}» – заголовок раздела должен быть полужирным")
            
            # Проверка выравнивания по центру
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                auto_issues.append(f"«{key}» – выровняйте заголовок по центру")
            
            # Проверка точки в конце
            if text.endswith("."):
                auto_issues.append(f"«{key}» – удалите точку в конце")
            
            # Проверка пустой строки после заголовка
            if idx + 1 < len(doc.paragraphs):
                next_para = doc.paragraphs[idx + 1]
                if not is_empty_paragraph(next_para):
                    auto_issues.append(f"«{key}» – после заголовка должна быть пустая строка")

        # === Проверка подразделов ===
        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s*', '', norm_text).strip()
            key = f"Подраздел «{sub_name[:50]}»"
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                auto_issues.append(f"{key} – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
            if not is_paragraph_bold(p):
                auto_issues.append(f"{key} – заголовок должен быть полужирным")
            if get_effective_alignment(p) != WD_ALIGN_PARAGRAPH.JUSTIFY:
                auto_issues.append(f"{key} – выровняйте по ширине")
            if text.endswith("."):
                auto_issues.append(f"{key} – удалите точку в конце")
            
            # Проверяем наличие реальной пустой строки перед подразделом
            has_empty_before = False
            if idx > 0:
                prev_para = doc.paragraphs[idx - 1]
                if is_empty_paragraph(prev_para):
                    has_empty_before = True
            
            # Проверяем, не является ли предыдущий текст техническим
            is_technical_prev = False
            if idx > 0:
                prev_text = doc.paragraphs[idx - 1].text.strip()
                if re.search(r'(?:Продолжение|Окончание)\s+таблицы', prev_text, re.IGNORECASE):
                    is_technical_prev = True
            
            # Ищем ПРЕДЫДУЩИЙ НЕПУСТОЙ параграф и проверяем, был ли он разделом
            prev_nonempty_was_section = False
            if idx > 0:
                for j in range(idx - 1, -1, -1):
                    prev_para_text = doc.paragraphs[j].text.strip()
                    if prev_para_text:
                        prev_nonempty_was_section = is_section_header(prev_para_text)
                        break
            
            # Ошибка: есть пустая строка И предыдущий непустой НЕ был разделом И не технический
            if has_empty_before and not prev_nonempty_was_section and not is_technical_prev:
                auto_issues.append(f"{key} – уберите пустую строку перед подразделом")

        # Сброс флагов
        if text:
            prev_para_empty = False
            prev_was_formula = False
        else:
            prev_para_empty = True

    # --- ЗАВЕРШАЮЩИЕ ПРОВЕРКИ ---
    if indent_issues:
        if len(indent_issues) > 2:
            first_key, first_indent = indent_issues[0]
            auto_issues.append(
                f"Основной текст, начиная со строки «{first_key}» – "
                f"установите абзацный отступ 1,0 см (сейчас {first_indent:.1f} см)"
            )
        else:
            for key, fl in indent_issues:
                auto_issues.append(f"«{key}» – установите абзацный отступ 1,0 см (сейчас {fl:.1f} см)")

    if list_errors:
        first_text = list_errors[0][1]
        auto_issues.append(f"Список начиная с «{first_text[:50]}» и далее – замените круглый маркер (•) на тире, букву или цифру")

    if figure_numbers_found:
        int_nums = sorted(set(int(n) for n in figure_numbers_found if n == int(n)))
        if int_nums:
            if int_nums[0] != 1:
                auto_issues.append("Рисунки – нумерация должна начинаться с 1")
            else:
                for expected in range(1, int_nums[-1] + 1):
                    if expected not in int_nums:
                        auto_issues.append(f"Рисунки – пропущен рисунок {expected}")
                        break

    table_seq_issues = []
    if table_numbers_found:
        int_tbl_nums = sorted(set(int(n) for n in table_numbers_found if n == int(n)))
        if int_tbl_nums:
            if int_tbl_nums[0] != 1 or any(expected not in int_tbl_nums for expected in range(1, int_tbl_nums[-1] + 1)):
                table_seq_issues.append("Таблицы – неверная нумерация")

    # --- ТАБЛИЦЫ ---
    end_body_pos = len(body_elems)
    if lit_start is not None:
        try:
            lit_element = doc.paragraphs[lit_start]._element
            end_body_pos = body_elems.index(lit_element)
        except:
            pass

    tables_in_range = []
    for table in doc.tables:
        if get_table_depth(table) > 0:
            continue
        try:
            tbl_pos = body_elems.index(table._element)
        except:
            continue
        if not (start_body_pos < tbl_pos < end_body_pos):
            continue
        if len(table.rows) == 0:
            continue
        tables_in_range.append((tbl_pos, table))

    table_issues = []
    for tbl_pos, table in tables_in_range:
        caption_info = None
        for cap in table_captions_info:
            if cap['body_idx'] is not None and cap['body_idx'] < tbl_pos:
                other_tables_between = any(
                    other_pos < tbl_pos and other_pos > cap['body_idx']
                    for other_pos, _ in tables_in_range
                )
                if not other_tables_between:
                    caption_info = cap
        if caption_info:
            key = caption_info['key']
            tbl_num = caption_info['number_str']

            if caption_info['body_idx'] is not None:
                empty_errors = check_empty_line_before_after(doc, caption_info['body_idx'], start_body_pos, key)
                table_issues.extend(empty_errors)

            bold_in_table = False
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if run.bold:
                                bold_in_table = True
                                break
                        if bold_in_table: break
                    if bold_in_table: break
                if bold_in_table: break
            if bold_in_table:
                table_issues.append(f"{key} – уберите полужирное начертание внутри таблицы")

            if len(table.rows) > 2:
                current_tbl_index = None
                for i, (pos, _) in enumerate(tables_in_range):
                    if pos == tbl_pos:
                        current_tbl_index = i
                        break
                if current_tbl_index is not None and current_tbl_index + 1 < len(tables_in_range):
                    next_tbl_pos = tables_in_range[current_tbl_index + 1][0]
                else:
                    next_tbl_pos = end_body_pos
                markers_found = False
                for i in range(tbl_pos + 1, next_tbl_pos):
                    if i < len(doc.paragraphs):
                        txt = doc.paragraphs[i].text.strip()
                        if txt and re.search(r'(?:Продолжение|Окончание)\s+таблицы?\s*' + re.escape(tbl_num), txt):
                            markers_found = True
                            break
                if not markers_found:
                    manual_checks.append(f"{key} – проверьте наличие «Продолжение таблицы {tbl_num}» / «Окончание таблицы {tbl_num}» при переносе на следующую страницу")
        else:
            prev_text = ""
            for i in range(tbl_pos - 1, start_body_pos - 1, -1):
                elem = body_elems[i]
                if elem.tag == qn('w:p'):
                    txt = ''.join(node.text or '' for node in elem.iter() if node.tag == qn('w:t')).strip()
                    if txt:
                        prev_text = txt[:50]
                        break
            table_issues.append(
                f"Таблица – название должно быть перед таблицей (расположена после абзаца: «{prev_text}…»)" if prev_text
                else "Таблица – название должно быть перед таблицей"
            )

    auto_issues.extend(table_seq_issues)
    auto_issues.extend(table_issues)

    # --- СПИСОК ИСТОЧНИКОВ ---
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

    all_issues = auto_issues
    if manual_checks:
        all_issues.append("📋 Для проверки человеком:")
        all_issues.extend(manual_checks)
    return group_issues(all_issues)
# ------------------------------------------------------------
# Функции для тестовых режимов
# ------------------------------------------------------------
def run_subsections_test(file):
    doc = docx.Document(file)
    return analyze_subsections(doc)

def run_section_headers_test(file):
    doc = docx.Document(file)
    return analyze_section_headers(doc)

# ------------------------------------------------------------
# ИНТЕРФЕЙС STREAMLIT
# ------------------------------------------------------------
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")

# Создаем вкладки
tab1, tab2, tab3 = st.tabs(["📊 Полная проверка документа", "🔍 Тест: анализ подразделов", "🔍 Тест: анализ заголовков разделов"])

with tab1:
    st.title("📊 Автоматическая проверка документов Word")
    st.write("Загрузите документ в формате .docx – проверка по полному чек-листу.")
    uploaded_file = st.file_uploader("Выберите файл", type=["docx"], key="full_check")

    if uploaded_file is not None:
        with st.spinner("Проверяем..."):
            results = check_word_document(uploaded_file)
        st.subheader("Результаты проверки:")
        for r in results:
            if r.startswith("📋"):
                st.markdown(f"**{r}**")
            else:
                st.write(f"• {r}")

with tab2:
    st.title("🔍 Тест определения подразделов и пустых строк")
    st.markdown("Загрузите документ .docx – программа покажет, для каких подразделов ошибочно требуется убрать пустую строку.")
    
    test_file = st.file_uploader("Выберите файл .docx", type=["docx"], key="test_check")

    if test_file is not None:
        try:
            subsections = run_subsections_test(test_file)

            if not subsections:
                st.info("В документе не найдено подразделов с номерами вида 1.1, 1.2 и т.д.")
            else:
                st.subheader("Результаты анализа")
                st.write(f"Всего подразделов: {len(subsections)}")

                with_error = [s for s in subsections if s["error_should_show"]]
                without_error = [s for s in subsections if not s["error_should_show"]]

                if with_error:
                    st.error(f"❌ {len(with_error)} подраздел(ов), перед которыми пустая строка НЕ после раздела (нужно убрать):")
                    for s in with_error:
                        st.markdown(f"- **{s['text']}**")
                        st.caption(f"  Пустая строка перед: {s['has_empty_before']}, предыдущий непустой был разделом: {s['prev_was_section']}, технический: {s.get('is_technical_prev', False)} (текст: «{s['prev_text']}»)")
                else:
                    st.success("✅ Нет ошибочных требований убрать пустую строку перед подразделами.")

                if without_error:
                    st.info(f"ℹ️ {len(without_error)} подраздел(ов) с корректным расположением:")
                    for s in without_error[:10]:
                        st.markdown(f"- **{s['text']}**")
                        st.caption(f"  Пустая строка перед: {s['has_empty_before']}, предыдущий раздел: {s['prev_was_section']}")

        except Exception as e:
            st.error(f"Ошибка: {e}")
    else:
        st.info("Загрузите файл .docx для анализа.")

with tab3:
    st.title("🔍 Тест определения проблем заголовков разделов")
    st.markdown("Загрузите документ .docx – программа покажет фактические отступы и начало новой страницы для заголовков разделов.")
    
    test_file3 = st.file_uploader("Выберите файл .docx", type=["docx"], key="test_check3")

    if test_file3 is not None:
        try:
            headers = run_section_headers_test(test_file3)

            if not headers:
                st.info("В документе не найдено заголовков разделов.")
            else:
                st.subheader("Результаты анализа заголовков разделов")
                st.write(f"Всего заголовков разделов: {len(headers)}")

                for h in headers:
                    st.markdown(f"**{h['text']}**")
                    
                    if h['first_line_indent_ok']:
                        st.success(f"✅ Отступ первой строки: {h['first_line_indent_cm']:.2f} см (должен быть 0 см)")
                    else:
                        st.error(f"❌ Отступ первой строки: {h['first_line_indent_cm']:.2f} см (должен быть 0 см)")
                    
                    if h['text'].upper() == "ВВЕДЕНИЕ":
                        st.info(f"ℹ️ Для ВВЕДЕНИЯ проверка новой страницы не требуется")
                    elif h['starts_new_page_ok']:
                        st.success(f"✅ Начинается с новой страницы: {h['starts_new_page']}")
                    else:
                        st.error(f"❌ Начинается с новой страницы: {h['starts_new_page']} (должно быть True)")
                    
                    if h['has_empty_after_ok']:
                        st.success(f"✅ Пустая строка после заголовка: {h['has_empty_after']}")
                    else:
                        st.error(f"❌ Пустая строка после заголовка: {h['has_empty_after']} (должна быть)")
                    
                    if h['is_bold']:
                        st.success(f"✅ Полужирное начертание: Да")
                    else:
                        st.error(f"❌ Полужирное начертание: Нет")
                    
                    align_text = str(h['alignment']) if h['alignment'] else "None"
                    if h['alignment'] == WD_ALIGN_PARAGRAPH.CENTER:
                        st.success(f"✅ Выравнивание: по центру")
                    else:
                        st.error(f"❌ Выравнивание: {align_text} (должно быть по центру)")
                    
                    st.write("---")
                    
        except Exception as e:
            st.error(f"Ошибка: {e}")
    else:
        st.info("Загрузите файл .docx для анализа.")
