import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import re
from collections import defaultdict, Counter
import zipfile
from lxml import etree

NSMAP = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# ------------------------------------------------------------
# Вспомогательные функции
# ------------------------------------------------------------
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
    if text.upper().strip() in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}:
        return True
    if re.match(r'^(?:ГЛАВА|РАЗДЕЛ)\s+\d+', text, re.IGNORECASE):
        return True

    # Нормализация пробельных символов (неразрывный, тонкий и т.п.)
    normalized = text.replace('\u00A0', ' ').replace('\u202F', ' ').replace('\u2009', ' ')

    if re.match(r'^\d+\.\s+[А-ЯЁ]', normalized):
        # Очистку от спецсимволов делаем по исходному text,
        # чтобы не потерять буквы, которые могут быть в оригинале
        clean = re.sub(r'[\d\s\.,;:!?\-–—()«»""''«»]', '', text)
        if not clean:
            return False
        upper_count = sum(1 for c in clean if c.isupper())
        return upper_count >= len(clean) * 0.8
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
    if code in [0xF02D]:   # PUA-тире
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

def is_on_new_page(doc, body_idx, start_body_pos=0, min_empty_paragraphs=10):
    body_elems = list(doc.element.body)
    if body_idx == start_body_pos:
        return True

    blank_count = 0
    for i in range(body_idx - 1, start_body_pos - 1, -1):
        elem = body_elems[i]

        if elem.tag == qn('w:sectPr'):
            if i == len(body_elems) - 1:
                continue
            type_el = elem.find(qn('w:type'))
            val = type_el.get(qn('w:val')) if type_el is not None else None
            if val == 'continuous':
                continue
            return True

        if elem.tag == qn('w:p'):
            for br in elem.findall('.//w:br', NSMAP):
                if br.get(qn('w:type')) == 'page':
                    return True

            pPr = elem.find(qn('w:pPr'))
            if pPr is not None:
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

def has_content(elem):
    texts = [node.text or '' for node in elem.iter() if node.tag == qn('w:t')]
    if any(t.strip() for t in texts):
        return True
    if elem.find('.//w:drawing', NSMAP) is not None:
        return True
    return False

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
        if re.match(r'^\d+\.\s+[А-Яа-я]', txt.replace('\u00A0', ' ')) and not is_section_header(txt.replace('\u00A0', ' ')):
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
                texts = []
                for t in elem.findall('.//w:t', qn('w')):
                    if t.text:
                        texts.append(t.text)
                full_text = ''.join(texts).strip().upper()
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
    prev_was_section_header = False
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

        if re.match(r'^\d+\.\s+[А-Яа-я]', text) and not is_section_header(text):
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

        if is_formula_where_line(text):
            errors = check_formula_explanation(text, p, prev_was_formula, prev_para_empty)
            auto_issues.extend(errors)
            prev_para_empty = False
            continue

        if is_formula_or_equation(text):
            prev_was_formula = True
            prev_para_empty = False
            continue

        is_level1 = is_section_header(text)
        is_subsection = False
        is_figure = text.startswith("Рисунок") or text.startswith("Рис.")
        is_table_caption = text.startswith("Таблица")

        if not is_level1:
            if re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]', text.replace('\u00A0', ' ').replace('\u202F', ' ')):
                is_subsection = True
            else:
                normalized = normalize_title(text)
                in_toc = any(normalize_title(e) == normalized for e in toc_entries) if toc_entries else False
                if in_toc and len(text) > 20:
                    is_subsection = True

        if is_level1:
            key = text[:80]
            if text.upper() != "ВВЕДЕНИЕ":
                body_idx = para_to_body_idx.get(idx)
                if body_idx is not None:
                    if not is_on_new_page(doc, body_idx, start_body_pos):
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

        elif is_subsection:
            sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s+', '', text).strip()
            key = f"Подраздел «{sub_name[:50]}»"
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                auto_issues.append(f"{key} – установите абзацный отступ 1,0 см (сейчас {first_line:.1f} см)")
            if not is_paragraph_bold(p):
                auto_issues.append(f"{key} – заголовок должен быть полужирным")
            if alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                auto_issues.append(f"{key} – выровняйте по ширине")
            if text.endswith("."):
                auto_issues.append(f"{key} – удалите точку в конце")
            # Пустая строка допустима только после заголовка раздела
            if prev_para_empty and not prev_was_section_header:
                auto_issues.append(f"{key} – уберите пустую строку перед подразделом")

        elif is_figure:
            num_match = re.search(r'(?:Рисунок|Рис\.)\s*:?\s*(\d+(?:\.\d+)?)', text)
            if num_match:
                fig_num = num_match.group(1)
                figure_numbers_found.append(float(fig_num))
            else:
                figure_counter += 1
                fig_num = str(figure_counter)
            fig_number = f"Рисунок {fig_num}"

            if text.startswith("Рис."):
                auto_issues.append(f"{fig_number} – измените «Рис.» на «Рисунок»")
            if re.match(r'Рисунок\s*:', text):
                auto_issues.append(f"{fig_number} – замените двоеточие на тире (формат: «Рисунок N — Название»)")
            if not re.search(r'Рисунок\s+\d+(?:\.\d+)?\s*[–—]', text) and not re.match(r'Рисунок\s*:', text):
                if re.search(r'Рисунок\s+\d+(?:\.\d+)?\s*[-]', text):
                    auto_issues.append(f"{fig_number} – замените дефис на тире (—)")
                else:
                    auto_issues.append(f"{fig_number} – должно быть тире после номера")
            if alignment != WD_ALIGN_PARAGRAPH.CENTER:
                auto_issues.append(f"{fig_number} – выровняйте подпись по центру")
            if text.endswith(".") and not re.search(r'\([^)]*\)\.$', text):
                auto_issues.append(f"{fig_number} – удалите точку в конце")
            m = re.match(r'^(?:Рисунок|Рис\.)\s+\d+(?:\.\d+)?\s*[–—]\s*(.+)$', text)
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

        elif is_table_caption and not is_table_continuation(text):
            tbl_match = re.match(r'Таблица\s+(\d+(?:\.\d+)?)', text)
            if tbl_match:
                tbl_num = tbl_match.group(1)
                tbl_num_float = float(tbl_num)
                table_numbers_found.append(tbl_num_float)
                key = f"Таблица {tbl_num}"

                if not re.match(r'Таблица\s+\d+(?:\.\d+)?\s+[–—]\s+\S', text):
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

        else:
            key = text[:50]
            first_line = get_effective_first_line_indent(p)
            if abs(first_line - 1.0) > 0.2:
                indent_issues.append((key, first_line))
            if pf.space_before and pf.space_before.pt > 0.5:
                auto_issues.append(f"«{key}» – интервал перед абзацем должен быть 0 пт")

        # Обновление флага предыдущего раздела (только для непустых параграфов)
        if text:
            if is_level1:
                prev_was_section_header = True
            else:
                prev_was_section_header = False
        prev_para_empty = False
        prev_was_formula = False

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

# --- ИНТЕРФЕЙС STREAMLIT ---
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
