def check_word_document(file):
    import docx
    import re
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = docx.Document(file)
    issues = []

    # -----------------------------
    # SAFE HELPERS
    # -----------------------------

    def is_empty(p):
        return not p.text or not p.text.strip()

    def safe_cm(val):
        try:
            if val is None:
                return None
            return val.cm
        except:
            return None

    def is_bold_paragraph(p):
        try:
            runs = [r for r in p.runs if r.text.strip()]
            if not runs:
                return False
            return all(r.bold is True for r in runs)
        except:
            return False

    def safe_space_before(pf):
        try:
            return pf.space_before.pt if pf.space_before else 0
        except:
            return 0

    def safe_line_spacing(pf):
        try:
            return pf.line_spacing
        except:
            return None

    # -----------------------------
    # RULES
    # -----------------------------

    level1_headings = {
        "ВВЕДЕНИЕ",
        "ЗАКЛЮЧЕНИЕ",
        "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"
    }

    def is_level1_heading(text):
        text = text.strip()
        if text in level1_headings:
            return True
        return bool(re.match(r'^\d+\.\s+[А-ЯЁ].*$', text))

    def is_subsection(text):
        return bool(re.match(r'^\d+\.\d+(\.\d+)?\s+.+$', text))

    def is_toc(text):
        return text.strip().upper() == "СОДЕРЖАНИЕ"

    def is_refs(text):
        return text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"

    # -----------------------------
    # FIND CONTENT START
    # -----------------------------

    content_idx = 0
    for i, p in enumerate(doc.paragraphs):
        if is_toc(p.text):
            content_idx = i
            break

    inside_contents = False
    figure_counter = 0

    # -----------------------------
    # MAIN LOOP
    # -----------------------------

    for idx, p in enumerate(doc.paragraphs):

        if idx < content_idx:
            continue

        text = p.text.strip()
        if not text:
            continue

        pf = p.paragraph_format

        # -------------------------
        # CONTENT HEADER
        # -------------------------

        if is_toc(text):
            inside_contents = True

            try:
                if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                    issues.append("Содержание – выравнивание по центру")
            except:
                pass

            if not is_bold_paragraph(p):
                issues.append("Содержание – заголовок должен быть полужирным")

            continue

        # -------------------------
        # CONTENT BLOCK (STRICT STYLE ONLY)
        # -------------------------

        if inside_contents:

            try:
                if p.alignment != WD_ALIGN_PARAGRAPH.LEFT:
                    issues.append("Содержание – выравнивание элементов")
            except:
                pass

            if safe_space_before(pf) > 0:
                issues.append("Содержание – интервал перед абзацем должен быть 0 пт")

            if safe_line_spacing(pf) and abs(safe_line_spacing(pf) - 1.2) > 0.05:
                issues.append("Содержание – межстрочный интервал 1.2")

            # НЕ используем эвристики выхода
            continue

        # -------------------------
        # LEVEL 1 HEADINGS
        # -------------------------

        if is_level1_heading(text):

            if not is_bold_paragraph(p):
                issues.append(f"Раздел {text} – должен быть полужирным")

            if text != text.upper():
                issues.append(f"Раздел {text} – только ПРОПИСНЫЕ буквы")

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Раздел {text} – выравнивание по центру")

            indent = safe_cm(pf.first_line_indent)

            if indent is not None and abs(indent) > 0.1:
                issues.append(f"Раздел {text} – убрать отступ")

            if text.endswith("."):
                issues.append(f"Раздел {text} – убрать точку")

            if idx + 1 < len(doc.paragraphs):
                if not is_empty(doc.paragraphs[idx + 1]):
                    issues.append(f"Раздел {text} – пустая строка после заголовка")

        # -------------------------
        # SUBSECTIONS
        # -------------------------

        if is_subsection(text):

            name = text.split(" ", 1)[-1]

            indent = safe_cm(pf.first_line_indent)

            if indent is None or abs(indent - 1.0) > 0.1:
                issues.append(f'Подраздел "{name}" – отступ 1.0 см')

            if not is_bold_paragraph(p):
                issues.append(f'Подраздел "{name}" – полужирный заголовок')

            if text.endswith("."):
                issues.append(f'Подраздел "{name}" – убрать точку')

            # SAFE prev check
            if idx > 0:
                prev = doc.paragraphs[idx - 1]
                if is_empty(prev):
                    issues.append(f'Подраздел "{name}" – нельзя пустую строку перед заголовком')

        # -------------------------
        # BODY TEXT
        # -------------------------

        is_structure = (
            is_level1_heading(text)
            or is_subsection(text)
            or is_toc(text)
            or is_refs(text)
        )

        if not is_structure:

            indent = safe_cm(pf.first_line_indent)
            if indent is None or abs(indent - 1.0) > 0.1:
                issues.append(f'Текст – отступ 1.0 см ({text[:30]})')

            if safe_space_before(pf) > 0:
                issues.append(f'Текст – интервал перед абзацем должен быть 0')

        # -------------------------
        # FIGURES
        # -------------------------

        if text.startswith("Рисунок"):

            figure_counter += 1

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – центрирование")

            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – убрать точку")

            m = re.match(r'^Рисунок\s+\d+\s*[–-]\s*(.+)$', text)

            if m and m.group(1):
                if m.group(1)[0].islower():
                    issues.append(f"Рисунок {figure_counter} – заглавная буква")

            if figure_counter == 3:
                if idx > 0 and not is_empty(doc.paragraphs[idx - 1]):
                    issues.append("Рисунок 3 – пустая строка перед рисунком")

    # -----------------------------
    # REFERENCES (SAFE BLOCK)
    # -----------------------------

    lit_start = None

    for i, p in enumerate(doc.paragraphs):
        if is_refs(p.text):
            lit_start = i
            break

    if lit_start is not None:

        refs = [
            p for p in doc.paragraphs[lit_start + 1:]
            if p.text and p.text.strip()
        ]

        if refs:

            first = refs[0]
            pf = first.paragraph_format

            if safe_cm(pf.left_indent) and abs(pf.left_indent.cm) > 0.1:
                issues.append("Источники – отступ слева 0 см")

            fi = safe_cm(pf.first_line_indent)
            if fi is None or abs(fi - 1.0) > 0.1:
                issues.append("Источники – отступ первой строки 1.0 см")

            if first.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append("Источники – выравнивание по ширине")

            if safe_line_spacing(pf) and abs(safe_line_spacing(pf) - 1.2) > 0.05:
                issues.append("Источники – межстрочный интервал 1.2")

    # -----------------------------
    # RESULT
    # -----------------------------

    issues = list(dict.fromkeys(issues))

    return issues or ["OK"]
