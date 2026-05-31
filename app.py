def check_word_document(file):
    import docx
    import re
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = docx.Document(file)
    issues = []

    # -----------------------------
    # СТАТУСЫ ДОКУМЕНТА
    # -----------------------------
    OUTSIDE = 0
    CONTENTS = 1
    BODY = 2
    REFERENCES = 3

    state = OUTSIDE

    figure_counter = 0

    # -----------------------------
    # ВСПОМОГАТЕЛЬНЫЕ
    # -----------------------------

    def is_empty(p):
        return not p.text.strip()

    def is_heading_lvl1(text):
        return bool(re.match(r'^\d+\.\s+[А-ЯЁ].*$', text.strip()))

    def is_subsection(text):
        return bool(re.match(r'^\d+\.\d+(\.\d+)?\s+.+$', text.strip()))

    def is_refs_title(text):
        return text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"

    def is_toc_title(text):
        return text.strip().upper() == "СОДЕРЖАНИЕ"

    # -----------------------------
    # ОСНОВНОЙ ЦИКЛ
    # -----------------------------

    for idx, p in enumerate(doc.paragraphs):

        text = p.text.strip()
        pf = p.paragraph_format

        # ---------------------------------
        # ПЕРЕХОДЫ СОСТОЯНИЙ (СТРОГО)
        # ---------------------------------

        if state == OUTSIDE and is_toc_title(text):
            state = CONTENTS
            continue

        if state == CONTENTS and is_refs_title(text):
            state = BODY
            continue

        if state == BODY and is_refs_title(text):
            state = REFERENCES
            continue

        # ---------------------------------
        # СОДЕРЖАНИЕ (ТОЛЬКО СТИЛЬ)
        # ---------------------------------

        if state == CONTENTS:

            # внутри содержания запрещены любые структурные проверки
            if p.alignment != WD_ALIGN_PARAGRAPH.LEFT:
                issues.append("Содержание – выравнивание по левому краю")

            if pf.space_before and pf.space_before.pt > 0:
                issues.append("Содержание – интервал перед абзацем должен быть 0 пт")

            if pf.line_spacing and abs(pf.line_spacing - 1.2) > 0.01:
                issues.append("Содержание – межстрочный интервал 1.2")

            continue

        # ---------------------------------
        # РАЗДЕЛЫ (ТОЛЬКО BODY)
        # ---------------------------------

        if state == BODY and is_heading_lvl1(text):

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Раздел {text} – центрирование")

            if not text.isupper():
                issues.append(f"Раздел {text} – только ПРОПИСНЫЕ буквы")

            if pf.first_line_indent and abs(pf.first_line_indent.cm) > 0.01:
                issues.append(f"Раздел {text} – убрать отступ")

            if text.endswith("."):
                issues.append(f"Раздел {text} – убрать точку")

            # строгая проверка пустой строки после заголовка
            if idx + 1 < len(doc.paragraphs):
                if not is_empty(doc.paragraphs[idx + 1]):
                    issues.append(f"Раздел {text} – пустая строка после заголовка")

        # ---------------------------------
        # ПОДРАЗДЕЛЫ (ТОЛЬКО BODY)
        # ---------------------------------

        if state == BODY and is_subsection(text):

            name = text.split(" ", 1)[1] if " " in text else text

            if not pf.first_line_indent or abs(pf.first_line_indent.cm - 1.0) > 0.01:
                issues.append(f"Подраздел {name} – отступ 1.0 см")

            if text.endswith("."):
                issues.append(f"Подраздел {name} – убрать точку")

            # СТРОГАЯ проверка: только предыдущий абзац
            if idx > 0:
                prev = doc.paragraphs[idx - 1]
                if is_empty(prev):
                    issues.append(f"Подраздел {name} – не допускается пустая строка перед заголовком")

        # ---------------------------------
        # ОСНОВНОЙ ТЕКСТ
        # ---------------------------------

        is_structure = (
            is_heading_lvl1(text)
            or is_subsection(text)
            or is_toc_title(text)
        )

        if state == BODY and not is_structure:

            if pf.first_line_indent and abs(pf.first_line_indent.cm - 1.0) > 0.01:
                issues.append(f"Текст – отступ 1.0 см ({text[:30]})")

            if pf.space_before and pf.space_before.pt > 0:
                issues.append(f"Текст – интервал перед абзацем должен быть 0")

        # ---------------------------------
        # РИСУНКИ
        # ---------------------------------

        if text.startswith("Рисунок"):

            figure_counter += 1

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_counter} – центрирование")

            if text.endswith("."):
                issues.append(f"Рисунок {figure_counter} – убрать точку")

            m = re.match(r'^Рисунок\s+\d+\s*[–-]\s*(.+)$', text)

            if m:
                if m.group(1) and m.group(1)[0].islower():
                    issues.append(f"Рисунок {figure_counter} – заглавная буква")

            if figure_counter == 3:
                if idx > 0 and not is_empty(doc.paragraphs[idx - 1]):
                    issues.append("Рисунок 3 – пустая строка перед рисунком")

        # ---------------------------------
        # СПИСОК ИСТОЧНИКОВ (СТРОГО ОДИН БЛОК)
        # ---------------------------------

        if state == REFERENCES:

            # проверяем только первый элемент списка
            refs = [
                p for p in doc.paragraphs[idx:]
                if p.text.strip()
            ]

            if refs:

                first = refs[0]
                pf2 = first.paragraph_format

                if pf2.left_indent and abs(pf2.left_indent.cm) > 0.01:
                    issues.append("Источники – отступ слева 0 см")

                if not pf2.first_line_indent or abs(pf2.first_line_indent.cm - 1.0) > 0.01:
                    issues.append("Источники – отступ первой строки 1.0 см")

                if first.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                    issues.append("Источники – выравнивание по ширине")

                if pf2.line_spacing and abs(pf2.line_spacing - 1.2) > 0.01:
                    issues.append("Источники – межстрочный интервал 1.2")

            break

    # -----------------------------
    # РЕЗУЛЬТАТ
    # -----------------------------

    issues = list(dict.fromkeys(issues))

    if not issues:
        return ["OK"]

    return issues
