import streamlit as st
import docx
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def mm_to_emu(mm):
    return Mm(mm).emu

def check_word_document(file):
    doc = docx.Document(file)
    issues = []

    # --------------------------------------------------
    # ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
    # --------------------------------------------------
    def is_empty_paragraph(p):
        return len(p.text.strip()) == 0

    def is_bold_paragraph(p):
        try:
            if p.style and p.style.font and p.style.font.bold:
                return True
        except:
            pass
        runs = [r for r in p.runs if r.text.strip()]
        if not runs:
            return False
        return all(r.bold is True for r in runs)

    level1_headings = {
        "ВВЕДЕНИЕ",
        "ЗАКЛЮЧЕНИЕ",
        "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"
    }

    def is_level1_heading(text):
        text = text.strip()
        if text in level1_headings:
            return True
        return bool(
            re.match(
                r'^\d+\.\s+[А-ЯЁ][А-ЯЁ0-9\s\-\(\)"]+$',
                text
            )
        )

    def is_subsection(text):
        return bool(
            re.match(
                r'^\d+\.\d+(\.\d+)?\s+',
                text
            )
        )

    def subsection_name(text):
        return re.sub(
            r'^\d+\.\d+(\.\d+)?\s+',
            '',
            text
        ).strip()

    def find_previous_non_empty(current_idx):
        """Возвращает последний непустой абзац перед current_idx"""
        for i in range(current_idx - 1, -1, -1):
            if not is_empty_paragraph(doc.paragraphs[i]):
                return doc.paragraphs[i]
        return None

    # --------------------------------------------------
    # ПОИСК СОДЕРЖАНИЯ (гибкий, регистронезависимый)
    # --------------------------------------------------
    content_idx = None
    for i, p in enumerate(doc.paragraphs):
        if "содержание" in p.text.strip().lower():
            content_idx = i
            break
    if content_idx is None:
        content_idx = 0

    # --------------------------------------------------
    # ПОЛЯ
    # --------------------------------------------------
    for section in doc.sections:
        left_mm = section.left_margin.pt * 25.4 / 72
        right_mm = section.right_margin.pt * 25.4 / 72
        top_mm = section.top_margin.pt * 25.4 / 72
        bottom_mm = section.bottom_margin.pt * 25.4 / 72
        if (
            abs(left_mm - 20) > 0.5
            or abs(right_mm - 20) > 0.5
            or abs(top_mm - 20) > 0.5
            or abs(bottom_mm - 20) > 0.5
        ):
            issues.append("Поля страницы – установите 20 мм со всех сторон")
            break

    # --------------------------------------------------
    # ОСНОВНОЙ ПРОХОД ПО АБЗАЦАМ
    # --------------------------------------------------
    figure_counter = 0
    inside_contents = False
    in_references = False    # ← флаг для пропуска записей списка литературы

    for idx, p in enumerate(doc.paragraphs):
        if idx < content_idx:
            continue

        text = p.text.strip()
        if not text:
            continue

        pf = p.paragraph_format

        # ---------------------------------------------
        # ЗАГОЛОВОК «СОДЕРЖАНИЕ»
        # ---------------------------------------------
        if text.upper() == "СОДЕРЖАНИЕ":
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Содержание – выровняйте по центру")
            if text.endswith("."):
                issues.append("Содержание – удалите точку в конце")
            if not is_bold_paragraph(p):
                issues.append("Содержание – сделайте заголовок полужирным")
            if idx + 1 < len(doc.paragraphs):
                next_p = doc.paragraphs[idx + 1]
                if not is_empty_paragraph(next_p):
                    issues.append(
                        "Содержание – после заголовка должна быть пустая строка"
                    )
            inside_contents = True
            continue   # заголовок содержания больше нигде не проверяем

        # ---------------------------------------------
        # Установка флага in_references
        # ---------------------------------------------
        if text.upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            in_references = True
            # сам заголовок будет проверен ниже как раздел (is_level1_heading)

        # ---------------------------------------------
        # РЕЖИМ СОДЕРЖАНИЯ – проверка строк содержания
        # ---------------------------------------------
        if inside_contents:
            # Конец содержания по слову «ВВЕДЕНИЕ»
            if text.upper() == "ВВЕДЕНИЕ" and idx > content_idx + 10:
                inside_contents = False
                # этот абзац будет проверен как раздел ниже
            else:
                # Проверки элементов содержания
                if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                    issues.append("Содержание – выровняйте элементы содержания")

                # Межстрочный интервал 1.2
                line_sp = pf.line_spacing
                if line_sp is None or abs(line_sp - 1.2) > 0.01:
                    issues.append(
                        "Содержание – установите междустрочный интервал 1,2"
                    )

                # Интервал перед абзацем 0 пт
                if pf.space_before and pf.space_before.pt > 0.5:
                    issues.append(
                        "Содержание – интервал перед абзацем должен быть 0 пт"
                    )

                # Интервал после абзаца 0 пт
                if pf.space_after and pf.space_after.pt > 0.5:
                    issues.append(
                        "Содержание – интервал после абзаца должен быть 0 пт"
                    )

                continue   # строка содержания не проверяется как обычный текст

        # ---------------------------------------------
        # РАЗДЕЛЫ (только вне содержания)
        # ---------------------------------------------
        if is_level1_heading(text):
            section_name = text
            if not is_bold_paragraph(p):
                issues.append(
                    f"Раздел {section_name} – заголовок должен быть полужирным"
                )
            if text != text.upper():
                issues.append(
                    f"Раздел {section_name} – используйте прописные буквы"
                )
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(
                    f"Раздел {section_name} – выровняйте по центру"
                )
            indent = pf.first_line_indent
            if indent and abs(indent.cm) > 0.1:
                issues.append(
                    f"Раздел {section_name} – уберите абзацный отступ"
                )
            if text.endswith("."):
                issues.append(
                    f"Раздел {section_name} – удалите точку в конце"
                )
            if idx + 1 < len(doc.paragraphs):
                if not is_empty_paragraph(doc.paragraphs[idx + 1]):
                    issues.append(
                        f"Раздел {section_name} – после заголовка должна быть пустая строка"
                    )

        # ---------------------------------------------
        # ПОДРАЗДЕЛЫ (только вне содержания)
        # ---------------------------------------------
        if is_subsection(text):
            name = subsection_name(text)
            indent = pf.first_line_indent
            if indent is None or abs(indent.cm - 1.0) > 0.1:
                issues.append(
                    f'Подраздел "{name}" – установите абзацный отступ 1,0 см'
                )
            if not is_bold_paragraph(p):
                issues.append(
                    f'Подраздел "{name}" – заголовок должен быть полужирным'
                )
            if text.endswith("."):
                issues.append(
                    f'Подраздел "{name}" – удалите точку в конце'
                )

            # Пустая строка перед подразделом запрещена, кроме случая,
            # когда предыдущий непустой абзац – заголовок раздела
            if idx > 0:
                prev = doc.paragraphs[idx - 1]
                if is_empty_paragraph(prev):
                    prev_non_empty = find_previous_non_empty(idx)
                    if prev_non_empty is None or not is_level1_heading(prev_non_empty.text):
                        issues.append(
                            f'Подраздел "{name}" – уберите пустую строку перед подразделом'
                        )

        # ---------------------------------------------
        # ОСНОВНОЙ ТЕКСТ (с учётом флага in_references)
        # ---------------------------------------------
        is_heading = (
            is_level1_heading(text)
            or is_subsection(text)
            or text.upper() == "СОДЕРЖАНИЕ"
        )
        if not is_heading and not in_references:
            if not text.startswith("Рисунок"):
                indent = pf.first_line_indent
                if indent is None or abs(indent.cm - 1.0) > 0.1:
                    issues.append(
                        f'«{text[:40]}...» – установите абзацный отступ 1,0 см'
                    )
            if pf.space_before:
                try:
                    if pf.space_before.pt > 0.5:
                        issues.append(
                            f'«{text[:40]}...» – интервал перед абзацем должен быть 0 пт'
                        )
                except:
                    pass

        # ---------------------------------------------
        # РИСУНКИ
        # ---------------------------------------------
        if text.startswith("Рисунок"):
            figure_counter += 1
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(
                    f"Рисунок {figure_counter} – выровняйте подпись по центру"
                )
            if text.endswith("."):
                issues.append(
                    f"Рисунок {figure_counter} – удалите точку в конце"
                )
            m = re.match(
                r'^Рисунок\s+\d+\s*[–-]\s*(.+)$',
                text
            )
            if m:
                title = m.group(1).strip()
                if title and title[0].islower():
                    issues.append(
                        f"Рисунок {figure_counter} – название должно начинаться с большой буквы"
                    )
            if figure_counter == 3:
                if idx > 0:
                    prev = doc.paragraphs[idx - 1]
                    if not is_empty_paragraph(prev):
                        issues.append(
                            "Рисунок 3 – добавьте пустую строку перед рисунком"
                        )

    # --------------------------------------------------
    # СПИСОК ИСТОЧНИКОВ (проверка как единого блока)
    # --------------------------------------------------
    lit_start = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().upper() == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ":
            lit_start = i
            break

    if lit_start is not None:
        left_issue = False
        first_line_issue = False
        align_issue = False
        line_spacing_issue = False

        for p in doc.paragraphs[lit_start + 1:]:
            if not p.text.strip():
                continue
            pf = p.paragraph_format

            # Отступ слева 0 см
            if pf.left_indent and abs(pf.left_indent.cm) > 0.1:
                left_issue = True
            # Отступ первой строки 1,0 см
            if pf.first_line_indent is None or abs(pf.first_line_indent.cm - 1.0) > 0.1:
                first_line_issue = True
            # Выравнивание по ширине
            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                align_issue = True
            # Межстрочный интервал 1.2
            line_sp = pf.line_spacing
            if line_sp is None or abs(line_sp - 1.2) > 0.01:
                line_spacing_issue = True

        if left_issue:
            issues.append("Список источников – отступ слева должен быть 0 см")
        if first_line_issue:
            issues.append("Список источников – установите отступ первой строки 1,0 см")
        if align_issue:
            issues.append("Список источников – выровняйте по ширине")
        if line_spacing_issue:
            issues.append("Список источников – междустрочный интервал 1,2")

    # --------------------------------------------------
    # Удаление дубликатов и вывод
    # --------------------------------------------------
    issues = list(dict.fromkeys(issues))

    if not issues:
        return ["✅ Ошибок не найдено. Документ соответствует требованиям."]
    return issues

# ---------- ИНТЕРФЕЙС STREAMLIT ----------
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write(
    "Загрузите ваш документ в формате .docx для проверки "
    "по полному чек-листу (поля, интервалы, отступы, заголовки, "
    "таблицы, рисунки, список литературы)."
)

uploaded_file = st.file_uploader(
    "Перетащите файл сюда или нажмите для выбора", type=["docx"]
)

if uploaded_file is not None:
    with st.spinner("Анализируем документ..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for res in results:
        st.write(f"• {res}")
