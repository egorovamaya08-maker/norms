
import streamlit as st
import docx
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
import re

def mm_to_emu(mm):
    return Mm(mm).emu

def check_word_document(file):
    import docx
    import re
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    doc = docx.Document(file)
    issues = []

    # --------------------------------------------------
    # ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
    # --------------------------------------------------

    def is_empty_paragraph(p):
        return len(p.text.strip()) == 0

    def is_bold_paragraph(p):

        try:
            if (
                p.style
                and p.style.font
                and p.style.font.bold
            ):
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

    # --------------------------------------------------
    # ПОИСК СОДЕРЖАНИЯ
    # --------------------------------------------------

    content_idx = None

    for i, p in enumerate(doc.paragraphs):

        if p.text.strip().upper() == "СОДЕРЖАНИЕ":
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
            issues.append(
                "Поля страницы – установите 20 мм со всех сторон"
            )
            break

    # --------------------------------------------------
    # ОСНОВНОЙ ПРОХОД ПО АБЗАЦАМ
    # --------------------------------------------------

    figure_counter = 0

    for idx, p in enumerate(doc.paragraphs):

        if idx < content_idx:
            continue

        text = p.text.strip()

        if not text:
            continue

        pf = p.paragraph_format

        # ---------------------------------------------
        # СОДЕРЖАНИЕ
        # ---------------------------------------------

        if text.upper() == "СОДЕРЖАНИЕ":

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(
                    "Содержание – выровняйте по центру"
                )

            if text.endswith("."):
                issues.append(
                    "Содержание – удалите точку в конце"
                )

            if not is_bold_paragraph(p):
                issues.append(
                    "Содержание – сделайте заголовок полужирным"
                )

            if idx + 1 < len(doc.paragraphs):

                next_p = doc.paragraphs[idx + 1]

                if not is_empty_paragraph(next_p):
                    issues.append(
                        "Содержание – после заголовка должна быть пустая строка"
                    )

        # ---------------------------------------------
        # РАЗДЕЛЫ
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

                if not is_empty_paragraph(
                    doc.paragraphs[idx + 1]
                ):
                    issues.append(
                        f"Раздел {section_name} – после заголовка должна быть пустая строка"
                    )

        # ---------------------------------------------
        # ПОДРАЗДЕЛЫ
        # ---------------------------------------------

        if is_subsection(text):

            name = subsection_name(text)

            indent = pf.first_line_indent

            if (
                indent is None
                or abs(indent.cm - 1.0) > 0.1
            ):
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

            if idx > 0:

                prev = doc.paragraphs[idx - 1]

                if is_empty_paragraph(prev):
                    issues.append(
                        f'Подраздел "{name}" – уберите пустую строку перед подразделом'
                    )

        # ---------------------------------------------
        # ОСНОВНОЙ ТЕКСТ
        # ---------------------------------------------

        is_heading = (
            is_level1_heading(text)
            or is_subsection(text)
            or text.upper() == "СОДЕРЖАНИЕ"
        )

        if not is_heading:

            if not text.startswith("Рисунок"):
                indent = pf.first_line_indent

                if (
                    indent is None
                    or abs(indent.cm - 1.0) > 0.1
                ):
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

                if (
                    title
                    and title[0].islower()
                ):
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
    # СПИСОК ИСТОЧНИКОВ
    # --------------------------------------------------

    lit_start = None

    for i, p in enumerate(doc.paragraphs):

        if (
            p.text.strip().upper()
            == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"
        ):
            lit_start = i
            break

    if lit_start:

        for p in doc.paragraphs[lit_start + 1:]:

            if not p.text.strip():
                continue

            pf = p.paragraph_format

            if (
                pf.left_indent
                and abs(pf.left_indent.cm) > 0.1
            ):
                issues.append(
                    "Список источников – отступ слева должен быть 0 см"
                )
                break

            if (
                pf.first_line_indent is None
                or abs(
                    pf.first_line_indent.cm - 1.0
                ) > 0.1
            ):
                issues.append(
                    "Список источников – установите отступ первой строки 1,0 см"
                )
                break

            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                issues.append(
                    "Список источников – выровняйте по ширине"
                )
                break

    issues = list(dict.fromkeys(issues))

    if not issues:
        return [
            "✅ Ошибок не найдено. Документ соответствует требованиям."
        ]

    return issues



# ---------- ИНТЕРФЕЙС STREAMLIT ----------
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write("Загрузите ваш документ в формате .docx для проверки по полному чек-листу (поля, интервалы, отступы, заголовки, таблицы, рисунки, список литературы).")

uploaded_file = st.file_uploader("Перетащите файл сюда или нажмите для выбора", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем документ..."):
        results = check_word_document(uploaded_file)
    st.subheader("Результаты проверки:")
    for res in results:
        st.write(f"• {res}")
