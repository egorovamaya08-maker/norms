import streamlit as st
import docx
import re

st.set_page_config(page_title="Проверка подразделов", layout="centered")
st.title("Проверка подразделов: лишняя пустая строка")

def is_section_header(text):
    """Заголовок раздела первого уровня"""
    cleaned = text.strip()
    if not cleaned:
        return False
    upper = cleaned.upper()
    if upper in ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]:
        return True
    if re.match(r'^(ГЛАВА|РАЗДЕЛ)\s+\d+', upper):
        return True
    if re.match(r'^\d+\.\s*[А-ЯЁ]', cleaned):
        return True
    # Заголовок без номера, но все заглавные буквы (длина > 3)
    letters = re.sub(r'[\d\s\.,;:!?\-–—()«»""\'\']', '', cleaned)
    if letters and letters == letters.upper() and len(letters) > 3:
        return True
    return False

def is_subsection(text):
    """Подраздел вида 1.1, 1.1.1"""
    cleaned = text.strip()
    if not cleaned:
        return False
    return bool(re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]', cleaned))

def starts_new_page(paragraph, doc):
    """Проверяет, начинается ли абзац с новой страницы"""
    # Получаем элемент абзаца в XML
    elem = paragraph._element
    # Ищем предыдущие элементы
    body = doc.element.body
    elements = list(body)
    try:
        idx = elements.index(elem)
    except ValueError:
        return False
    # Идём назад до начала документа или до предыдущего разрыва страницы
    for i in range(idx - 1, -1, -1):
        prev = elements[i]
        # Если встретили разрыв раздела (sectPr) с типом не continuous
        if prev.tag.endswith('sectPr'):
            for child in prev:
                if child.tag.endswith('type') and child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') != 'continuous':
                    return True
            continue
        # Если встретили разрыв страницы <w:br w:type="page"/>
        if prev.tag.endswith('p'):
            for br in prev.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br'):
                if br.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page':
                    return True
            # Если есть текст или рисунок, значит это не просто пустые строки
            texts = prev.xpath('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t/text()')
            if any(t.strip() for t in texts):
                return False
            # Если это пустая строка, продолжаем
    return False

def analyze(doc):
    results = []
    prev_empty = False
    prev_was_section = False
    prev_text = ""

    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            prev_empty = True
            continue

        is_section = is_section_header(text)
        is_sub = is_subsection(text)

        if is_sub:
            on_new_page = starts_new_page(para, doc)
            # Ошибка: есть пустая строка, предыдущий не раздел, и подраздел не на новой странице
            has_error = prev_empty and (not prev_was_section) and (not on_new_page)
            results.append({
                "text": text,
                "index": idx,
                "prev_empty": prev_empty,
                "prev_was_section": prev_was_section,
                "prev_text": prev_text,
                "on_new_page": on_new_page,
                "error": has_error
            })

        # Обновляем состояние
        if is_section:
            prev_was_section = True
        else:
            prev_was_section = False

        prev_text = text
        prev_empty = False

    return results

uploaded = st.file_uploader("Загрузите файл .docx", type=["docx"])

if uploaded is not None:
    try:
        doc = docx.Document(uploaded)
        subsections = analyze(doc)

        if not subsections:
            st.info("Подразделы не найдены.")
        else:
            errors = [s for s in subsections if s["error"]]
            if errors:
                st.error(f"Найдено {len(errors)} подраздел(ов) с лишней пустой строкой:")
                for e in errors:
                    st.markdown(f"- **{e['text']}**")
                    st.caption(f"  Пустая строка перед: {e['prev_empty']}, предыдущий раздел: {e['prev_was_section']}, новая страница: {e['on_new_page']}")
            else:
                st.success("Все подразделы оформлены верно.")

            with st.expander("Детали по всем подразделам"):
                for s in subsections:
                    status = "ОШИБКА" if s["error"] else "ОК"
                    st.write(f"{status}: {s['text']}")
                    st.write(f"  - пустая строка перед: {s['prev_empty']}")
                    st.write(f"  - предыдущий непустой был разделом: {s['prev_was_section']}")
                    st.write(f"  - предыдущий текст: {s['prev_text'][:60]}")
                    st.write(f"  - начинается с новой страницы: {s['on_new_page']}")
                    st.write("---")
    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")
else:
    st.info("Загрузите документ.")
