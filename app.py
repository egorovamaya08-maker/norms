import streamlit as st
import docx
import re

st.set_page_config(page_title="Проверка подразделов", layout="centered")
st.title("🔍 Проверка подразделов: лишняя пустая строка")

def is_section_header(text):
    """Заголовок раздела первого уровня"""
    cleaned = text.strip()
    if not cleaned:
        return False
    upper_cleaned = cleaned.upper()
    if upper_cleaned in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}:
        return True
    if re.match(r'^(ГЛАВА|РАЗДЕЛ)\s+\d+', upper_cleaned):
        return True
    if re.match(r'^\d+\.\s*[А-ЯЁ]', cleaned):
        return True
    # Заголовок без номера, но все заглавные буквы
    only_letters = re.sub(r'[\d\s\.,;:!?\-–—()«»""''«»]', '', cleaned)
    if only_letters and only_letters == only_letters.upper() and len(only_letters) > 3:
        return True
    return False

def is_subsection(text):
    """Подраздел вида 1.1, 1.1.1 и т.д."""
    cleaned = text.strip()
    if not cleaned:
        return False
    return bool(re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]', cleaned))

def get_body_elements(doc):
    """Возвращает список body-элементов документа для определения начала страницы"""
    return list(doc.element.body)

def is_on_new_page(doc, body_idx, start_body_pos=0):
    """Проверяет, начинается ли элемент (параграф) с новой страницы"""
    body_elems = get_body_elements(doc)
    if body_idx <= start_body_pos:
        return False
    blank_count = 0
    for i in range(body_idx - 1, start_body_pos - 1, -1):
        elem = body_elems[i]
        # Если встретили разрыв раздела (sectPr) не continuous — новая страница
        if elem.tag.endswith('sectPr'):
            type_elem = elem.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
            if type_elem is not None and type_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') != 'continuous':
                return True
            continue
        # Если встретили разрыв страницы <w:br w:type="page"/>
        if elem.tag.endswith('p'):
            for br in elem.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br'):
                if br.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page':
                    return True
            # Если есть содержимое и много пустых строк — считаем, что это начало новой страницы
            if has_content(elem):
                return blank_count >= 10
            else:
                blank_count += 1
                continue
        if elem.tag.endswith('tbl'):
            return blank_count >= 10
    return False

def has_content(elem):
    """Проверяет, есть ли у элемента текстовое содержимое или рисунок"""
    texts = elem.xpath('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t/text()')
    if any(t.strip() for t in texts):
        return True
    if elem.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing') is not None:
        return True
    return False

def analyze_subsections(doc):
    results = []
    prev_para_empty = False
    prev_nonempty_was_section = False
    prev_nonempty_text = ""
    para_to_body_idx = {}
    body_elems = list(doc.element.body)
    # Составим отображение индекса параграфа -> индекс в body_elems
    for i, elem in enumerate(body_elems):
        if elem.tag.endswith('p'):
            for j, p in enumerate(doc.paragraphs):
                if p._element is elem:
                    para_to_body_idx[j] = i
                    break
    start_body_pos = 0
    # Найдём первый заголовок раздела, чтобы определить начало основного текста (опционально)
    for i, p in enumerate(doc.paragraphs):
        if is_section_header(p.text.strip()):
            start_body_pos = para_to_body_idx.get(i, 0)
            break

    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            prev_para_empty = True
            continue

        is_section = is_section_header(text)
        is_sub = is_subsection(text)

        if is_sub:
            body_idx = para_to_body_idx.get(idx, -1)
            starts_new_page = is_on_new_page(doc, body_idx, start_body_pos) if body_idx != -1 else False
            error_condition = prev_para_empty and not prev_nonempty_was_section and not starts_new_page

            results.append({
                "text": text[:80],
                "index": idx,
                "has_empty_before": prev_para_empty,
                "prev_was_section": prev_nonempty_was_section,
                "prev_text": prev_nonempty_text[:60] if prev_nonempty_text else "—",
                "starts_new_page": starts_new_page,
                "error": error_condition
            })

        # Обновляем состояние для следующего абзаца
        if is_section:
            prev_nonempty_was_section = True
        else:
            prev_nonempty_was_section = False

        prev_nonempty_text = text
        prev_para_empty = False

    return results

uploaded_file = st.file_uploader("Загрузите документ .docx", type=["docx"])

if uploaded_file is not None:
    try:
        doc = docx.Document(uploaded_file)
        subsections = analyze_subsections(doc)

        if not subsections:
            st.info("В документе не найдено подразделов с номерами 1.1, 1.2 и т.д.")
        else:
            st.subheader(f"Найдено подразделов: {len(subsections)}")

            errors = [s for s in subsections if s["error"]]
            if errors:
                st.error(f"❌ {len(errors)} подраздел(ов) с лишней пустой строкой (нужно убрать):")
                for err in errors:
                    st.markdown(f"- **{err['text']}**")
                    st.caption(f"  Пустая строка перед: {err['has_empty_before']}, предыдущий раздел: {err['prev_was_section']} (текст: «{err['prev_text']}»), новая страница: {err['starts_new_page']}")
            else:
                st.success("✅ Нет ошибок: все подразделы оформлены верно.")

            with st.expander("📋 Детальная информация по всем подразделам"):
                for s in subsections:
                    status = "🔴 ОШИБКА" if s["error"] else "🟢 ОК"
                    st.write(f"{status}: {s['text']}")
                    st.write(f"   - Индекс: {s['index']}, пустая строка перед: {s['has_empty_before']}, предыдущий раздел: {s['prev_was_section']}, новая страница: {s['starts_new_page']}")
                    st.write(f"   - Текст предыдущего: «{s['prev_text']}»")
                    st.write("---")
    except Exception as e:
        st.error(f"Ошибка: {e}")
else:
    st.info("Загрузите файл .docx для проверки.")
