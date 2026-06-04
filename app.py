import streamlit as st
import docx
import re

st.set_page_config(page_title="Тест: подразделы после разделов", layout="centered")
st.title("🔍 Тест определения подразделов и пустых строк")
st.markdown("Загрузите документ .docx – программа покажет, для каких подразделов ошибочно требуется убрать пустую строку.")

def is_section_header(text):
    """
    Заголовок раздела первого уровня.
    Распознаёт:
      - Нумерованные: 1. ТЕКСТ, 2. ТЕКСТ (все заглавные после номера)
      - Служебные: ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ
      - Заголовки без номера, написанные заглавными буквами (длина > 3 и не содержат строчных)
    """
    cleaned = text.strip()
    if not cleaned:
        return False
    # Прямые названия
    upper_cleaned = cleaned.upper()
    if upper_cleaned in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}:
        return True
    # Начинается с "ГЛАВА" или "РАЗДЕЛ"
    if re.match(r'^(ГЛАВА|РАЗДЕЛ)\s+\d+', upper_cleaned):
        return True
    # Нумерованные заголовки: 1. ТЕКСТ, 2. ТЕКСТ и т.п.
    if re.match(r'^\d+\.\s*[А-ЯЁ]', cleaned):
        return True
    # Заголовок без номера, но все символы (кроме цифр, пробелов, знаков препинания) – заглавные
    # Убираем цифры, пробелы, тире, точки, запятые, скобки, кавычки
    only_letters = re.sub(r'[\d\s\.,;:!?\-–—()«»""''«»]', '', cleaned)
    if only_letters and only_letters == only_letters.upper() and len(only_letters) > 3:
        return True
    return False

def is_subsection(text):
    """Подраздел вида 1.1 Текст, 1.1.1 Текст (первая буква может быть заглавной или строчной)"""
    cleaned = text.strip()
    if not cleaned:
        return False
    # Шаблон: цифра, точка, цифра, пробел, буква
    if re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]', cleaned):
        return True
    return False

def analyze_subsections(doc):
    results = []
    prev_para_empty = False
    prev_nonempty_was_section = False
    prev_nonempty_text = ""

    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            prev_para_empty = True
            continue

        is_section = is_section_header(text)
        is_sub = is_subsection(text)

        if is_sub:
            # Ошибка: пустая строка перед подразделом И предыдущий непустой НЕ был разделом
            error_condition = prev_para_empty and not prev_nonempty_was_section
            results.append({
                "index": idx,
                "text": text[:80],
                "has_empty_before": prev_para_empty,
                "prev_was_section": prev_nonempty_was_section,
                "prev_text": prev_nonempty_text[:60] if prev_nonempty_text else "—",
                "error_should_show": error_condition,
                "error_msg": f"Подраздел «{text[:50]}» – уберите пустую строку перед подразделом" if error_condition else None
            })

        # Обновляем состояние для следующего абзаца (только для непустых)
        if is_section:
            prev_nonempty_was_section = True
        else:
            prev_nonempty_was_section = False

        prev_nonempty_text = text
        prev_para_empty = False

    return results

uploaded_file = st.file_uploader("Выберите файл .docx", type=["docx"])

if uploaded_file is not None:
    try:
        doc = docx.Document(uploaded_file)
        subsections = analyze_subsections(doc)

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
                    st.caption(f"  Пустая строка перед: {s['has_empty_before']}, предыдущий непустой был разделом: {s['prev_was_section']} (текст: «{s['prev_text']}»)")
            else:
                st.success("✅ Нет ошибочных требований убрать пустую строку перед подразделами.")

            if without_error:
                st.info(f"ℹ️ {len(without_error)} подраздел(ов) с корректным расположением (пустая строка либо отсутствует, либо после раздела):")
                for s in without_error[:10]:
                    st.markdown(f"- **{s['text']}**")
                    st.caption(f"  Пустая строка перед: {s['has_empty_before']}, предыдущий раздел: {s['prev_was_section']}")
                if len(without_error) > 10:
                    st.caption(f"... и ещё {len(without_error)-10} подразделов.")

            with st.expander("📋 Детальная информация по каждому подразделу"):
                for s in subsections:
                    status = "🔴 ОШИБКА" if s["error_should_show"] else "🟢 ОК"
                    st.write(f"{status}: {s['text']}")
                    st.write(f"   - Индекс параграфа: {s['index']}")
                    st.write(f"   - Пустая строка перед: {s['has_empty_before']}")
                    st.write(f"   - Предыдущий непустой был разделом: {s['prev_was_section']}")
                    st.write(f"   - Текст предыдущего: «{s['prev_text']}»")
                    st.write("---")

    except Exception as e:
        st.error(f"Ошибка: {e}")
else:
    st.info("Загрузите файл .docx для анализа.")
