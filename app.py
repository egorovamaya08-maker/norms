import streamlit as st
import docx
import re

st.set_page_config(page_title="Тест: подразделы после разделов", layout="centered")
st.title("🔍 Тест: определение подразделов и пустых строк")
st.markdown("Загрузите документ .docx, и я покажу, для каких подразделов ошибочно требуется убрать пустую строку.")

uploaded_file = st.file_uploader("Выберите файл", type=["docx"])

def is_section_header(text):
    """Заголовок раздела первого уровня (ВВЕДЕНИЕ, 1. НАЗВАНИЕ, ЗАКЛЮЧЕНИЕ и т.п.)"""
    cleaned = text.strip()
    if not cleaned:
        return False
    if cleaned.upper() in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}:
        return True
    if re.match(r'^\d+\.\s*[А-ЯЁ]', cleaned):
        return True
    return False

def is_subsection(text):
    """Подраздел вида 1.1 Текст, 1.1.1 Текст и т.д."""
    cleaned = text.strip()
    if not cleaned:
        return False
    if re.match(r'^\d+\.\d+(\.\d+)?\s+[А-Яа-я]', cleaned):
        return True
    return False

def analyze_subsections(doc):
    """
    Анализирует документ и возвращает список словарей с информацией о каждом подразделе:
    - text: текст подраздела
    - has_empty_line_before: bool, была ли пустая строка перед подразделом
    - prev_nonempty_was_section: был ли предыдущий непустой абзац заголовком раздела
    - should_report_error: должна ли выдаваться ошибка (пустая строка есть И предыдущий непустой НЕ раздел)
    """
    results = []
    prev_para_empty = False
    # Для отслеживания последнего непустого абзаца
    prev_nonempty_was_section = False
    # Для хранения предыдущего непустого текста (отладка)
    prev_nonempty_text = ""
    
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            prev_para_empty = True
            continue
        
        # Определяем тип текущего абзаца
        is_section = is_section_header(text)
        is_sub = is_subsection(text)
        
        # Если текущий абзац — подраздел, анализируем условия перед ним
        if is_sub:
            error_condition = prev_para_empty and not prev_nonempty_was_section
            results.append({
                "index": idx,
                "text": text[:80],
                "has_empty_line_before": prev_para_empty,
                "prev_nonempty_was_section": prev_nonempty_was_section,
                "prev_nonempty_text": prev_nonempty_text[:60] if prev_nonempty_text else "—",
                "should_report_error": error_condition,
                "error_message": f"Подраздел «{text[:50]}» – уберите пустую строку перед подразделом" if error_condition else None
            })
        
        # Обновляем состояние для следующего абзаца
        if is_section:
            prev_nonempty_was_section = True
        else:
            prev_nonempty_was_section = False
        
        prev_nonempty_text = text
        prev_para_empty = False
    
    return results

if uploaded_file is not None:
    try:
        doc = docx.Document(uploaded_file)
        subsections = analyze_subsections(doc)
        
        if not subsections:
            st.info("В документе не найдено подразделов (с номерами вида 1.1, 1.2 и т.д.)")
        else:
            st.subheader("Результаты анализа подразделов")
            st.write(f"Найдено подразделов: {len(subsections)}")
            
            # Группируем: с ошибкой и без
            with_error = [s for s in subsections if s["should_report_error"]]
            without_error = [s for s in subsections if not s["should_report_error"]]
            
            if with_error:
                st.error(f"❌ {len(with_error)} подраздел(ов), перед которыми пустая строка НЕ после раздела (нужно убрать):")
                for s in with_error:
                    st.markdown(f"- **{s['text']}**")
                    st.caption(f"  Перед ним пустая строка: {s['has_empty_line_before']}, предыдущий непустой был разделом: {s['prev_nonempty_was_section']} (текст: «{s['prev_nonempty_text']}»)")
            else:
                st.success("✅ Нет ошибочных требований убрать пустую строку перед подразделами.")
            
            if without_error:
                st.info(f"ℹ️ {len(without_error)} подраздел(ов), перед которыми пустая строка либо отсутствует, либо она после раздела (всё правильно):")
                for s in without_error[:10]:  # показываем не более 10
                    st.markdown(f"- **{s['text']}**")
                    st.caption(f"  Пустая строка перед: {s['has_empty_line_before']}, предыдущий непустой раздел: {s['prev_nonempty_was_section']}")
                if len(without_error) > 10:
                    st.caption(f"... и ещё {len(without_error)-10} подразделов.")
            
            # Детальная таблица для отладки (опционально)
            with st.expander("📋 Детальная таблица всех подразделов"):
                for s in subsections:
                    status = "⚠️ ОШИБКА" if s["should_report_error"] else "✅ ОК"
                    st.write(f"{status}: {s['text']}")
                    st.write(f"   - Пустая строка перед: {s['has_empty_line_before']}")
                    st.write(f"   - Предыдущий непустой был разделом: {s['prev_nonempty_was_section']}")
                    st.write(f"   - Текст предыдущего: «{s['prev_nonempty_text']}»")
                    st.write("---")
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {e}")
else:
    st.info("Ожидание загрузки файла...")
