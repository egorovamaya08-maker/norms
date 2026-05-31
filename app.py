import streamlit as st
import docx
from docx.shared import Pt, Cm, Mm
import re

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # 1. Проверка полей (переводим EMU в миллиметры)
    for i, section in enumerate(doc.sections, start=1):
        left = round(section.left_margin.mm) if section.left_margin else 0
        right = round(section.right_margin.mm) if section.right_margin else 0
        top = round(section.top_margin.mm) if section.top_margin else 0
        bottom = round(section.bottom_margin.mm) if section.bottom_margin else 0
        
        if left != 20 or right != 20 or top != 20 or bottom != 20:
            issues.append(f"Раздел {i} - установите поля: левое 20, правое 20, верх 20, низ 20 мм")

    # Переменные для отслеживания структуры
    table_count = 0
    figure_count = 0
    
    # 2. Проверка абзацев и текста
    for idx, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            continue
            
        p_format = p.paragraph_format
        
        # Проверка абзацного отступа (должен быть 1.0 см)
        indent = p_format.first_line_indent
        if indent and round(indent.cm, 1) != 1.0 and not text.isupper() and not text.startswith("Рисунок"):
            issues.append(f"Подраздел {text[:20]}... - установите абзацный отступ 1,0 см везде")

        # Проверка междустрочного интервала (1,2)
        spacing = p_format.line_spacing
        if spacing and round(spacing, 1) != 1.2:
            issues.append(f"Подраздел {text[:20]}... - замените интервал на 1,2")

        # Проверка шрифта и начертания внутри абзаца
        for run in p.runs:
            if run.font.name and run.font.name != "Times New Roman":
                issues.append(f"Подраздел {text[:20]}... - смените шрифт на Times New Roman")
            if run.font.size and run.font.size != Pt(14):
                issues.append(f"Подраздел {text[:20]}... - установите размер шрифта 14")
            if run.underline:
                issues.append(f"Подраздел {text[:20]}... - удалите подчеркивания")

        # Проверка СОДЕРЖАНИЯ
        if "СОДЕРЖАНИЕ" in text.upper():
            if text.endswith("."):
                issues.append("Содержание - удалите точку в конце")
            if p.alignment != docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER:
                issues.append("Содержание - выровняйте слово СОДЕРЖАНИЕ по центру")

        # Проверка подписей к рисункам
        if text.startswith("Рисунок"):
            figure_count += 1
            if "." in text[-2:]:
                issues.append(f"Рисунок {figure_count} - удалите точку в конце названия")
            if p.alignment != docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER:
                issues.append(f"Рисунок {figure_count} - выровняйте подпись по центру")

        # Проверка названий таблиц (над таблицей)
        if text.startswith("Таблица") and not any(t in text for t in ["в таблице", "из таблицы"]):
            if "." in text[-2:]:
                issues.append(f"Таблица {text.split()[1]} - удалите точку в конце названия")

    # 3. Проверка таблиц (начертание внутри)
    for t_idx, table in enumerate(doc.tables, start=1):
        bold_inside = False
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if run.bold:
                            bold_inside = True
        if bold_inside:
            issues.append(f"Таблица {t_idx} - уберите полужирное начертание")

    # Очистка дубликатов и финальный вывод
    issues = list(dict.fromkeys(issues))
    if not issues:
        return ["Ошибок не найдено. Документ соответствует чек-листу."]
    return issues

# Настройка интерфейса сайта
st.set_page_config(page_title="Нормоконтроль документов", layout="centered")
st.title("📊 Автоматическая проверка документов Word")
st.write("Загрузите ваш документ в формате .docx для мгновенной проверки по чек-листу.")

uploaded_file = st.file_uploader("Перетащите файл сюда или нажмите для выбора", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализируем структуру, интервалы и разметку..."):
        results = check_word_document(uploaded_file)
        
    st.subheader("Результаты проверки:")
    for res in results:
        st.write(f"• {res}")
