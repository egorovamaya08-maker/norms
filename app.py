import streamlit as st
import docx
from lxml import etree

st.set_page_config(page_title="Отладка колонтитулов", layout="wide")
st.title("🔧 Отладка: что в колонтитулах?")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    doc = docx.Document(uploaded_file)
    
    for i, section in enumerate(doc.sections, 1):
        st.write(f"---")
        st.write(f"## Секция {i}")
        
        # Проверяем, есть ли вообще колонтитулы
        st.write(f"**header:** {'есть' if section.header else 'НЕТ'}")
        st.write(f"**footer:** {'есть' if section.footer else 'НЕТ'}")
        st.write(f"**different_first_page:** {section.different_first_page_header_footer}")
        
        # Нижний колонтитул
        if section.footer:
            footer = section.footer
            st.write(f"**Нижний колонтитул — параграфов:** {len(footer.paragraphs)}")
            
            for j, para in enumerate(footer.paragraphs):
                st.write(f"  **Параграф {j+1}:**")
                st.write(f"    text: '{para.text}'")
                st.write(f"    runs: {len(para.runs)}")
                
                for k, run in enumerate(para.runs):
                    st.write(f"      Run {k+1}: text='{run.text}', bold={run.bold}, size={run.font.size}")
                
                # Показываем XML параграфа
                st.write(f"    XML:")
                st.code(etree.tostring(para._element, encoding='unicode'), language='xml')
            
            # Полный XML колонтитула
            st.write(f"  **Полный XML колонтитула:**")
            st.code(etree.tostring(footer._element, encoding='unicode')[:2000], language='xml')
        else:
            st.write(f"**Нижний колонтитул ОТСУТСТВУЕТ**")
        
        # Верхний колонтитул
        if section.header:
            header = section.header
            st.write(f"**Верхний колонтитул — параграфов:** {len(header.paragraphs)}")
            for j, para in enumerate(header.paragraphs):
                st.write(f"  Параграф {j+1}: '{para.text}'")
        else:
            st.write(f"**Верхний колонтитул ОТСУТСТВУЕТ**")
