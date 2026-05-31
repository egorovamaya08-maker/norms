import streamlit as st
import zipfile
from lxml import etree

st.set_page_config(page_title="Просмотр XML колонтитулов", layout="wide")
st.title("🔧 Прямой просмотр колонтитулов из docx")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with zipfile.ZipFile(uploaded_file, 'r') as z:
        # Показываем структуру архива
        st.write("**Файлы в архиве docx:**")
        for name in sorted(z.namelist()):
            if 'header' in name.lower() or 'footer' in name.lower():
                st.write(f"  📄 {name}")
        
        st.write("---")
        
        # Ищем и показываем все колонтитулы
        footer_files = [f for f in z.namelist() if 'footer' in f.lower()]
        header_files = [f for f in z.namelist() if 'header' in f.lower()]
        
        # Нижние колонтитулы
        st.subheader("📥 Нижние колонтитулы (footer):")
        for fname in sorted(footer_files):
            st.write(f"**{fname}:**")
            try:
                content = z.read(fname)
                # Парсим XML
                tree = etree.fromstring(content)
                # Ищем все параграфы
                nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                
                for i, para in enumerate(tree.findall('.//w:p', nsmap), 1):
                    # Собираем весь текст из параграфа
                    texts = []
                    for t in para.findall('.//w:t', nsmap):
                        if t.text:
                            texts.append(t.text)
                    full_text = ''.join(texts)
                    
                    # Ищем поле PAGE
                    has_page_field = False
                    for instr in para.findall('.//w:instrText', nsmap):
                        if instr.text and 'PAGE' in instr.text:
                            has_page_field = True
                            break
                    
                    has_fldChar = len(para.findall('.//w:fldChar', nsmap)) > 0
                    
                    # Выравнивание
                    jc = para.find('.//w:jc', nsmap)
                    alignment = jc.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if jc is not None else 'не задано'
                    
                    # Размер шрифта
                    sz = para.find('.//w:sz', nsmap)
                    font_size = sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if sz is not None else None
                    
                    st.write(f"  Параграф {i}:")
                    st.write(f"    Текст: '{full_text}'")
                    st.write(f"    Поле PAGE: {'✅' if has_page_field else '❌'}")
                    st.write(f"    fldChar: {'✅' if has_fldChar else '❌'}")
                    st.write(f"    Выравнивание: {alignment}")
                    if font_size:
                        st.write(f"    Размер шрифта: {int(font_size)/2} pt (w:sz={font_size})")
                    
                    # Показываем XML если есть что-то интересное
                    if full_text or has_page_field or has_fldChar:
                        st.code(etree.tostring(para, encoding='unicode', pretty_print=True)[:500], language='xml')
                
                # Если параграфов нет
                if len(tree.findall('.//w:p', nsmap)) == 0:
                    st.write("  (пусто)")
                    st.code(content.decode('utf-8')[:500], language='xml')
            
            except Exception as e:
                st.error(f"Ошибка чтения: {e}")
        
        # Верхние колонтитулы
        st.subheader("📥 Верхние колонтитулы (header):")
        for fname in sorted(header_files):
            st.write(f"**{fname}:**")
            try:
                content = z.read(fname)
                tree = etree.fromstring(content)
                nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                
                for i, para in enumerate(tree.findall('.//w:p', nsmap), 1):
                    texts = []
                    for t in para.findall('.//w:t', nsmap):
                        if t.text:
                            texts.append(t.text)
                    full_text = ''.join(texts)
                    
                    if full_text:
                        st.write(f"  Параграф {i}: '{full_text}'")
                    else:
                        st.write(f"  Параграф {i}: (пустой)")
                
                if len(tree.findall('.//w:p', nsmap)) == 0:
                    st.write("  (пусто)")
            
            except Exception as e:
                st.error(f"Ошибка чтения: {e}")
