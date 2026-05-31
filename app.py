import streamlit as st
import zipfile
from lxml import etree

def analyze_docx(file):
    results = []
    
    with zipfile.ZipFile(file, 'r') as z:
        # Читаем document.xml чтобы понять связи секций с колонтитулами
        doc_xml = etree.fromstring(z.read('word/document.xml'))
        nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
        
        # Находим все секции
        sections = doc_xml.findall('.//w:sectPr', nsmap)
        
        # Для каждой секции находим footerReference
        for i, sect in enumerate(sections, 1):
            # Ищем ссылку на колонтитул
            footer_refs = sect.findall('.//w:footerReference', nsmap)
            header_refs = sect.findall('.//w:headerReference', nsmap)
            
            # Начальный номер страницы
            pgNumType = sect.find('.//w:pgNumType', nsmap)
            start_page = pgNumType.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}start') if pgNumType is not None else None
            
            section_info = {
                'section': i,
                'headers': [],
                'footers': [],
                'start_page': int(start_page) if start_page else 1
            }
            
            # Собираем информацию о колонтитулах
            for ref in footer_refs:
                ref_type = ref.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'default')
                ref_id = ref.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                section_info['footers'].append({'type': ref_type, 'id': ref_id})
            
            for ref in header_refs:
                ref_type = ref.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'default')
                ref_id = ref.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                section_info['headers'].append({'type': ref_type, 'id': ref_id})
            
            results.append(section_info)
        
        # Читаем связи (relationships)
        rels_xml = etree.fromstring(z.read('word/_rels/document.xml.rels'))
        
        # Строим карту: id -> файл колонтитула
        rels_map = {}
        for rel in rels_xml:
            rel_id = rel.get('Id')
            target = rel.get('Target')
            rel_type = rel.get('Type')
            if 'footer' in rel_type.lower() or 'header' in rel_type.lower():
                rels_map[rel_id] = target
        
        # Для каждой секции читаем содержимое колонтитулов
        for section_info in results:
            for footer_ref in section_info['footers']:
                file_name = rels_map.get(footer_ref['id'])
                if file_name:
                    full_path = 'word/' + file_name
                    try:
                        content = z.read(full_path)
                        tree = etree.fromstring(content)
                        
                        paragraphs_data = []
                        for para in tree.findall('.//w:p', nsmap):
                            # Текст
                            texts = []
                            for t in para.findall('.//w:t', nsmap):
                                if t.text:
                                    texts.append(t.text)
                            full_text = ''.join(texts)
                            
                            # Поле PAGE
                            has_page = False
                            for instr in para.findall('.//w:instrText', nsmap):
                                if instr.text and 'PAGE' in instr.text:
                                    has_page = True
                            if not has_page:
                                has_page = len(para.findall('.//w:fldChar', nsmap)) > 0
                            
                            # Выравнивание
                            jc = para.find('.//w:jc', nsmap)
                            align = jc.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if jc is not None else None
                            
                            # Размер шрифта
                            sz = para.find('.//w:sz', nsmap)
                            font_size = int(sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')) if sz is not None else None
                            if font_size:
                                font_size = font_size / 2  # половинные пункты -> pt
                            
                            paragraphs_data.append({
                                'text': full_text,
                                'has_page': has_page,
                                'alignment': align,
                                'font_size': font_size,
                                'is_empty': not bool(full_text)
                            })
                        
                        footer_ref['paragraphs'] = paragraphs_data
                    except:
                        footer_ref['paragraphs'] = []
            
            for header_ref in section_info['headers']:
                file_name = rels_map.get(header_ref['id'])
                if file_name:
                    full_path = 'word/' + file_name
                    try:
                        content = z.read(full_path)
                        tree = etree.fromstring(content)
                        
                        texts = []
                        for t in tree.findall('.//w:t', nsmap):
                            if t.text:
                                texts.append(t.text)
                        header_ref['text'] = ''.join(texts)
                    except:
                        header_ref['text'] = ''
    
    return results

def test_document(file):
    results = analyze_docx(file)
    
    st.header("📄 Анализ нумерации страниц (прямое чтение docx)")
    
    st.write("**Требования:**")
    st.write("• Автонумерация (поле PAGE)")
    st.write("• Размер шрифта: 14 pt")
    st.write("• Выравнивание: по центру")
    st.write("• Без пустых строк после номера")
    st.write("• Начальный номер введения = 5")
    
    st.write("---")
    
    # Ищем введение
    with zipfile.ZipFile(file, 'r') as z:
        doc_xml = etree.fromstring(z.read('word/document.xml'))
        nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        intro_found = False
        for para in doc_xml.findall('.//w:p', nsmap):
            texts = []
            for t in para.findall('.//w:t', nsmap):
                if t.text:
                    texts.append(t.text)
            full_text = ''.join(texts).strip().upper()
            if full_text == 'ВВЕДЕНИЕ':
                intro_found = True
                break
    
    st.write(f"Введение: {'✅ Найдено' if intro_found else '❌ Не найдено'}")
    
    issues = []
    
    for section_info in results:
        sec_num = section_info['section']
        st.write(f"**Секция {sec_num}:**")
        st.write(f"  Начальный номер: {section_info['start_page']}")
        
        # Проверяем нижние колонтитулы
        for footer_ref in section_info['footers']:
            footer_type = footer_ref.get('type', 'default')
            st.write(f"  Нижний колонтитул ({footer_type}):")
            
            paragraphs = footer_ref.get('paragraphs', [])
            
            if not paragraphs:
                st.write(f"    ❌ Пустой")
                issues.append(f"Секция {sec_num}: нижний колонтитул пуст")
                continue
            
            has_page_in_footer = any(p['has_page'] for p in paragraphs)
            
            if has_page_in_footer:
                # Анализируем параграфы
                for j, p in enumerate(paragraphs, 1):
                    if p['has_page']:
                        st.write(f"    ✅ Параграф {j}: автонумерация PAGE")
                        
                        # Проверяем размер шрифта
                        if p['font_size']:
                            if abs(p['font_size'] - 14) < 0.5:
                                st.write(f"      ✅ Размер шрифта: {p['font_size']} pt")
                            else:
                                st.write(f"      ❌ Размер шрифта: {p['font_size']} pt (должен быть 14)")
                                issues.append(f"Секция {sec_num}: размер шрифта {p['font_size']} pt")
                        else:
                            st.write(f"      ⚠️ Размер шрифта не определён")
                        
                        # Проверяем выравнивание
                        if p['alignment'] == 'center':
                            st.write(f"      ✅ Выравнивание: по центру")
                        else:
                            st.write(f"      ❌ Выравнивание: {p['alignment']} (должно быть center)")
                            issues.append(f"Секция {sec_num}: выравнивание {p['alignment']}")
                    
                    elif p['is_empty']:
                        # Проверяем, после ли это номера
                        prev_has_page = j > 1 and paragraphs[j-2]['has_page']
                        if prev_has_page or has_page_in_footer:
                            st.write(f"    ❌ Параграф {j}: пустая строка после номера")
                            issues.append(f"Секция {sec_num}: пустая строка после номера")
                    else:
                        st.write(f"    Параграф {j}: '{p['text']}'")
            else:
                # Проверяем статичные номера
                static_numbers = [p for p in paragraphs if p['text'].isdigit()]
                if static_numbers:
                    for p in static_numbers:
                        st.write(f"    ❌ Статичный номер: '{p['text']}' (нужно поле PAGE)")
                        issues.append(f"Секция {sec_num}: статичный номер '{p['text']}'")
                else:
                    st.write(f"    ❌ Нет нумерации")
                    issues.append(f"Секция {sec_num}: нет нумерации")
        
        # Проверяем начальный номер
        if section_info['start_page'] != 5:
            st.write(f"  ⚠️ Начальный номер: {section_info['start_page']} (должен быть 5)")
        
        st.write("---")
    
    # Итоги
    if issues:
        st.subheader("❌ Найденные проблемы:")
        for issue in issues:
            st.write(f"• {issue}")
    else:
        st.success("✅ Нумерация настроена правильно!")

# Интерфейс
st.set_page_config(page_title="Проверка нумерации v3", layout="wide")
st.title("📄 Проверка нумерации страниц (прямое чтение)")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализ..."):
        test_document(uploaded_file)
