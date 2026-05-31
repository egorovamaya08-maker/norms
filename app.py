import streamlit as st
import docx
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree
import re

def get_header_footer_info(section, header_type='default', location='header'):
    """
    Получает информацию из колонтитула.
    location: 'header' или 'footer'
    """
    info = {
        'has_page_number': False,
        'page_number_text': '',
        'font_size': None,
        'alignment': None,
        'paragraphs_count': 0,
        'all_text': '',
        'xml_info': '',
        'location': location
    }
    
    try:
        # Получаем колонтитул
        if location == 'header':
            if header_type == 'default':
                container = section.header
            elif header_type == 'first':
                container = section.first_page_header
            elif header_type == 'even':
                container = section.even_page_header
            else:
                return info
        else:  # footer
            if header_type == 'default':
                container = section.footer
            elif header_type == 'first':
                container = section.first_page_footer
            elif header_type == 'even':
                container = section.even_page_footer
            else:
                return info
        
        if container is None:
            return info
        
        # Сохраняем XML для отладки
        info['xml_info'] = etree.tostring(container._element, encoding='unicode')
        
        paragraphs = container.paragraphs
        info['paragraphs_count'] = len(paragraphs)
        
        for para in paragraphs:
            text = para.text.strip()
            info['all_text'] += text + ' '
            
            # Проверяем каждый run
            for run in para.runs:
                run_xml = run._element.xml
                
                # Способ 1: w:fldChar (символы поля)
                if 'w:fldChar' in run_xml:
                    info['has_page_number'] = True
                    info['page_number_text'] = run.text if run.text else '[поле]'
                    
                    if run.font.size:
                        info['font_size'] = run.font.size.pt
                    
                    if para.alignment is not None:
                        info['alignment'] = para.alignment
                
                # Способ 2: w:instrText с PAGE
                if 'w:instrText' in run_xml and 'PAGE' in run.text:
                    info['has_page_number'] = True
                    info['page_number_text'] = run.text
                    
                    if run.font.size and not info['font_size']:
                        info['font_size'] = run.font.size.pt
            
            # Получаем размер шрифта, если ещё нет
            if not info['font_size']:
                for run in para.runs:
                    if run.font.size:
                        info['font_size'] = run.font.size.pt
                        break
                
                if not info['font_size']:
                    try:
                        if para.style and para.style.font.size:
                            info['font_size'] = para.style.font.size.pt
                    except:
                        pass
            
            # Выравнивание
            if info['alignment'] is None and para.alignment is not None:
                info['alignment'] = para.alignment
        
        # Проверяем XML на PAGE, если не нашли через runs
        if not info['has_page_number']:
            if 'PAGE' in info['xml_info']:
                info['has_page_number'] = True
                info['page_number_text'] = '[найдено в XML]'
    
    except Exception as e:
        info['xml_info'] = f'Ошибка: {str(e)[:200]}'
    
    return info

def analyze_page_numbering(doc):
    """Анализирует нумерацию страниц в колонтитулах"""
    results = {
        'sections': [],
        'issues': [],
        'has_numbering_in_header': False,
        'has_numbering_in_footer': False
    }
    
    for i, section in enumerate(doc.sections, 1):
        # Проверяем хедеры
        header_default = get_header_footer_info(section, 'default', 'header')
        footer_default = get_header_footer_info(section, 'default', 'footer')
        
        header_first = get_header_footer_info(section, 'first', 'header') if section.different_first_page_header_footer else None
        footer_first = get_header_footer_info(section, 'first', 'footer') if section.different_first_page_header_footer else None
        
        if header_default['has_page_number']:
            results['has_numbering_in_header'] = True
        if footer_default['has_page_number']:
            results['has_numbering_in_footer'] = True
        
        section_data = {
            'section_num': i,
            'different_first_page': section.different_first_page_header_footer,
            'header': header_default,
            'footer': footer_default,
            'header_first': header_first,
            'footer_first': footer_first,
        }
        
        results['sections'].append(section_data)
        
        # Определяем, где находится нумерация
        numbering_container = None
        if header_default['has_page_number']:
            numbering_container = header_default
            location_text = 'верхнем колонтитуле'
        elif footer_default['has_page_number']:
            numbering_container = footer_default
            location_text = 'нижнем колонтитуле'
        
        if numbering_container:
            # Проверяем размер шрифта
            if numbering_container['font_size']:
                if abs(numbering_container['font_size'] - 14) > 0.5:
                    results['issues'].append({
                        'section': i,
                        'type': 'font_size',
                        'message': f"Секция {i}: размер шрифта {numbering_container['font_size']} pt в {location_text} (должен быть 14 pt)",
                    })
            
            # Проверяем выравнивание
            if numbering_container['alignment'] is not None:
                if numbering_container['alignment'] != WD_ALIGN_PARAGRAPH.CENTER:
                    align_names = {0: 'по левому краю', 1: 'по центру', 2: 'по правому краю', 3: 'по ширине'}
                    results['issues'].append({
                        'section': i,
                        'type': 'alignment',
                        'message': f"Секция {i}: выравнивание {align_names.get(numbering_container['alignment'], 'неизвестно')} в {location_text} (должно быть по центру)",
                    })
        else:
            results['issues'].append({
                'section': i,
                'type': 'no_numbering',
                'message': f"Секция {i}: нумерация не найдена ни в верхнем, ни в нижнем колонтитуле",
            })
        
        # Проверяем первую страницу
        first_page_container = None
        if section.different_first_page_header_footer:
            if header_first and header_first['has_page_number']:
                first_page_container = header_first
                location_text = 'верхнем колонтитуле первой страницы'
            elif footer_first and footer_first['has_page_number']:
                first_page_container = footer_first
                location_text = 'нижнем колонтитуле первой страницы'
            
            if first_page_container:
                results['issues'].append({
                    'section': i,
                    'type': 'first_page_number',
                    'message': f"Секция {i}: есть номер в {location_text} (должна быть пустой)",
                })
    
    # Начальный номер
    try:
        for section in doc.sections:
            sectPr = section._sectPr
            pgNumType = sectPr.find(qn('w:pgNumType'))
            if pgNumType is not None:
                start = pgNumType.get(qn('w:start'))
                if start:
                    results['start_page'] = int(start)
                    break
    except:
        results['start_page'] = None
    
    return results

def test_document(file):
    doc = docx.Document(file)
    
    st.header("📄 Проверка нумерации страниц")
    st.write("**Требования:**")
    st.write("• Размер шрифта: 14 pt")
    st.write("• Выравнивание: по центру")
    st.write("• Нумерация в нижнем или верхнем колонтитуле")
    st.write("• Первая страница секции — без номера (если разные колонтитулы)")
    
    st.write("---")
    
    results = analyze_page_numbering(doc)
    
    # Общая информация
    st.subheader(f"📊 Секций: {len(results['sections'])}")
    st.write(f"Нумерация в верхнем колонтитуле: {'✅ Да' if results['has_numbering_in_header'] else '❌ Нет'}")
    st.write(f"Нумерация в нижнем колонтитуле: {'✅ Да' if results['has_numbering_in_footer'] else '❌ Нет'}")
    
    st.write("---")
    
    # Детали по секциям
    for section_data in results['sections']:
        sec_num = section_data['section_num']
        header = section_data['header']
        footer = section_data['footer']
        
        st.write(f"**Секция {sec_num}:**")
        
        # Верхний колонтитул
        st.write(f"  🔼 Верхний колонтитул:")
        st.write(f"    Нумерация: {'✅ Есть' if header['has_page_number'] else '❌ Нет'}")
        st.write(f"    Текст: '{header['all_text'].strip()[:80]}'")
        st.write(f"    Параграфов: {header['paragraphs_count']}")
        
        # Нижний колонтитул
        st.write(f"  🔽 Нижний колонтитул:")
        st.write(f"    Нумерация: {'✅ Есть' if footer['has_page_number'] else '❌ Нет'}")
        st.write(f"    Текст: '{footer['all_text'].strip()[:80]}'")
        st.write(f"    Параграфов: {footer['paragraphs_count']}")
        
        # Определяем, где номер
        if header['has_page_number'] or footer['has_page_number']:
            container = header if header['has_page_number'] else footer
            location = 'верхнем' if header['has_page_number'] else 'нижнем'
            
            st.write(f"  ✅ Номер найден в **{location}** колонтитуле")
            
            if container['font_size']:
                font_ok = abs(container['font_size'] - 14) < 0.5
                st.write(f"  Размер шрифта: {container['font_size']} pt {'✅' if font_ok else '❌ (14 pt)'}")
            
            if container['alignment'] is not None:
                align_names = {0: 'LEFT', 1: 'CENTER', 2: 'RIGHT', 3: 'JUSTIFY'}
                align_ok = container['alignment'] == 1
                st.write(f"  Выравнивание: {align_names.get(container['alignment'], '?')} {'✅' if align_ok else '❌ (CENTER)'}")
        
        # Первая страница
        if section_data['different_first_page']:
            first_has_number = False
            if section_data['header_first'] and section_data['header_first']['has_page_number']:
                first_has_number = True
            if section_data['footer_first'] and section_data['footer_first']['has_page_number']:
                first_has_number = True
            
            if first_has_number:
                st.write(f"  ❌ Первая страница: есть номер")
            else:
                st.write(f"  ✅ Первая страница: без номера")
        
        st.write("---")
    
    # Проблемы
    real_issues = [i for i in results['issues']]
    
    if real_issues:
        st.subheader("❌ Найденные проблемы:")
        for issue in real_issues:
            st.write(f"• {issue['message']}")
    else:
        st.success("✅ Нумерация страниц настроена правильно")
    
    # Начальный номер
    st.subheader("🔢 Начальный номер:")
    if results.get('start_page'):
        st.write(f"Установлен: **{results['start_page']}**")
        if results['start_page'] == 5:
            st.write("✅ Правильно (4 страницы до введения)")
        elif results['start_page'] == 1:
            st.write("⚠️ Начинается с 1 (должно быть 5, если 4 страницы до введения)")
    else:
        st.write("⚠️ Не определен")
    
    # Отладка
    with st.expander("🔧 Отладка: XML колонтитулов"):
        for i, section_data in enumerate(results['sections'], 1):
            st.write(f"**Секция {i} - Верхний колонтитул:**")
            st.code(section_data['header']['xml_info'][:1000], language='xml')
            st.write(f"**Секция {i} - Нижний колонтитул:**")
            st.code(section_data['footer']['xml_info'][:1000], language='xml')


# Интерфейс
st.set_page_config(page_title="Проверка нумерации", layout="wide")
st.title("📄 Проверка нумерации страниц")
st.write("Анализ верхних и нижних колонтитулов")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализ..."):
        test_document(uploaded_file)
