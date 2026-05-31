import streamlit as st
import docx
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def get_header_footer_info(doc, section, header_type='default'):
    """Получает информацию из колонтитула"""
    info = {
        'has_page_number': False,
        'page_number_text': '',
        'font_size': None,
        'alignment': None,
        'xml_info': '',
        'paragraphs_count': 0
    }
    
    try:
        # Получаем колонтитул
        if header_type == 'default':
            header = section.header
        elif header_type == 'first':
            header = section.first_page_header
        elif header_type == 'even':
            header = section.even_page_header
        else:
            return info
        
        if header is None:
            return info
        
        # Проверяем параграфы в колонтитуле
        paragraphs = header.paragraphs
        info['paragraphs_count'] = len(paragraphs)
        
        for para in paragraphs:
            text = para.text.strip()
            
            # Ищем номера страниц (поле PAGE)
            for run in para.runs:
                if 'PAGE' in run._element.xml or 'w:fldChar' in run._element.xml:
                    info['has_page_number'] = True
                    info['page_number_text'] = text if text else '[поле PAGE]'
                    
                    # Размер шрифта
                    if run.font.size:
                        info['font_size'] = run.font.size.pt
                    elif para.style and para.style.font.size:
                        info['font_size'] = para.style.font.size.pt
                    
                    # Выравнивание
                    if para.alignment is not None:
                        info['alignment'] = para.alignment
                    elif para.style and para.style.paragraph_format.alignment is not None:
                        info['alignment'] = para.style.paragraph_format.alignment
                    
                    # XML
                    info['xml_info'] = etree.tostring(run._element, encoding='unicode')[:200]
                    break
            
            # Если нашли номер, выходим
            if info['has_page_number']:
                break
        
        # Если не нашли через runs, проверяем XML напрямую
        if not info['has_page_number']:
            xml_str = etree.tostring(header._element, encoding='unicode')
            if 'w:fldChar' in xml_str or 'PAGE' in xml_str:
                info['has_page_number'] = True
                info['page_number_text'] = '[найдено в XML]'
    
    except Exception as e:
        info['xml_info'] = f'Ошибка: {str(e)[:100]}'
    
    return info

def get_section_info(doc, section, section_num):
    """Получает полную информацию о секции"""
    info = {
        'section_num': section_num,
        'page_width': section.page_width.cm if section.page_width else None,
        'page_height': section.page_height.cm if section.page_height else None,
        'left_margin': section.left_margin.cm if section.left_margin else None,
        'right_margin': section.right_margin.cm if section.right_margin else None,
        'start_type': section.start_type if section.start_type else None,
        'different_first_page': section.different_first_page_header_footer,
    }
    return info

def analyze_page_numbering(doc):
    """Анализирует нумерацию страниц"""
    results = {
        'sections': [],
        'issues': [],
        'summary': {}
    }
    
    for i, section in enumerate(doc.sections, 1):
        section_info = get_section_info(doc, section, i)
        
        # Проверяем все типы колонтитулов
        header_default = get_header_footer_info(doc, section, 'default')
        header_first = get_header_footer_info(doc, section, 'first') if section.different_first_page_header_footer else None
        header_even = get_header_footer_info(doc, section, 'even') if section.different_first_page_header_footer else None
        
        section_data = {
            'section_info': section_info,
            'header_default': header_default,
            'header_first': header_first,
            'header_even': header_even,
        }
        
        results['sections'].append(section_data)
        
        # Проверяем основной колонтитул
        if header_default['has_page_number']:
            # Проверяем размер шрифта
            if header_default['font_size']:
                if abs(header_default['font_size'] - 14) > 0.5:
                    results['issues'].append({
                        'section': i,
                        'type': 'font_size',
                        'message': f"Секция {i}: размер шрифта нумерации {header_default['font_size']} pt (должен быть 14 pt)",
                        'found': header_default['font_size'],
                        'expected': 14
                    })
            
            # Проверяем выравнивание
            if header_default['alignment'] is not None:
                if header_default['alignment'] != WD_ALIGN_PARAGRAPH.CENTER:
                    results['issues'].append({
                        'section': i,
                        'type': 'alignment',
                        'message': f"Секция {i}: нумерация не по центру",
                        'found': header_default['alignment'],
                        'expected': 'CENTER'
                    })
        else:
            results['issues'].append({
                'section': i,
                'type': 'no_numbering',
                'message': f"Секция {i}: нет нумерации страниц в колонтитуле",
                'found': 'отсутствует',
                'expected': 'поле PAGE'
            })
        
        # Проверяем разные колонтитулы для первой страницы
        if section.different_first_page_header_footer and header_first:
            if header_first['has_page_number']:
                results['issues'].append({
                    'section': i,
                    'type': 'first_page_number',
                    'message': f"Секция {i}: на первой странице не должно быть номера",
                    'found': 'есть номер',
                    'expected': 'без номера'
                })
    
    # Определяем начальный номер страницы
    try:
        # Проверяем pageNumType в sectPr
        for section in doc.sections:
            sectPr = section._sectPr
            pgNumType = sectPr.find(qn('w:pgNumType'))
            if pgNumType is not None:
                start = pgNumType.get(qn('w:start'))
                if start:
                    results['start_page'] = int(start)
    except:
        results['start_page'] = None
    
    return results

def test_document(file):
    doc = docx.Document(file)
    
    st.header("📄 Проверка нумерации страниц (колонтитулы)")
    st.write("**Требования:**")
    st.write("• Нумерация должна быть во всех секциях")
    st.write("• Размер шрифта: 14 pt")
    st.write("• Выравнивание: по центру")
    st.write("• На первой странице секции номера быть не должно (если включены разные колонтитулы)")
    st.write("• Отсчёт страниц начинается с титульного листа (первая страница введения = 5)")
    
    st.write("---")
    
    # Анализируем нумерацию
    results = analyze_page_numbering(doc)
    
    # Показываем информацию о секциях
    st.subheader("📊 Информация о секциях документа:")
    st.write(f"Всего секций: {len(results['sections'])}")
    
    for section_data in results['sections']:
        sec = section_data['section_info']
        h_default = section_data['header_default']
        
        st.write(f"**Секция {sec['section_num']}:**")
        st.write(f"  Размер страницы: {sec['page_width']:.1f} × {sec['page_height']:.1f} см")
        st.write(f"  Поля: левое {sec['left_margin']:.1f} см, правое {sec['right_margin']:.1f} см")
        st.write(f"  Разные колонтитулы для первой страницы: {'Да' if sec['different_first_page'] else 'Нет'}")
        
        st.write(f"  **Колонтитул (основной):**")
        if h_default['has_page_number']:
            st.write(f"    ✅ Нумерация есть: {h_default['page_number_text']}")
            st.write(f"    Размер шрифта: {h_default['font_size']} pt {'✅' if h_default['font_size'] and abs(h_default['font_size'] - 14) < 0.5 else '❌ должен быть 14 pt'}")
            
            if h_default['alignment'] is not None:
                align_name = {0: 'LEFT', 1: 'CENTER', 2: 'RIGHT', 3: 'JUSTIFY'}.get(h_default['alignment'], str(h_default['alignment']))
                st.write(f"    Выравнивание: {align_name} {'✅' if h_default['alignment'] == 1 else '❌ должно быть по центру'}")
        else:
            st.write(f"    ❌ Нумерация отсутствует")
        
        # Проверяем первую страницу
        if section_data['header_first']:
            h_first = section_data['header_first']
            st.write(f"  **Колонтитул (первая страница):**")
            if h_first['has_page_number']:
                st.write(f"    ❌ Есть номер на первой странице (должен быть пустым)")
            else:
                st.write(f"    ✅ Пустой (правильно)")
        
        st.write(f"  Параграфов в колонтитуле: {h_default['paragraphs_count']}")
        st.write("---")
    
    # Показываем найденные проблемы
    if results['issues']:
        st.subheader("❌ Найденные проблемы:")
        for issue in results['issues']:
            st.write(f"• {issue['message']}")
    else:
        st.success("✅ Нумерация страниц настроена правильно")
    
    # Проверяем начальный номер страницы
    st.subheader("🔢 Проверка начального номера страницы:")
    if results.get('start_page'):
        st.write(f"Начальный номер страницы: {results['start_page']}")
        if results['start_page'] == 1:
            st.write("💡 Нумерация начинается с 1 — проверьте, что первая страница введения = 5 (если титульный лист и содержание занимают 4 страницы)")
    else:
        st.write("⚠️ Не удалось определить начальный номер страницы (возможно, нумерация сквозная с 1)")
    
    # Дополнительная отладка
    st.subheader("🔧 Отладка (проверка колонтитулов):")
    for i, section in enumerate(doc.sections, 1):
        try:
            default_header = section.header
            if default_header:
                st.write(f"**Секция {i} - XML колонтитула:**")
                xml_str = etree.tostring(default_header._element, encoding='unicode')
                # Показываем только первые 1000 символов
                if len(xml_str) > 1000:
                    xml_str = xml_str[:1000] + "..."
                st.code(xml_str, language='xml')
        except Exception as e:
            st.write(f"**Секция {i}:** Ошибка чтения колонтитула: {e}")

# Импорт etree
from lxml import etree

# Интерфейс
st.set_page_config(page_title="Проверка нумерации страниц", layout="wide")
st.title("📄 Проверка нумерации страниц")
st.write("Анализ колонтитулов и проверка правильности нумерации")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализ колонтитулов..."):
        test_document(uploaded_file)
