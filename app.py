import streamlit as st
import docx
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree
import re

def get_header_footer_info(section, header_type='default'):
    """Получает информацию из колонтитула"""
    info = {
        'has_page_number': False,
        'page_number_text': '',
        'font_size': None,
        'alignment': None,
        'xml_info': '',
        'paragraphs_count': 0,
        'all_text': '',
        'has_xml_number': False
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
        
        # Сохраняем полный XML для отладки
        info['xml_info'] = etree.tostring(header._element, encoding='unicode')
        
        paragraphs = header.paragraphs
        info['paragraphs_count'] = len(paragraphs)
        
        # Проверяем каждый параграф
        for para in paragraphs:
            text = para.text.strip()
            info['all_text'] += text + ' '
            
            # Способ 1: ищем через runs с w:fldChar
            for run in para.runs:
                run_xml = run._element.xml
                
                # Проверяем w:fldChar (начало/конец поля)
                if 'w:fldChar' in run_xml:
                    info['has_page_number'] = True
                    info['has_xml_number'] = True
                    
                    # Размер шрифта
                    if run.font.size:
                        info['font_size'] = run.font.size.pt
                    
                    # Выравнивание
                    if para.alignment is not None:
                        info['alignment'] = para.alignment
            
            # Способ 2: ищем w:instrText="PAGE" или w:instrText=" PAGE "
            for run in para.runs:
                run_xml = run._element.xml
                if 'w:instrText' in run_xml and ('PAGE' in run_xml or 'Page' in run_xml):
                    info['has_page_number'] = True
                    info['has_xml_number'] = True
                    info['page_number_text'] = run.text
                    
                    if not info['font_size'] and run.font.size:
                        info['font_size'] = run.font.size.pt
            
            # Способ 3: проверяем весь XML параграфа на наличие PAGE
            para_xml = etree.tostring(para._element, encoding='unicode')
            if 'PAGE' in para_xml or 'w:fldChar' in para_xml:
                info['has_xml_number'] = True
                if not info['has_page_number']:
                    info['has_page_number'] = True
                    info['page_number_text'] = '[найдено через XML]'
            
            # Получаем размер шрифта, если ещё не получили
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
            if info['alignment'] is None:
                if para.alignment is not None:
                    info['alignment'] = para.alignment
                elif para.style and para.style.paragraph_format.alignment is not None:
                    info['alignment'] = para.style.paragraph_format.alignment
        
        # Если всё ещё не нашли, проверяем XML на PAGE
        if not info['has_page_number']:
            xml_str = info['xml_info']
            if 'PAGE' in xml_str or 'Page' in xml_str:
                info['has_page_number'] = True
                info['page_number_text'] = '[найдено в XML]'
    
    except Exception as e:
        info['xml_info'] = f'Ошибка: {str(e)[:200]}'
    
    return info

def analyze_page_numbering(doc):
    """Анализирует нумерацию страниц"""
    results = {
        'sections': [],
        'issues': [],
        'summary': {}
    }
    
    for i, section in enumerate(doc.sections, 1):
        # Проверяем все типы колонтитулов
        header_default = get_header_footer_info(section, 'default')
        header_first = get_header_footer_info(section, 'first') if section.different_first_page_header_footer else None
        header_even = get_header_footer_info(section, 'even') if section.different_first_page_header_footer else None
        
        section_data = {
            'section_num': i,
            'different_first_page': section.different_first_page_header_footer,
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
                else:
                    results['issues'].append({
                        'section': i,
                        'type': 'font_size_ok',
                        'message': f"Секция {i}: размер шрифта {header_default['font_size']} pt ✅",
                        'found': header_default['font_size'],
                        'expected': 14
                    })
            else:
                results['issues'].append({
                    'section': i,
                    'type': 'no_font_size',
                    'message': f"Секция {i}: не удалось определить размер шрифта нумерации",
                    'found': None,
                    'expected': 14
                })
            
            # Проверяем выравнивание
            if header_default['alignment'] is not None:
                if header_default['alignment'] != WD_ALIGN_PARAGRAPH.CENTER:
                    align_names = {0: 'LEFT', 1: 'CENTER', 2: 'RIGHT', 3: 'JUSTIFY'}
                    results['issues'].append({
                        'section': i,
                        'type': 'alignment',
                        'message': f"Секция {i}: нумерация не по центру (выравнивание: {align_names.get(header_default['alignment'], str(header_default['alignment']))})",
                        'found': header_default['alignment'],
                        'expected': 'CENTER'
                    })
            else:
                results['issues'].append({
                    'section': i,
                    'type': 'no_alignment',
                    'message': f"Секция {i}: не удалось определить выравнивание нумерации",
                    'found': None,
                    'expected': 'CENTER'
                })
        else:
            if header_default['has_xml_number']:
                results['issues'].append({
                    'section': i,
                    'type': 'xml_found',
                    'message': f"Секция {i}: нумерация найдена через XML, но не через runs (проверьте вручную)",
                    'found': 'XML',
                    'expected': 'видимая нумерация'
                })
            else:
                results['issues'].append({
                    'section': i,
                    'type': 'no_numbering',
                    'message': f"Секция {i}: нумерация страниц не найдена в колонтитуле",
                    'found': 'отсутствует',
                    'expected': 'поле PAGE'
                })
        
        # Проверяем первую страницу
        if section.different_first_page_header_footer and header_first:
            if header_first['has_page_number']:
                results['issues'].append({
                    'section': i,
                    'type': 'first_page_number',
                    'message': f"Секция {i}: на первой странице есть номер (должна быть пустой)",
                    'found': 'есть номер',
                    'expected': 'пусто'
                })
    
    # Определяем начальный номер страницы
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
    
    st.header("📄 Проверка нумерации страниц (колонтитулы)")
    st.write("**Требования:**")
    st.write("• Размер шрифта: 14 pt")
    st.write("• Выравнивание: по центру")
    st.write("• Нумерация сквозная, начинается с титульного листа")
    
    st.write("---")
    
    # Анализируем нумерацию
    results = analyze_page_numbering(doc)
    
    # Показываем информацию о секциях
    st.subheader(f"📊 Секций в документе: {len(results['sections'])}")
    
    for section_data in results['sections']:
        sec_num = section_data['section_num']
        h_default = section_data['header_default']
        
        st.write(f"**Секция {sec_num}:**")
        st.write(f"  Разные колонтитулы: {'Да' if section_data['different_first_page'] else 'Нет'}")
        
        # Статус нумерации
        if h_default['has_page_number']:
            st.write(f"  ✅ Нумерация найдена: {h_default['page_number_text']}")
        else:
            if h_default['has_xml_number']:
                st.write(f"  ⚠️ Нумерация найдена в XML, но не распознана полностью")
            else:
                st.write(f"  ❌ Нумерация НЕ найдена")
        
        # Размер шрифта
        if h_default['font_size']:
            font_ok = abs(h_default['font_size'] - 14) < 0.5
            st.write(f"  Размер шрифта: {h_default['font_size']} pt {'✅' if font_ok else '❌ (должен быть 14 pt)'}")
        else:
            st.write(f"  Размер шрифта: не определён ❌")
        
        # Выравнивание
        if h_default['alignment'] is not None:
            align_names = {0: 'LEFT', 1: 'CENTER', 2: 'RIGHT', 3: 'JUSTIFY'}
            align_ok = h_default['alignment'] == 1
            st.write(f"  Выравнивание: {align_names.get(h_default['alignment'], 'неизвестно')} {'✅' if align_ok else '❌ (должно быть по центру)'}")
        else:
            st.write(f"  Выравнивание: не определено")
        
        # Первая страница
        if section_data['header_first']:
            h_first = section_data['header_first']
            if h_first['has_page_number']:
                st.write(f"  ❌ Первая страница: есть номер (должен быть пустым)")
            else:
                st.write(f"  ✅ Первая страница: без номера")
        
        st.write(f"  Параграфов в колонтитуле: {h_default['paragraphs_count']}")
        st.write(f"  Текст в колонтитуле: '{h_default['all_text'].strip()[:100]}'")
        st.write("---")
    
    # Показываем проблемы
    real_issues = [i for i in results['issues'] if not i['type'].endswith('_ok')]
    info_messages = [i for i in results['issues'] if i['type'].endswith('_ok')]
    
    if real_issues:
        st.subheader("❌ Найденные проблемы:")
        for issue in real_issues:
            st.write(f"• {issue['message']}")
    
    if info_messages:
        st.subheader("✅ Правильные настройки:")
        for info in info_messages:
            st.write(f"• {info['message']}")
    
    if not real_issues and not info_messages:
        st.warning("⚠️ Не удалось проверить нумерацию")
    
    # Начальный номер
    st.subheader("🔢 Начальный номер страницы:")
    if results.get('start_page'):
        st.write(f"Установлен начальный номер: **{results['start_page']}**")
        if results['start_page'] == 1:
            st.write("💡 Нумерация с 1 — проверьте, что первая страница введения = 5")
        elif results['start_page'] == 5:
            st.write("✅ Начальный номер 5 — правильно (4 страницы до введения)")
    else:
        st.write("⚠️ Не удалось определить начальный номер")
    
    # Отладка: показываем XML колонтитулов
    with st.expander("🔧 Отладка: XML колонтитулов"):
        for i, section_data in enumerate(results['sections'], 1):
            st.write(f"**Секция {i} - XML колонтитула:**")
            xml_str = section_data['header_default']['xml_info']
            if len(xml_str) > 2000:
                xml_str = xml_str[:2000] + "..."
            st.code(xml_str, language='xml')


# Интерфейс
st.set_page_config(page_title="Проверка нумерации страниц", layout="wide")
st.title("📄 Проверка нумерации страниц")
st.write("Анализ колонтитулов и проверка правильности нумерации")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализ колонтитулов..."):
        test_document(uploaded_file)
