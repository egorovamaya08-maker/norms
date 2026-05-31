import streamlit as st
import docx
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree
import re

def get_header_footer_info(section, header_type='default', location='header'):
    """
    Получает информацию из колонтитула.
    """
    info = {
        'has_page_number': False,
        'is_auto_numbering': False,  # True если это поле PAGE (автонумерация)
        'is_static_number': False,   # True если это просто текст с цифрой
        'static_number_value': None, # Значение статичного номера
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
            
            # Проверяем, есть ли в XML поле PAGE
            para_xml = etree.tostring(para._element, encoding='unicode')
            has_page_field = 'PAGE' in para_xml and ('w:fldChar' in para_xml or 'w:instrText' in para_xml)
            
            if has_page_field:
                # Это правильная автонумерация через поле PAGE
                info['has_page_number'] = True
                info['is_auto_numbering'] = True
                info['page_number_text'] = '[автонумерация PAGE]'
            elif text and re.match(r'^\d+$', text):
                # Это просто текст с цифрой (статичный номер)
                info['has_page_number'] = True
                info['is_static_number'] = True
                info['static_number_value'] = int(text)
                info['page_number_text'] = f'статичный номер: {text}'
            elif text:
                # Любой другой текст в колонтитуле
                info['page_number_text'] = text
            
            # Получаем размер шрифта
            for run in para.runs:
                if run.font.size and not info['font_size']:
                    info['font_size'] = run.font.size.pt
                
                # Выравнивание
                if info['alignment'] is None and para.alignment is not None:
                    info['alignment'] = para.alignment
            
            if not info['font_size']:
                try:
                    if para.style and para.style.font.size:
                        info['font_size'] = para.style.font.size.pt
                except:
                    pass
    
    except Exception as e:
        info['xml_info'] = f'Ошибка: {str(e)[:200]}'
    
    return info

def analyze_page_numbering(doc):
    """Анализирует нумерацию страниц"""
    results = {
        'sections': [],
        'issues': [],
        'has_auto_numbering': False,
        'has_static_numbering': False,
        'static_numbers': []
    }
    
    for i, section in enumerate(doc.sections, 1):
        footer_default = get_header_footer_info(section, 'default', 'footer')
        header_default = get_header_footer_info(section, 'default', 'header')
        
        section_data = {
            'section_num': i,
            'different_first_page': section.different_first_page_header_footer,
            'header': header_default,
            'footer': footer_default,
        }
        
        results['sections'].append(section_data)
        
        # Проверяем, где есть нумерация
        for container, location in [(footer_default, 'нижнем'), (header_default, 'верхнем')]:
            if container['has_page_number']:
                if container['is_auto_numbering']:
                    results['has_auto_numbering'] = True
                    
                    # Проверяем размер шрифта
                    if container['font_size'] and abs(container['font_size'] - 14) > 0.5:
                        results['issues'].append({
                            'section': i,
                            'type': 'font_size',
                            'message': f"Секция {i}: размер шрифта {container['font_size']} pt в {location} колонтитуле (должен быть 14 pt)"
                        })
                    
                    # Проверяем выравнивание
                    if container['alignment'] is not None and container['alignment'] != WD_ALIGN_PARAGRAPH.CENTER:
                        align_names = {0: 'по левому краю', 1: 'по центру', 2: 'по правому краю', 3: 'по ширине'}
                        results['issues'].append({
                            'section': i,
                            'type': 'alignment',
                            'message': f"Секция {i}: выравнивание {align_names.get(container['alignment'], 'неизвестно')} в {location} колонтитуле (должно быть по центру)"
                        })
                
                elif container['is_static_number']:
                    results['has_static_numbering'] = True
                    results['static_numbers'].append({
                        'section': i,
                        'value': container['static_number_value'],
                        'location': location
                    })
                    
                    # Это ошибка - должно быть поле PAGE
                    results['issues'].append({
                        'section': i,
                        'type': 'static_number',
                        'message': f"Секция {i}: в {location} колонтитуле статичный номер '{container['static_number_value']}' вместо автонумерации (нужно вставить поле PAGE)"
                    })
    
    # Проверяем, одинаковые ли статичные номера
    if len(results['static_numbers']) > 1:
        values = [n['value'] for n in results['static_numbers']]
        if len(set(values)) == 1:
            results['issues'].append({
                'section': 0,
                'type': 'same_static_number',
                'message': f"⚠️ Во всех секциях одинаковый статичный номер '{values[0]}' — нужно заменить на поле PAGE (автонумерацию)"
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
    st.write("• Автоматическая нумерация (поле PAGE)")
    st.write("• Размер шрифта: 14 pt")
    st.write("• Выравнивание: по центру")
    st.write("• Нумерация сквозная, начинается с титульного листа")
    
    st.write("---")
    
    results = analyze_page_numbering(doc)
    
    # Общая информация
    st.subheader(f"📊 Секций: {len(results['sections'])}")
    st.write(f"Автонумерация (поле PAGE): {'✅ Есть' if results['has_auto_numbering'] else '❌ Нет'}")
    st.write(f"Статичные номера (просто цифры): {'⚠️ Есть' if results['has_static_numbering'] else '✅ Нет'}")
    
    st.write("---")
    
    # Детали по секциям
    for section_data in results['sections']:
        sec_num = section_data['section_num']
        header = section_data['header']
        footer = section_data['footer']
        
        st.write(f"**Секция {sec_num}:**")
        
        # Нижний колонтитул
        st.write(f"  🔽 Нижний колонтитул:")
        if footer['has_page_number']:
            if footer['is_auto_numbering']:
                st.write(f"    ✅ Автонумерация (поле PAGE)")
            elif footer['is_static_number']:
                st.write(f"    ❌ Статичный номер: **{footer['static_number_value']}** (должно быть поле PAGE)")
            st.write(f"    Текст: '{footer['all_text'].strip()}'")
            
            if footer['font_size']:
                font_ok = abs(footer['font_size'] - 14) < 0.5
                st.write(f"    Размер шрифта: {footer['font_size']} pt {'✅' if font_ok else '❌ (14 pt)'}")
            
            if footer['alignment'] is not None:
                align_names = {0: 'LEFT', 1: 'CENTER', 2: 'RIGHT', 3: 'JUSTIFY'}
                align_ok = footer['alignment'] == 1
                st.write(f"    Выравнивание: {align_names.get(footer['alignment'], '?')} {'✅' if align_ok else '❌ (CENTER)'}")
        else:
            st.write(f"    ❌ Нет нумерации")
        
        # Верхний колонтитул
        st.write(f"  🔼 Верхний колонтитул:")
        if header['has_page_number']:
            st.write(f"    Текст: '{header['all_text'].strip()}'")
        else:
            st.write(f"    Пустой")
        
        st.write("---")
    
    # Проблемы
    real_issues = [i for i in results['issues']]
    
    if real_issues:
        st.subheader("❌ Найденные проблемы:")
        for issue in real_issues:
            icon = "•"
            if issue['type'] == 'static_number':
                icon = "🔴"
            elif issue['type'] == 'same_static_number':
                icon = "⚠️"
            st.write(f"{icon} {issue['message']}")
        
        # Рекомендация
        if results['has_static_numbering']:
            st.write("---")
            st.info("💡 **Как исправить:** В Word удалите статичную цифру и вставьте автонумерацию: Вставка → Номер страницы")
    else:
        st.success("✅ Нумерация страниц настроена правильно")
    
    # Начальный номер
    st.subheader("🔢 Начальный номер:")
    if results.get('start_page'):
        st.write(f"Установлен: **{results['start_page']}**")
        if results['start_page'] == 5:
            st.write("✅ Правильно (4 страницы до введения)")
        elif results['start_page'] == 1:
            st.write("⚠️ Начинается с 1 (должно быть 5)")
    else:
        st.write("⚠️ Не определён")
    
    # Отладка
    with st.expander("🔧 Отладка: XML колонтитулов"):
        for i, section_data in enumerate(results['sections'], 1):
            st.write(f"**Секция {i} - Нижний колонтитул:**")
            st.code(section_data['footer']['xml_info'][:1500], language='xml')


# Интерфейс
st.set_page_config(page_title="Проверка нумерации", layout="wide")
st.title("📄 Проверка нумерации страниц")
st.write("Анализ колонтитулов: автонумерация vs статичные номера")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализ..."):
        test_document(uploaded_file)
