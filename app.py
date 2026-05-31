import streamlit as st
import docx
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree
import re

def has_page_number(text):
    if re.search(r'[\t\s\.]{2,}\d+$', text):
        return True
    return False

def is_all_caps(text):
    clean_text = re.sub(r'[\d\s\.,;:!?\-–—()«»""''\-\+–—]', '', text)
    if not clean_text:
        return False
    return clean_text == clean_text.upper()

def find_intro_start(doc):
    """Находит индекс начала введения"""
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt or has_page_number(txt):
            continue
        if txt.upper() == "ВВЕДЕНИЕ":
            return i
    return None

def count_pages_before_intro(doc, intro_idx):
    """Считает примерное количество страниц до введения"""
    # Считаем все непустые параграфы до введения
    # Грубая оценка: ~35 строк на страницу
    lines = 0
    for i in range(intro_idx):
        txt = doc.paragraphs[i].text.strip()
        if txt:
            lines += 1
    
    # + таблицы
    for table in doc.tables:
        try:
            tbl_pos = list(doc.element.body).index(table._element)
            if tbl_pos < intro_idx:
                lines += len(table.rows)
        except:
            pass
    
    pages = max(1, lines // 35 + 1)
    return pages

def get_footer_info(section):
    """Получает информацию из нижнего колонтитула"""
    info = {
        'has_content': False,
        'has_page_number': False,
        'is_auto_numbering': False,
        'is_static_number': False,
        'static_number_value': None,
        'page_number_text': '',
        'font_size': None,
        'alignment': None,
        'paragraphs': [],
        'empty_paragraphs_after': 0,
        'all_text': '',
        'xml_info': ''
    }
    
    try:
        footer = section.footer
        if footer is None:
            return info
        
        info['xml_info'] = etree.tostring(footer._element, encoding='unicode')
        paragraphs = footer.paragraphs
        
        for para in paragraphs:
            text = para.text.strip()
            para_info = {
                'text': text,
                'is_empty': not bool(text),
                'alignment': para.alignment
            }
            info['paragraphs'].append(para_info)
            info['all_text'] += text + ' '
            
            if text:
                info['has_content'] = True
            
            # Проверяем XML на поле PAGE
            para_xml = etree.tostring(para._element, encoding='unicode')
            has_page_field = 'PAGE' in para_xml and ('w:fldChar' in para_xml or 'w:instrText' in para_xml)
            
            if has_page_field:
                info['has_page_number'] = True
                info['is_auto_numbering'] = True
                info['page_number_text'] = '[автонумерация PAGE]'
                
                # Размер шрифта
                for run in para.runs:
                    if run.font.size:
                        info['font_size'] = run.font.size.pt
                        break
                
                # Выравнивание
                if para.alignment is not None:
                    info['alignment'] = para.alignment
            
            elif text and re.match(r'^\d+$', text):
                info['has_page_number'] = True
                info['is_static_number'] = True
                info['static_number_value'] = int(text)
                info['page_number_text'] = f'статичный номер: {text}'
                
                for run in para.runs:
                    if run.font.size and not info['font_size']:
                        info['font_size'] = run.font.size.pt
                
                if info['alignment'] is None and para.alignment is not None:
                    info['alignment'] = para.alignment
        
        # Проверяем пустые параграфы после номера
        found_number = False
        for para_info in info['paragraphs']:
            if para_info['text'] and ('PAGE' in str(para_info) or re.match(r'^\d+$', para_info['text'])):
                found_number = True
            elif found_number and para_info['is_empty']:
                info['empty_paragraphs_after'] += 1
    
    except Exception as e:
        info['xml_info'] = f'Ошибка: {str(e)[:200]}'
    
    return info

def analyze_page_numbering(doc):
    """Анализирует нумерацию страниц"""
    results = {
        'sections': [],
        'issues': [],
        'intro_idx': None,
        'expected_start_page': None
    }
    
    # Ищем введение
    intro_idx = find_intro_start(doc)
    results['intro_idx'] = intro_idx
    
    if intro_idx:
        pages_before = count_pages_before_intro(doc, intro_idx)
        results['expected_start_page'] = pages_before + 1
    else:
        results['expected_start_page'] = None
    
    # Анализируем секции
    for i, section in enumerate(doc.sections, 1):
        footer_info = get_footer_info(section)
        
        section_data = {
            'section_num': i,
            'footer': footer_info
        }
        
        results['sections'].append(section_data)
        
        # Проверки
        if footer_info['has_page_number']:
            if footer_info['is_static_number']:
                results['issues'].append({
                    'section': i,
                    'type': 'static_number',
                    'message': f"Секция {i}: статичный номер '{footer_info['static_number_value']}' (нужно поле PAGE для автонумерации)"
                })
            
            # Размер шрифта
            if footer_info['font_size']:
                if abs(footer_info['font_size'] - 14) > 0.5:
                    results['issues'].append({
                        'section': i,
                        'type': 'font_size',
                        'message': f"Секция {i}: размер шрифта нумерации {footer_info['font_size']} pt (должен быть 14 pt)"
                    })
            else:
                results['issues'].append({
                    'section': i,
                    'type': 'no_font_size',
                    'message': f"Секция {i}: не удалось определить размер шрифта нумерации"
                })
            
            # Выравнивание
            if footer_info['alignment'] is not None:
                if footer_info['alignment'] != WD_ALIGN_PARAGRAPH.CENTER:
                    align_names = {0: 'по левому краю', 1: 'по центру', 2: 'по правому краю', 3: 'по ширине'}
                    results['issues'].append({
                        'section': i,
                        'type': 'alignment',
                        'message': f"Секция {i}: нумерация выровнена {align_names.get(footer_info['alignment'], 'неизвестно')} (должно быть по центру)"
                    })
            
            # Пустые строки после номера
            if footer_info['empty_paragraphs_after'] > 0:
                results['issues'].append({
                    'section': i,
                    'type': 'empty_after',
                    'message': f"Секция {i}: есть пустая строка после номера в колонтитуле (нужно удалить)"
                })
        else:
            results['issues'].append({
                'section': i,
                'type': 'no_numbering',
                'message': f"Секция {i}: нумерация не найдена в нижнем колонтитуле"
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
    st.write("• Без пустых строк после номера в колонтитуле")
    st.write("• Начальный номер: первая страница введения = 5 (или другая, в зависимости от страниц до введения)")
    
    st.write("---")
    
    results = analyze_page_numbering(doc)
    
    # Информация о введении
    st.subheader("📑 Введение:")
    if results['intro_idx'] is not None:
        st.write(f"Найдено на строке: {results['intro_idx']}")
        st.write(f"Примерное количество страниц до введения: {results['expected_start_page'] - 1 if results['expected_start_page'] else '?'}")
        st.write(f"Ожидаемый номер первой страницы введения: **{results['expected_start_page']}**")
    else:
        st.warning("⚠️ Введение не найдено!")
    
    st.write("---")
    
    # Детали по секциям
    st.subheader(f"📊 Секций: {len(results['sections'])}")
    
    for section_data in results['sections']:
        sec_num = section_data['section_num']
        footer = section_data['footer']
        
        st.write(f"**Секция {sec_num}:**")
        
        if footer['has_content']:
            if footer['is_auto_numbering']:
                st.write(f"  ✅ Автонумерация (поле PAGE)")
            elif footer['is_static_number']:
                st.write(f"  ❌ Статичный номер: **{footer['static_number_value']}**")
            
            # Размер шрифта
            if footer['font_size']:
                font_ok = abs(footer['font_size'] - 14) < 0.5
                st.write(f"  Размер шрифта: {footer['font_size']} pt {'✅' if font_ok else '❌'}")
            else:
                st.write(f"  Размер шрифта: не определён")
            
            # Выравнивание
            if footer['alignment'] is not None:
                align_names = {0: 'LEFT', 1: 'CENTER', 2: 'RIGHT', 3: 'JUSTIFY'}
                st.write(f"  Выравнивание: {align_names.get(footer['alignment'], '?')} {'✅' if footer['alignment'] == 1 else '❌'}")
            
            # Пустые строки
            if footer['empty_paragraphs_after'] > 0:
                st.write(f"  ❌ Пустых строк после номера: {footer['empty_paragraphs_after']}")
            
            # Все параграфы колонтитула
            st.write(f"  Параграфов в колонтитуле: {len(footer['paragraphs'])}")
            for j, p in enumerate(footer['paragraphs']):
                status = "пустой" if p['is_empty'] else f"'{p['text']}'"
                st.write(f"    {j+1}. {status}")
        else:
            st.write(f"  ❌ Нижний колонтитул пуст")
        
        st.write("---")
    
    # Проблемы
    real_issues = [i for i in results['issues']]
    
    if real_issues:
        st.subheader("❌ Найденные проблемы:")
        for issue in real_issues:
            st.write(f"• {issue['message']}")
    else:
        st.success("✅ Нумерация настроена правильно!")
    
    # Начальный номер
    st.subheader("🔢 Начальный номер страницы:")
    if results.get('start_page'):
        st.write(f"Установлен в документе: **{results['start_page']}**")
        expected = results.get('expected_start_page')
        if expected and results['start_page'] != expected:
            st.error(f"❌ Должен быть **{expected}** (первая страница введения), а установлен **{results['start_page']}**")
        elif expected:
            st.success(f"✅ Правильно ({expected})")
    else:
        expected = results.get('expected_start_page')
        if expected:
            st.warning(f"⚠️ Не удалось определить. Ожидаемый: **{expected}**")
    
    # Отладка
    with st.expander("🔧 Отладка: XML колонтитулов"):
        for i, section_data in enumerate(results['sections'], 1):
            st.write(f"**Секция {i}:**")
            st.code(section_data['footer']['xml_info'][:1500], language='xml')


# Интерфейс
st.set_page_config(page_title="Проверка нумерации", layout="wide")
st.title("📄 Проверка нумерации страниц")

uploaded_file = st.file_uploader("Загрузите .docx", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Анализ..."):
        test_document(uploaded_file)
