import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import re

def get_effective_alignment(paragraph):
    if paragraph.alignment is not None:
        return paragraph.alignment
    try:
        style = paragraph.style
        if style and style.paragraph_format.alignment is not None:
            return style.paragraph_format.alignment
    except:
        pass
    return None

def is_paragraph_bold(paragraph):
    try:
        if paragraph.style and paragraph.style.font and paragraph.style.font.bold:
            return True
    except:
        pass
    runs = [r for r in paragraph.runs if r.text.strip()]
    if not runs:
        return False
    return all(r.bold for r in runs)

def is_empty_paragraph(paragraph):
    return len(paragraph.text.strip()) == 0

def is_all_caps(text):
    return bool(re.match(r'^[А-ЯЁ\s\-]+$', text)) and len(text) > 3

def find_content_heading_ultimate(doc):
    """
    УЛЬТИМАТИВНЫЙ поиск заголовка «СОДЕРЖАНИЕ»
    Использует все возможные методы
    """
    
    print(f"Всего параграфов в документе: {len(doc.paragraphs)}")
    
    # Метод 1: Прямой поиск в тексте параграфа
    for i, p in enumerate(doc.paragraphs):
        text = p.text
        # Ищем слово "Содержание" в любом виде
        if 'Содержание' in text or 'СОДЕРЖАНИЕ' in text or 'содержание' in text:
            print(f"Метод 1: Найдено в параграфе {i}, текст: '{text[:100]}'")
            return i
    
    # Метод 2: Поиск через XML (самый надежный)
    for i, p in enumerate(doc.paragraphs):
        # Получаем XML элемента
        xml_str = p._element.xml if hasattr(p, '_element') else ''
        if 'Содержание' in xml_str or 'СОДЕРЖАНИЕ' in xml_str:
            print(f"Метод 2: Найдено в XML параграфа {i}")
            return i
    
    # Метод 3: Поиск по runs с объединением
    for i, p in enumerate(doc.paragraphs):
        full_text = ''
        for run in p.runs:
            full_text += run.text
        if 'Содержание' in full_text or 'СОДЕРЖАНИЕ' in full_text:
            print(f"Метод 3: Найдено в runs параграфа {i}, текст: '{full_text[:100]}'")
            return i
    
    # Метод 4: Удаляем все не-буквы и ищем
    for i, p in enumerate(doc.paragraphs):
        # Удаляем всё, кроме букв и цифр
        cleaned = re.sub(r'[^А-Яа-яA-Za-z0-9]', '', p.text)
        if 'Содержание' in cleaned or 'СОДЕРЖАНИЕ' in cleaned:
            print(f"Метод 4: Найдено после очистки в параграфе {i}")
            return i
    
    # Метод 5: Поиск по словам (разбиваем на слова)
    for i, p in enumerate(doc.paragraphs):
        words = p.text.split()
        for word in words:
            # Очищаем слово от знаков препинания
            clean_word = re.sub(r'[^\w]', '', word)
            if clean_word in ['Содержание', 'СОДЕРЖАНИЕ', 'содержание']:
                print(f"Метод 5: Найдено слово '{word}' в параграфе {i}")
                return i
    
    # Метод 6: Поиск по шаблону (регулярное выражение)
    pattern = re.compile(r'[СC][ОO][ДD][ЕE][РP][ЖЖ][АA][НH][ИI][ЕE]', re.IGNORECASE)
    for i, p in enumerate(doc.paragraphs):
        if pattern.search(p.text):
            print(f"Метод 6: Найдено по regex в параграфе {i}")
            return i
    
    # Метод 7: Проверка первых 50 параграфов (содержание обычно в начале)
    for i in range(min(50, len(doc.paragraphs))):
        text_lower = doc.paragraphs[i].text.lower()
        if 'содерж' in text_lower or 'оглавл' in text_lower:
            print(f"Метод 7: Найдено по части слова в параграфе {i}")
            return i
    
    return None

def check_word_document(file):
    doc = docx.Document(file)
    issues = []
    
    # Создаем placeholder для отладки
    debug_info = []
    
    # ---------- 1. ПОЛЯ СТРАНИЦ ----------
    margins_ok = True
    for section in doc.sections:
        if (abs(section.left_margin.mm - 20) > 0.5 or
            abs(section.right_margin.mm - 20) > 0.5 or
            abs(section.top_margin.mm - 20) > 0.5 or
            abs(section.bottom_margin.mm - 20) > 0.5):
            margins_ok = False
            break
    if not margins_ok:
        issues.append("Поля страниц – установите левое 20 мм, правое 20 мм, верхнее 20 мм, нижнее 20 мм")

    # ---------- 2. ПОИСК "СОДЕРЖАНИЕ" (УЛЬТИМАТИВНЫЙ) ----------
    content_idx = find_content_heading_ultimate(doc)
    
    # Добавляем отладочную информацию
    debug_info.append(f"📊 Всего параграфов в документе: {len(doc.paragraphs)}")
    debug_info.append("")
    debug_info.append("📋 Первые 20 параграфов документа:")
    
    for i in range(min(20, len(doc.paragraphs))):
        p = doc.paragraphs[i]
        text_preview = p.text[:100] if p.text else "[пусто]"
        # Показываем также repr для выявления скрытых символов
        debug_info.append(f"  {i:3d}: {repr(text_preview)}")
    
    if content_idx is None:
        issues.append("❌ Не найден заголовок «СОДЕРЖАНИЕ» — проверка невозможна.")
        
        # Показываем отладочную информацию
        for line in debug_info:
            issues.append(line)
        
        # Дополнительный анализ XML первых 5 параграфов
        issues.append("")
        issues.append("🔍 Детальный XML-анализ первых 5 параграфов:")
        for i in range(min(5, len(doc.paragraphs))):
            p = doc.paragraphs[i]
            if p.text:
                # Получаем XML параграфа
                xml_str = p._element.xml if hasattr(p, '_element') else ""
                # Ищем в XML слово "Содержание"
                if 'Содержание' in xml_str or 'содержание' in xml_str:
                    issues.append(f"  Параграф {i}: НАЙДЕНО 'Содержание' в XML!")
                    issues.append(f"    XML фрагмент: {xml_str[:200]}")
                else:
                    # Показываем первые символы XML
                    issues.append(f"  Параграф {i}: XML начинается с {xml_str[:100]}")
        
        return issues
    
    issues.append(f"✅ Заголовок «СОДЕРЖАНИЕ» найден в абзаце {content_idx}")
    issues.append(f"   Текст абзаца: {repr(doc.paragraphs[content_idx].text[:100])}")
    
    # ---------- 3. ПРОВЕРКА ЗАГОЛОВКА "СОДЕРЖАНИЕ" ----------
    p_content = doc.paragraphs[content_idx]
    text = p_content.text.strip()
    
    # Очищаем от скрытых символов
    clean_text = re.sub(r'[^\w\s]', '', text).strip()
    
    if text.endswith("."):
        issues.append("Содержание – удалите точку в конце")
    
    # Проверяем выравнивание (игнорируем если не удалось определить)
    alignment = get_effective_alignment(p_content)
    if alignment is not None and alignment != WD_ALIGN_PARAGRAPH.CENTER:
        issues.append("Содержание – выровняйте слово СОДЕРЖАНИЕ по центру")
    
    # Проверка жирности
    if not is_paragraph_bold(p_content):
        issues.append("Содержание – сделайте заголовок полужирным")
    
    # Проверка пустой строки после
    empty_found = False
    for i in range(content_idx + 1, min(content_idx + 5, len(doc.paragraphs))):
        if is_empty_paragraph(doc.paragraphs[i]):
            empty_found = True
            break
    
    if not empty_found and content_idx + 1 < len(doc.paragraphs):
        issues.append("Содержание – после заголовка должна быть пустая строка")

    # ---------- 4. ПОИСК ПЕРВОГО ЗАГОЛОВКА ----------
    start_idx = None
    level1_keywords = {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"}
    
    for i in range(content_idx + 1, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        txt = p.text.strip()
        if not txt:
            continue
        
        if re.match(r'^\d+\.\s+[А-Я]', txt):
            start_idx = i
            break
        
        if txt.upper() in level1_keywords:
            start_idx = i
            break
        
        if is_all_caps(txt) and get_effective_alignment(p) == WD_ALIGN_PARAGRAPH.CENTER:
            start_idx = i
            break
    
    if start_idx is None:
        issues.append("❌ Не найден ни один заголовок раздела после содержания.")
        return issues

    # ---------- ОСТАЛЬНЫЕ ПРОВЕРКИ ----------
    issues = list(dict.fromkeys(issues))
    return issues

# ========== ИНТЕРФЕЙС STREAMLIT ==========
st.set_page_config(page_title="Нормоконтроль документов", layout="wide")

st.title("📊 Автоматическая проверка документов Word")
st.markdown("---")

uploaded_file = st.file_uploader("📂 Выберите файл для проверки", type=["docx"])

if uploaded_file is not None:
    with st.spinner("🔍 Проверяем документ..."):
        try:
            results = check_word_document(uploaded_file)
            
            st.subheader("📋 Результаты проверки:")
            
            for r in results:
                if r.startswith('✅'):
                    st.success(r)
                elif r.startswith('❌'):
                    st.error(r)
                elif r.startswith('📊') or r.startswith('📋') or r.startswith('🔍'):
                    st.info(r)
                else:
                    st.warning(f"⚠️ {r}")
                    
        except Exception as e:
            st.error(f"❌ Ошибка при проверке документа: {str(e)}")
            st.exception(e)
else:
    st.info("👈 Загрузите документ для начала проверки")
