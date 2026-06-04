# test_subsection_empty_line.py

import re

# Скопируем ТОЛЬКО необходимые функции из основного кода (упрощённо для тестов)
# В реальности можно импортировать из модуля, но для наглядности приведём минимум.

def normalize_text(text):
    return text.strip()

def is_section_header(norm_text):
    if norm_text.upper() in ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]:
        return True
    if re.match(r'^\d+\.\s*[А-ЯЁ]', norm_text):
        return True
    return False

def is_subsection(norm_text):
    # Упрощённая проверка: номер с точкой и буквой или слово без номера в оглавлении
    if re.match(r'^\d+\.\d+(\.\d+)?\s*[А-Яа-я]', norm_text):
        return True
    # Для тестов примем, что любой текст, начинающийся с цифр и точек, не являющийся разделом — подраздел
    if re.match(r'^\d+\.\d+', norm_text) and not is_section_header(norm_text):
        return True
    return False

# Сама логика, которую мы тестируем (взята из цикла по параграфам)
def process_paragraphs(paragraphs):
    """
    paragraphs: список кортежей (text, is_empty)
    Возвращает список ошибок вида "Подраздел «...» – уберите пустую строку перед подразделом"
    """
    errors = []
    prev_para_empty = False
    prev_nonempty_was_section_header = False   # последний непустой абзац был заголовком раздела
    
    for idx, (text, is_empty) in enumerate(paragraphs):
        if is_empty:
            prev_para_empty = True
            continue
        
        norm_text = normalize_text(text)
        
        # Определяем тип
        is_level1 = is_section_header(norm_text)
        is_sub = is_subsection(norm_text)
        
        # Обновляем флаг последнего непустого заголовка раздела ДО обработки текущего абзаца
        # (это важно: перед подразделом мы смотрим на предыдущий непустой)
        # Сначала обработаем ошибку для подраздела
        if is_sub:
            # Копия логики из кода: для подраздела проверяем пустую строку перед ним
            if prev_para_empty and not prev_nonempty_was_section_header:
                sub_name = re.sub(r'^\d+\.\d+(\.\d+)?\s*', '', norm_text).strip()
                key = f"Подраздел «{sub_name[:50]}»"
                errors.append(f"{key} – уберите пустую строку перед подразделом")
        
        # После обработки текущего абзаца обновляем флаг для следующего
        if is_level1:
            prev_nonempty_was_section_header = True
        else:
            # Если текущий абзац не является заголовком раздела, сбрасываем флаг
            # Но важно: сам подраздел не считается заголовком раздела
            prev_nonempty_was_section_header = False
        
        prev_para_empty = False
    
    return errors


# ------------------------------------------------------------
# ТЕСТЫ
# ------------------------------------------------------------

def run_test(case_name, paragraphs, expected_error_count, expected_substring=None):
    errors = process_paragraphs(paragraphs)
    if expected_error_count == 0:
        assert len(errors) == 0, f"{case_name}: expected 0 errors, got {errors}"
    else:
        assert len(errors) == expected_error_count, f"{case_name}: expected {expected_error_count} errors, got {len(errors)}"
        if expected_substring and errors:
            assert expected_substring in errors[0], f"{case_name}: error message does not contain '{expected_substring}'"
    print(f"✅ {case_name} passed")

# Сценарий 1: Раздел, пустая строка, подраздел
test1 = [
    ("1. ЛИТЕРАТУРНЫЙ ОБЗОР", False),
    ("", True),   # пустая строка
    ("1.1 Методология обзора", False),
]
run_test("Сценарий 1 (раздел -> пустая -> подраздел)", test1, 0)

# Сценарий 2: Раздел, подраздел без пустой строки
test2 = [
    ("1. ЛИТЕРАТУРНЫЙ ОБЗОР", False),
    ("1.1 Методология обзора", False),
]
run_test("Сценарий 2 (раздел -> подраздел без пустой)", test2, 0)

# Сценарий 3: Обычный текст, пустая строка, подраздел
test3 = [
    ("Какой-то обычный текст", False),
    ("", True),
    ("1.1 Методология обзора", False),
]
run_test("Сценарий 3 (обычный текст -> пустая -> подраздел)", test3, 1, "уберите пустую строку перед подразделом")

# Сценарий 4: Раздел, две пустые строки, подраздел (лишняя пустая строка не должна вызывать ошибку)
test4 = [
    ("1. ЛИТЕРАТУРНЫЙ ОБЗОР", False),
    ("", True),
    ("", True),
    ("1.1 Методология обзора", False),
]
run_test("Сценарий 4 (раздел -> две пустые -> подраздел)", test4, 0)

# Сценарий 5: Подраздел, пустая строка, вложенный подраздел (ошибка должна быть)
test5 = [
    ("1.1 Методология обзора", False),
    ("", True),
    ("1.1.1 Детализация", False),
]
run_test("Сценарий 5 (подраздел -> пустая -> вложенный подраздел)", test5, 1, "уберите пустую строку перед подразделом")

# Сценарий 6: Раздел, пустая строка, не подраздел (обычный текст) – никакой ошибки
test6 = [
    ("1. ЛИТЕРАТУРНЫЙ ОБЗОР", False),
    ("", True),
    ("Это просто абзац", False),
]
run_test("Сценарий 6 (раздел -> пустая -> обычный текст)", test6, 0)

print("\n🎉 Все тесты пройдены!")
