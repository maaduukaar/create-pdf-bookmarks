#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Извлечение структуры заголовков из СОДЕРЖАНИЯ DOCX в JSON для PDF закладок.

Формат JSON совместим с PyMuPDF для создания закладок в PDF.

Режимы работы:
1. Автоматический (по умолчанию) - парсинг DOCX с регуляркой
2. Ручной (--toc-pages) - указание страниц PDF с оглавлением для извлечения
"""

import sys
import os
import json
import re
import argparse

import docx

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False


# Регулярка для строк содержания (оригинальный формат)
TOC_LINE_PATTERN = re.compile(
    r"^(\d+(?:\.\d+)*)\.\s+(.+?)\s+(\d+)$"
)

# Дополнительные паттерны для разных форматов оглавления
# Формат: "Заголовок ...... 15" или "Заголовок 15"
TOC_LINE_PATTERN_ALT1 = re.compile(
    r"^(.+?)\s*\.{2,}\s*(\d+)\s*$"  # С точками-лидерами
)

TOC_LINE_PATTERN_ALT2 = re.compile(
    r"^(.+?)\s{2,}(\d+)\s*$"  # Без точек, но с пробелами
)

# Паттерн для номера раздела в начале заголовка
SECTION_NUMBER_PATTERN = re.compile(
    r"^(\d+(?:\.\d+)*)\.\s*(.+)$"
)


def parse_toc_line_from_pdf(text: str):
    """
    Распарсить строку содержания из PDF с более гибкой логикой.
    
    Поддерживаемые форматы:
    - "1.2.3 Название раздела ...... 42"
    - "Название раздела ...... 42"
    - "1.2.3 Название раздела    42"
    
    Возвращает (заголовок, уровень, страница) или (None, None, None)
    """
    text = text.strip()
    if not text or len(text) < 3:
        return None, None, None
    
    title = None
    page_number = None
    section_number = None
    
    # Пробуем оригинальный паттерн
    m = TOC_LINE_PATTERN.match(text)
    if m:
        section_number = m.group(1)
        title = m.group(2).strip()
        page_number = int(m.group(3))
    else:
        # Пробуем паттерн с точками-лидерами
        m = TOC_LINE_PATTERN_ALT1.match(text)
        if m:
            title = m.group(1).strip()
            page_number = int(m.group(2))
        else:
            # Пробуем паттерн с пробелами
            m = TOC_LINE_PATTERN_ALT2.match(text)
            if m:
                title = m.group(1).strip()
                page_number = int(m.group(2))
    
    if title is None or page_number is None:
        return None, None, None
    
    # Убираем точки-лидеры из середины если остались
    title = re.sub(r'\.{2,}', '', title).strip()
    
    # Пробуем извлечь номер раздела из заголовка
    if section_number is None:
        m_section = SECTION_NUMBER_PATTERN.match(title)
        if m_section:
            section_number = m_section.group(1)
            title = m_section.group(2).strip()
    
    # Определяем уровень
    if section_number:
        level = section_number.count(".") + 1
        full_title = f"{section_number} {title}"
    else:
        level = 1  # Без номера считаем уровнем 1
        full_title = title
    
    return full_title, level, page_number


def extract_toc_from_pdf_pages(pdf_path: str, page_numbers: list, show_output: bool = True):
    """
    Извлечь строки оглавления из указанных страниц PDF.
    Определяет уровень вложенности по отступам (X-координатам).
    
    Args:
        pdf_path: путь к PDF файлу
        page_numbers: список номеров страниц (1-indexed)
        show_output: показывать ли отладочный вывод
    
    Returns:
        список записей [{title, level, page}, ...]
    """
    if not PYMUPDF_AVAILABLE:
        print("[!] Библиотека PyMuPDF не установлена!")
        print("[*] Установи: pip install PyMuPDF")
        return []
    
    if not os.path.isfile(pdf_path):
        print(f"[!] PDF файл не найден: {pdf_path}")
        return []
    
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"[!] Ошибка открытия PDF: {e}")
        return []
    
    total_pages = len(doc)
    
    if show_output:
        print(f"\n[INFO] Всего страниц в PDF: {total_pages}")
        print(f"[*] Читаю страницы оглавления: {page_numbers}")
        print("[*] Определяю уровень вложенности по отступам...")
    
    # Собираем все строки со всех указанных страниц
    all_lines = []
    
    for page_num in page_numbers:
        if page_num < 1 or page_num > total_pages:
            print(f"[!] Страница {page_num} вне диапазона (1-{total_pages})")
            continue
        
        # PyMuPDF использует 0-индексацию
        page = doc[page_num - 1]
        
        # Получаем текст с координатами блоков
        blocks = page.get_text("dict")["blocks"]
        
        for block in blocks:
            if "lines" not in block:
                continue
            
            for line in block["lines"]:
                # Собираем текст из всех spans в линии
                line_text = ""
                x_coord = None
                
                for span in line["spans"]:
                    line_text += span["text"]
                    # Берём X-координату первого символа
                    if x_coord is None:
                        x_coord = span["bbox"][0]  # bbox = [x0, y0, x1, y1]
                
                line_text = line_text.strip()
                
                if line_text and x_coord is not None:
                    all_lines.append({
                        "text": line_text,
                        "x": x_coord,
                        "page_num": page_num
                    })
    
    doc.close()
    
    if not all_lines:
        return []
    
    # Определяем уровни по отступам
    entries = []
    x_to_level = {}  # Маппинг X-координаты на уровень
    current_level = 1
    
    # Сортируем уникальные X-координаты
    unique_x = sorted(set(line["x"] for line in all_lines))
    
    # Группируем близкие X-координаты (разница < 5 пикселей)
    x_groups = []
    if unique_x:
        current_group = [unique_x[0]]
        for x in unique_x[1:]:
            if x - current_group[-1] < 5:
                current_group.append(x)
            else:
                x_groups.append(current_group)
                current_group = [x]
        x_groups.append(current_group)
    
    # Назначаем уровни группам (чем левее, тем меньше уровень)
    for level, group in enumerate(x_groups, start=1):
        for x in group:
            x_to_level[x] = level
    
    if show_output:
        print(f"\n[DEBUG] Найдено уровней отступов: {len(x_groups)}")
        for level, group in enumerate(x_groups, start=1):
            avg_x = sum(group) / len(group)
            print(f"  Уровень {level}: X ≈ {avg_x:.1f}px")
    
    # Парсим каждую строку
    if show_output:
        print(f"\n[*] Разбор строк оглавления:")
        print("-" * 60)
    
    for line_info in all_lines:
        text = line_info["text"]
        x = line_info["x"]
        
        # Парсим строку
        full_title, auto_level, page_number = parse_toc_line_from_pdf(text)
        
        if full_title is not None:
            # Если у строки есть номер раздела, используем его уровень
            # Иначе используем уровень по отступу
            if auto_level > 1:
                # Есть номер раздела (1.2.3) - доверяем ему
                final_level = auto_level
            else:
                # Нет номера - используем отступ
                final_level = x_to_level.get(x, 1)
            
            entries.append({
                "title": full_title,
                "level": final_level,
                "page": page_number
            })
            
            if show_output:
                indent = "  " * (final_level - 1)
                print(f"{indent}[L{final_level}] {full_title} -> стр. {page_number}")
        elif show_output and len(text) > 3:
            # Показываем строки, которые не распознались
            print(f"[-] Пропущено: {text[:60]}...")
    
    if show_output:
        print("-" * 60)
    
    return entries


def get_toc_pages_interactively(pdf_path: str):
    """
    Интерактивно запросить у пользователя номера страниц с оглавлением.
    
    Returns:
        список номеров страниц или None при отмене
    """
    print("\n" + "=" * 60)
    print("РЕЖИМ РУЧНОГО УКАЗАНИЯ СТРАНИЦ ОГЛАВЛЕНИЯ")
    print("=" * 60)
    
    if not PYMUPDF_AVAILABLE:
        print("[!] Библиотека PyMuPDF не установлена!")
        print("[*] Установи: pip install PyMuPDF")
        return None
    
    if pdf_path and os.path.isfile(pdf_path):
        try:
            doc = fitz.open(pdf_path)
            total = len(doc)
            doc.close()
            print(f"\n[PDF] {os.path.basename(pdf_path)}")
            print(f"[INFO] Всего страниц: {total}")
        except:
            pass
    
    print("\n[?] Введи номера страниц с оглавлением.")
    print("    Форматы: '2' или '2,3' или '2-4' или '2,3,5-7'")
    print("    Введи 'q' для отмены.\n")
    
    while True:
        answer = input("[СТРАНИЦЫ] > ").strip()
        
        if answer.lower() in ('q', 'quit', 'exit', 'н', 'нет'):
            print("[-] Отмена.")
            return None
        
        if not answer:
            print("[!] Введи хотя бы один номер страницы")
            continue
        
        # Парсим номера страниц
        try:
            pages = parse_page_range(answer)
            if pages:
                print(f"[+] Выбраны страницы: {pages}")
                return pages
            else:
                print("[!] Не удалось разобрать номера страниц")
        except ValueError as e:
            print(f"[!] Ошибка: {e}")


def parse_page_range(page_str: str):
    """
    Разобрать строку с номерами страниц.
    
    Форматы: '2', '2,3', '2-4', '2,3,5-7'
    
    Returns:
        отсортированный список уникальных номеров страниц
    """
    pages = set()
    
    parts = page_str.replace(' ', '').split(',')
    for part in parts:
        if '-' in part:
            # Диапазон
            try:
                start, end = part.split('-', 1)
                start = int(start)
                end = int(end)
                if start > end:
                    start, end = end, start
                for p in range(start, end + 1):
                    pages.add(p)
            except ValueError:
                raise ValueError(f"Неверный диапазон: {part}")
        else:
            # Одиночный номер
            try:
                pages.add(int(part))
            except ValueError:
                raise ValueError(f"Неверный номер страницы: {part}")
    
    return sorted(pages)


def parse_toc_line(text: str):
    """
    Распарсить строку содержания.
    
    Возвращает (номер_раздела, заголовок, страница) или (None, None, None)
    """
    m = TOC_LINE_PATTERN.match(text.strip())
    if not m:
        return None, None, None
    
    section_number = m.group(1)
    title_text = m.group(2).strip()
    page_number = int(m.group(3))
    
    # Убираем точки-лидеры
    title_text = re.sub(r'\.{2,}', '', title_text).strip()
    
    # Полный заголовок с номером
    full_title = f"{section_number} {title_text}"
    
    # Уровень = количество точек в номере + 1
    level = section_number.count(".") + 1
    
    return full_title, level, page_number


def extract_toc_entries(doc_path: str):
    """Извлечь заголовки из содержания DOCX."""
    document = docx.Document(doc_path)
    
    entries = []
    
    for para in document.paragraphs:
        text = (para.text or "").strip()
        if not text:
            continue
        
        full_title, level, page = parse_toc_line(text)
        
        if full_title is not None:
            entries.append({
                "title": full_title,
                "level": level,
                "page": page
            })
    
    return entries


def build_bookmark_tree(entries):
    """
    Построить дерево закладок в формате PyMuPDF.
    
    Структура каждого узла:
    {
        "title": "Название",
        "dest": [page, "XYZ", x, y, 0] или [page, "Fit"],
        "color": {"0": 0, "1": 0, "2": 0},
        "bold": false,
        "italic": false,
        "children": [...]
    }
    """
    root = []
    last_nodes = {}
    
    for entry in entries:
        level = max(1, min(int(entry["level"]), 9))
        
        # Определяем destination: с координатами или без
        if "coords" in entry and entry["coords"]:
            # Точный переход к заголовку
            coords = entry["coords"]
            # Формат: [page, "XYZ", x, y, zoom]
            # y координата в PDF начинается снизу, но фактически используется top
            dest = [entry["page"], "XYZ", coords["x"], coords["y"], 0]
        else:
            # Fallback: переход к странице целиком
            dest = [entry["page"], "Fit"]
        
        # Создаем узел закладки
        node = {
            "title": entry["title"],
            "dest": dest,
            "color": {
                "0": 0,
                "1": 0,
                "2": 0
            },
            "bold": False,
            "italic": False,
            "children": []
        }
        
        # Строим иерархию
        if level == 1:
            root.append(node)
        else:
            parent = last_nodes.get(level - 1)
            if parent is not None:
                parent["children"].append(node)
            else:
                # Если нет родителя - добавляем в корень
                root.append(node)
        
        last_nodes[level] = node
    
    return root


def find_pdf_for_docx(docx_path: str):
    """
    Найти PDF файл с тем же именем в той же папке.
    
    Возвращает путь к PDF или None.
    """
    base, ext = os.path.splitext(docx_path)
    pdf_path = base + ".pdf"
    
    if os.path.isfile(pdf_path):
        return pdf_path
    return None


def embed_bookmarks_to_pdf(pdf_path: str, json_path: str, show_output: bool = True):
    """
    Встроить закладки из JSON в PDF файл.
    
    Args:
        pdf_path: путь к PDF файлу
        json_path: путь к JSON с закладками
        show_output: показывать ли вывод
    
    Returns:
        True если успешно, False иначе
    """
    if not PYMUPDF_AVAILABLE:
        print("\n[!] Библиотека PyMuPDF не установлена!")
        print("\n[*] Установи её командой:")
        print("   pip install PyMuPDF")
        return False
    
    if not os.path.isfile(pdf_path):
        print(f"\n[!] PDF файл не найден: {pdf_path}")
        return False
    
    if not os.path.isfile(json_path):
        print(f"\n[!] JSON файл не найден: {json_path}")
        return False
    
    if show_output:
        print("\n" + "=" * 60)
        print("ВСТРАИВАНИЕ ЗАКЛАДОК В PDF")
        print("=" * 60)
        print(f"\n[PDF] {os.path.basename(pdf_path)}")
        print(f"[JSON] {os.path.basename(json_path)}")
    
    # Читаем закладки из JSON
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            bookmarks = json.load(f)
    except Exception as e:
        print(f"\n[!] Ошибка чтения JSON: {e}")
        return False
    
    # Открываем PDF
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"\n[!] Ошибка открытия PDF: {e}")
        return False
    
    if show_output:
        print(f"\n[INFO] Страниц в PDF: {len(doc)}")
        print("\n[*] Встраиваю закладки...")
    
    # Удаляем старые закладки
    try:
        doc.set_toc([])  # Очищаем оглавление
    except:
        pass
    
    # Конвертируем дерево закладок в формат PyMuPDF TOC
    def tree_to_toc(nodes, toc_list, parent_level=0):
        """
        Рекурсивно конвертировать дерево закладок в список для PyMuPDF.
        
        Формат TOC для PyMuPDF: [level, title, page, dest_dict]
        dest_dict может содержать:
          - {"kind": 1, "to": fitz.Point(x, y)} для точного перехода
          - или просто пустой для перехода к странице
        """
        for node in nodes:
            level = parent_level + 1
            title = node.get("title", "Untitled")
            
            # Получаем destination
            dest = node.get("dest", [])
            page = 1
            dest_dict = {}
            
            if isinstance(dest, list) and len(dest) > 0:
                page = dest[0]
                
                # Проверяем тип destination
                if len(dest) >= 4 and dest[1] == "XYZ":
                    # Формат: [page, "XYZ", x, y, zoom]
                    x = dest[2] if len(dest) > 2 else 0
                    y = dest[3] if len(dest) > 3 else 0
                    # Для PyMuPDF используем kind=1 (goto) и точку
                    dest_dict = {
                        "kind": 1,  # LINK_GOTO
                        "page": page - 1,  # 0-indexed для fitz
                        "to": fitz.Point(x, y),
                        "zoom": 0  # 0 = сохранить текущий зум
                    }
            
            # Преобразуем номер страницы
            page = max(1, min(page, len(doc)))
            
            # Добавляем закладку в список TOC
            # PyMuPDF: [level, title, page] или [level, title, page, dest_dict]
            if dest_dict:
                toc_list.append([level, title, page, dest_dict])
            else:
                toc_list.append([level, title, page])
            
            # Рекурсивно добавляем детей
            children = node.get("children", [])
            if children:
                tree_to_toc(children, toc_list, level)

    
    toc = []
    tree_to_toc(bookmarks, toc)
    
    if show_output:
        print(f"[+] Подготовлено закладок: {len(toc)}")
    
    # Встраиваем закладки в PDF
    try:
        doc.set_toc(toc)
    except Exception as e:
        print(f"\n[!] Ошибка встраивания закладок: {e}")
        doc.close()
        return False
    
    # Сохраняем PDF с закладками
    base, ext = os.path.splitext(pdf_path)
    output_path = base + "_with_bookmarks.pdf"
    
    try:
        doc.save(output_path, garbage=4, deflate=True)
        doc.close()
    except Exception as e:
        print(f"\n[!] Ошибка сохранения PDF: {e}")
        doc.close()
        return False
    
    if show_output:
        print("\n" + "=" * 60)
        print("[OK] ЗАКЛАДКИ ВСТРОЕНЫ!")
        print("=" * 60)
        print(f"\n[>>] Создан файл: {output_path}")
        print(f"\n[STATS] Статистика:")
        print(f"   - Встроено закладок: {len(toc)}")
        print(f"   - Исходный PDF: {os.path.basename(pdf_path)}")
        print(f"   - Новый PDF: {os.path.basename(output_path)}")
    
    return True


def ask_embed_bookmarks(docx_path: str, json_path: str):
    """
    Автоматически встроить закладки в PDF или предложить пользователю указать путь.
    
    Args:
        docx_path: путь к исходному DOCX
        json_path: путь к созданному JSON
    """
    if not PYMUPDF_AVAILABLE:
        print("\n[!] PyMuPDF не установлен - встраивание закладок недоступно.")
        print("[*] Установи: pip install PyMuPDF")
        return
    
    # Ищем PDF с тем же именем
    pdf_path = find_pdf_for_docx(docx_path)
    
    if pdf_path:
        # PDF найден - встраиваем автоматически
        print("\n" + "=" * 60)
        print("АВТОМАТИЧЕСКОЕ ВСТРАИВАНИЕ ЗАКЛАДОК")
        print("=" * 60)
        print(f"\n[+] Найден PDF файл: {os.path.basename(pdf_path)}")
        print("[*] Автоматически встраиваю закладки...")
        embed_bookmarks_to_pdf(pdf_path, json_path)
    else:
        # PDF не найден - запрашиваем у пользователя
        print("\n" + "=" * 60)
        print("ВСТРАИВАНИЕ ЗАКЛАДОК В PDF")
        print("=" * 60)
        print(f"\n[!] PDF файл с именем '{os.path.splitext(os.path.basename(docx_path))[0]}.pdf' не найден.")
        
        while True:
            answer = input("\n[?] Введи путь к PDF файлу (или 'n' для отказа): ").strip()
            
            if answer.lower() in ('n', 'no', 'н', 'нет', ''):
                print("[-] Пропускаю встраивание закладок.")
                return
            
            pdf_path = answer.strip('"\'')
            if os.path.isfile(pdf_path) and pdf_path.lower().endswith('.pdf'):
                embed_bookmarks_to_pdf(pdf_path, json_path)
                return
            else:
                print(f"[!] Файл не найден или не является PDF: {pdf_path}")
                print("[*] Попробуй ещё раз или введи 'n' для отказа")


def process_docx(docx_path: str, show_output: bool = True):
    """Основная логика обработки DOCX файла."""
    
    if not os.path.isfile(docx_path):
        print(f"[!] Файл не найден: {docx_path}")
        return False
    
    if not docx_path.lower().endswith(".docx"):
        print("[!] Ожидается DOCX-файл (.docx).")
        return False
    
    if show_output:
        print("=" * 60)
        print("ИЗВЛЕЧЕНИЕ ЗАКЛАДОК ИЗ СОДЕРЖАНИЯ DOCX")
        print("=" * 60)
        print(f"\n[FILE] {os.path.basename(docx_path)}")
    
    if show_output:
        print("\n[*] Читаю содержание из DOCX...")
    
    try:
        entries = extract_toc_entries(docx_path)
    except Exception as e:
        print(f"\n[!] Ошибка при чтении файла: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    if not entries:
        print("\n[!] Строки содержания не найдены!")
        print("\n[*] Формат строк должен быть:")
        print("   '3.4.2.1 Название раздела 69'")
        print("   где 3.4.2.1 - номер раздела, 69 - номер страницы")
        return False
    
    if show_output:
        print(f"\n[+] Найдено заголовков: {len(entries)}\n")
        
        print("Структура закладок:")
        print("-" * 60)
        for entry in entries[:15]:
            indent = "  " * (entry["level"] - 1)
            print(f"{indent}[>>] {entry['title']} -> стр. {entry['page']}")
        
        if len(entries) > 15:
            print(f"   ... и ещё {len(entries) - 15} заголовков")
        
        print("-" * 60)
    
    # Строим дерево закладок
    if show_output:
        print("\n[*] Строю иерархическое дерево закладок...")
    
    tree = build_bookmark_tree(entries)
    
    # Сохраняем JSON
    base, ext = os.path.splitext(docx_path)
    out_path = base + "_bookmarks.json"
    
    if show_output:
        print(f"\n[*] Сохраняю JSON: {os.path.basename(out_path)}")
    
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(tree, f, ensure_ascii=False, indent=2)
    
    if show_output:
        print("\n" + "=" * 60)
        print("[OK] ГОТОВО!")
        print("=" * 60)
        print(f"\n[>>] Создан файл: {out_path}")
        print("\n[INFO] Структура JSON:")
        print("   - title: название закладки")
        print("   - dest: [страница, 'Fit'] - переход к странице")
        print("   - color, bold, italic: стиль закладки")
        print("   - children: вложенные закладки")
        print("\n[!] Примечание: координаты не установлены (используется 'Fit')")
        print("   Для точного позиционирования нужен PDF файл.")
    
    # Предлагаем встроить закладки в PDF
    if show_output:
        ask_embed_bookmarks(docx_path, out_path)
    
    return True


def get_file_interactively():
    """Запросить путь к файлу интерактивно."""
    print("=" * 60)
    print("ИЗВЛЕЧЕНИЕ ЗАКЛАДОК ИЗ СОДЕРЖАНИЯ DOCX")
    print("=" * 60)
    print("\nРежимы запуска:")
    print("  1. Drag & Drop: перетащи DOCX на скрипт")
    print("  2. Командная строка: python script.py файл.docx")
    print("  3. Интерактивный: введи путь ниже\n")
    
    while True:
        file_path = input("[?] Введи путь к DOCX-файлу (или 'q' для выхода): ").strip()
        
        if file_path.lower() in ('q', 'quit', 'exit'):
            print("[-] Выход...")
            return None
        
        file_path = file_path.strip('"\'')
        
        if os.path.isfile(file_path):
            return file_path
        else:
            print(f"[!] Файл не найден: {file_path}")
            print("[*] Попробуй ещё раз или введи 'q' для выхода\n")


def main():
    parser = argparse.ArgumentParser(
        description="Извлечение закладок из содержания DOCX/PDF в JSON",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python %(prog)s document.docx              # Автоматический режим (DOCX)
  python %(prog)s document.pdf --toc-pages   # Ручной режим (интерактивный)
  python %(prog)s document.pdf -t 2          # Указать страницу 2
  python %(prog)s document.pdf -t 2,3,5-7    # Указать несколько страниц
  python %(prog)s --quiet file.docx          # Тихий режим

Формат содержания (автоматический):
  1. Название раздела 5
  1.1 Подраздел 12
  3.4.2.1 Интерфейс раздела 69

Режим --toc-pages:
  Используй когда регулярка не справляется с форматом оглавления.
  Укажи страницы PDF с оглавлением, и скрипт извлечёт заголовки.
  Поддерживает форматы: "Заголовок...... 15", "1.2 Раздел    42"
   
Выходной JSON содержит:
  - Заголовки с правильной нумерацией
  - Номера страниц для навигации
  - Иерархическую структуру (children)
  - Базовые атрибуты (color, bold, italic)
        """
    )
    
    parser.add_argument(
        'file',
        nargs='?',
        help='Путь к DOCX или PDF файлу'
    )
    
    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Тихий режим (минимум вывода)'
    )
    
    parser.add_argument(
        '-t', '--toc-pages',
        nargs='?',
        const='interactive',
        metavar='PAGES',
        help='Режим указания страниц оглавления в PDF. '
             'Без значения - интерактивный запрос. '
             'С значением: номера страниц (2 или 2,3 или 2-4)'
    )
    
    args = parser.parse_args()
    
    # Определяем режим работы
    if args.toc_pages is not None:
        # Режим работы с PDF и ручным указанием страниц
        success = run_toc_pages_mode(args)
    else:
        # Стандартный режим работы с DOCX
        if args.file:
            docx_path = args.file
        else:
            docx_path = get_file_interactively()
            if docx_path is None:
                return
        
        success = process_docx(docx_path, show_output=not args.quiet)
    
    if not sys.stdin.isatty():
        input("\n[PAUSE] Нажми Enter для выхода...")
    
    sys.exit(0 if success else 1)


def find_exact_coordinates(pdf_path: str, entries: list, show_output: bool = True):
    """
    Найти точные координаты заголовков в PDF документе.
    
    Для каждой записи из оглавления ищет соответствующий заголовок
    на указанной странице и получает его координаты.
    
    Args:
        pdf_path: путь к PDF файлу
        entries: список записей [{title, level, page}, ...]
        show_output: показывать ли отладочный вывод
    
    Returns:
        обновлённый список entries с добавленными координатами
    """
    if not PYMUPDF_AVAILABLE:
        return entries
    
    if not os.path.isfile(pdf_path):
        return entries
    
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        if show_output:
            print(f"[!] Ошибка открытия PDF для поиска координат: {e}")
        return entries
    
    if show_output:
        print(f"\n[*] Поиск точных координат заголовков в PDF...")
        print("[*] Это позволит переходить прямо к заголовкам, а не к началу страницы")
    
    found_count = 0
    total = len(entries)
    
    for i, entry in enumerate(entries):
        title = entry["title"]
        page_num = entry["page"]
        
        # Проверяем диапазон страниц
        if page_num < 1 or page_num > len(doc):
            continue
        
        page = doc[page_num - 1]  # 0-indexed
        
        # Генерируем варианты поиска
        search_variants = generate_search_variants(title)
        
        coords = None
        found_text = None
        
        # Пробуем найти каждый вариант
        for variant in search_variants:
            instances = page.search_for(variant)
            if instances:
                # Берём первое вхождение (самое верхнее на странице)
                rect = instances[0]
                coords = {
                    "x": rect.x0,
                    "y": rect.y0,
                    "width": rect.width,
                    "height": rect.height
                }
                found_text = variant
                break
        
        if coords:
            entry["coords"] = coords
            found_count += 1
            if show_output and i < 5:  # Показываем первые 5 для примера
                print(f"  [✓] '{found_text[:50]}...' -> ({coords['x']:.1f}, {coords['y']:.1f})")
        elif show_output and i < 5:
            print(f"  [?] '{title[:50]}...' -> координаты не найдены (fallback: страница)")
    
    doc.close()
    
    if show_output:
        print(f"\n[+] Найдены координаты для {found_count}/{total} заголовков")
        if found_count < total:
            print(f"[INFO] Для {total - found_count} заголовков используется переход к странице")
    
    return entries


def generate_search_variants(title: str):
    """
    Генерирует варианты текста для поиска заголовка в PDF.
    
    Пробует разные варианты:
    - Полный заголовок
    - Без номера раздела
    - С точкой после номера
    - Только первые слова
    
    Args:
        title: заголовок из оглавления
    
    Returns:
        список вариантов для поиска (от точного к общему)
    """
    variants = []
    
    # 1. Полный заголовок как есть
    variants.append(title.strip())
    
    # 2. Попробуем убрать/добавить точку после номера раздела
    # "1.2 Введение" -> "1.2. Введение"
    m = re.match(r'^(\d+(?:\.\d+)*)\s+(.+)$', title)
    if m:
        section_num = m.group(1)
        section_title = m.group(2)
        
        # С точкой
        variants.append(f"{section_num}. {section_title}")
        
        # Без номера (только название)
        variants.append(section_title)
        
        # Первые 3-4 слова названия (для длинных заголовков)
        words = section_title.split()
        if len(words) > 3:
            variants.append(' '.join(words[:3]))
    else:
        # Если нет номера, пробуем первые слова
        words = title.split()
        if len(words) > 3:
            variants.append(' '.join(words[:3]))
    
    # Убираем дубликаты, сохраняя порядок
    seen = set()
    unique_variants = []
    for v in variants:
        v_clean = v.strip()
        if v_clean and v_clean not in seen:
            seen.add(v_clean)
            unique_variants.append(v_clean)
    
    return unique_variants


def run_toc_pages_mode(args):
    """
    Запуск в режиме ручного указания страниц оглавления.
    
    Args:
        args: аргументы командной строки
    
    Returns:
        True если успешно, False иначе
    """
    show_output = not args.quiet
    
    # Получаем путь к PDF
    if args.file:
        pdf_path = args.file.strip('"\'')
    else:
        pdf_path = get_pdf_interactively()
        if pdf_path is None:
            return False
    
    if not os.path.isfile(pdf_path):
        print(f"[!] Файл не найден: {pdf_path}")
        return False
    
    if not pdf_path.lower().endswith('.pdf'):
        print("[!] В режиме --toc-pages требуется PDF файл.")
        print("[*] Для DOCX используй стандартный режим без --toc-pages")
        return False
    
    # Определяем страницы
    if args.toc_pages == 'interactive':
        # Интерактивный режим запроса страниц
        page_numbers = get_toc_pages_interactively(pdf_path)
        if page_numbers is None:
            return False
    else:
        # Страницы указаны в командной строке
        try:
            page_numbers = parse_page_range(args.toc_pages)
            if not page_numbers:
                print("[!] Не удалось разобрать номера страниц")
                return False
        except ValueError as e:
            print(f"[!] Ошибка разбора страниц: {e}")
            return False
    
    if show_output:
        print("\n" + "=" * 60)
        print("ИЗВЛЕЧЕНИЕ ЗАКЛАДОК ИЗ СТРАНИЦ ОГЛАВЛЕНИЯ PDF")
        print("=" * 60)
        print(f"\n[PDF] {os.path.basename(pdf_path)}")
        print(f"[PAGES] Страницы оглавления: {page_numbers}")
    
    # Извлекаем записи из указанных страниц
    entries = extract_toc_from_pdf_pages(pdf_path, page_numbers, show_output)
    
    if not entries:
        print("\n[!] Не удалось извлечь записи оглавления!")
        print("[*] Попробуй указать другие страницы или проверь формат.")
        return False
    
    if show_output:
        print(f"\n[+] Найдено заголовков: {len(entries)}")
    
    # Ищем точные координаты заголовков в PDF
    entries = find_exact_coordinates(pdf_path, entries, show_output)
    
    # Строим дерево закладок
    tree = build_bookmark_tree(entries)
    
    # Сохраняем JSON
    base, ext = os.path.splitext(pdf_path)
    out_path = base + "_bookmarks.json"
    
    if show_output:
        print(f"\n[*] Сохраняю JSON: {os.path.basename(out_path)}")
    
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(tree, f, ensure_ascii=False, indent=2)
    
    if show_output:
        print("\n" + "=" * 60)
        print("[OK] ГОТОВО!")
        print("=" * 60)
        print(f"\n[>>] Создан файл: {out_path}")
    
    # Предлагаем встроить закладки в PDF
    if show_output:
        embed_bookmarks_to_pdf(pdf_path, out_path, show_output=True)
    
    return True


def get_pdf_interactively():
    """Запросить путь к PDF файлу интерактивно."""
    print("=" * 60)
    print("РЕЖИМ ИЗВЛЕЧЕНИЯ ЗАКЛАДОК ИЗ PDF")
    print("=" * 60)
    print("\nВведи путь к PDF файлу с оглавлением.\n")
    
    while True:
        file_path = input("[?] PDF файл (или 'q' для выхода): ").strip()
        
        if file_path.lower() in ('q', 'quit', 'exit'):
            print("[-] Выход...")
            return None
        
        file_path = file_path.strip('"\'')
        
        if os.path.isfile(file_path):
            if file_path.lower().endswith('.pdf'):
                return file_path
            else:
                print("[!] Файл должен быть в формате PDF")
        else:
            print(f"[!] Файл не найден: {file_path}")
            print("[*] Попробуй ещё раз или введи 'q' для выхода\n")


if __name__ == "__main__":
    main()

