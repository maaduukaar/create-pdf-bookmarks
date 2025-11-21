#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Извлечение структуры заголовков из СОДЕРЖАНИЯ DOCX в JSON для PDF закладок.

Формат JSON совместим с PyMuPDF для создания закладок в PDF.
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


# Регулярка для строк содержания
TOC_LINE_PATTERN = re.compile(
    r"^(\d+(?:\.\d+)*)\.?\s+(.+?)\s+(\d+)$"
)

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
        "dest": [page, "Fit"],  # Навигация к странице (без координат)
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
        
        # Создаем узел закладки
        node = {
            "title": entry["title"],
            "dest": [entry["page"], "Fit"],  # Простая навигация к странице
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
        где dest_dict может быть пустым словарём для простого перехода.
        """
        for node in nodes:
            level = parent_level + 1
            title = node.get("title", "Untitled")
            
            # Получаем номер страницы
            dest = node.get("dest", [])
            if isinstance(dest, list) and len(dest) > 0:
                page = dest[0]
            else:
                page = 1
            
            # Преобразуем номер страницы (в JSON нумерация может начинаться с 1)
            # PyMuPDF использует нумерацию страниц с 1
            page = max(1, min(page, len(doc)))
            
            # Добавляем закладку в список TOC
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
        description="Извлечение закладок из содержания DOCX в JSON",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python %(prog)s document.docx
  python %(prog)s "C:\\Docs\\file.docx"
  python %(prog)s --quiet file.docx

Формат содержания:
  1. Название раздела 5
  1.1 Подраздел 12
  3.4.2.1 Интерфейс раздела 69
  
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
        help='Путь к DOCX-файлу с содержанием'
    )
    
    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Тихий режим (минимум вывода)'
    )
    
    args = parser.parse_args()
    
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


if __name__ == "__main__":
    main()
