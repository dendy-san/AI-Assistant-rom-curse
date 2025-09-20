#!/usr/bin/env python3
"""
Скрипт-обертка для обработки реальных файлов пользователя
"""

import os
import sys
from process_excel_files import ExcelProcessor
from pathlib import Path

def main():
    print("=" * 60)
    print("ОБРАБОТКА EXCEL ФАЙЛОВ С ОРГАНИЗАЦИЯМИ")
    print("=" * 60)
    
    # Возможные пути к файлам
    possible_paths = [
        r"d:\001\Kursis\Kurs_from_Cursor\Организации\00",
        "/mnt/d/001/Kursis/Kurs_from_Cursor/Организации/00",
        "/workspace/excel_files",
        "/workspace"
    ]
    
    # Если передан путь в аргументах
    if len(sys.argv) > 1:
        user_path = sys.argv[1]
        possible_paths.insert(0, user_path)
    
    directory_path = None
    
    # Ищем директорию с Excel файлами
    for path in possible_paths:
        if os.path.exists(path):
            xlsx_files = list(Path(path).glob("*.xlsx"))
            if xlsx_files:
                directory_path = path
                print(f"✅ Найдена директория с Excel файлами: {path}")
                print(f"   Найдено файлов: {len(xlsx_files)}")
                for f in xlsx_files[:5]:  # Показываем первые 5 файлов
                    print(f"   - {f.name}")
                if len(xlsx_files) > 5:
                    print(f"   ... и еще {len(xlsx_files) - 5} файлов")
                break
    
    if not directory_path:
        print("❌ ФАЙЛЫ НЕ НАЙДЕНЫ!")
        print("\nИНСТРУКЦИЯ:")
        print("1. Скопируйте ваши Excel файлы в одну из директорий:")
        for path in possible_paths[2:]:  # Показываем только доступные пути
            print(f"   - {path}")
        print("2. Или запустите: python3 run_for_real_files.py /путь/к/вашим/файлам")
        print(f"3. Ваши файлы должны быть в формате .xlsx")
        print(f"   Пример: 'ольгинка + 3км.xlsx'")
        return
    
    print("\n" + "=" * 60)
    print("НАЧИНАЕМ ОБРАБОТКУ...")
    print("=" * 60)
    
    try:
        processor = ExcelProcessor(directory_path)
        processor.process()
        
        result_file = Path(directory_path) / "ИТОГ.xlsx"
        if result_file.exists():
            print("\n" + "=" * 60)
            print("✅ УСПЕШНО ЗАВЕРШЕНО!")
            print("=" * 60)
            print(f"Результат сохранен в: {result_file}")
            print(f"Размер файла: {result_file.stat().st_size} байт")
            
            # Показываем краткую статистику
            try:
                import pandas as pd
                df = pd.read_excel(result_file)
                print(f"Количество организаций в итоговом файле: {len(df)}")
                if 'Веб-сайт' in df.columns:
                    valid_sites = (df['Веб-сайт'] != '---').sum()
                    print(f"Организаций с валидными сайтами: {valid_sites}")
                    print(f"Организаций без сайтов: {len(df) - valid_sites}")
            except:
                pass
        else:
            print("❌ Файл результата не был создан!")
            
    except Exception as e:
        print(f"❌ ОШИБКА: {e}")
        print("Проверьте формат ваших Excel файлов и повторите попытку.")

if __name__ == "__main__":
    main()