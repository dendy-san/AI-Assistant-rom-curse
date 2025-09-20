#!/usr/bin/env python3
"""
Скрипт для просмотра содержимого Excel файлов
"""

import pandas as pd
import sys
import os

def view_excel_file(file_path):
    """Просмотр содержимого Excel файла"""
    try:
        df = pd.read_excel(file_path)
        print(f"\n=== Файл: {os.path.basename(file_path)} ===")
        print(f"Количество записей: {len(df)}")
        print(f"Колонки: {list(df.columns)}")
        print("\nСодержимое:")
        print(df.to_string(index=False))
        return df
    except Exception as e:
        print(f"Ошибка при чтении файла {file_path}: {e}")
        return None

def main():
    if len(sys.argv) > 1:
        file_or_dir = sys.argv[1]
    else:
        file_or_dir = "/workspace/test_data"
    
    if os.path.isfile(file_or_dir):
        view_excel_file(file_or_dir)
    elif os.path.isdir(file_or_dir):
        excel_files = [f for f in os.listdir(file_or_dir) if f.endswith('.xlsx')]
        for file_name in sorted(excel_files):
            file_path = os.path.join(file_or_dir, file_name)
            view_excel_file(file_path)
    else:
        print(f"Файл или директория не найдены: {file_or_dir}")

if __name__ == "__main__":
    main()