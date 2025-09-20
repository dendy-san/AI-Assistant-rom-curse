#!/usr/bin/env python3
"""
Помощник для создания архива из Excel файлов
"""

import os
import zipfile
import sys
from pathlib import Path

def create_archive_from_folder(folder_path, archive_name="мои_организации.zip"):
    """Создает ZIP архив из всех Excel файлов в папке"""
    folder = Path(folder_path)
    
    if not folder.exists():
        print(f"❌ Папка не найдена: {folder_path}")
        return None
    
    excel_files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xls"))
    
    if not excel_files:
        print(f"❌ В папке {folder_path} не найдено Excel файлов")
        return None
    
    archive_path = folder.parent / archive_name
    
    print(f"📦 Создаем архив: {archive_path}")
    print(f"📁 Из папки: {folder_path}")
    print(f"📄 Файлов для архивации: {len(excel_files)}")
    
    with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path in excel_files:
            zipf.write(file_path, file_path.name)
            print(f"  ✅ Добавлен: {file_path.name}")
    
    print(f"\n🎉 Архив создан: {archive_path}")
    print(f"📊 Размер архива: {archive_path.stat().st_size:,} байт")
    
    return str(archive_path)

def main():
    if len(sys.argv) < 2:
        print("Использование:")
        print(f"  python3 {sys.argv[0]} <папка_с_excel_файлами> [имя_архива.zip]")
        print("\nПример:")
        print(f"  python3 {sys.argv[0]} /path/to/excel/files")
        print(f"  python3 {sys.argv[0]} ./ваши_файлы мой_архив.zip")
        return
    
    folder_path = sys.argv[1]
    archive_name = sys.argv[2] if len(sys.argv) > 2 else "мои_организации.zip"
    
    archive_path = create_archive_from_folder(folder_path, archive_name)
    
    if archive_path:
        print(f"\n💡 Теперь можете обработать архив командой:")
        print(f"   python3 universal_processor.py {archive_path}")

if __name__ == "__main__":
    main()