#!/usr/bin/env python3
"""
Универсальный процессор для Excel файлов организаций
Поддерживает: отдельные файлы, папки с файлами, ZIP архивы
"""

import os
import sys
import zipfile
from pathlib import Path
import logging

# Импортируем наши модули
from process_excel_files import ExcelProcessor
from simple_archive_processor import process_zip_archive

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def detect_input_type(input_path: str) -> str:
    """
    Определяет тип входных данных
    """
    path = Path(input_path)
    
    if not path.exists():
        return "not_found"
    
    if path.is_file():
        if path.suffix.lower() == '.zip':
            return "zip_archive"
        elif path.suffix.lower() in ['.xlsx', '.xls']:
            return "excel_file"
        else:
            return "unknown_file"
    elif path.is_dir():
        # Проверяем есть ли Excel файлы в директории
        excel_files = list(path.glob("*.xlsx")) + list(path.glob("*.xls"))
        if excel_files:
            return "excel_directory"
        else:
            return "empty_directory"
    
    return "unknown"

def process_single_excel_file(file_path: str, output_dir: str = None) -> str:
    """
    Обрабатывает один Excel файл (просто копирует его как результат)
    """
    input_path = Path(file_path)
    
    if output_dir:
        output_path = Path(output_dir)
    else:
        output_path = input_path.parent
    
    output_path.mkdir(exist_ok=True, parents=True)
    result_file = output_path / "ИТОГ.xlsx"
    
    # Читаем файл и сохраняем с валидацией
    processor = ExcelProcessor(str(input_path.parent))
    
    # Создаем временную обработку только для этого файла
    import pandas as pd
    df = pd.read_excel(input_path)
    
    # Применяем валидацию URL если есть соответствующие колонки
    site_columns = [col for col in df.columns if any(word in col.lower() 
                   for word in ['сайт', 'site', 'web', 'url'])]
    
    if site_columns:
        site_col = site_columns[0]
        logger.info(f"Валидируем веб-адреса в колонке: {site_col}")
        df[site_col] = df[site_col].apply(processor.validate_and_fix_url)
        
        valid_count = (df[site_col] != '---').sum()
        logger.info(f"Валидных URL: {valid_count} из {len(df)}")
    
    # Сохраняем результат
    df.to_excel(result_file, index=False)
    logger.info(f"Результат сохранен: {result_file}")
    
    return str(result_file)

def main():
    print("=" * 70)
    print("УНИВЕРСАЛЬНЫЙ ОБРАБОТЧИК EXCEL ФАЙЛОВ С ОРГАНИЗАЦИЯМИ")
    print("=" * 70)
    
    if len(sys.argv) < 2:
        print("Использование:")
        print(f"  python3 {sys.argv[0]} <путь> [папка_для_результата]")
        print()
        print("Поддерживаемые типы входных данных:")
        print("  📁 Папка с Excel файлами")
        print("  📦 ZIP архив с Excel файлами") 
        print("  📄 Отдельный Excel файл")
        print()
        print("Примеры:")
        print(f"  python3 {sys.argv[0]} /path/to/excel/files/")
        print(f"  python3 {sys.argv[0]} организации.zip")
        print(f"  python3 {sys.argv[0]} single_file.xlsx")
        print()
        print("💡 Просто перетащите папку, архив или файл как аргумент!")
        return
    
    input_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Определяем тип входных данных
    input_type = detect_input_type(input_path)
    
    print(f"📂 Анализируем входные данные: {input_path}")
    
    try:
        if input_type == "not_found":
            print(f"❌ ОШИБКА: Путь не найден: {input_path}")
            return
        
        elif input_type == "zip_archive":
            print("📦 Обнаружен ZIP архив")
            result_file = process_zip_archive(input_path, output_dir)
            
        elif input_type == "excel_directory":
            print("📁 Обнаружена папка с Excel файлами")
            processor = ExcelProcessor(input_path)
            processor.process()
            
            # Определяем путь результата
            if output_dir:
                import shutil
                source = Path(input_path) / "ИТОГ.xlsx"
                dest = Path(output_dir) / "ИТОГ.xlsx"
                Path(output_dir).mkdir(exist_ok=True, parents=True)
                shutil.copy2(source, dest)
                result_file = str(dest)
            else:
                result_file = str(Path(input_path) / "ИТОГ.xlsx")
        
        elif input_type == "excel_file":
            print("📄 Обнаружен отдельный Excel файл")
            result_file = process_single_excel_file(input_path, output_dir)
        
        elif input_type == "empty_directory":
            print(f"❌ ОШИБКА: В папке {input_path} не найдено Excel файлов")
            return
        
        elif input_type == "unknown_file":
            print(f"❌ ОШИБКА: Неподдерживаемый тип файла: {input_path}")
            print("Поддерживаются: .xlsx, .xls, .zip")
            return
        
        else:
            print(f"❌ ОШИБКА: Неизвестный тип входных данных")
            return
        
        # Показываем результат
        print("\n" + "=" * 70)
        print("✅ ОБРАБОТКА УСПЕШНО ЗАВЕРШЕНА!")
        print("=" * 70)
        print(f"📄 Результат: {result_file}")
        
        if os.path.exists(result_file):
            file_size = os.path.getsize(result_file)
            print(f"📊 Размер: {file_size:,} байт")
            
            # Статистика
            try:
                import pandas as pd
                df = pd.read_excel(result_file)
                print(f"🏢 Организаций: {len(df)}")
                
                # Проверяем наличие колонки с сайтами
                site_col = None
                for col in df.columns:
                    if any(word in col.lower() for word in ['сайт', 'site', 'web', 'url']):
                        site_col = col
                        break
                
                if site_col:
                    valid = (df[site_col] != '---').sum()
                    invalid = len(df) - valid
                    print(f"🌐 С сайтами: {valid}, без сайтов: {invalid}")
                
            except Exception as e:
                logger.warning(f"Не удалось получить статистику: {e}")
        
        print(f"\n🎉 Готово! Все организации обработаны и сохранены в 'ИТОГ.xlsx'")
        
    except Exception as e:
        print(f"\n❌ ОШИБКА: {e}")
        logger.error(f"Подробности: {e}", exc_info=True)

if __name__ == "__main__":
    main()