#!/usr/bin/env python3
"""
Упрощенный скрипт для обработки ZIP архивов с Excel файлами
Работает только с ZIP архивами (наиболее распространенный формат)
"""

import os
import sys
import tempfile
import shutil
import zipfile
from pathlib import Path
import logging
from process_excel_files import ExcelProcessor

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def extract_zip_archive(zip_path: str) -> str:
    """
    Извлекает ZIP архив во временную директорию
    """
    if not os.path.exists(zip_path):
        raise FileNotFoundError(f"Архив не найден: {zip_path}")
    
    # Создаем временную директорию
    temp_dir = tempfile.mkdtemp(prefix="excel_zip_")
    logger.info(f"Извлекаем архив во временную директорию: {temp_dir}")
    
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            logger.info(f"Извлечено файлов: {len(zip_ref.namelist())}")
            
            # Показываем список файлов
            excel_count = 0
            for name in zip_ref.namelist()[:10]:  # Первые 10 файлов
                if name.lower().endswith(('.xlsx', '.xls')):
                    excel_count += 1
                logger.info(f"  - {name}")
            
            if len(zip_ref.namelist()) > 10:
                logger.info(f"  ... и еще {len(zip_ref.namelist()) - 10} файлов")
            
            total_excel = sum(1 for name in zip_ref.namelist() 
                            if name.lower().endswith(('.xlsx', '.xls')))
            logger.info(f"Найдено Excel файлов: {total_excel}")
    
    except Exception as e:
        shutil.rmtree(temp_dir)
        raise Exception(f"Ошибка при извлечении ZIP архива: {e}")
    
    return temp_dir

def find_and_prepare_excel_files(temp_dir: str) -> str:
    """
    Находит все Excel файлы и подготавливает их для обработки
    """
    processing_dir = Path(temp_dir) / "excel_processing"
    processing_dir.mkdir(exist_ok=True)
    
    excel_files = []
    
    # Рекурсивно ищем Excel файлы
    for root, dirs, files in os.walk(temp_dir):
        for file in files:
            if file.lower().endswith(('.xlsx', '.xls')):
                source_path = Path(root) / file
                excel_files.append(source_path)
    
    if not excel_files:
        raise ValueError("В архиве не найдено Excel файлов")
    
    logger.info(f"Подготавливаем {len(excel_files)} Excel файлов:")
    
    # Копируем файлы в директорию обработки
    for i, file_path in enumerate(excel_files):
        # Создаем уникальное имя если есть дубликаты
        new_name = file_path.name
        counter = 1
        while (processing_dir / new_name).exists():
            name_parts = file_path.stem, counter, file_path.suffix
            new_name = f"{name_parts[0]}_{name_parts[1]}{name_parts[2]}"
            counter += 1
        
        shutil.copy2(file_path, processing_dir / new_name)
        logger.info(f"  {i+1}. {new_name}")
    
    return str(processing_dir)

def process_zip_archive(zip_path: str, output_dir: str = None) -> str:
    """
    Основная функция обработки ZIP архива
    """
    temp_dir = None
    
    try:
        logger.info(f"Начинаем обработку ZIP архива: {zip_path}")
        
        # Извлекаем архив
        temp_dir = extract_zip_archive(zip_path)
        
        # Подготавливаем Excel файлы
        processing_dir = find_and_prepare_excel_files(temp_dir)
        
        # Обрабатываем файлы
        logger.info("Начинаем объединение файлов и удаление дубликатов...")
        processor = ExcelProcessor(processing_dir)
        processor.process()
        
        # Определяем где сохранить результат
        if output_dir:
            output_path = Path(output_dir)
        else:
            output_path = Path(zip_path).parent
        
        output_path.mkdir(exist_ok=True, parents=True)
        result_file = output_path / "ИТОГ.xlsx"
        
        # Копируем результат
        temp_result = Path(processing_dir) / "ИТОГ.xlsx"
        if temp_result.exists():
            shutil.copy2(temp_result, result_file)
            return str(result_file)
        else:
            raise Exception("Файл результата не был создан")
    
    finally:
        # Очищаем временные файлы
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            logger.info("Временные файлы очищены")

def main():
    print("=" * 70)
    print("ОБРАБОТКА ZIP АРХИВА С EXCEL ФАЙЛАМИ")
    print("=" * 70)
    print("Поддерживается: ZIP архивы с файлами .xlsx и .xls")
    print("=" * 70)
    
    # Проверяем аргументы
    if len(sys.argv) < 2:
        print("\n❌ Не указан путь к архиву!")
        print("\nИспользование:")
        print(f"  python3 {sys.argv[0]} <путь_к_zip_архиву> [папка_для_результата]")
        print("\nПримеры:")
        print(f"  python3 {sys.argv[0]} организации.zip")
        print(f"  python3 {sys.argv[0]} файлы.zip /home/user/results")
        print(f"  python3 {sys.argv[0]} excel_files.zip .")
        print("\n📁 Поместите ваш ZIP архив с Excel файлами рядом со скриптом")
        print("   и запустите команду выше")
        return
    
    zip_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Проверяем существование файла
    if not os.path.exists(zip_path):
        print(f"❌ ОШИБКА: ZIP архив не найден: {zip_path}")
        print(f"\nПроверьте правильность пути к файлу.")
        return
    
    # Проверяем что это ZIP файл
    if not zip_path.lower().endswith('.zip'):
        print(f"❌ ОШИБКА: Файл должен быть в формате ZIP")
        print(f"Получен файл: {zip_path}")
        print(f"Для других форматов используйте process_archive.py")
        return
    
    try:
        # Обрабатываем архив
        result_file = process_zip_archive(zip_path, output_dir)
        
        print("\n" + "=" * 70)
        print("✅ ОБРАБОТКА УСПЕШНО ЗАВЕРШЕНА!")
        print("=" * 70)
        print(f"📄 Результат сохранен: {result_file}")
        
        # Показываем статистику
        if os.path.exists(result_file):
            file_size = os.path.getsize(result_file)
            print(f"📊 Размер файла: {file_size:,} байт")
            
            try:
                import pandas as pd
                df = pd.read_excel(result_file)
                print(f"🏢 Количество организаций: {len(df)}")
                
                # Ищем колонку с сайтами
                site_col = None
                for col in df.columns:
                    if any(word in col.lower() for word in ['сайт', 'site', 'web', 'url']):
                        site_col = col
                        break
                
                if site_col:
                    valid_sites = (df[site_col] != '---').sum()
                    invalid_sites = len(df) - valid_sites
                    print(f"🌐 С валидными сайтами: {valid_sites}")
                    print(f"❌ Без сайтов/невалидные: {invalid_sites}")
                    
            except Exception as e:
                logger.warning(f"Не удалось получить детальную статистику: {e}")
        
        print("\n🎉 Готово! Файл 'ИТОГ.xlsx' содержит все уникальные")
        print("   организации с проверенными веб-адресами.")
        
    except Exception as e:
        print(f"\n❌ ОШИБКА ПРИ ОБРАБОТКЕ: {e}")
        logger.error(f"Детали ошибки: {e}", exc_info=True)
        print("\n💡 Советы по устранению:")
        print("   - Проверьте, что архив не поврежден")
        print("   - Убедитесь, что в архиве есть файлы .xlsx или .xls")
        print("   - Проверьте права доступа к файлам")

if __name__ == "__main__":
    main()