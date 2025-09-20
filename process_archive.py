#!/usr/bin/env python3
"""
Скрипт для обработки архива с Excel файлами организаций
Поддерживает форматы: .zip, .rar, .7z, .tar, .tar.gz, .tar.bz2
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
import logging
from process_excel_files import ExcelProcessor

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class ArchiveProcessor:
    def __init__(self):
        self.temp_dir = None
        self.supported_archives = ['.zip', '.rar', '.7z', '.tar', '.gz', '.bz2']
        self.supported_excel = ['.xlsx', '.xls']
    
    def extract_archive(self, archive_path: str) -> str:
        """
        Извлекает архив во временную директорию
        """
        archive_path = Path(archive_path)
        
        if not archive_path.exists():
            raise FileNotFoundError(f"Архив не найден: {archive_path}")
        
        # Создаем временную директорию
        self.temp_dir = tempfile.mkdtemp(prefix="excel_archive_")
        logger.info(f"Создана временная директория: {self.temp_dir}")
        
        archive_ext = archive_path.suffix.lower()
        archive_name = archive_path.name.lower()
        
        try:
            if archive_ext == '.zip':
                import zipfile
                with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                    zip_ref.extractall(self.temp_dir)
                    logger.info(f"Извлечен ZIP архив: {len(zip_ref.namelist())} файлов")
            
            elif archive_ext == '.rar':
                try:
                    import rarfile
                    with rarfile.RarFile(archive_path, 'r') as rar_ref:
                        rar_ref.extractall(self.temp_dir)
                        logger.info(f"Извлечен RAR архив")
                except ImportError:
                    # Пробуем через системную команду
                    import subprocess
                    result = subprocess.run(['unrar', 'x', str(archive_path), self.temp_dir], 
                                          capture_output=True, text=True)
                    if result.returncode != 0:
                        raise Exception(f"Ошибка извлечения RAR: {result.stderr}")
                    logger.info("Извлечен RAR архив (через unrar)")
            
            elif archive_ext == '.7z':
                try:
                    import py7zr
                    with py7zr.SevenZipFile(archive_path, mode="r") as z:
                        z.extractall(self.temp_dir)
                        logger.info("Извлечен 7Z архив")
                except ImportError:
                    # Пробуем через системную команду
                    import subprocess
                    result = subprocess.run(['7z', 'x', str(archive_path), f'-o{self.temp_dir}'], 
                                          capture_output=True, text=True)
                    if result.returncode != 0:
                        raise Exception(f"Ошибка извлечения 7Z: {result.stderr}")
                    logger.info("Извлечен 7Z архив (через 7z)")
            
            elif archive_ext in ['.tar', '.gz', '.bz2'] or 'tar' in archive_name:
                import tarfile
                with tarfile.open(archive_path, 'r:*') as tar_ref:
                    tar_ref.extractall(self.temp_dir)
                    logger.info(f"Извлечен TAR архив: {len(tar_ref.getnames())} файлов")
            
            else:
                raise ValueError(f"Неподдерживаемый формат архива: {archive_ext}")
        
        except Exception as e:
            self.cleanup()
            raise Exception(f"Ошибка при извлечении архива: {e}")
        
        return self.temp_dir
    
    def find_excel_files(self, directory: str) -> list:
        """
        Рекурсивно ищет все Excel файлы в директории
        """
        excel_files = []
        
        for root, dirs, files in os.walk(directory):
            for file in files:
                file_path = Path(root) / file
                if file_path.suffix.lower() in self.supported_excel:
                    excel_files.append(file_path)
        
        logger.info(f"Найдено Excel файлов: {len(excel_files)}")
        for f in excel_files[:10]:  # Показываем первые 10
            logger.info(f"  - {f.name}")
        if len(excel_files) > 10:
            logger.info(f"  ... и еще {len(excel_files) - 10} файлов")
        
        return excel_files
    
    def prepare_files_for_processing(self, excel_files: list) -> str:
        """
        Копирует все Excel файлы в одну директорию для обработки
        """
        processing_dir = Path(self.temp_dir) / "processing"
        processing_dir.mkdir(exist_ok=True)
        
        for i, file_path in enumerate(excel_files):
            # Создаем уникальное имя файла если есть дубликаты
            new_name = file_path.name
            counter = 1
            while (processing_dir / new_name).exists():
                name_parts = file_path.stem, counter, file_path.suffix
                new_name = f"{name_parts[0]}_{name_parts[1]}{name_parts[2]}"
                counter += 1
            
            shutil.copy2(file_path, processing_dir / new_name)
        
        logger.info(f"Подготовлено файлов для обработки: {len(excel_files)}")
        return str(processing_dir)
    
    def cleanup(self):
        """
        Очищает временные файлы
        """
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
            logger.info("Временные файлы очищены")
    
    def process_archive(self, archive_path: str, output_dir: str = None) -> str:
        """
        Основной метод обработки архива
        """
        try:
            logger.info(f"Начинаем обработку архива: {archive_path}")
            
            # Извлекаем архив
            temp_dir = self.extract_archive(archive_path)
            
            # Ищем Excel файлы
            excel_files = self.find_excel_files(temp_dir)
            
            if not excel_files:
                raise ValueError("В архиве не найдено Excel файлов (.xlsx, .xls)")
            
            # Подготавливаем файлы для обработки
            processing_dir = self.prepare_files_for_processing(excel_files)
            
            # Обрабатываем файлы
            logger.info("Начинаем объединение и обработку файлов...")
            processor = ExcelProcessor(processing_dir)
            processor.process()
            
            # Определяем путь для сохранения результата
            if output_dir:
                output_path = Path(output_dir)
            else:
                output_path = Path(archive_path).parent
            
            output_path.mkdir(exist_ok=True, parents=True)
            result_file = output_path / "ИТОГ.xlsx"
            
            # Копируем результат
            temp_result = Path(processing_dir) / "ИТОГ.xlsx"
            if temp_result.exists():
                shutil.copy2(temp_result, result_file)
                logger.info(f"Результат сохранен: {result_file}")
                
                # Показываем статистику
                try:
                    import pandas as pd
                    df = pd.read_excel(result_file)
                    logger.info(f"Итоговое количество организаций: {len(df)}")
                    
                    # Статистика по сайтам если есть соответствующая колонка
                    site_columns = [col for col in df.columns if any(word in col.lower() 
                                  for word in ['сайт', 'site', 'web', 'url'])]
                    if site_columns:
                        site_col = site_columns[0]
                        valid_sites = (df[site_col] != '---').sum()
                        logger.info(f"Организаций с валидными сайтами: {valid_sites}")
                        logger.info(f"Организаций без сайтов: {len(df) - valid_sites}")
                
                except Exception as e:
                    logger.warning(f"Не удалось получить статистику: {e}")
                
                return str(result_file)
            else:
                raise Exception("Файл результата не был создан")
        
        finally:
            # Очищаем временные файлы
            self.cleanup()


def main():
    print("=" * 70)
    print("ОБРАБОТКА АРХИВА С EXCEL ФАЙЛАМИ ОРГАНИЗАЦИЙ")
    print("=" * 70)
    
    if len(sys.argv) < 2:
        print("Использование:")
        print(f"  {sys.argv[0]} <путь_к_архиву> [путь_для_сохранения_результата]")
        print()
        print("Поддерживаемые форматы архивов:")
        print("  - ZIP (.zip)")
        print("  - RAR (.rar)")
        print("  - 7-Zip (.7z)")
        print("  - TAR (.tar, .tar.gz, .tar.bz2)")
        print()
        print("Поддерживаемые форматы Excel:")
        print("  - .xlsx (рекомендуется)")
        print("  - .xls")
        print()
        print("Примеры:")
        print(f"  {sys.argv[0]} организации.zip")
        print(f"  {sys.argv[0]} файлы.rar /home/user/results")
        return
    
    archive_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(archive_path):
        print(f"❌ ОШИБКА: Архив не найден: {archive_path}")
        return
    
    try:
        processor = ArchiveProcessor()
        result_file = processor.process_archive(archive_path, output_dir)
        
        print("\n" + "=" * 70)
        print("✅ ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО!")
        print("=" * 70)
        print(f"Результат сохранен в: {result_file}")
        
        # Показываем размер файла
        if os.path.exists(result_file):
            file_size = os.path.getsize(result_file)
            print(f"Размер файла: {file_size:,} байт")
        
        print("\nВы можете найти объединенный файл 'ИТОГ.xlsx' со всеми")
        print("уникальными организациями и проверенными веб-адресами.")
        
    except Exception as e:
        print(f"❌ ОШИБКА: {e}")
        logger.error(f"Детали ошибки: {e}", exc_info=True)


if __name__ == "__main__":
    main()