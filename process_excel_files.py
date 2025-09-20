#!/usr/bin/env python3
"""
Скрипт для объединения файлов Excel с организациями
Удаляет дубликаты, проверяет валидность веб-адресов
"""

import pandas as pd
import os
import re
import requests
from urllib.parse import urlparse
from pathlib import Path
import logging
from typing import List, Dict, Any, Optional

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class ExcelProcessor:
    def __init__(self, directory_path: str):
        self.directory_path = Path(directory_path)
        self.all_data = pd.DataFrame()
        self.processed_organizations = {}
        
    def is_valid_url(self, url: str) -> bool:
        """
        Проверяет валидность URL
        """
        if not url or pd.isna(url) or str(url).strip() == '':
            return False
            
        url_str = str(url).strip()
        
        # Проверяем базовый формат URL
        url_pattern = re.compile(
            r'^https?://'  # http:// или https://
            r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+[A-Z]{2,6}\.?|'  # домен
            r'localhost|'  # localhost
            r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})'  # IP
            r'(?::\d+)?'  # порт
            r'(?:/?|[/?]\S+)$', re.IGNORECASE)
        
        if not url_pattern.match(url_str):
            # Попробуем добавить http:// если его нет
            if not url_str.startswith(('http://', 'https://')):
                url_str = 'http://' + url_str
                if not url_pattern.match(url_str):
                    return False
            else:
                return False
        
        # Дополнительная проверка через urlparse
        try:
            parsed = urlparse(url_str)
            return all([parsed.scheme, parsed.netloc])
        except Exception:
            return False
    
    def validate_and_fix_url(self, url: str) -> str:
        """
        Валидирует URL и возвращает исправленный или "---"
        """
        if not url or pd.isna(url):
            return "---"
            
        url_str = str(url).strip()
        
        if url_str == "---":
            return "---"
            
        # Попробуем разные варианты URL
        urls_to_try = [
            url_str,
            'http://' + url_str if not url_str.startswith(('http://', 'https://')) else url_str,
            'https://' + url_str.replace('http://', '') if url_str.startswith('http://') else None
        ]
        
        for test_url in urls_to_try:
            if test_url and self.is_valid_url(test_url):
                return test_url
        
        return "---"
    
    def read_excel_files(self) -> None:
        """
        Читает все Excel файлы из директории
        """
        xlsx_files = list(self.directory_path.glob("*.xlsx"))
        
        if not xlsx_files:
            logger.warning(f"Не найдено файлов .xlsx в директории {self.directory_path}")
            return
        
        logger.info(f"Найдено {len(xlsx_files)} файлов Excel")
        
        all_dataframes = []
        
        for file_path in xlsx_files:
            try:
                logger.info(f"Обрабатываем файл: {file_path.name}")
                df = pd.read_excel(file_path)
                
                # Добавляем информацию об источнике файла
                df['_source_file'] = file_path.name
                all_dataframes.append(df)
                
                logger.info(f"Прочитано {len(df)} записей из {file_path.name}")
                
            except Exception as e:
                logger.error(f"Ошибка при чтении файла {file_path}: {e}")
        
        if all_dataframes:
            self.all_data = pd.concat(all_dataframes, ignore_index=True)
            logger.info(f"Общее количество записей: {len(self.all_data)}")
    
    def find_organization_columns(self) -> Dict[str, str]:
        """
        Автоматически определяет колонки с названием организации и веб-сайтом
        """
        columns = self.all_data.columns.tolist()
        
        # Возможные названия колонок для организаций
        org_patterns = [
            r'.*организ.*', r'.*company.*', r'.*наименование.*', 
            r'.*название.*', r'.*name.*', r'.*firm.*'
        ]
        
        # Возможные названия колонок для сайтов
        site_patterns = [
            r'.*сайт.*', r'.*site.*', r'.*web.*', r'.*url.*', 
            r'.*website.*', r'.*адрес.*сайт.*'
        ]
        
        org_col = None
        site_col = None
        
        # Ищем колонку организации
        for col in columns:
            for pattern in org_patterns:
                if re.match(pattern, col.lower()):
                    org_col = col
                    break
            if org_col:
                break
        
        # Ищем колонку сайта
        for col in columns:
            for pattern in site_patterns:
                if re.match(pattern, col.lower()):
                    site_col = col
                    break
            if site_col:
                break
        
        return {'organization': org_col, 'website': site_col}
    
    def remove_duplicates_and_validate(self) -> pd.DataFrame:
        """
        Удаляет дубликаты организаций и валидирует веб-адреса
        """
        if self.all_data.empty:
            logger.warning("Нет данных для обработки")
            return pd.DataFrame()
        
        # Определяем колонки
        column_mapping = self.find_organization_columns()
        org_col = column_mapping['organization']
        site_col = column_mapping['website']
        
        if not org_col:
            logger.error("Не удалось найти колонку с названиями организаций")
            # Попробуем использовать первую колонку
            org_col = self.all_data.columns[0]
            logger.info(f"Используем первую колонку как название организации: {org_col}")
        
        if not site_col:
            logger.error("Не удалось найти колонку с веб-адресами")
            # Попробуем найти колонку с URL-подобными данными
            for col in self.all_data.columns:
                sample_data = self.all_data[col].dropna().astype(str).head(10)
                if any('.' in str(val) and ('http' in str(val) or 'www' in str(val)) for val in sample_data):
                    site_col = col
                    logger.info(f"Найдена колонка с веб-адресами: {site_col}")
                    break
        
        logger.info(f"Колонка организаций: {org_col}")
        logger.info(f"Колонка веб-сайтов: {site_col}")
        
        # Создаем копию данных для обработки
        df_processed = self.all_data.copy()
        
        # Валидируем веб-адреса если есть соответствующая колонка
        if site_col and site_col in df_processed.columns:
            logger.info("Валидируем веб-адреса...")
            df_processed[site_col] = df_processed[site_col].apply(self.validate_and_fix_url)
            
            # Подсчитываем статистику валидации
            valid_urls = (df_processed[site_col] != "---").sum()
            total_urls = len(df_processed)
            logger.info(f"Валидных URL: {valid_urls} из {total_urls}")
        
        # Обрабатываем дубликаты по названию организации
        if org_col and org_col in df_processed.columns:
            logger.info("Удаляем дубликаты организаций...")
            
            # Группируем по названию организации
            grouped = df_processed.groupby(org_col)
            
            result_rows = []
            duplicates_count = 0
            
            for org_name, group in grouped:
                if len(group) > 1:
                    duplicates_count += len(group) - 1
                    logger.debug(f"Найдено {len(group)} дубликатов для '{org_name}'")
                    
                    # Если есть колонка с сайтами, выбираем запись с валидным URL
                    if site_col and site_col in group.columns:
                        valid_site_rows = group[group[site_col] != "---"]
                        if not valid_site_rows.empty:
                            # Берем первую запись с валидным сайтом
                            result_rows.append(valid_site_rows.iloc[0])
                        else:
                            # Если нет валидных сайтов, берем первую запись
                            result_rows.append(group.iloc[0])
                    else:
                        # Если нет колонки с сайтами, просто берем первую запись
                        result_rows.append(group.iloc[0])
                else:
                    result_rows.append(group.iloc[0])
            
            df_result = pd.DataFrame(result_rows)
            logger.info(f"Удалено дубликатов: {duplicates_count}")
            logger.info(f"Осталось уникальных организаций: {len(df_result)}")
            
        else:
            logger.warning("Не удалось обработать дубликаты - колонка организаций не найдена")
            df_result = df_processed
        
        return df_result
    
    def save_result(self, df: pd.DataFrame, output_filename: str = "ИТОГ.xlsx") -> None:
        """
        Сохраняет результат в Excel файл
        """
        if df.empty:
            logger.error("Нет данных для сохранения")
            return
        
        output_path = self.directory_path / output_filename
        
        try:
            # Удаляем служебную колонку перед сохранением
            if '_source_file' in df.columns:
                df_to_save = df.drop('_source_file', axis=1)
            else:
                df_to_save = df
            
            df_to_save.to_excel(output_path, index=False)
            logger.info(f"Результат сохранен в файл: {output_path}")
            logger.info(f"Количество записей в итоговом файле: {len(df_to_save)}")
            
        except Exception as e:
            logger.error(f"Ошибка при сохранении файла: {e}")
    
    def process(self) -> None:
        """
        Основной метод обработки
        """
        logger.info(f"Начинаем обработку файлов в директории: {self.directory_path}")
        
        # Читаем все Excel файлы
        self.read_excel_files()
        
        if self.all_data.empty:
            logger.error("Нет данных для обработки")
            return
        
        # Удаляем дубликаты и валидируем
        result_df = self.remove_duplicates_and_validate()
        
        # Сохраняем результат
        self.save_result(result_df)
        
        logger.info("Обработка завершена!")


def main():
    """
    Основная функция
    """
    import sys
    
    # Если передан аргумент командной строки, используем его как путь
    if len(sys.argv) > 1:
        directory_path = sys.argv[1]
    else:
        # Путь к директории с файлами (замените на актуальный)
        directory_path = r"d:\001\Kursis\Kurs_from_Cursor\Организации\00"
        
        # Для Linux/Unix систем путь может быть другим
        if os.name != 'nt':  # Если не Windows
            # Попробуем найти смонтированный диск или используем рабочую директорию
            possible_paths = [
                "/mnt/d/001/Kursis/Kurs_from_Cursor/Организации/00",
                "/workspace/excel_files",  # Директория для пользовательских файлов
                "/workspace/test_data",    # Для тестирования
                "/workspace"
            ]
            
            for path in possible_paths:
                if os.path.exists(path) and any(Path(path).glob("*.xlsx")):
                    directory_path = path
                    logger.info(f"Найдена директория с Excel файлами: {path}")
                    break
            else:
                print(f"""
ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ:

1. Скопируйте ваши Excel файлы в директорию /workspace/excel_files/
2. Или запустите скрипт с указанием пути: python3 process_excel_files.py /path/to/your/files
3. Для тестирования будут созданы примеры файлов в /workspace/test_data

Пример файлов, которые ожидает скрипт:
- ольгинка + 3км.xlsx
- другие файлы с расширением .xlsx

Создаем тестовые данные для демонстрации...
""")
                directory_path = "/workspace/test_data"
                create_test_data(directory_path)
    
    if not os.path.exists(directory_path):
        print(f"ОШИБКА: Директория {directory_path} не существует!")
        print("Создайте директорию и поместите в неё файлы Excel, или укажите правильный путь.")
        return
    
    processor = ExcelProcessor(directory_path)
    processor.process()


def create_test_data(directory_path: str):
    """
    Создает тестовые данные для демонстрации работы скрипта
    """
    import os
    os.makedirs(directory_path, exist_ok=True)
    
    # Тестовые данные
    test_data_1 = {
        'Название организации': [
            'ООО "Рога и Копыта"',
            'ИП Иванов И.И.',
            'ЗАО "Технологии будущего"',
            'ООО "Рога и Копыта"',  # дубликат
            'ООО "Мир компьютеров"'
        ],
        'Веб-сайт': [
            'www.roga-kopyta.ru',
            'invalid-url',
            'https://tech-future.com',
            'http://roga-kopyta.ru',  # дубликат с другим URL
            'mir-comp.net'
        ],
        'Телефон': [
            '+7 (495) 123-45-67',
            '+7 (926) 987-65-43',
            '+7 (812) 555-12-34',
            '+7 (495) 123-45-67',
            '+7 (499) 777-88-99'
        ]
    }
    
    test_data_2 = {
        'Название организации': [
            'ООО "Новые технологии"',
            'ИП Петров П.П.',
            'ООО "Мир компьютеров"',  # дубликат из первого файла
            'ООО "Строительная компания"'
        ],
        'Веб-сайт': [
            'https://new-tech.ru',
            'petrov-ip.com',
            '---',  # дубликат без сайта
            'not-a-valid-url-at-all'
        ],
        'Телефон': [
            '+7 (495) 111-22-33',
            '+7 (903) 444-55-66',
            '+7 (499) 777-88-99',
            '+7 (812) 999-00-11'
        ]
    }
    
    # Сохраняем тестовые файлы
    df1 = pd.DataFrame(test_data_1)
    df2 = pd.DataFrame(test_data_2)
    
    df1.to_excel(os.path.join(directory_path, 'организации_1.xlsx'), index=False)
    df2.to_excel(os.path.join(directory_path, 'организации_2.xlsx'), index=False)
    
    print(f"Созданы тестовые файлы в {directory_path}")


if __name__ == "__main__":
    main()