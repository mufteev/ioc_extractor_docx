

import csv
from pathlib import Path
from ioc.extractor import ExtractionResult


def export_to_csv(
    results: dict[str, ExtractionResult],
    output_path: str | Path,
    include_context: bool = True
):
    """
    Экспортирует результаты извлечения IoC в CSV формат.
    
    Функция разворачивает вложенную структуру данных в плоскую таблицу,
    где каждая строка представляет один найденный индикатор компрометации.
    
    Args:
        results: Словарь с результатами извлечения (ключ - имя файла)
        output_path: Путь к выходному CSV файлу
        include_errors: Создавать ли отдельный файл с ошибками
        encoding: Кодировка для CSV (utf-8-sig добавляет BOM для Excel)
    """
    output_path = Path(output_path)
    
    # Преобразуем вложенную структуру в список словарей
    # Каждый IoC становится отдельной строкой с информацией об источнике
    fieldnames = [
        'filepath', 'value', 'ioc_type', 'source_context', 
        'defanged', 'original_value']
    ioc_records = []
    
    for result in results.values():
        for ioc in result.iocs:
            ioc_records.append({
                'filepath': result.filepath,
                'value': ioc.value,
                'ioc_type': ioc.ioc_type.name,  # Преобразуем Enum в строку
                'source_context': ioc.source_context,
                'defanged': ioc.defanged,
                'original_value': ioc.original_value
            })
    
    with open(output_path, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, quoting=csv.QUOTE_MINIMAL, escapechar='\\')
        writer.writeheader()
        writer.writerows(ioc_records)
