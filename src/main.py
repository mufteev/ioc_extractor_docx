#!/usr/bin/env python3

import argparse
import re
import sys

from ioc.extractor import extract_iocs_from_files
from out.json import export_to_json
from out.txt import export_to_text
from out.xlsx import export_to_excel
from out.csv import export_to_csv

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Извлечение индикаторов компрометации из DOCX-документов"
    )
    parser.add_argument(
        "files",
        nargs="+",
        help="Пути к DOCX-файлам для обработки"
    )
    parser.add_argument(
        "-o", "--output",
        required=True,
        help="Файл для сохранения результатов в JSON (по умолчанию: stdout)"
    )
    parser.add_argument(
        "-f", "--format",
        choices=["json", "txt", "csv", "xlsx"],
        default="txt",
        help="Формат вывода (по умолчанию: txt)"
    )
    parser.add_argument(
        "--unknown",
        action="store_true",
        help="Выводить unknown IoC (по умолчанию: не выводить)"
    )
    parser.add_argument(
        "--hash-original",
        action="store_true",
        help="Выводить оригинальные хэши (по умолчанию: приведение к верхнему регистру)"
    )
    parser.add_argument(
        "--url-original",
        action="store_true",
        help="Выводить оригинальные URL (по умолчанию: извлечение доменов и IP-адресов)"
    )
    args = parser.parse_args()

                                      hash_original=args.hash_original,
                                      url_original=args.url_original)

    if args.format:
        if args.format == "json":
            output_file = args.output or "ioc_results.json"
            export_to_json(results, output_file)
        elif args.format == "txt":
            output_file = args.output or "ioc_results.txt"
            export_to_text(results, output_file)
        elif args.format == "xlsx":
            output_file = args.output or "ioc_results.xlsx"
            if not output_file.endswith('.xlsx'):
                output_file += '.xlsx'
            export_to_excel(results, output_file)
        elif args.format == "csv":
            output_file = args.output or "ioc_results.csv"
            export_to_csv(results, output_file)

    print(f"Результаты сохранены в {output_file}")
    
    # Выводим краткую статистику
    total = sum(len(r) for r in results.values())
    print(f"Всего извлечено IoC: {total} из {len(results)} файлов")
    sys.exit(0)