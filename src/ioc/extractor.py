#!/usr/bin/env python3
"""
IoC Extractor - Профессиональный инструмент для извлечения индикаторов компрометации из DOCX-документов.

Архитектура построена на паттерне "Стратегия", что позволяет легко добавлять новые правила
извлечения без изменения основного кода.

Использование:
    from ioc_extractor import IoCExtractor, extract_iocs_from_files
    
    # Простой вариант - извлечение из нескольких файлов
    results = extract_iocs_from_files(["doc1.docx", "doc2.docx"])
    
    # Расширенный вариант - с пользовательскими правилами
    extractor = IoCExtractor()
    extractor.add_rule(MyCustomRule())
    results = extractor.extract_from_files(["doc1.docx"])
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator, Optional
from urllib.parse import urlparse

from docx import Document as DocxDocument
from docx.document import Document

from rules.list_after_colon import ListAfterColonRule
from rules.table_after_header import TableAfterHeaderRule
from rules.regex_pattern import RegexPatternRule

from ioc.normalizer import IoC, IoCType
from rules.rule import IoCExtractionRule

@dataclass
class ExtractionResult:
    """
    Результат извлечения IoC из одного файла.
    
    Attributes:
        filepath: Путь к обработанному файлу
        iocs: Список найденных индикаторов
        errors: Список ошибок, возникших при обработке
    """
    filepath: str
    iocs: list[IoC] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    
    def __len__(self):
        return len(self.iocs)
    
    def by_type(self, ioc_type: IoCType) -> list[IoC]:
        """Возвращает IoC определённого типа."""
        return [ioc for ioc in self.iocs if ioc.ioc_type == ioc_type]
    
    def unique_values(self) -> set[str]:
        """Возвращает уникальные значения IoC."""
        return {ioc.value for ioc in self.iocs}

class IoCExtractor:
    """
    Основной класс для извлечения IoC из документов.
    
    Координирует работу правил извлечения и обеспечивает:
    - Применение нескольких правил
    - Дедупликацию результатов
    - Обработку ошибок
    
    Примеры использования:
        # Базовое использование
        extractor = IoCExtractor()
        result = extractor.extract("document.docx")
        
        # Добавление пользовательского правила
        extractor.add_rule(MyCustomRule())
        
        # Обработка нескольких файлов
        results = extractor.extract_from_files(["doc1.docx", "doc2.docx"])
    """
    
    def __init__(self, use_default_rules: bool = True):
        """
        Args:
            use_default_rules: Если True, добавляет стандартные правила извлечения.
        """
        self._rules: list[IoCExtractionRule] = []
        
        if use_default_rules:
            self._rules.extend([
                ListAfterColonRule(),
                TableAfterHeaderRule(),
                RegexPatternRule(),
            ])
    
    def add_rule(self, rule: IoCExtractionRule) -> "IoCExtractor":
        """
        Добавляет правило извлечения.
        
        Args:
            rule: Экземпляр правила
            
        Returns:
            self для цепочки вызовов
        """
        self._rules.append(rule)
        return self
    
    def remove_rule(self, rule_name: str) -> "IoCExtractor":
        """
        Удаляет правило по имени.
        
        Args:
            rule_name: Имя правила для удаления
            
        Returns:
            self для цепочки вызовов
        """
        self._rules = [r for r in self._rules if r.name != rule_name]
        return self
    
    def get_rules(self) -> list[str]:
        """Возвращает список имён активных правил."""
        return [r.name for r in self._rules]
    

    def extract_base_domain_from_url(self, url: str) -> str:
        """
        Извлекает базовый домен из URL.
        
        Args:
            url: Полный URL
            
        Returns:
            Базовый домен (например, example.com)
        """
        parsed = urlparse(url)
        host_with_port = parsed.netloc
        
        if ':' in host_with_port:
            # Проверяем, что после : идут только цифры (это порт)
            parts = host_with_port.rsplit(':', 1)
            if parts[1].isdigit():
                return parts[0]
        
        return host_with_port
    
    def extract_base_domain_from_domain(self, domain: str) -> str:
        """
        Извлекает базовый домен из доменного имени.
        
        Args:
            domain: Полное доменное имя
            
        Returns:
            Базовый домен (например, example.com)
        """
        # Убираем порт, если он присутствует
        if ':' in domain:
            # Отделяем порт (всё после последнего двоеточия)
            return domain.rsplit(':', 1)[0]
        
        return domain

    def extract(self, filepath: str | Path, pass_unknown: bool = False, url_original: bool = False) -> ExtractionResult:
        """
        Извлекает IoC из одного файла.
        
        Args:
            filepath: Путь к DOCX-файлу
            
        Returns:
            ExtractionResult с найденными IoC и возможными ошибками
        """
        filepath = Path(filepath)
        result = ExtractionResult(filepath=str(filepath))
        
        if not filepath.exists():
            result.errors.append(f"Файл не найден: {filepath}")
            return result
        
        if not filepath.suffix.lower() == '.docx':
            result.errors.append(f"Неподдерживаемый формат файла: {filepath.suffix}")
            return result
        
        try:
            document = DocxDocument(str(filepath))
        except Exception as e:
            result.errors.append(f"Ошибка открытия документа: {e}")
            return result
        
        # Применяем все правила и собираем IoC
        seen_iocs = set()  # Для дедупликации по (value, type)
        
        for rule in self._rules:
            try:
                for ioc in rule.extract(document):
                    if not pass_unknown and ioc.ioc_type == IoCType.UNKNOWN:
                        continue

                    if not url_original:
                        if ioc.ioc_type == IoCType.URL:
                            ioc.value = self.extract_base_domain_from_url(ioc.value)
                            ioc.ioc_type = IoCType.DOMAIN
                        elif ioc.ioc_type == IoCType.DOMAIN:
                            ioc.value = self.extract_base_domain_from_domain(ioc.value)

                    # Дедупликация
                    key = (ioc.value, ioc.ioc_type)
                    if key not in seen_iocs:
                        seen_iocs.add(key)
                        ioc.rule_extracted = rule.name
                        result.iocs.append(ioc)
            except Exception as e:
                result.errors.append(f"Ошибка в правиле '{rule.name}': {e}")
        
        return result
    
    def extract_from_files(
        self, 
        filepaths: list[str | Path],
        pass_unknown: bool = False,
        url_original: bool = False,
    ) -> dict[str, ExtractionResult]:
        """
        Извлекает IoC из нескольких файлов.
        
        Args:
            filepaths: Список путей к файлам
            
        Returns:
            Словарь {путь_к_файлу: ExtractionResult}
        """
        results = {}
        
        for filepath in filepaths:
            filepath_str = str(filepath)
            results[filepath_str] = self.extract(filepath,
                                                 pass_unknown=pass_unknown,
                                                 url_original=url_original)
        
        return results


def extract_iocs_from_files(
    filepaths: list[str | Path],
    custom_rules: Optional[list[IoCExtractionRule]] = None,
    pass_unknown: bool = False,
    hash_original: bool = False,
    url_original: bool = False,
) -> dict[str, ExtractionResult]:
    """
    Удобная функция для извлечения IoC из нескольких файлов.
    
    Args:
        filepaths: Список путей к DOCX-файлам
        custom_rules: Дополнительные правила извлечения (опционально)
        
    Returns:
        Словарь {путь_к_файлу: ExtractionResult}
        
    Пример:
        results = extract_iocs_from_files(["report1.docx", "report2.docx"])
        
        for filepath, result in results.items():
            print(f"\\nФайл: {filepath}")
            print(f"Найдено IoC: {len(result)}")
            
            for ioc in result.iocs:
                print(f"  [{ioc.ioc_type.name}] {ioc.value}")
    """
    extractor = IoCExtractor(use_default_rules=True,
                             hash_original=hash_original)
    
    if custom_rules:
        for rule in custom_rules:
            extractor.add_rule(rule)
    
    return extractor.extract_from_files(filepaths,
                                        pass_unknown=pass_unknown,
                                        url_original=url_original)


# ============================================================================
# Пример создания пользовательского правила
# ============================================================================

class CustomMitreATTCKRule(IoCExtractionRule):
    """
    Пример пользовательского правила для извлечения идентификаторов MITRE ATT&CK.
    
    Демонстрирует, как легко расширить функциональность экстрактора.
    """
    
    MITRE_PATTERN = re.compile(r'\b(T\d{4}(?:\.\d{3})?)\b')
    
    @property
    def name(self) -> str:
        return "mitre_attck"
    
    def extract(self, document: Document) -> Iterator[IoC]:
        for paragraph in document.paragraphs:
            text = self._get_paragraph_text(paragraph)
            
            for match in self.MITRE_PATTERN.finditer(text):
                yield IoC(
                    value=match.group(1),
                    ioc_type=IoCType.UNKNOWN,  # Можно создать свой тип
                    source_context=text[:100]
                )
