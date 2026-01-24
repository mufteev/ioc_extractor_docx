import re
from typing import Iterator
from docx.document import Document
from ioc.normalizer import IoC, IoCNormalizer
from rules.rule import IoCExtractionRule
from typing import Iterator, Optional


class RegexPatternRule(IoCExtractionRule):
    """
    Правило для извлечения IoC с помощью регулярных выражений.
    
    Ищет хэши, URL, IP-адреса и другие паттерны во всём тексте документа.
    Это правило полезно как "сеть безопасности" для IoC, которые не попали
    в структурированные списки или таблицы.
    """
    
    # Паттерны для поиска (компилируем заранее для производительности)
    SEARCH_PATTERNS = {
        'hash_sha256': re.compile(r'\b[a-fA-F0-9]{64}\b'),
        'hash_sha1': re.compile(r'\b[a-fA-F0-9]{40}\b'),
        'hash_md5': re.compile(r'\b[a-fA-F0-9]{32}\b'),
        'hash_sha512': re.compile(r'\b[a-fA-F0-9]{128}\b'),
        'url_defanged': re.compile(
            r'hxxps?\[:\]//[^\s<>"\';,\n]+(?<![.,!?;:])',
            re.IGNORECASE
        ),
        'url_normal': re.compile(
            r'https?://[^\s<>"\';,\n]+(?<![.,!?;:])',
            re.IGNORECASE
        ),
        'ip_defanged': re.compile(
            r'\b\d{1,3}\[\.\]\d{1,3}\[\.\]\d{1,3}\[\.\]\d{1,3}\b'
        ),
        'ip_normal': re.compile(
            r'\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}'
            r'(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'
        ),
        'domain_defanged': re.compile(
            r'\b[a-zA-Z0-9][a-zA-Z0-9-]*(?:\[\.\][a-zA-Z0-9][a-zA-Z0-9-]*)+\[\.\][a-zA-Z]{2,}\b'
        ),
        'cve': re.compile(r'\bCVE-\d{4}-\d{4,}\b', re.IGNORECASE),
        'email': re.compile(r'\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b'),
    }
    
    def __init__(self, patterns: Optional[dict] = None):
        """
        Args:
            patterns: Словарь дополнительных паттернов {имя: regex}
                     Если None, используются только встроенные паттерны.
        """
        self._custom_patterns = patterns or {}
    
    @property
    def name(self) -> str:
        return "regex_pattern"
    
    def extract(self, document: Document) -> Iterator[IoC]:
        # Собираем весь текст из параграфов
        all_text_parts = []
        
        for paragraph in document.paragraphs:
            text = self._get_paragraph_text(paragraph)
            if text:
                all_text_parts.append(text)
        
        full_text = self.normalize_text_for_url_extraction('\n'.join(all_text_parts))
        
        # Применяем все паттерны
        seen = set()  # Для дедупликации
        
        all_patterns = {**self.SEARCH_PATTERNS, **self._custom_patterns}
        
        for pattern_name, pattern in all_patterns.items():
            for match in pattern.finditer(full_text):
                value = match.group()
                
                # Дедупликация по значению
                if value in seen:
                    continue
                seen.add(value)
                
                # Определяем контекст (несколько слов вокруг найденного значения)
                start = max(0, match.start() - 50)
                end = min(len(full_text), match.end() + 50)
                context = full_text[start:end].replace('\n', ' ').strip()
                
                yield IoCNormalizer.normalize_and_classify(value, f"[{pattern_name}] ...{context}...")

    def normalize_text_for_url_extraction(self, text: str) -> str:
        """
        Удаляет переносы строк, которые находятся внутри URL-подобных конструкций.
        
        Логика: если перенос строки окружён символами, типичными для URL
        (буквы, цифры, /, -, _, .), то это скорее всего "мусорный" перенос.
        """

        return re.sub(r'(?<=[a-zA-Z0-9/\-_.\[\]])\n(?=[a-zA-Z0-9/\-_.\[\]])', '', text)