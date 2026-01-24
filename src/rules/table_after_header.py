import re
from docx.table import Table
from docx.document import Document
from typing import Iterator
from ioc.normalizer import IoC, IoCNormalizer

from rules.rule import IoCExtractionRule


class TableAfterHeaderRule(IoCExtractionRule):
    """
    Правило для извлечения IoC из таблиц после заголовка "индикаторы компрометации".
    
    Ищет заголовки (выровненные по центру или содержащие ключевые слова),
    за которыми следует таблица с IoC в ячейках.
    """
    
    # Ключевые слова для поиска заголовка (без учёта регистра)
    HEADER_KEYWORDS = [
        "индикаторы компрометации",
        "indicators of compromise",
        "ioc",
        "iocs",
    ]
    
    @property
    def name(self) -> str:
        return "table_after_header"
    
    def extract(self, document: Document) -> Iterator[IoC]:
        # Собираем все элементы документа в порядке появления
        # (параграфы и таблицы могут чередоваться)
        elements = []
        
        for element in document.element.body:
            if element.tag.endswith('p'):
                # Это параграф - ищем соответствующий объект Paragraph
                for p in document.paragraphs:
                    if p._element is element:
                        elements.append(('paragraph', p))
                        break
            elif element.tag.endswith('tbl'):
                # Это таблица
                for t in document.tables:
                    if t._tbl is element:
                        elements.append(('table', t))
                        break
        
        # Ищем паттерн: заголовок -> таблица
        for i, (elem_type, elem) in enumerate(elements):
            if elem_type == 'paragraph':
                text = self._get_paragraph_text(elem).strip().lower()
                
                # Проверяем, содержит ли параграф ключевые слова
                if any(kw in text for kw in self.HEADER_KEYWORDS):
                    context = self._get_paragraph_text(elem).strip()
                    has_bold = any(run.bold is True for run in elem.runs)
                    
                    # Ищем следующую таблицу
                    for j in range(i + 1, len(elements)):
                        next_type, next_elem = elements[j]
                        
                        if next_type == 'table':
                            yield from self._extract_from_table(next_elem, context)
                            break
                        elif next_type == 'paragraph':
                            # Продолжаем, если параграф выделен жирным шрифтом
                            if any(run.bold is True for run in next_elem.runs) or self._get_paragraph_text(next_elem).strip() == "":
                                continue

                            # Если следующий параграф не выделен жирным и пустой, прерываем поиск
                            break
    
    def _extract_from_table(self, table: Table, context: str) -> Iterator[IoC]:
        """Извлекает IoC из всех ячеек таблицы, пропуская заголовки."""
        # Типичные слова в заголовках таблиц, которые нужно пропустить
        HEADER_WORDS = {
            'тип', 'type', 'значение', 'value', 'описание', 'description',
            'индикатор', 'indicator', 'hash', 'хэш', 'ip', 'domain', 'домен',
            'url', 'email', 'comment', 'комментарий', 'название', 'name',
            'sha256', 'sha1', 'md5', 'sha512'
        }
        
        for row_idx, row in enumerate(table.rows):
            for cell in row.cells:
                cell_text = self._get_cell_text(cell)
                
                if not cell_text:
                    continue
                
                # Пропускаем типичные заголовки
                if cell_text.lower().strip() in HEADER_WORDS:
                    continue 
                
                # Ячейка может содержать несколько IoC, разделённых переносами
                for line in cell_text.split('\n'):
                    line = line.strip()
                    if line and line.lower() not in HEADER_WORDS:
                        yield IoCNormalizer.normalize_and_classify(line, context)
