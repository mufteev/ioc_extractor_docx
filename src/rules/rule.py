
from abc import ABC, abstractmethod
from docx.document import Document
from typing import Iterator, Optional
from docx.text.paragraph import Paragraph
from ioc.normalizer import IoC

class IoCExtractionRule(ABC):
    """
    Абстрактный базовый класс для правил извлечения IoC.
    
    Для добавления нового правила создайте класс-наследник и реализуйте метод extract().
    
    Пример:
        class MyCustomRule(IoCExtractionRule):
            @property
            def name(self) -> str:
                return "my_custom_rule"
            
            def extract(self, document: Document) -> Iterator[IoC]:
                # Ваша логика извлечения
                yield IoC(value="...")
    """
    
    @property
    @abstractmethod
    def name(self) -> str:
        """Уникальное имя правила для идентификации."""
        pass
    
    @abstractmethod
    def extract(self, document: Document) -> Iterator[IoC]:
        """
        Извлекает IoC из документа.
        
        Args:
            document: Объект Document из python-docx
            
        Yields:
            Объекты IoC
        """
        pass
    
    def _get_paragraph_text(self, paragraph: Paragraph) -> str:
        """
        Извлекает текст из параграфа, игнорируя форматирование.
        
        Собирает текст из всех runs (фрагментов с одинаковым форматированием),
        что позволяет игнорировать стилизацию.
        """
        return "".join(run.text for run in paragraph.runs)
    
    def _get_cell_text(self, cell) -> str:
        """Извлекает текст из ячейки таблицы."""
        return "\n".join(
            self._get_paragraph_text(p) for p in cell.paragraphs
        ).strip()

