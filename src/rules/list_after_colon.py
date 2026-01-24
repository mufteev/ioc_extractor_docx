import re
from docx.document import Document
from typing import Iterator
from ioc.normalizer import IoC, IoCNormalizer
from rules.rule import IoCExtractionRule

class ListAfterColonRule(IoCExtractionRule):
    """
    Правило для извлечения IoC из списков, следующих после двоеточия.
    
    Формат:
        Заголовок с двоеточием:
        индикатор1;
        индикатор2;
        последний_индикатор.
    
    Поддерживает:
    - Списки с точкой с запятой в конце каждого элемента
    - Последний элемент с точкой
    - Списки из одного элемента
    """
    
    @property
    def name(self) -> str:
        return "list_after_colon"
    
    def extract(self, document: Document) -> Iterator[IoC]:
        paragraphs = list(document.paragraphs)
        i = 0
        
        while i < len(paragraphs):
            text = self._get_paragraph_text(paragraphs[i]).strip()
            
            # Ищем параграф, заканчивающийся двоеточием
            if text.endswith(':'):
                context = text
                i += 1
                
                # Собираем элементы списка
                while i < len(paragraphs):
                    item_text = self._get_paragraph_text(paragraphs[i]).strip()
                    
                    if not item_text:
                        i += 1
                        continue
                    
                    # Элемент с точкой с запятой - продолжаем собирать
                    if item_text.endswith(';'):
                        value = item_text[:-1].strip()
                        if value:
                            yield IoCNormalizer.normalize_and_classify(value, context)
                        i += 1
                        continue
                    
                    # Элемент с точкой - последний в списке
                    if item_text.endswith('.'):
                        value = item_text[:-1].strip()
                        if value:
                            yield IoCNormalizer.normalize_and_classify(value, context)
                        i += 1
                        break
                    
                    # Элемент без разделителя - возможно, часть того же списка
                    # или уже начался новый контекст
                    # Проверяем, похоже ли это на IoC
                    if self._looks_like_ioc(item_text):
                        yield IoCNormalizer.normalize_and_classify(item_text, context)
                        i += 1
                        continue
                    
                    # Это уже другой контекст
                    break
            else:
                i += 1
    
    def _looks_like_ioc(self, text: str) -> bool:
        """Проверяет, похоже ли значение на IoC."""
        # Хэши (32, 40, 64, 128 символов hex)
        if re.match(r'^[a-fA-F0-9]{32,128}$', text):
            return True
        # URL-подобные строки
        if re.match(r'^(?:hxxps?|https?|ftp)', text, re.IGNORECASE):
            return True
        # IP-адреса
        if re.match(r'^\d{1,3}(?:\[?\.\]?\d{1,3}){3}', text):
            return True
        return False

