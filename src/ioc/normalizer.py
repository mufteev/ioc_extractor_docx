
import re
from enum import Enum, auto
from dataclasses import dataclass



class IoCType(Enum):
    """Типы индикаторов компрометации."""
    HASH_MD5 = auto()
    HASH_SHA1 = auto()
    HASH_SHA256 = auto()
    HASH_SHA512 = auto()
    URL = auto()
    DOMAIN = auto()
    IP_ADDRESS = auto()
    EMAIL = auto()
    FILE_PATH = auto()
    REGISTRY_KEY = auto()
    CVE = auto()
    UNKNOWN = auto()


@dataclass
class IoC:
    """
    Представление одного индикатора компрометации.
    
    Attributes:
        value: Значение индикатора (например, хэш или URL)
        ioc_type: Тип индикатора (см. IoCType)
        source_context: Контекст, в котором был найден индикатор
        defanged: Был ли индикатор "обезврежен" (например, hxxps вместо https)
        original_value: Исходное значение до нормализации (если применимо)
    """
    value: str
    ioc_type: IoCType = IoCType.UNKNOWN
    source_context: str = ""
    defanged: bool = False
    original_value: str = ""
    rule_extracted: str = ""
    
    def __post_init__(self):
        # Если original_value не указан, используем value
        if not self.original_value:
            self.original_value = self.value
    
    def __hash__(self):
        return hash((self.value, self.ioc_type))
    
    def __eq__(self, other):
        if not isinstance(other, IoC):
            return False
        return self.value == other.value and self.ioc_type == other.ioc_type




class IoCNormalizer:
    """
    Утилитарный класс для нормализации и определения типов IoC.
    
    Содержит методы для:
    - "Дефангинга" и "рефангинга" URL/доменов
    - Определения типа хэша по длине
    - Классификации IoC по паттернам
    """
    
    # Паттерны для "дефангинга" (обезвреживания) URL
    DEFANG_PATTERNS = [
        (r'hxxps', 'https'),              # hxxps -> https
        (r'hxxp', 'http'),                # hxxp -> http
        (r'\[:\]', ':'),                   # [:] -> :
        (r'\[\.\]', '.'),                  # [.] -> .
        (r'\[dot\]', '.', re.IGNORECASE),  # [dot] -> .
        (r'\[@\]', '@'),                   # [@] -> @
        (r'\[at\]', '@', re.IGNORECASE),   # [at] -> @
    ]
    
    # Регулярные выражения для определения типов IoC
    PATTERNS = {
        IoCType.HASH_MD5: re.compile(r'^[a-fA-F0-9]{32}$'),
        IoCType.HASH_SHA1: re.compile(r'^[a-fA-F0-9]{40}$'),
        IoCType.HASH_SHA256: re.compile(r'^[a-fA-F0-9]{64}$'),
        IoCType.HASH_SHA512: re.compile(r'^[a-fA-F0-9]{128}$'),
        IoCType.IP_ADDRESS: re.compile(
            r'^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}'
            r'(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
        ),
        IoCType.EMAIL: re.compile(
            r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        ),
        IoCType.CVE: re.compile(r'^CVE-\d{4}-\d{4,}$', re.IGNORECASE),
        IoCType.URL: re.compile(
            r'^(?:hxxps?|https?|ftp)://[^\s;.]+(?:\.[^\s;.]+)*(?:/[^\s]*)?$',
            re.IGNORECASE
        ),
        IoCType.DOMAIN: re.compile(
            r'^(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?'
            r'(?:\[\.\]|\.)){1,}[a-zA-Z]{2,}$'
        ),
    }
    
    @classmethod
    def refang(cls, value: str) -> tuple[str, bool]:
        """
        Преобразует "обезвреженный" URL/домен в нормальный вид.
        
        Args:
            value: Строка, возможно содержащая дефангированные элементы
            
        Returns:
            Кортеж (нормализованное_значение, был_ли_дефангирован)
        """
        result = value
        was_defanged = False
        
        for pattern_tuple in cls.DEFANG_PATTERNS:
            pattern = pattern_tuple[0]
            replacement = pattern_tuple[1]
            flags = pattern_tuple[2] if len(pattern_tuple) > 2 else 0
            
            new_result = re.sub(pattern, replacement, result, flags=flags)
            if new_result != result:
                was_defanged = True
                result = new_result
        
        return result, was_defanged
    
    @classmethod
    def classify(cls, value: str) -> IoCType:
        """
        Определяет тип IoC на основе паттернов.
        
        Args:
            value: Значение индикатора
            
        Returns:
            Тип индикатора (IoCType)
        """
        # Сначала пробуем определить как хэш (наиболее строгие паттерны)
        for ioc_type in [IoCType.HASH_MD5, IoCType.HASH_SHA1, 
                         IoCType.HASH_SHA256, IoCType.HASH_SHA512]:
            if cls.PATTERNS[ioc_type].match(value):
                return ioc_type
        
        # Затем другие типы
        for ioc_type in [IoCType.CVE, IoCType.IP_ADDRESS, IoCType.EMAIL,
                         IoCType.URL, IoCType.DOMAIN]:
            if cls.PATTERNS[ioc_type].match(value):
                return ioc_type
        
        return IoCType.UNKNOWN
    
    @classmethod
    def normalize_and_classify(cls, value: str, context: str = "") -> IoC:
        """
        Нормализует значение и создаёт объект IoC с определённым типом.
        
        Args:
            value: Исходное значение
            context: Контекст, в котором найден индикатор
            
        Returns:
            Объект IoC
        """
        original = value.strip()
        
        # Очищаем от разделителей списков (;.) в конце
        cleaned = original.rstrip(';.')
        
        normalized, was_defanged = cls.refang(cleaned)
        ioc_type = cls.classify(normalized)
        
        return IoC(
            value=normalized,
            ioc_type=ioc_type,
            source_context=context,
            defanged=was_defanged,
            original_value=original
        )

