from pathlib import Path
from ioc.extractor import ExtractionResult


def export_to_text(
    results: dict[str, ExtractionResult],
    output_path: str | Path
):
    
    lines = []
    for filepath, result in results.items():
        lines.append(f"\n{'='*60}")
        lines.append(f"Файл: {filepath}")
        lines.append(f"Найдено IoC: {len(result)}")
            
        if result.errors:
            lines.append(f"Ошибки: {', '.join(result.errors)}")
        
        lines.append("-" * 60)

        for ioc in result.iocs:
            lines.append(ioc.value)

    Path(output_path).write_text('\n'.join(lines) + '\n', encoding="utf-8")
