

import json
from pathlib import Path
from ioc.extractor import ExtractionResult


def export_to_json(
    results: dict[str, ExtractionResult],
    output_path: str | Path,
    include_context: bool = True
):
    output = {}

    for filepath, result in results.items():
        iocs = []
        for ioc in result.iocs:
            ioc_dict = {
                "value": ioc.value,
                "type": ioc.ioc_type.name,
            }
            if include_context:
                ioc_dict["original"] = ioc.original_value
                if ioc.source_context:
                    ioc_dict["context"] = ioc.source_context
            iocs.append(ioc_dict)
        
        output[filepath] = {
            "iocs": iocs,
            "errors": result.errors
        }

    json_output = json.dumps(output, indent=4)

    Path(output_path).write_text(json_output, encoding="utf-8")
