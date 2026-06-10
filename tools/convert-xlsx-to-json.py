import argparse
import json
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

import openpyxl


LANGUAGE_COLUMNS = [
    ("German", "DE", "German"),
    ("French", "FR", "French"),
    ("English", "EN", "English"),
    ("Englisch", "EN", "English"),
    ("Italian", "IT", "Italian"),
]


def clean_header(value):
    return str(value or "").strip()


def clean_text(value):
    if value is None:
        return ""

    text = str(value).replace("\r\n", "\n").replace("\r", "\n").strip()
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    if len(text) >= 2 and text[0] == text[-1] == '"':
        text = text[1:-1].strip()

    return text


def parse_weight(value):
    if value is None or str(value).strip() == "":
        return 0

    try:
        number = float(str(value).strip().replace(",", "."))
    except ValueError:
        return 0

    return int(number) if number.is_integer() else number


def slug(value):
    normalized = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    normalized = re.sub(r"[^a-zA-Z0-9]+", "-", normalized).strip("-").lower()
    return normalized[:70] or "textblock"


def find_header_row(sheet):
    for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 20), values_only=False):
        row_headers = [clean_header(cell.value) for cell in row]

        if "Category" in row_headers and "Topic" in row_headers:
            headers = {clean_header(cell.value): cell.column for cell in row if clean_header(cell.value)}
            return row[0].row, headers

    return None, {}


def convert(source_path, output_path):
    workbook = openpyxl.load_workbook(source_path, data_only=True)
    blocks = []
    source_rows = 0

    for sheet in workbook.worksheets:
        header_row_index, headers = find_header_row(sheet)

        if header_row_index is None:
            continue

        category_column = headers.get("Category")
        topic_column = headers.get("Topic")
        usage_column = headers.get("Usage")
        weight_column = headers.get("Weight") or headers.get("weight")
        language_map = []
        seen_language_codes = set()

        for field, code, label in LANGUAGE_COLUMNS:
            if field in headers and code not in seen_language_codes:
                language_map.append((field, code, label, headers[field]))
                seen_language_codes.add(code)

        for row_index in range(header_row_index + 1, sheet.max_row + 1):
            category = clean_text(sheet.cell(row_index, category_column).value) if category_column else ""
            topic = clean_text(sheet.cell(row_index, topic_column).value) if topic_column else ""
            usage = clean_text(sheet.cell(row_index, usage_column).value) if usage_column else ""
            weight = parse_weight(sheet.cell(row_index, weight_column).value) if weight_column else 0
            language_values = [
                (field, code, label, clean_text(sheet.cell(row_index, column).value))
                for field, code, label, column in language_map
            ]

            if not category and not topic and not usage and not any(text for *_, text in language_values):
                continue

            source_rows += 1
            base_id = f"{slug(category)}-{slug(topic)}-{row_index}"

            for _field, code, label, text in language_values:
                if not text:
                    continue

                blocks.append(
                    {
                        "id": f"{base_id}-{code.lower()}",
                        "sourceItemId": f"{sheet.title}!{row_index}",
                        "sourceSheet": sheet.title,
                        "sourceRow": row_index,
                        "category": category,
                        "topic": topic,
                        "usage": usage,
                        "language": code,
                        "languageLabel": label,
                        "weight": weight,
                        "text": text,
                    }
                )

    payload = {
        "schemaVersion": 1,
        "source": {
            "type": "xlsx",
            "fileName": source_path.name,
            "sheetCount": len(workbook.worksheets),
            "dataRowCount": source_rows,
        },
        "generatedAt": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "blockCount": len(blocks),
        "blocks": blocks,
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    return payload


def main():
    parser = argparse.ArgumentParser(description="Convert the STC email text block workbook to JSON.")
    parser.add_argument("source", type=Path)
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("public/data/textblocks.json"),
    )
    args = parser.parse_args()
    payload = convert(args.source, args.output)
    print(json.dumps({"rows": payload["source"]["dataRowCount"], "blocks": payload["blockCount"]}))


if __name__ == "__main__":
    main()
