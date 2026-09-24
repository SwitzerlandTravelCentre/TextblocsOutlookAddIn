"""Download the configured SharePoint workbook and validate a candidate before use.

No workbook content or download credentials are printed. Graph access is read-only.
--source provides an explicit local validation path; it is never a download fallback.
"""
import argparse
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import re
import tempfile
from urllib.parse import quote, urlsplit

import requests


def validate(data):
    if data.get("schemaVersion") != 1:
        raise ValueError("Unsupported data schema")
    blocks = data.get("blocks")
    if not isinstance(blocks, list) or not blocks or data.get("blockCount") != len(blocks):
        raise ValueError("Empty or inconsistent conversion")
    seen = set()
    for block in blocks:
        for field in ("id", "sourceItemId", "category", "topic", "usage", "languageLabel", "text"):
            if not isinstance(block.get(field), str):
                raise ValueError("Invalid text block field: " + field)
        if not block["id"] or block["id"] in seen:
            raise ValueError("Duplicate or empty text block ID")
        seen.add(block["id"])
        if block.get("language") not in {"DE", "FR", "EN", "IT"} or not block["text"].strip():
            raise ValueError("Invalid language or empty text")
        weight = block.get("weight", 0)
        if isinstance(weight, bool) or not isinstance(weight, (int, float)) or not math.isfinite(weight):
            raise ValueError("Weight is not a finite number")
        if "formattedText" in block:
            runs = block["formattedText"]
            if not isinstance(runs, list) or not all(isinstance(r, dict) for r in runs):
                raise ValueError("Invalid formatted text")
            if "".join(r.get("text", "") for r in runs) != block["text"]:
                raise ValueError("Formatted text differs from plain text")
            for run in runs:
                if not isinstance(run.get("text"), str):
                    raise ValueError("Invalid formatted text run")
                if "bold" in run and not isinstance(run["bold"], bool):
                    raise ValueError("Invalid bold flag")
                if "highlight" in run and run["highlight"] != "required":
                    raise ValueError("Invalid highlight")
                if "href" in run and not re.match(r"^(https?://|mailto:)", run["href"], re.I):
                    raise ValueError("Unsupported link scheme")
    return data


def download(destination):
    token = os.environ["TEXTBLOCKS_GRAPH_TOKEN"]
    drive = quote(os.environ["TEXTBLOCKS_DRIVE_ID"], safe="")
    item = quote(os.environ["TEXTBLOCKS_ITEM_ID"], safe="")
    url = f"https://graph.microsoft.com/v1.0/drives/{drive}/items/{item}"
    headers = {"Authorization": "Bearer " + token}
    response = requests.get(url, headers=headers, timeout=30)
    if response.status_code != 200:
        raise RuntimeError(f"Source metadata unavailable: HTTP {response.status_code}")
    meta = response.json()
    if meta.get("name") != "STC_Textblocs_Source.xlsx" or "file" not in meta:
        raise ValueError("Configured item is not the expected source workbook")
    if not 0 < meta.get("size", 0) <= 25 * 1024 * 1024:
        raise ValueError("Workbook size is empty or exceeds the download limit")
    content = requests.get(url + "/content", headers=headers, allow_redirects=False, timeout=30)
    if content.status_code == 302:
        location = content.headers.get("Location", "")
        host = urlsplit(location).hostname or ""
        if urlsplit(location).scheme != "https" or not host.endswith(".sharepoint.com"):
            raise ValueError("Unexpected SharePoint download destination")
        # Never forward a Microsoft Graph bearer token to the download host.
        content = requests.get(location, timeout=60)
    if content.status_code != 200:
        raise RuntimeError(f"Workbook download failed: HTTP {content.status_code}")
    if len(content.content) != meta["size"] or not content.content.startswith(b"PK"):
        raise ValueError("Workbook download is incomplete or not an XLSX file")
    check = requests.get(url, headers=headers, timeout=30)
    if check.status_code != 200 or check.json().get("eTag") != meta.get("eTag"):
        raise ValueError("Source changed during download; retry on the next run")
    destination.write_bytes(content.content)


def refresh(source, output):
    spec = importlib.util.spec_from_file_location("textblocks_converter", Path(__file__).with_name("convert-xlsx-to-json.py"))
    converter = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(converter)
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=output.parent) as stage:
        candidate = Path(stage) / "textblocks.json"
        data = validate(converter.convert(source, candidate))
        # The previous output survives every failed download/conversion/validation.
        os.replace(candidate, output)
    print(json.dumps({"rows": data["source"]["dataRowCount"], "blocks": data["blockCount"],
                      "sha256": hashlib.sha256(output.read_bytes()).hexdigest()}))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, help="Explicit local test input (not used by the scheduled job)")
    parser.add_argument("--output", type=Path, default=Path("public/data/textblocks.json"))
    args = parser.parse_args()
    if args.source:
        refresh(args.source, args.output)
    else:
        with tempfile.TemporaryDirectory() as work:
            source = Path(work) / "STC_Textblocs_Source.xlsx"
            download(source)
            refresh(source, args.output)


if __name__ == "__main__":
    try:
        main()
    except (requests.RequestException, KeyError):
        raise SystemExit("Source access failed or configuration is incomplete. No publication performed.")
