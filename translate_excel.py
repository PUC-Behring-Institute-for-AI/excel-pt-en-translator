#!/usr/bin/env python3
"""
Translate Excel file from Portuguese to English using Google Gemini API.

Usage:
    python translate_excel.py input.xlsx output.xlsx --api-key YOUR_KEY
    python translate_excel.py input.xlsx output.xlsx  # reads GEMINI_API_KEY env var
"""

import argparse
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta

import openpyxl
from openpyxl.cell.cell import MergedCell

try:
    from google import genai
except ImportError:
    print("ERROR: google-genai not installed. Run: pip install google-genai")
    sys.exit(1)


# ─── Progress logger ────────────────────────────────────────────────────────

class Progress:
    def __init__(self, total_cells: int, total_batches: int):
        self.total_cells = total_cells
        self.total_batches = total_batches
        self.cells_done = 0
        self.batches_done = 0
        self.start_time = time.time()
        self.current_sheet = ""

    def _elapsed(self) -> str:
        s = int(time.time() - self.start_time)
        return f"{s // 60:02d}:{s % 60:02d}"

    def _eta(self, batch_size: int) -> str:
        elapsed = time.time() - self.start_time
        if self.cells_done == 0:
            return "--:--"
        rate = self.cells_done / elapsed          # cells/sec
        remaining = self.total_cells - self.cells_done
        eta_s = int(remaining / rate) if rate > 0 else 0
        return f"{eta_s // 60:02d}:{eta_s % 60:02d}"

    def start_batch(self, batch_num: int, batch_size: int, sheet: str):
        pct = self.cells_done / self.total_cells * 100 if self.total_cells else 0
        bar_len = 25
        filled = int(bar_len * pct / 100)
        bar = "█" * filled + "░" * (bar_len - filled)
        print(
            f"\r[{bar}] {pct:5.1f}%  "
            f"batch {batch_num}/{self.total_batches}  "
            f"cells {self.cells_done}/{self.total_cells}  "
            f"elapsed {self._elapsed()}  "
            f"ETA {self._eta(batch_size)}  "
            f"sheet: {sheet[:30]}",
            end="  ",
            flush=True,
        )

    def finish_batch(self, batch_size: int):
        self.cells_done += batch_size
        self.batches_done += 1

    def done(self):
        elapsed = int(time.time() - self.start_time)
        print(
            f"\r[{'█' * 25}] 100.0%  "
            f"batch {self.total_batches}/{self.total_batches}  "
            f"cells {self.total_cells}/{self.total_cells}  "
            f"elapsed {elapsed // 60:02d}:{elapsed % 60:02d}  "
            f"ETA 00:00" + " " * 40
        )


# ─── Translation helpers ────────────────────────────────────────────────────

def should_translate(cell) -> bool:
    if isinstance(cell, MergedCell):
        return False
    value = cell.value
    if value is None or not isinstance(value, str):
        return False
    stripped = value.strip()
    if not stripped:
        return False
    if stripped.startswith(("http://", "https://")):
        return False
    return True


def translate_batch(client, model_name: str, texts: list) -> list:
    texts_json = json.dumps(texts, ensure_ascii=False, indent=2)

    prompt = f"""You are a professional translator. Translate the following texts from Portuguese to English.

STRICT RULES:
1. Proper nouns (person names, place names) → keep unchanged.
2. Institution / brand names → keep unchanged.
   Examples: PUC-Rio, ICA, CCEC, ECOA, BI MASTER, ExperIA, AcademIA, AI MASTER,
             AI LAB, Lattes, MCTI, IRI, Power BI, AutoGen, CrewAI, LlamaIndex,
             LangChain, HuggingFace, Google Analytics, Figma, CapCut, Zoom, Adobe.
3. Course codes → keep unchanged (e.g. BIA1001, MEC2007, CIS2114, CTN1408).
4. URLs → keep exactly unchanged.
5. Standard English / institution acronyms → keep unchanged.
   Examples: AI, ML, LLM, RAG, NLP, MoE, MCP, XAI, IDP, API, BI, ChatGPT, GPT.
6. English technical terms already in common use → keep in English.
   Examples: Machine Learning, Deep Learning, Bootcamp, Big Data, Dashboard,
             Streaming, Prompt Engineering, Embeddings, Pipeline, Chatbot.
7. Preserve internal line breaks (\\n) within each segment.
8. Return ONLY a valid JSON array with EXACTLY {len(texts)} strings, same order.
   No markdown, no commentary outside the JSON.

Input ({len(texts)} texts):
{texts_json}

Output (JSON array, {len(texts)} strings):"""

    max_retries = 6
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(model=model_name, contents=prompt)
            raw = response.text.strip()
            if raw.startswith("```"):
                lines = raw.splitlines()
                raw = "\n".join(lines[1:-1]).strip()

            translations = json.loads(raw)
            if not isinstance(translations, list) or len(translations) != len(texts):
                raise ValueError(f"Got {len(translations) if isinstance(translations, list) else '?'} items, expected {len(texts)}")
            return [str(t) for t in translations]

        except Exception as exc:
            err_str = str(exc)

            # Rate limit — read suggested retry delay and wait
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                delay = 62.0
                match = re.search(r"retry[^\d]*(\d+(?:\.\d+)?)\s*s", err_str, re.I)
                if match:
                    delay = float(match.group(1)) + 3
                if attempt < max_retries - 1:
                    print(f"\n  [rate limit] waiting {delay:.0f}s (attempt {attempt + 1}/{max_retries})…", flush=True)
                    time.sleep(delay)
                    continue
                else:
                    print(f"\n  WARNING: rate limit persisted after {max_retries} retries — keeping originals for this batch.")
                    return texts

            # Fatal errors — abort immediately
            if any(code in err_str for code in ("404", "403", "NOT_FOUND", "PERMISSION_DENIED")):
                print(f"\n\nFATAL: {exc}\n")
                sys.exit(1)

            # Other unexpected error
            print(f"\n  WARNING: unexpected error ({exc}) — keeping originals for this batch.")
            return texts

    return texts


# ─── Core workbook translation ───────────────────────────────────────────────

def translate_workbook(wb: openpyxl.Workbook, client, model_name: str, batch_size: int) -> None:
    # Collect all translatable cells grouped by sheet
    pending = []  # (sheet_name, row, col, text)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if should_translate(cell):
                    pending.append((sheet_name, cell.row, cell.column, cell.value))

    total = len(pending)
    texts = [item[3] for item in pending]
    total_batches = (total + batch_size - 1) // batch_size

    print(f"\n{'─' * 60}")
    print(f"  Sheets       : {len(wb.sheetnames)}")
    print(f"  Text cells   : {total}")
    print(f"  Batches      : {total_batches}  (batch size = {batch_size})")
    print(f"  Model        : {model_name}")
    print(f"  Started      : {datetime.now().strftime('%H:%M:%S')}")
    print(f"{'─' * 60}\n")

    progress = Progress(total, total_batches)
    translated = []

    for i in range(0, total, batch_size):
        batch = texts[i : i + batch_size]
        batch_num = i // batch_size + 1
        sheet_name = pending[i][0]

        progress.start_batch(batch_num, len(batch), sheet_name)
        result = translate_batch(client, model_name, batch)
        translated.extend(result)
        progress.finish_batch(len(batch))

        if i + batch_size < total:
            time.sleep(1.5)

    progress.done()

    # Write translations back
    for (sheet_name, row, col, _), new_value in zip(pending, translated):
        cell = wb[sheet_name].cell(row=row, column=col)
        if not isinstance(cell, MergedCell):
            cell.value = new_value

    print(f"\n{'─' * 60}")
    print(f"  Finished     : {datetime.now().strftime('%H:%M:%S')}")
    print(f"{'─' * 60}\n")


# ─── Entry point ────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Translate an Excel (.xlsx) file from Portuguese to English using Gemini."
    )
    parser.add_argument("input", help="Path to the source .xlsx file")
    parser.add_argument("output", help="Path for the translated .xlsx file")
    parser.add_argument(
        "--api-key",
        default=None,
        help="Google Gemini API key (falls back to GEMINI_API_KEY env var)",
    )
    parser.add_argument(
        "--model",
        default="gemini-2.5-flash",
        help="Gemini model name (default: gemini-2.5-flash)",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=20,
        help="Number of texts per API call (default: 20)",
    )
    args = parser.parse_args()

    api_key = args.api_key or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print(
            "ERROR: Gemini API key not found.\n"
            "  Pass --api-key YOUR_KEY  or  set the GEMINI_API_KEY environment variable.\n"
            "  Get a free key at: https://aistudio.google.com"
        )
        sys.exit(1)

    client = genai.Client(api_key=api_key)

    print(f"Loading: {args.input}")
    wb = openpyxl.load_workbook(args.input)

    translate_workbook(wb, client, args.model, args.batch_size)

    print(f"Saving: {args.output}")
    wb.save(args.output)
    print("Done!")


if __name__ == "__main__":
    main()
