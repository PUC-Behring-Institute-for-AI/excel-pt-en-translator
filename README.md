# Excel PT → EN Translator

Translates an Excel (`.xlsx`) file from **Portuguese to English** using the **Google Gemini API**.

- Preserves proper nouns, institution names, course codes, URLs, and English technical terms
- Handles multiple sheets automatically
- Retries automatically on rate-limit errors with the delay suggested by the API
- Real-time progress bar with elapsed time and ETA

---

## Getting a Gemini API Key

1. Go to **[aistudio.google.com](https://aistudio.google.com)**
2. Sign in with your Google account
3. Click **"Get API key"** → **"Create API key"**
4. Copy the key — it starts with `AIza…`

> The free tier of `gemini-2.5-flash` is enough to translate a spreadsheet with a few thousand cells.

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Usage

```bash
# Pass the key directly
python translate_excel.py input.xlsx output.xlsx --api-key AIzaSy...

# Or export it as an environment variable (recommended)
export GEMINI_API_KEY=AIzaSy...
python translate_excel.py input.xlsx output.xlsx
```

### Options

| Option | Default | Description |
|---|---|---|
| `--api-key` | `GEMINI_API_KEY` env var | Your Gemini API key |
| `--model` | `gemini-2.5-flash` | Gemini model to use |
| `--batch-size` | `20` | Number of cells sent per API call |

### Testing available models

If you hit quota errors, run this to see which models your key supports:

```python
from google import genai, types
import time

client = genai.Client(api_key="YOUR_KEY")
for m in client.models.list():
    if "generateContent" in (m.supported_actions or []):
        try:
            r = client.models.generate_content(model=m.name, contents="Say hi")
            print(f"OK   {m.name}")
        except Exception as e:
            print(f"FAIL {m.name}  →  {str(e)[:60]}")
        time.sleep(2)
```

---

## Expected spreadsheet structure

The script translates **any** `.xlsx` file — it does not require a specific structure. It iterates over every sheet and every cell, translating all non-empty string values.

**Cells that are skipped (left unchanged):**

| Type | Example |
|---|---|
| Numbers | `42`, `3.14` |
| Dates | `2024-11-14` |
| Bare URLs | `https://example.com` |
| Empty / blank cells | — |
| Merged cells (non-master) | — |

The spreadsheet this script was originally built for (`Levantamento Dados Iniciativas IA PUC-Rio.xlsx`) contains 7 sheets:

| Sheet | Content |
|---|---|
| Educação Continuada | Continuing education courses (title, syllabus, professor, department, hours, dates) |
| Disciplinas de graduação e pós | Undergraduate and graduate subjects (code, type, name, syllabus, professor) |
| Palestras e eventos | Lectures and events (name, type, speaker, summary, date, source) |
| Professores | Faculty list (name, Lattes CV link, department) |
| Iniciativas acadêmicas | AI Bachelor's curriculum (code, credits, subject, professor, syllabus) |
| Iniciativas administrativas | Administrative AI initiatives (name, unit, usage, tools) |
| Documentos administrativos | Administrative documents (name, function) |

The script handles any similar tabular structure without modification.

---

## How it works

1. **Loads** the workbook with `openpyxl`
2. **Collects** all translatable string cells across every sheet
3. **Batches** them into groups of `--batch-size` and sends each batch to Gemini in a single API call using a structured JSON prompt
4. **Parses** the JSON array response and writes translations back into the workbook
5. **Saves** the translated workbook to the output path

The prompt instructs Gemini to preserve:
- Person names and place names
- Institution names: `PUC-Rio`, `ICA`, `CCEC`, `ECOA`, `BI MASTER`, `AcademIA`, etc.
- Course codes: `BIA1001`, `MEC2007`, `CIS2114`, etc.
- Technical acronyms: `LLM`, `RAG`, `NLP`, `MoE`, `MCP`, `XAI`, `IDP`, etc.
- Common English terms: `Machine Learning`, `Deep Learning`, `Bootcamp`, `Pipeline`, etc.
