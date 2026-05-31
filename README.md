# Excel PT → EN Translator

Translates an Excel (`.xlsx`) file from **Portuguese to English** using the **OpenCode Zen API** — a curated gateway that gives access to 30+ models (Claude, GPT, Gemini, Qwen…) with a single API key.

- Preserves proper nouns, institution names, course codes, URLs, and English technical terms
- Handles multiple sheets automatically
- Retries automatically on rate-limit errors with the delay suggested by the API
- Real-time progress bar with elapsed time and ETA
- Output filename automatically includes the date and model used

---

## Getting an OpenCode Zen API Key

1. Go to **[opencode.ai](https://opencode.ai)**
2. Sign in and go to your account settings
3. Click **"Create API key"** and copy it

> OpenCode Zen is pay-as-you-go. Translation of a ~2000-cell spreadsheet costs a few cents depending on the model used.

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Usage

```bash
# Pass the key directly
python translate_excel.py input.xlsx output.xlsx --api-key YOUR_KEY

# Or export it as an environment variable (recommended)
export OPENCODE_ZEN_API_KEY=YOUR_KEY
python translate_excel.py input.xlsx output.xlsx
```

The output filename automatically includes the date and model:

```
output_2026-05-31_opencode-claude-sonnet-4-5.xlsx
```

### Options

| Option | Default | Description |
|---|---|---|
| `--api-key` | `OPENCODE_ZEN_API_KEY` env var | Your OpenCode Zen API key |
| `--model` | `opencode/claude-sonnet-4-5` | Model to use (OpenCode Zen format) |
| `--batch-size` | `20` | Number of cells sent per API call |
| `--list-models` | — | Print all available models and exit |

### Listing available models

```bash
python translate_excel.py dummy.xlsx dummy_out.xlsx --api-key YOUR_KEY --list-models
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

---

## How it works

1. **Loads** the workbook with `openpyxl`
2. **Collects** all translatable string cells across every sheet
3. **Batches** them into groups of `--batch-size` and sends each batch to OpenCode Zen in a single API call using a structured JSON prompt
4. **Parses** the JSON array response and writes translations back into the workbook
5. **Saves** the translated workbook with `_<date>_<model>` appended to the filename

The prompt instructs the model to preserve:
- Person names and place names
- Institution names: `PUC-Rio`, `ICA`, `CCEC`, `ECOA`, `BI MASTER`, `AcademIA`, etc.
- Course codes: `BIA1001`, `MEC2007`, `CIS2114`, etc.
- Technical acronyms: `LLM`, `RAG`, `NLP`, `MoE`, `MCP`, `XAI`, `IDP`, etc.
- Common English terms: `Machine Learning`, `Deep Learning`, `Bootcamp`, `Pipeline`, etc.
