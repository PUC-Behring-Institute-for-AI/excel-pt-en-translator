# Excel PT → EN Translator

Translates an Excel (`.xlsx`) file from **Portuguese to English** using your choice of AI provider.

- Preserves proper nouns, institution names, course codes, URLs, and English technical terms
- Handles multiple sheets automatically
- Retries automatically on rate-limit errors
- Real-time progress bar with elapsed time and ETA
- Output filename automatically includes the date and model used

---

## Providers

| Provider | Flag | Env var | Default model | Cost |
|---|---|---|---|---|
| **OpenCode Zen** (default) | `--provider opencode-zen` | `OPENCODE_ZEN_API_KEY` | `opencode/big-pickle` | **Free** |
| **Google Gemini** | `--provider gemini` | `GEMINI_API_KEY` | `gemini-2.5-flash` | ~$0.02 per spreadsheet |

---

## Getting an API Key

### OpenCode Zen (free)
1. Go to **[opencode.ai](https://opencode.ai)**
2. Sign in → account settings → **"Create API key"**

Free models available in OpenCode Zen:
- `opencode/big-pickle` ← default (best free performance)
- `opencode/minimax-m2.5-free`
- `opencode/mimo-v2-pro-free`
- `opencode/mimo-v2-omni-free`
- `opencode/nemotron-3-super-free`

### Google Gemini
1. Go to **[aistudio.google.com](https://aistudio.google.com)**
2. Sign in → **"Get API key"** → **"Create API key"**

Tested working model: `gemini-2.5-flash`

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Usage

```bash
# OpenCode Zen (free, default)
export OPENCODE_ZEN_API_KEY=YOUR_KEY
python translate_excel.py input.xlsx output.xlsx

# Google Gemini
export GEMINI_API_KEY=YOUR_KEY
python translate_excel.py input.xlsx output.xlsx --provider gemini

# Pass key directly
python translate_excel.py input.xlsx output.xlsx --api-key YOUR_KEY

# Use a specific model
python translate_excel.py input.xlsx output.xlsx --model opencode/mimo-v2-pro-free

# List available models for a provider
python translate_excel.py dummy.xlsx dummy.xlsx --list-models
python translate_excel.py dummy.xlsx dummy.xlsx --provider gemini --list-models
```

The output filename automatically includes the date and model:
```
output_2026-05-31_opencode-big-pickle.xlsx
```

### All options

| Option | Default | Description |
|---|---|---|
| `--provider` | `opencode-zen` | `opencode-zen` or `gemini` |
| `--api-key` | env var | API key for the selected provider |
| `--model` | provider default | Override the default model |
| `--batch-size` | `20` | Number of cells per API call |
| `--list-models` | — | Print available models and exit |

---

## Expected spreadsheet structure

The script translates **any** `.xlsx` file — no specific structure required. It iterates over every sheet and every cell, translating all non-empty string values.

**Cells skipped (left unchanged):**

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
3. **Batches** them and sends each batch to the selected provider in one API call using a structured JSON prompt
4. **Parses** the JSON array response and writes translations back into the workbook
5. **Saves** the translated workbook with `_<date>_<model>` appended to the filename

The prompt instructs the model to preserve:
- Person names and place names
- Institution names: `PUC-Rio`, `ICA`, `CCEC`, `ECOA`, `BI MASTER`, `AcademIA`, etc.
- Course codes: `BIA1001`, `MEC2007`, `CIS2114`, etc.
- Technical acronyms: `LLM`, `RAG`, `NLP`, `MoE`, `MCP`, `XAI`, `IDP`, etc.
- Common English terms: `Machine Learning`, `Deep Learning`, `Bootcamp`, `Pipeline`, etc.
