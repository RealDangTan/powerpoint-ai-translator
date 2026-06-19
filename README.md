# 🌐 PPTX Translator

A Python script that translates PowerPoint (`.pptx`) presentations between **Vietnamese ↔ English** using any OpenAI-compatible LLM API. Formatting, fonts, and styles are fully preserved in the output file.

---

## ✨ Features

- **Bidirectional translation** — Vietnamese → English or English → Vietnamese
- **Formatting preservation** — fonts, sizes, colors, and styles are kept intact
- **Smart batching** — all text on a slide is translated in a single API call
- **Global context awareness** — extracts topic/industry/tone before translating for consistent terminology
- **Domain glossary** — built-in banking & CX glossary ensures correct industry-specific terms
- **Translation cache** — SHA-256 keyed cache avoids redundant API calls across runs
- **Graceful fallback** — on any API failure, original text is used so the output is never broken
- **Retry with backoff** — 3 automatic retries with exponential backoff on transient errors
- **Group & table support** — handles nested grouped shapes and table cells

---

## 📋 Requirements

- Python 3.8+
- An OpenAI-compatible LLM API endpoint and key

Install dependencies:

```bash
pip install openai python-pptx httpx
```

---

## 🚀 Quick Start

1. **Clone the repo**

   ```bash
   git clone https://github.com/your-username/pptx-translator.git
   cd pptx-translator
   ```

2. **Configure the script** — open `Translator.py` and set these variables at the top:

   ```python
   ORIMISE_BASE_URL = "https://your-api-endpoint.com"
   ORIMISE_API_KEY  = "your-api-key-here"
   MODEL            = "claude-sonnet-4-6"          # or any model your endpoint supports
   TRANSLATION_DIRECTION = "VI_TO_EN"              # "VI_TO_EN" or "EN_TO_VI"
   ```

3. **Place your file** — copy your PowerPoint file into the project directory and set:

   ```python
   INPUT_FILE  = "input.pptx"
   OUTPUT_FILE = "output.pptx"
   ```

4. **Run**

   ```bash
   python Translator.py
   ```

   The translated file is saved to `output.pptx`. A cache file (`translation_cache_v2.json`) is created automatically to speed up future runs on the same content.

---

## ⚙️ Configuration Reference

All configuration lives at the top of `Translator.py` — no CLI arguments or config files needed.

| Variable | Default | Description |
|---|---|---|
| `ORIMISE_BASE_URL` | *(required)* | OpenAI-compatible API base URL |
| `ORIMISE_API_KEY` | *(required)* | API key for authentication |
| `MODEL` | `claude-sonnet-4-6` | Model name to use for translation |
| `TRANSLATION_DIRECTION` | `VI_TO_EN` | `"VI_TO_EN"` or `"EN_TO_VI"` |
| `INPUT_FILE` | `input.pptx` | Path to the source PowerPoint file |
| `OUTPUT_FILE` | `output.pptx` | Path for the translated output file |
| `CACHE_FILE` | `translation_cache_v2.json` | Path to the JSON translation cache |
| `SHOW_USAGE` | `False` | Set `True` to call `/v1/usage` on startup |
| `SHOW_GEMINI` | `False` | Set `True` to run a live LLM smoke-test on startup |

---

## 🏗️ How It Works

Translation runs in two sequential steps:

```
┌─────────────────────────────────────────────────────────────┐
│  Step 1 — Global Context Extraction  (1 API call)           │
│  Slide titles + first-slide content → Topic / Industry / Tone│
└───────────────────────────┬─────────────────────────────────┘
                            │ context injected into every prompt
┌───────────────────────────▼─────────────────────────────────┐
│  Step 2 — Per-Slide Batch Translation  (1 API call / slide) │
│  Paragraphs tagged T0…Tn → translated → written back to PPTX│
└─────────────────────────────────────────────────────────────┘
```

**Step 1** sends all slide titles and the first slide's content to the LLM to extract a 20-word context summary (`Topic / Industry / Tone`). This grounds all subsequent translation prompts.

**Step 2** iterates slide by slide. Every text paragraph is assigned a sequential ID (`T0:`, `T1:`, …), sent as a single batched prompt, and parsed back by ID. The translation is written only into `runs[0]` of each paragraph while all other runs are cleared — this is how formatting is preserved without re-applying any styles.

A **partial glossary** is built per slide by scanning the slide text for matching entries in `MASTER_GLOSSARY`, keeping prompts short.

---

## 📦 Caching

Translations are cached in `translation_cache_v2.json` with a key derived from:

```
SHA-256( original_text + global_context + model + translation_direction )
```

On subsequent runs, cached paragraphs are skipped entirely, saving both time and API cost. The cache is saved after every slide.

---

## 📖 Domain Glossary

`MASTER_GLOSSARY` contains hardcoded mappings for banking and customer-experience terminology (e.g. *customer journey*, *touchpoint*, *omnichannel*, *VoC*). The direction is automatically flipped based on `TRANSLATION_DIRECTION`.

To add your own terms, extend the dictionary at the top of the file:

```python
MASTER_GLOSSARY = {
    "your english term": "thuật ngữ tiếng Việt",
    ...
}
```

---

## 📊 Sample Output

```
============================================================
🚀 ADVANCED TRANSLATOR — input.pptx
   Model : claude-sonnet-4-6
============================================================
📦 Cache loaded: 142 entries.

🔍 Step 1: Extracting global context...
   ✅ Global Context (cached): Topic: Banking CX, Industry: Finance, Tone: Formal

📝 Step 2: Translating 18 slides (batch mode)...

   📄 Slide 1/18
      📋 6 text items
      ✅ 6 paragraphs from cache

   📄 Slide 2/18
      📋 11 text items
      🌐 11 paragraphs translated (1 API call)
   ...

============================================================
✨ DONE! → output.pptx
============================================================
📊 PERFORMANCE:
   Total operations : 94
   ✅ Cache hits    : 78 (83.0%)
   🌐 API calls     : 17
   💾 Cache entries  : 219
   🪙 Tokens saved  : ~8,400 (est.)
============================================================
```

---

## 🤝 Contributing

Contributions are welcome. Feel free to open an issue or submit a pull request for:

- Additional language pair support
- CLI argument parsing
- A `requirements.txt` or `pyproject.toml`
- Extended glossary entries

---

## 📄 License

MIT License — see [LICENSE](LICENSE) for details.
