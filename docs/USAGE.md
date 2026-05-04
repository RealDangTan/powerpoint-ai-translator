# Usage Guide

## Desktop UI

Start the application:

```bash
python3 translation_tool_ui.py
```

Recommended workflow:

1. Choose a `.pptx` input file.
2. Choose the output path.
3. Confirm provider, model, source language, and target language.
4. Review glossary terms and prompts.
5. Run translation and inspect the generated deck.

## Terminal

Basic translation:

```bash
python3 translation_tool_terminal.py --input input.pptx --output output.pptx
```

Vietnamese to English:

```bash
python3 translation_tool_terminal.py \
  --input input.pptx \
  --output output.pptx \
  --source-language VI \
  --target-language EN
```

Format-intact mode:

```bash
python3 translation_tool_terminal.py \
  --input input.pptx \
  --output output.pptx \
  --formatting-mode format_intact \
  --keep-intermediate
```

## Custom Provider Template

For custom APIs, `request_template` can describe the outbound JSON payload. The template supports these variables:

- `$prompt` as a JSON-encoded prompt string
- `$prompt_text` as the raw prompt text
- `$model` as the configured model name
- `$api_key` as the configured API key

The response body is parsed using `response_text_path`, for example `choices.0.message.content`.
