# AI-based Review of Word Documents

A Python package to review a `.docx` Word document using LLMs to check grammar and spelling.
All the suggestions are added as comments to the document, which is then saved with the suffix `_reviewed`.

## Requirements

- Python 3.10 or higher
- This project uses `uv` for development dependency management

## Usage

This tool uses the [`LiteLLM`](https://github.com/BerriAI/litellm) library to interact with a wide range of LLM providers. You can configure the model, API key, and base URL to connect to different services.

### Arguments

- `document_path`: The path to the `.docx` document to review.
- `--model`: The model name to use for the review (e.g., `ollama/gemma3:12b`, `openrouter/openai/gpt-4o`).
- `--api-key`: Your API key for the LLM provider. Can also be set with the `LITELLM_API_KEY` environment variable.
- `--base-url`: The base URL for the LLM provider's API. Can also be set with the `LITELLM_BASE_URL` environment variable.
- `--context`: Optional context or instructions to add to the LLM prompt.
- `--cache-location`: The path to the directory where to store the cache of model responses.
- `--verbose`: Enable verbose DEBUG level logging to console.

### Examples

#### Using a local Ollama model

To use a model served by a local [Ollama](https://ollama.com/) instance:
```bash
uv run ai-review-docx data/example.docx --model ollama/gpt-oss:20b
```

To also set the url of the Ollama server:
```bash
uv run ai-review-docx data/example.docx --model ollama/gemma3:12b --base-url "http://localhost:11434"
```

#### Using OpenRouter

You can use a model from [OpenRouter](https://openrouter.ai/) by providing your API key.

As an environment variable:
```bash
export LITELLM_API_KEY="your_openrouter_api_key"
uv run ai-review-docx data/example.docx --model openrouter/openai/gpt-4o
```

Or as a command-line argument:
```bash
uv run ai-review-docx data/example.docx --model openrouter/openai/gpt-4o --api-key "your_openrouter_api_key"
```
