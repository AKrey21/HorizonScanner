# HorizonScanner
Horizon Scanner

## AI provider toggle (Gemini/OpenAI)
Set these Script Properties (Apps Script → Project Settings → Script Properties) to switch providers and control model usage:

| Property | Purpose | Example |
| --- | --- | --- |
| `GEMINI_API_KEY` | Gemini API key | `AIza...` |
| `OPENAI_API_KEY` | OpenAI API key | `sk-...` |
| `AI_PROVIDER` | Active provider (`gemini` or `openai`) | `openai` |
| `AI_MODEL` | Optional shared model override | `gpt-4o-mini` |
| `GEMINI_MODEL` | Optional Gemini-specific override | `gemini-1.5-flash` |
| `OPENAI_MODEL` | Optional OpenAI-specific override | `gpt-4o-mini` |
| `AI_MAX_OUTPUT_TOKENS` | Optional global cap for max output tokens | `600` |

Defaults are conservative (`gemini-1.5-flash` / `gpt-4o-mini`) with a 600-token output cap unless overridden. The toggle is provider-aware, so you can switch by changing `AI_PROVIDER` without code edits.
