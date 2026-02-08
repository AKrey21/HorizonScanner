# HorizonScanner
Horizon Scanner

## Daily raw article ingestion
Raw article ingestion runs on a time-based trigger in Apps Script. There are two ways to ensure it runs daily:

1. Open the spreadsheet and use the **HorizonScanner → Install daily raw ingest trigger** menu item.
2. Keep the spreadsheet opened periodically; on open it attempts to ensure the daily trigger exists.

Note: Apps Script requires authorization the first time you install triggers. Use the menu item above to grant access once.

You can also use **HorizonScanner → Run daily raw ingest now** or the **Raw Articles → Run daily ingest now** button to test a run on demand.
Use **HorizonScanner → Check raw ingest trigger status** to verify the trigger is installed and see the last run time, last status, and any recent errors.

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
