# AGENTS.md — Instructions for AI Coding Agents

## Project Context
**LemmaCheck** is a single-file Python desktop app (`app.py`, ~1400 lines) for full-text lemma-based document search. GUI: PyQt6. Trilingual lemmatization: pymorphy3 (Russian), nltk (English), built-in suffix stemmer (Kazakh).

## File Layout
| File | Purpose |
|---|---|
| `app.py` | Entire application: parsers, search engine, threads, UI |
| `requirements.txt` | pip dependencies |
| `instruction.md` | Original specification (reference only, may be outdated) |

## Setup
```bash
source venv/bin/activate   # venv already exists
pip install -r requirements.txt
python3 app.py
```

## Code Map (app.py)
| Lines (approx.) | Section |
|---|---|
| 1–85 | Imports and optional dependency guards |
| 87–160 | `KazakhStemmer` class and Kazakh character constants |
| 162–320 | Document parser functions |
| 322–750 | `LemmaSearchEngine` class |
| 752–790 | `IndexingThread(QThread)` |
| 792–815 | `SearchThread(QThread)` |
| 817–1370 | `LemmaCheckApp(QMainWindow)` — UI and event handlers |
| 1372–1400 | `main()` entry point |

## Rules for Agents

### General
- Keep all code in `app.py` unless explicitly asked to split
- Use Python 3.10+ features and type hints
- UI strings must be in Russian
- Always wrap optional imports in `try/except ImportError` with a `None` fallback
- Guard platform-specific code with `platform.system()` checks

### When editing search logic
- `LemmaSearchEngine` is the core class — understand its inverted index before modifying
- Language detection priority: Kazakh-specific chars → KazakhStemmer, other Cyrillic → pymorphy3 (Russian), Latin → nltk (English)
- `KazakhStemmer` is a built-in suffix stemmer with no external dependencies
- Phrase search uses positional matching (`_find_phrase_in_document`) — don't break sequential word ordering
- The lemma cache (`_lemma_cache`) improves performance — keep it

### When editing UI
- Never block the main thread — use `QThread` + `pyqtSignal`
- Results are rendered via `QTextCursor` with `QTextCharFormat` — not plain HTML
- The found-words panel is on the right side of a `QSplitter`

### When adding features
1. Add parser functions near the other parsers (top of file)
2. Register new formats in `extract_text()` and the file dialog filter
3. New dependencies: add to `requirements.txt`, import with try/except guard
4. New UI controls: add in `init_ui()`, keep the existing layout structure

### Quality checks before finishing
- Verify no bare `except:` clauses (use specific exceptions)
- Confirm all new functions have type hints
- Ensure graceful degradation when optional deps are missing
- Check that `save_index` / `load_index` still work if you changed data structures

## Testing
No automated tests. Manual verification:
1. `python3 app.py` — app launches without errors
2. Add files (DOCX, PDF, TXT) — indexing completes with progress bar
3. Search Russian phrase (e.g., "учебный план") — results are highlighted
4. Search English phrase (e.g., "machine learning") — results are highlighted
5. Search Kazakh phrase (e.g., "білім беру") — results are highlighted via suffix stemmer
6. Save index → restart → load index → search produces same results
