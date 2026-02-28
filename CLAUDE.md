# CLAUDE.md — Project Guide for Claude Code

## What is this project?
LemmaCheck — a PyQt6 desktop app for full-text search with lemmatization. Users load documents (DOCX, DOC, PDF, TXT, XLSX, XLS), the app indexes them by lemmas, and provides ranked phrase search with highlighted results. Supports Russian, English, and Kazakh languages.

## Quick Start
```bash
cd /Users/narimanatakanov/Documents/projects/lemmacheck
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 -c "import nltk; nltk.download('wordnet', quiet=True); nltk.download('averaged_perceptron_tagger', quiet=True)"
python3 app.py
```

## Project Structure
```
app.py              # Single-file application (~1300 lines): parsers, engine, threads, UI
instruction.md      # Original feature spec (reference only)
requirements.txt    # Python dependencies
.github/            # GitHub config (workflows, copilot instructions)
```

## Architecture at a Glance
- **Document parsers** (top of file): `extract_text_from_docx`, `_pdf`, `_txt`, `_doc`, `_xlsx`, `_xls`
- **`LemmaSearchEngine`**: inverted index, lemma cache, TF-IDF scoring, phrase matching
- **`KazakhStemmer`**: rule-based suffix stemmer for Kazakh (agglutinative language)
- **`IndexingThread` / `SearchThread`**: QThread subclasses for non-blocking operations
- **`LemmaCheckApp(QMainWindow)`**: all UI setup and event handling

## Conventions
- Python 3.10+, type hints everywhere
- UI text in Russian
- All optional dependencies wrapped in `try/except ImportError`
- Platform-specific code guarded by `platform.system()`
- Single `app.py` file — do not split unless explicitly asked

## Common Tasks

### Add a new document format
1. Add a `extract_text_from_<ext>()` function in the parsers section
2. Register it in the `extract_text()` dispatcher
3. Add the file extension to the `QFileDialog` filter in `add_files()`
4. Add any new dependency to `requirements.txt` with a try/except import

### Modify search behavior
- Lemmatization: `_lemmatize_word()` and `lemmatize()` in `LemmaSearchEngine`
- Ranking: `_calculate_tfidf()` and `search()`
- Phrase matching: `_find_phrase_in_document()` with `max_distance` param

### Modify UI
- All widgets are created in `LemmaCheckApp.init_ui()`
- Results rendering: `display_results()` and `_highlight_phrases_in_context()`
- Found-words panel: right side of the results splitter

## Gotchas
- `morph` and `lemmatizer` can be `None` if pymorphy3/nltk aren't installed — always guard
- `kaz_stemmer` is always available (pure Python, no external deps)
- Kazakh is detected by Kazakh-specific Cyrillic chars (Ә, Ғ, Қ, Ң, Ө, Ұ, Ү, Һ, І); words with only standard Cyrillic go to Russian
- The inverted index uses `defaultdict(lambda: defaultdict(list))` — be careful when serializing to JSON
- Hyphenated words are indexed as both the full compound and individual parts
- `chardet` is used for TXT encoding detection with a fallback chain
- NLTK data is auto-downloaded on first run if missing

## Testing
No test suite yet. To verify manually:
1. Run `python3 app.py`
2. Add a mix of DOCX/PDF/TXT files
3. Search for Russian and English terms
4. Search for Kazakh terms (e.g., "білім беру") — verify Kazakh stemmer works
5. Verify phrase search returns correct sequential matches
6. Check save/load index round-trips correctly

## Do Not
- Remove graceful import fallbacks
- Block the main thread with heavy computation
- Hardcode file paths — use `QFileDialog` and `os.path`
- Change PyQt6 to another framework unless explicitly asked
