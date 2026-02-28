# CLAUDE.md — Project Guide for Claude Code

## What is this project?
LemmaCheck — a PyQt6 desktop app for full-text search with lemmatization. Users load documents (DOCX, DOC, PDF, TXT, XLSX, XLS), the app indexes them by lemmas, and provides ranked phrase search with highlighted results. Supports Russian, English, and Kazakh languages. Includes KWIC concordance, export to CSV/XLSX, and a Barbie pink UI theme.

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
app.py              # Single-file application (~2100 lines): parsers, engine, threads, export, UI
instruction.md      # Original feature spec (reference only)
requirements.txt    # Python dependencies
docs/               # Russian documentation for articles
.github/            # GitHub config (copilot instructions)
```

## Architecture at a Glance
- **Document parsers** (top of file): `extract_text_from_docx`, `_pdf`, `_txt`, `_doc`, `_xlsx`, `_xls`
- **`KazakhStemmer`**: rule-based suffix stemmer for Kazakh (agglutinative language)
- **`LemmaSearchEngine`**: inverted index, lemma cache, TF-IDF scoring, phrase matching, KWIC concordance
- **`IndexingThread`**: QThread for background document indexing
- **`SearchThread`**: QThread for search + KWIC generation (emits `result_ready` and `kwic_ready` signals)
- **`ExportDialog`**: QDialog for choosing export type (results/concordance/summary) and format (CSV/XLSX)
- **`ExportThread`**: QThread for background file export with progress
- **`LemmaCheckApp(QMainWindow)`**: all UI setup and event handling
- **UI Theme**: Barbie pink palette — `QPalette` + global QSS stylesheet applied in `main()`

## Conventions
- Python 3.10+, type hints everywhere
- UI text in Russian
- All optional dependencies wrapped in `try/except ImportError`
- Platform-specific code guarded by `platform.system()`
- Single `app.py` file — do not split unless explicitly asked
- Barbie pink color scheme — all UI colors should stay within the pink/purple/plum palette

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
- KWIC: `get_kwic_concordance()` returns 5-tuples `(filename, left, keyword, right, position)`

### Modify UI
- All widgets are created in `LemmaCheckApp.init_ui()`
- Results rendering: `display_results()` and `_highlight_phrases_in_context()`
- KWIC table: `_populate_kwic_table()` — data comes from `SearchThread.kwic_ready` signal
- Found-words panel: right side of the results splitter
- Export: `export_results()` → `ExportDialog` → `ExportThread`
- Colors: global theme in `main()`, button-specific styles in `init_ui()`, text colors in `display_results()`

### Modify export
- `_compute_results_table()` — columns: Лемма, Словоформа, Документ, Абс.частота, %, IP10K, IPM, TF-IDF
- `_compute_concordance_table()` — columns: Левый контекст, Ключевое слово, Правый контекст, Документ, Позиция
- `_compute_summary()` — key-value pairs: date, query, language mode, corpus stats, frequency metrics

## Gotchas
- `morph` and `lemmatizer` can be `None` if pymorphy3/nltk aren't installed — always guard
- `kaz_stemmer` is always available (pure Python, no external deps)
- Kazakh is detected by Kazakh-specific Cyrillic chars (Ә, Ғ, Қ, Ң, Ө, Ұ, Ү, Һ, І); words with only standard Cyrillic go to Russian
- The inverted index uses `defaultdict(lambda: defaultdict(list))` — be careful when serializing to JSON
- Hyphenated words are indexed as both the full compound and individual parts
- `chardet` is used for TXT encoding detection with a fallback chain
- NLTK data is auto-downloaded on first run if missing
- KWIC data is stored in `self.last_kwic_data` — must be non-empty for concordance export
- Export file naming convention: `lemmaCheck_export_{type}_{YYYYMMDD_HHMM}.{csv|xlsx}`

## Testing
No test suite yet. To verify manually:
1. Run `python3 app.py` — verify Barbie pink theme renders correctly
2. Add a mix of DOCX/PDF/TXT files — indexing completes with progress bar
3. Search for Russian and English terms — results highlighted in pink
4. Search for Kazakh terms (e.g., "білім беру") — verify Kazakh stemmer works
5. Enable KWIC checkbox → search — KWIC table populates with Left|Keyword|Right|Document columns
6. Use context type combo (±5 words, ±10 words, Предложение) and context filter
7. Click Export → verify all 3 export types work (results table, concordance, summary)
8. Save index → restart → load index → search produces same results

## Do Not
- Remove graceful import fallbacks
- Block the main thread with heavy computation
- Hardcode file paths — use `QFileDialog` and `os.path`
- Change PyQt6 to another framework unless explicitly asked
- Break the Barbie pink theme — keep colors consistent
