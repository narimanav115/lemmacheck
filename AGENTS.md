# AGENTS.md — Instructions for AI Coding Agents

## Project Context
**LemmaCheck** is a single-file Python desktop app (`app.py`, ~2100 lines) for full-text lemma-based document search. GUI: PyQt6 with Barbie pink theme. Trilingual lemmatization: pymorphy3 (Russian), nltk (English), built-in suffix stemmer (Kazakh). Features KWIC concordance and export to CSV/XLSX.

## File Layout
| File | Purpose |
|---|---|
| `app.py` | Entire application: parsers, search engine, threads, export, UI |
| `requirements.txt` | pip dependencies |
| `instruction.md` | Original specification (reference only, may be outdated) |
| `docs/documentation_ru.md` | Russian-language documentation for articles |

## Setup
```bash
source venv/bin/activate   # venv already exists
pip install -r requirements.txt
python3 app.py
```

## Code Map (app.py)
| Lines (approx.) | Section |
|---|---|
| 1–92 | Imports and optional dependency guards |
| 93–172 | `KazakhStemmer` class and Kazakh character constants |
| 174–320 | Document parser functions |
| 322–830 | `LemmaSearchEngine` class (incl. KWIC concordance, save/load) |
| 832–860 | `IndexingThread(QThread)` |
| 863–920 | `SearchThread(QThread)` — search + KWIC generation |
| 922–996 | `ExportDialog(QDialog)` — export type/format selection |
| 998–1050 | `ExportThread(QThread)` — background CSV/XLSX export |
| 1052–1965 | `LemmaCheckApp(QMainWindow)` — UI, event handlers, export methods |
| 1968–2126 | `main()` — Barbie pink QPalette + global QSS stylesheet |

## Rules for Agents

### General
- Keep all code in `app.py` unless explicitly asked to split
- Use Python 3.10+ features and type hints
- UI strings must be in Russian
- Always wrap optional imports in `try/except ImportError` with a `None` fallback
- Guard platform-specific code with `platform.system()` checks
- Maintain the Barbie pink color theme (pink/purple/plum palette)

### When editing search logic
- `LemmaSearchEngine` is the core class — understand its inverted index before modifying
- Language detection priority: Kazakh-specific chars → KazakhStemmer, other Cyrillic → pymorphy3 (Russian), Latin → nltk (English)
- `KazakhStemmer` is a built-in suffix stemmer with no external dependencies
- Phrase search uses positional matching (`_find_phrase_in_document`) — don't break sequential word ordering
- The lemma cache (`_lemma_cache`) improves performance — keep it
- KWIC concordance: `get_kwic_concordance()` returns 5-tuples `(filename, left, keyword, right, position)`

### When editing UI
- Never block the main thread — use `QThread` + `pyqtSignal`
- Results are rendered via `QTextCursor` with `QTextCharFormat` — not plain HTML
- The found-words panel is on the right side of a `QSplitter`
- KWIC table columns: Левый контекст | Ключевое слово | Правый контекст | Документ
- KWIC data comes from `SearchThread.kwic_ready` signal → `_populate_kwic_table()`
- New buttons should use the pink/purple styles from `init_ui()`

### When editing export
- 3 export types: results table (0), concordance (1), summary report (2)
- `_compute_results_table()`, `_compute_concordance_table()`, `_compute_summary()`
- `ExportThread` handles CSV (via `csv` module) and XLSX (via `openpyxl`)
- File naming: `lemmaCheck_export_{type}_{YYYYMMDD_HHMM}.{ext}`
- Data stored in `self.last_results`, `self.last_query`, `self.last_kwic_data`

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
- Verify the Barbie pink theme is not broken by new elements

## Testing
No automated tests. Manual verification:
1. `python3 app.py` — app launches with Barbie pink theme, no errors
2. Add files (DOCX, PDF, TXT) — indexing completes with pink progress bar
3. Search Russian phrase (e.g., "учебный план") — results highlighted in pink
4. Search English phrase (e.g., "machine learning") — results highlighted
5. Search Kazakh phrase (e.g., "білім беру") — results highlighted via suffix stemmer
6. Enable KWIC → search → KWIC table shows Left|Keyword|Right|Document
7. Test all 3 context modes: ±5 words, ±10 words, Предложение
8. Export → Results Table → CSV — file created with correct data
9. Export → Concordance → XLSX — file created with KWIC data
10. Export → Summary → CSV — file has corpus statistics
11. Save index → restart → load index → search produces same results
