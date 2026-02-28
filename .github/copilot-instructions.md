# Copilot Instructions for LemmaCheck

## Project Overview
LemmaCheck is a Python desktop application for full-text lemma-based search across documents (DOCX, DOC, PDF, TXT, XLSX, XLS). It uses PyQt6 for the GUI and supports trilingual lemmatization (Russian via pymorphy3, English via nltk/WordNetLemmatizer, Kazakh via built-in suffix stemmer). Features KWIC concordance, data export (CSV/XLSX), and a Barbie pink UI theme.

## Architecture
- **Single-file app**: All code lives in `app.py` (~2100 lines)
- **GUI framework**: PyQt6 with Fusion style + Barbie pink QPalette & QSS theme
- **Key classes**:
  - `LemmaSearchEngine` — inverted index, TF-IDF ranking, phrase search, KWIC concordance
  - `KazakhStemmer` — rule-based suffix stemmer for Kazakh
  - `IndexingThread(QThread)` — background document indexing
  - `SearchThread(QThread)` — background search + KWIC concordance generation
  - `ExportDialog(QDialog)` — export type/format selection dialog
  - `ExportThread(QThread)` — background CSV/XLSX export
  - `LemmaCheckApp(QMainWindow)` — main window and UI logic
- **Document parsers**: standalone functions (`extract_text_from_*`)

## Coding Conventions
- Language: Python 3.10+
- Type hints on all function signatures (`Dict`, `List`, `Tuple`, `Set`, `Optional` from `typing`)
- Docstrings in Russian or English (mixed is acceptable)
- UI strings and labels in Russian
- Use `defaultdict` for inverted index structures
- Graceful import fallbacks with `try/except ImportError` for optional dependencies
- Error messages shown via `QMessageBox`

## Key Patterns
- **Language detection**: Kazakh-specific Cyrillic chars (Ә, Ғ, Қ, Ң, Ө, Ұ, Ү, Һ, І) → KazakhStemmer, other Cyrillic → pymorphy3, Latin → nltk
- **Hyphenated words**: Index both the full compound and individual parts
- **Phrase search**: Sequential lemma matching with `max_distance` parameter
- **Threading**: `QThread` with `pyqtSignal` for progress/results — never block the UI thread
- **KWIC concordance**: Built in `SearchThread`, emits `kwic_ready` signal with `(filename, left, keyword, right, position)` tuples
- **Export**: `ExportDialog` → `ExportThread` pipeline; 3 types (results table, concordance, summary report); CSV/XLSX formats
- **Index persistence**: JSON save/load for the inverted index
- **UI theme**: Barbie pink palette via `QPalette` + global `QSS` stylesheet in `main()`

## Dependencies
See `requirements.txt`: pymorphy3, nltk, python-docx, PyMuPDF, PyQt6, chardet, openpyxl, xlrd

## When Making Changes
- Keep everything in `app.py` unless explicitly told to split
- Maintain graceful degradation — optional libs should not crash the app if missing
- Test with both Russian and English text
- Preserve the existing TF-IDF and phrase-search logic unless asked to change it
- Use `QApplication.processEvents()` sparingly — prefer `QThread` + signals
- Platform-specific code (Windows COM, macOS textutil) must be guarded by `platform.system()` checks
- Keep the Barbie pink color palette consistent — all new UI elements should use the pink/purple/plum color scheme
