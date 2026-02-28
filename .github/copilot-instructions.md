# Copilot Instructions for LemmaCheck

## Project Overview
LemmaCheck is a Python desktop application for full-text lemma-based search across documents (DOCX, DOC, PDF, TXT, XLSX, XLS). It uses PyQt6 for the GUI and supports trilingual lemmatization (Russian via pymorphy3, English via nltk/WordNetLemmatizer, Kazakh via built-in suffix stemmer).

## Architecture
- **Single-file app**: All code lives in `app.py` (~1300 lines)
- **GUI framework**: PyQt6 (not tkinter despite the instruction.md reference)
- **Key classes**:
  - `LemmaSearchEngine` — inverted index, TF-IDF ranking, phrase search
  - `IndexingThread(QThread)` — background document indexing
  - `SearchThread(QThread)` — background search execution
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
- **Index persistence**: JSON save/load for the inverted index

## Dependencies
See `requirements.txt`: pymorphy3, nltk, python-docx, PyMuPDF, PyQt6, chardet, openpyxl, xlrd

## When Making Changes
- Keep everything in `app.py` unless explicitly told to split
- Maintain graceful degradation — optional libs should not crash the app if missing
- Test with both Russian and English text
- Preserve the existing TF-IDF and phrase-search logic unless asked to change it
- Use `QApplication.processEvents()` sparingly — prefer `QThread` + signals
- Platform-specific code (Windows COM, macOS textutil) must be guarded by `platform.system()` checks
