#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LemmaCheck - Full-text search application with lemmatization support
Supports Russian (pymorphy3) and English (nltk) lemmatization
Uses PyQt6 for GUI
"""

import sys
import os
import re
import json
import math
import subprocess
import platform
from collections import defaultdict
from typing import Dict, List, Tuple, Set, Optional
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLineEdit, QListWidget, QTextEdit, QLabel,
    QFileDialog, QMessageBox, QProgressBar, QGroupBox, QSplitter,
    QAbstractItemView, QListWidgetItem, QFrame
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QTextCharFormat, QColor, QTextCursor, QBrush

import chardet

# Document parsers
try:
    from docx import Document
except ImportError:
    Document = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

# Excel support
try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import xlrd
except ImportError:
    xlrd = None

# Windows COM support for .doc files
win32com_client = None
if platform.system() == 'Windows':
    try:
        import win32com.client as win32com_client
    except ImportError:
        pass

# Lemmatization
try:
    import pymorphy3
    morph = pymorphy3.MorphAnalyzer()
except ImportError:
    morph = None

try:
    import nltk
    from nltk.stem import WordNetLemmatizer
    from nltk.corpus import wordnet
    try:
        nltk.data.find('corpora/wordnet')
    except LookupError:
        nltk.download('wordnet', quiet=True)
    try:
        nltk.data.find('taggers/averaged_perceptron_tagger')
    except LookupError:
        nltk.download('averaged_perceptron_tagger', quiet=True)
    try:
        nltk.data.find('taggers/averaged_perceptron_tagger_eng')
    except LookupError:
        nltk.download('averaged_perceptron_tagger_eng', quiet=True)
    lemmatizer = WordNetLemmatizer()
except ImportError:
    lemmatizer = None


# ==================== Document Parsers ====================

def extract_text_from_docx(filepath: str) -> str:
    if Document is None:
        raise ImportError("python-docx не установлен")
    doc = Document(filepath)
    paragraphs = [para.text for para in doc.paragraphs if para.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    paragraphs.append(cell.text)
    return '\n'.join(paragraphs)


def extract_text_from_pdf(filepath: str) -> str:
    if fitz is None:
        raise ImportError("PyMuPDF не установлен")
    text_parts = []
    with fitz.open(filepath) as doc:
        for page in doc:
            text_parts.append(page.get_text())
    return '\n'.join(text_parts)


def extract_text_from_txt(filepath: str) -> str:
    with open(filepath, 'rb') as f:
        raw_data = f.read()
    detected = chardet.detect(raw_data)
    encoding = detected.get('encoding', 'utf-8')
    for enc in [encoding, 'utf-8', 'cp1251', 'cp1252', 'latin-1']:
        if enc:
            try:
                return raw_data.decode(enc)
            except (UnicodeDecodeError, LookupError):
                continue
    return raw_data.decode('utf-8', errors='ignore')


def extract_text(filepath: str) -> str:
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.docx':
        return extract_text_from_docx(filepath)
    elif ext == '.doc':
        return extract_text_from_doc(filepath)
    elif ext == '.pdf':
        return extract_text_from_pdf(filepath)
    elif ext == '.txt':
        return extract_text_from_txt(filepath)
    elif ext == '.xlsx':
        return extract_text_from_xlsx(filepath)
    elif ext == '.xls':
        return extract_text_from_xls(filepath)
    raise ValueError(f"Неподдерживаемый формат: {ext}")


def extract_text_from_doc(filepath: str) -> str:
    """Extract text from old .doc format using platform-specific methods"""
    import subprocess
    import tempfile
    
    system = platform.system()
    
    # Windows: use pywin32 COM automation
    if system == 'Windows':
        if win32com_client is not None:
            try:
                word = win32com_client.Dispatch('Word.Application')
                word.Visible = False
                doc = word.Documents.Open(os.path.abspath(filepath))
                text = doc.Content.Text
                doc.Close(False)
                word.Quit()
                return text
            except Exception:
                pass
    
    # macOS: use textutil (built-in)
    if system == 'Darwin':
        try:
            with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
                tmp_path = tmp.name
            subprocess.run(['textutil', '-convert', 'txt', '-output', tmp_path, filepath], 
                          check=True, capture_output=True)
            with open(tmp_path, 'r', encoding='utf-8') as f:
                text = f.read()
            os.unlink(tmp_path)
            return text
        except Exception:
            pass
    
    # Cross-platform fallback: antiword
    try:
        result = subprocess.run(['antiword', filepath], capture_output=True, text=True, check=True)
        return result.stdout
    except Exception:
        pass
    
    # Platform-specific error messages
    if system == 'Windows':
        raise ImportError("Не удалось прочитать .doc файл. Установите pywin32: pip install pywin32")
    elif system == 'Darwin':
        raise ImportError("Не удалось прочитать .doc файл с помощью textutil.")
    else:
        raise ImportError("Не удалось прочитать .doc файл. Установите antiword.")


def extract_text_from_xlsx(filepath: str) -> str:
    """Extract text from .xlsx files"""
    if openpyxl is None:
        raise ImportError("openpyxl не установлен. Установите: pip install openpyxl")
    
    wb = openpyxl.load_workbook(filepath, data_only=True)
    texts = []
    
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            row_texts = []
            for cell in row:
                if cell.value is not None:
                    row_texts.append(str(cell.value))
            if row_texts:
                texts.append(' '.join(row_texts))
    
    return '\n'.join(texts)


def extract_text_from_xls(filepath: str) -> str:
    """Extract text from old .xls files"""
    if xlrd is None:
        raise ImportError("xlrd не установлен. Установите: pip install xlrd")
    
    wb = xlrd.open_workbook(filepath)
    texts = []
    
    for sheet in wb.sheets():
        for row_idx in range(sheet.nrows):
            row_texts = []
            for col_idx in range(sheet.ncols):
                cell_value = sheet.cell_value(row_idx, col_idx)
                if cell_value:
                    row_texts.append(str(cell_value))
            if row_texts:
                texts.append(' '.join(row_texts))
    
    return '\n'.join(texts)


# ==================== Search Engine ====================

class LemmaSearchEngine:
    def __init__(self):
        self.documents: Dict[str, dict] = {}
        self.inverted_index: Dict[str, Dict[str, List[int]]] = defaultdict(lambda: defaultdict(list))
        self.doc_frequency: Dict[str, int] = defaultdict(int)
        self.total_docs: int = 0
        self.total_words: int = 0
        self._lemma_cache: Dict[str, str] = {}

    def _is_cyrillic(self, word: str) -> bool:
        return bool(re.search(r'[а-яёА-ЯЁ]', word))

    def _is_latin(self, word: str) -> bool:
        return bool(re.search(r'[a-zA-Z]', word))

    def _get_wordnet_pos(self, word: str) -> str:
        try:
            tag = nltk.pos_tag([word])[0][1][0].upper()
            tag_dict = {'J': wordnet.ADJ, 'N': wordnet.NOUN, 'V': wordnet.VERB, 'R': wordnet.ADV}
            return tag_dict.get(tag, wordnet.NOUN)
        except:
            return wordnet.NOUN

    def _lemmatize_word(self, word: str) -> str:
        word_lower = word.lower().strip()
        if not word_lower or len(word_lower) < 2:
            return word_lower
        if word_lower in self._lemma_cache:
            return self._lemma_cache[word_lower]
        lemma = word_lower
        if self._is_cyrillic(word_lower) and morph:
            parsed = morph.parse(word_lower)
            if parsed:
                lemma = parsed[0].normal_form
        elif self._is_latin(word_lower) and lemmatizer:
            pos = self._get_wordnet_pos(word_lower)
            lemma = lemmatizer.lemmatize(word_lower, pos)
        self._lemma_cache[word_lower] = lemma
        return lemma

    def lemmatize(self, text: str) -> List[Tuple[str, str, int]]:
        """Tokenize and lemmatize text, supporting hyphenated words"""
        words = []
        # Match words including hyphenated compounds (e.g., "Three-cycle", "научно-технический")
        for match in re.finditer(r'[а-яёА-ЯЁa-zA-Z]+(?:-[а-яёА-ЯЁa-zA-Z]+)*', text):
            word = match.group()
            position = match.start()
            
            # Handle hyphenated words
            if '-' in word:
                # Index the full compound word
                full_lemma = self._lemmatize_compound(word)
                words.append((word, full_lemma, position))
                
                # Also index individual parts for broader matching
                parts = word.split('-')
                offset = 0
                for part in parts:
                    if part:
                        part_lemma = self._lemmatize_word(part)
                        words.append((part, part_lemma, position + offset))
                    offset += len(part) + 1  # +1 for hyphen
            else:
                lemma = self._lemmatize_word(word)
                words.append((word, lemma, position))
        return words

    def _lemmatize_compound(self, word: str) -> str:
        """Lemmatize a hyphenated compound word"""
        parts = word.split('-')
        lemmatized_parts = [self._lemmatize_word(part) for part in parts if part]
        return '-'.join(lemmatized_parts)

    def add_document(self, doc_id: str, text: str, filename: str) -> int:
        if doc_id in self.documents:
            self.remove_document(doc_id)
        words = self.lemmatize(text)
        doc_lemmas: Set[str] = set()
        for original, lemma, position in words:
            if lemma:
                self.inverted_index[lemma][doc_id].append(position)
                doc_lemmas.add(lemma)
        for lemma in doc_lemmas:
            self.doc_frequency[lemma] += 1
        self.documents[doc_id] = {'filename': filename, 'text': text, 'word_count': len(words)}
        self.total_docs += 1
        self.total_words += len(words)
        return len(words)

    def remove_document(self, doc_id: str):
        if doc_id not in self.documents:
            return
        lemmas_to_check = [l for l, docs in self.inverted_index.items() if doc_id in docs]
        for lemma in lemmas_to_check:
            del self.inverted_index[lemma][doc_id]
            self.doc_frequency[lemma] -= 1
            if self.doc_frequency[lemma] <= 0:
                del self.doc_frequency[lemma]
            if not self.inverted_index[lemma]:
                del self.inverted_index[lemma]
        self.total_words -= self.documents[doc_id]['word_count']
        self.total_docs -= 1
        del self.documents[doc_id]

    def _calculate_tfidf(self, lemma: str, doc_id: str) -> float:
        if lemma not in self.inverted_index or doc_id not in self.inverted_index[lemma]:
            return 0.0
        tf = len(self.inverted_index[lemma][doc_id])
        doc_word_count = self.documents[doc_id]['word_count']
        if doc_word_count > 0:
            tf = tf / doc_word_count
        df = self.doc_frequency.get(lemma, 0)
        idf = math.log(self.total_docs / df) + 1 if df > 0 and self.total_docs > 0 else 1
        return tf * idf

    def search(self, query: str) -> List[Tuple[str, int, str, List[int]]]:
        """Search for exact phrase (words in sequence) in documents"""
        if not query.strip():
            return []
        
        query_words = self.lemmatize(query)
        query_lemmas = [lemma for _, lemma, _ in query_words if lemma]
        if not query_lemmas:
            return []
        
        query_lemmas_set = set(query_lemmas)
        
        # Find candidate documents that have ALL query lemmas
        candidate_docs: Set[str] = set()
        for lemma in query_lemmas_set:
            if lemma in self.inverted_index:
                if not candidate_docs:
                    candidate_docs = set(self.inverted_index[lemma].keys())
                else:
                    candidate_docs &= set(self.inverted_index[lemma].keys())
        
        if not candidate_docs:
            return []
        
        results = []
        
        # For single word queries, just count occurrences
        if len(query_lemmas) == 1:
            lemma = query_lemmas[0]
            for doc_id in candidate_docs:
                positions = self.inverted_index[lemma][doc_id]
                count = len(positions)
                results.append((doc_id, count, self.documents[doc_id]['filename'], positions))
        else:
            # For multi-word queries, find exact phrase matches
            for doc_id in candidate_docs:
                phrase_positions = self._find_phrase_in_document(doc_id, query_lemmas)
                if phrase_positions:
                    count = len(phrase_positions)
                    results.append((doc_id, count, self.documents[doc_id]['filename'], phrase_positions))
        
        results.sort(key=lambda x: x[1], reverse=True)
        return results

    def _find_phrase_in_document(self, doc_id: str, query_lemmas: List[str], max_distance: int = 15) -> List[int]:
        """Find all occurrences of a phrase in a document.
        Returns list of starting positions where the phrase was found.
        max_distance is the maximum character distance between consecutive words."""
        if doc_id not in self.documents:
            return []
        
        # Get positions for each lemma in the query
        lemma_positions = []
        for lemma in query_lemmas:
            if lemma in self.inverted_index and doc_id in self.inverted_index[lemma]:
                lemma_positions.append(sorted(self.inverted_index[lemma][doc_id]))
            else:
                return []  # If any lemma is missing, no phrase match possible
        
        if not lemma_positions or not lemma_positions[0]:
            return []
        
        phrase_starts = []
        first_positions = lemma_positions[0]
        
        for start_pos in first_positions:
            current_pos = start_pos
            matched = True
            
            for i in range(1, len(lemma_positions)):
                next_positions = lemma_positions[i]
                found_next = False
                
                # Find the next lemma position that comes after current_pos within max_distance
                for pos in next_positions:
                    if current_pos < pos <= current_pos + max_distance:
                        current_pos = pos
                        found_next = True
                        break
                
                if not found_next:
                    matched = False
                    break
            
            if matched:
                phrase_starts.append(start_pos)
        
        return phrase_starts

    def get_query_lemmas(self, query: str) -> List[str]:
        """Get ordered list of lemmas from query"""
        return [lemma for _, lemma, _ in self.lemmatize(query) if lemma]

    def _parse_phrases(self, query: str) -> List[List[str]]:
        """Parse query into phrases (multi-word sequences)"""
        phrases = []
        
        # Handle quoted phrases first
        quoted_pattern = r'"([^"]+)"|\'([^\']+)\''
        for match in re.finditer(quoted_pattern, query):
            phrase_text = match.group(1) or match.group(2)
            phrase_lemmas = [lemma for _, lemma, _ in self.lemmatize(phrase_text) if lemma]
            if len(phrase_lemmas) >= 2:
                phrases.append(phrase_lemmas)
        
        # Remove quoted parts from query
        remaining_query = re.sub(quoted_pattern, ' ', query)
        
        # Split by common phrase delimiters: semicolons, commas
        parts = re.split(r'[;,]', remaining_query)
        
        for part in parts:
            part = part.strip()
            if not part:
                continue
            
            # Each part with 2+ words is treated as a potential phrase
            part_lemmas = [lemma for _, lemma, _ in self.lemmatize(part) if lemma]
            if len(part_lemmas) >= 2:
                phrases.append(part_lemmas)
        
        return phrases

    def _calculate_phrase_boost(self, doc_id: str, phrases: List[List[str]], max_distance: int = 10) -> float:
        """Calculate boost factor based on how well phrases match in the document"""
        if doc_id not in self.documents:
            return 0.0
        
        total_boost = 0.0
        
        for phrase_lemmas in phrases:
            if len(phrase_lemmas) < 2:
                continue
            
            # Get positions for each lemma in the phrase
            lemma_positions = []
            for lemma in phrase_lemmas:
                if lemma in self.inverted_index and doc_id in self.inverted_index[lemma]:
                    lemma_positions.append(self.inverted_index[lemma][doc_id])
                else:
                    lemma_positions.append([])
            
            # Check if all lemmas are present
            if any(len(pos) == 0 for pos in lemma_positions):
                continue
            
            # Find sequences where words appear close together
            phrase_matches = self._find_phrase_matches(lemma_positions, max_distance)
            
            if phrase_matches > 0:
                # Boost based on number of phrase matches and phrase length
                phrase_boost = phrase_matches * (len(phrase_lemmas) ** 1.5) * 0.5
                total_boost += phrase_boost
        
        return min(total_boost, 5.0)  # Cap the boost

    def _find_phrase_matches(self, lemma_positions: List[List[int]], max_distance: int) -> int:
        """Find how many times words appear in sequence within max_distance"""
        if not lemma_positions or not lemma_positions[0]:
            return 0
        
        matches = 0
        first_positions = lemma_positions[0]
        
        for start_pos in first_positions:
            current_pos = start_pos
            matched = True
            
            for i in range(1, len(lemma_positions)):
                # Find next lemma position that's after current_pos but within max_distance
                next_positions = lemma_positions[i]
                found_next = False
                
                for pos in next_positions:
                    if current_pos < pos <= current_pos + max_distance:
                        current_pos = pos
                        found_next = True
                        break
                
                if not found_next:
                    matched = False
                    break
            
            if matched:
                matches += 1
        
        return matches

    def get_context(self, doc_id: str, positions: List[int], context_size: int = 50) -> str:
        if doc_id not in self.documents:
            return ""
        text = self.documents[doc_id]['text']
        if not positions:
            return text[:200] + "..." if len(text) > 200 else text
        contexts = []
        used_ranges = []
        for pos in positions[:3]:
            start = max(0, pos - context_size)
            end = min(len(text), pos + context_size)
            overlap = any(start < r_end and end > r_start for r_start, r_end in used_ranges)
            if not overlap:
                context = text[start:end]
                if start > 0:
                    context = "..." + context
                if end < len(text):
                    context = context + "..."
                contexts.append(context)
                used_ranges.append((start, end))
        return " | ".join(contexts)

    def get_sentences_with_matches(self, doc_id: str, positions: List[int]) -> List[Tuple[int, str]]:
        """Get sentences containing matches with 1 sentence before and after.
        Returns list of (match_number, context_text) tuples."""
        if doc_id not in self.documents:
            return []
        
        text = self.documents[doc_id]['text']
        if not positions:
            return []
        
        # Split text into sentences
        # Pattern matches sentence endings: . ! ? followed by space or newline
        sentence_pattern = r'[^.!?]*[.!?]+(?:\s|$)|[^.!?\n]+$'
        sentences = []
        for match in re.finditer(sentence_pattern, text):
            sent_text = match.group().strip()
            if sent_text:
                sentences.append((match.start(), match.end(), sent_text))
        
        if not sentences:
            return [(1, text[:500] + "..." if len(text) > 500 else text)]
        
        # Find which sentences contain matches
        matched_contexts = []
        used_positions = set()
        
        for match_idx, pos in enumerate(positions):
            if pos in used_positions:
                continue
                
            # Find sentence containing this position
            sent_idx = None
            for idx, (start, end, sent_text) in enumerate(sentences):
                if start <= pos < end:
                    sent_idx = idx
                    break
            
            if sent_idx is None:
                continue
            
            # Mark this position as used
            used_positions.add(pos)
            
            # Get 1 sentence before, the match sentence, and 1 sentence after
            context_parts = []
            
            # Previous sentence
            if sent_idx > 0:
                context_parts.append(sentences[sent_idx - 1][2])
            
            # Current sentence (with match)
            context_parts.append(sentences[sent_idx][2])
            
            # Next sentence
            if sent_idx < len(sentences) - 1:
                context_parts.append(sentences[sent_idx + 1][2])
            
            context = ' '.join(context_parts)
            
            # Check if this context overlaps with already added ones
            is_duplicate = False
            for _, existing_context in matched_contexts:
                if sentences[sent_idx][2] in existing_context:
                    is_duplicate = True
                    break
            
            if not is_duplicate:
                matched_contexts.append((match_idx + 1, context))
        
        return matched_contexts

    def save_index(self, filepath: str):
        data = {
            'documents': self.documents,
            'inverted_index': {k: dict(v) for k, v in self.inverted_index.items()},
            'doc_frequency': dict(self.doc_frequency),
            'total_docs': self.total_docs,
            'total_words': self.total_words
        }
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def load_index(self, filepath: str):
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self.documents = data['documents']
        self.inverted_index = defaultdict(lambda: defaultdict(list))
        for lemma, docs in data['inverted_index'].items():
            for doc_id, positions in docs.items():
                self.inverted_index[lemma][doc_id] = positions
        self.doc_frequency = defaultdict(int, data['doc_frequency'])
        self.total_docs = data['total_docs']
        self.total_words = data['total_words']
        self._lemma_cache.clear()


# ==================== Indexing Thread ====================

class IndexingThread(QThread):
    progress = pyqtSignal(str, int, int)
    finished_file = pyqtSignal(str, str)
    error = pyqtSignal(str)
    done = pyqtSignal()

    def __init__(self, files: List[str], engine: LemmaSearchEngine):
        super().__init__()
        self.files = files
        self.engine = engine
        self.doc_paths: Dict[str, str] = {}
        self.cancelled = False

    def run(self):
        for i, filepath in enumerate(self.files):
            if self.cancelled:
                break
            filename = os.path.basename(filepath)
            self.progress.emit(f"Индексация: {filename}", i + 1, len(self.files))
            try:
                text = extract_text(filepath)
                self.engine.add_document(filepath, text, filename)
                self.doc_paths[filepath] = filepath
                self.finished_file.emit(filepath, filename)
            except Exception as e:
                self.error.emit(f"{filename}: {str(e)}")
        self.done.emit()


# ==================== Search Thread ====================

class SearchThread(QThread):
    """Thread for performing search and collecting results"""
    progress = pyqtSignal(str)
    result_ready = pyqtSignal(object, str)  # (results, query)
    done = pyqtSignal()

    def __init__(self, engine: LemmaSearchEngine, query: str):
        super().__init__()
        self.engine = engine
        self.query = query
        self.results = []

    def run(self):
        self.progress.emit("Поиск совпадений...")
        self.results = self.engine.search(self.query)
        self.result_ready.emit(self.results, self.query)
        self.done.emit()


# ==================== Main Window ====================

class LemmaCheckApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.engine = LemmaSearchEngine()
        self.doc_paths: Dict[str, str] = {}
        self.result_doc_ids: List[str] = []
        self.indexing_thread: Optional[IndexingThread] = None
        self.search_thread: Optional[SearchThread] = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("LemmaCheck - Полнотекстовый поиск по леммам")
        self.setGeometry(100, 100, 900, 700)
        self.setMinimumSize(800, 600)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)

        # === Documents Section ===
        doc_group = QGroupBox("📁 Документы")
        doc_layout = QVBoxLayout(doc_group)

        btn_layout = QHBoxLayout()
        self.btn_add_files = QPushButton("Добавить файлы")
        self.btn_add_files.clicked.connect(self.add_files)
        btn_layout.addWidget(self.btn_add_files)

        self.btn_add_folder = QPushButton("Добавить папку")
        self.btn_add_folder.clicked.connect(self.add_folder)
        btn_layout.addWidget(self.btn_add_folder)

        self.btn_remove = QPushButton("Удалить")
        self.btn_remove.setStyleSheet("background-color: #d9534f; color: white;")
        self.btn_remove.clicked.connect(self.remove_selected)
        btn_layout.addWidget(self.btn_remove)

        self.btn_clear = QPushButton("Очистить")
        self.btn_clear.setStyleSheet("background-color: #d9534f; color: white;")
        self.btn_clear.clicked.connect(self.clear_all)
        btn_layout.addWidget(self.btn_clear)

        btn_layout.addStretch()

        self.btn_save = QPushButton("Сохранить индекс")
        self.btn_save.setStyleSheet("background-color: #5cb85c; color: white;")
        self.btn_save.clicked.connect(self.save_index)
        btn_layout.addWidget(self.btn_save)

        self.btn_load = QPushButton("Загрузить индекс")
        self.btn_load.setStyleSheet("background-color: #5cb85c; color: white;")
        self.btn_load.clicked.connect(self.load_index)
        btn_layout.addWidget(self.btn_load)

        doc_layout.addLayout(btn_layout)

        self.doc_list = QListWidget()
        self.doc_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.doc_list.setMaximumHeight(100)
        doc_layout.addWidget(self.doc_list)

        # Progress
        self.progress_widget = QWidget()
        progress_layout = QHBoxLayout(self.progress_widget)
        progress_layout.setContentsMargins(0, 0, 0, 0)
        self.progress_label = QLabel("")
        progress_layout.addWidget(self.progress_label)
        self.progress_bar = QProgressBar()
        progress_layout.addWidget(self.progress_bar)
        self.btn_cancel = QPushButton("Отмена")
        self.btn_cancel.clicked.connect(self.cancel_indexing)
        progress_layout.addWidget(self.btn_cancel)
        self.progress_widget.hide()
        doc_layout.addWidget(self.progress_widget)

        layout.addWidget(doc_group)

        # === Search Section ===
        search_group = QGroupBox("🔍 Поиск")
        search_layout = QHBoxLayout(search_group)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Введите поисковый запрос...")
        self.search_input.setFont(QFont("Helvetica", 14))
        self.search_input.returnPressed.connect(self.search)
        search_layout.addWidget(self.search_input)

        self.btn_search = QPushButton("Искать")
        self.btn_search.setStyleSheet("background-color: #337ab7; color: white;")
        self.btn_search.clicked.connect(self.search)
        search_layout.addWidget(self.btn_search)

        layout.addWidget(search_group)

        # === Results Section ===
        results_group = QGroupBox("📋 Результаты")
        results_layout = QHBoxLayout(results_group)
        
        # Main results text
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        self.results_text.setFont(QFont("Helvetica", 12))
        
        # Found words panel
        words_frame = QFrame()
        words_frame.setMaximumWidth(250)
        words_frame.setMinimumWidth(180)
        words_layout = QVBoxLayout(words_frame)
        words_layout.setContentsMargins(0, 0, 0, 0)
        
        words_label = QLabel("🔤 Найденные слова:")
        words_label.setFont(QFont("Helvetica", 11, QFont.Weight.Bold))
        words_layout.addWidget(words_label)
        
        self.found_words_text = QTextEdit()
        self.found_words_text.setReadOnly(True)
        self.found_words_text.setFont(QFont("Helvetica", 11))
        words_layout.addWidget(self.found_words_text)
        
        self.btn_copy_words = QPushButton("📋 Копировать список")
        self.btn_copy_words.clicked.connect(self.copy_found_words)
        words_layout.addWidget(self.btn_copy_words)
        
        # Use splitter for resizable panels
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self.results_text)
        splitter.addWidget(words_frame)
        splitter.setSizes([700, 200])
        
        results_layout.addWidget(splitter)

        layout.addWidget(results_group, stretch=1)

        # === Status Bar ===
        self.status_label = QLabel("Готово. Добавьте документы для начала работы.")
        layout.addWidget(self.status_label)

    def update_status(self):
        self.status_label.setText(
            f"Документов: {len(self.engine.documents)} | Индексировано слов: {self.engine.total_words}"
        )

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выберите документы", "",
            "Все поддерживаемые (*.docx *.doc *.pdf *.txt *.xlsx *.xls);;"
            "Word (*.docx *.doc);;PDF (*.pdf);;Text (*.txt);;Excel (*.xlsx *.xls)"
        )
        if files:
            self.index_files(files)

    def add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if folder:
            files = []
            for root, _, filenames in os.walk(folder):
                for fn in filenames:
                    if fn.lower().endswith(('.docx', '.doc', '.pdf', '.txt', '.xlsx', '.xls')):
                        files.append(os.path.join(root, fn))
            if files:
                self.index_files(files)
            else:
                QMessageBox.information(self, "Информация", "Поддерживаемых документов не найдено.")

    def index_files(self, files: List[str]):
        if self.indexing_thread and self.indexing_thread.isRunning():
            QMessageBox.warning(self, "Внимание", "Индексация уже выполняется.")
            return

        self.progress_widget.show()
        self.progress_bar.setValue(0)
        self.set_buttons_enabled(False)

        self.indexing_thread = IndexingThread(files, self.engine)
        self.indexing_thread.progress.connect(self.on_progress)
        self.indexing_thread.finished_file.connect(self.on_file_indexed)
        self.indexing_thread.error.connect(self.on_indexing_error)
        self.indexing_thread.done.connect(self.on_indexing_done)
        self.indexing_thread.start()

    def on_progress(self, text: str, current: int, total: int):
        self.progress_label.setText(text)
        self.progress_bar.setValue(int(current / total * 100))

    def on_file_indexed(self, doc_id: str, filename: str):
        self.doc_paths[doc_id] = doc_id
        self.doc_list.addItem(filename)

    def on_indexing_error(self, error: str):
        QMessageBox.warning(self, "Ошибка", error)

    def on_indexing_done(self):
        self.progress_widget.hide()
        self.set_buttons_enabled(True)
        self.update_status()

    def cancel_indexing(self):
        if self.indexing_thread:
            self.indexing_thread.cancelled = True

    def set_buttons_enabled(self, enabled: bool):
        for btn in [self.btn_add_files, self.btn_add_folder, self.btn_remove,
                    self.btn_clear, self.btn_save, self.btn_load, self.btn_search]:
            btn.setEnabled(enabled)

    def remove_selected(self):
        for item in self.doc_list.selectedItems():
            filename = item.text()
            for doc_id, doc in list(self.engine.documents.items()):
                if doc['filename'] == filename:
                    self.engine.remove_document(doc_id)
                    if doc_id in self.doc_paths:
                        del self.doc_paths[doc_id]
                    break
            self.doc_list.takeItem(self.doc_list.row(item))
        self.update_status()

    def clear_all(self):
        if not self.engine.documents:
            return
        if QMessageBox.question(self, "Подтверждение", "Удалить все документы?") == QMessageBox.StandardButton.Yes:
            self.engine = LemmaSearchEngine()
            self.doc_paths.clear()
            self.doc_list.clear()
            self.results_text.clear()
            self.found_words_text.clear()
            self.update_status()

    def save_index(self):
        if not self.engine.documents:
            QMessageBox.information(self, "Информация", "Нет документов для сохранения.")
            return
        filepath, _ = QFileDialog.getSaveFileName(self, "Сохранить индекс", "", "JSON (*.json)")
        if filepath:
            try:
                self.engine.save_index(filepath)
                paths_file = filepath.replace('.json', '_paths.json')
                with open(paths_file, 'w', encoding='utf-8') as f:
                    json.dump({'doc_paths': self.doc_paths}, f, ensure_ascii=False)
                QMessageBox.information(self, "Успешно", "Индекс сохранён.")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", str(e))

    def load_index(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Загрузить индекс", "", "JSON (*.json)")
        if filepath:
            try:
                self.engine.load_index(filepath)
                paths_file = filepath.replace('.json', '_paths.json')
                if os.path.exists(paths_file):
                    with open(paths_file, 'r', encoding='utf-8') as f:
                        self.doc_paths = json.load(f).get('doc_paths', {})
                else:
                    self.doc_paths = {doc_id: doc_id for doc_id in self.engine.documents}
                self.doc_list.clear()
                for doc in self.engine.documents.values():
                    self.doc_list.addItem(doc['filename'])
                self.update_status()
                QMessageBox.information(self, "Успешно", "Индекс загружен.")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", str(e))

    def search(self):
        query = self.search_input.text().strip()
        if not query:
            return
        if not self.engine.documents:
            QMessageBox.information(self, "Информация", "Сначала добавьте документы.")
            return
        
        # Show loading state
        self.results_text.clear()
        self.found_words_text.clear()
        self.results_text.setPlainText("⏳ Поиск...")
        self.btn_search.setEnabled(False)
        self.status_label.setText("Выполняется поиск...")
        QApplication.processEvents()
        
        # Run search in thread
        self.search_thread = SearchThread(self.engine, query)
        self.search_thread.result_ready.connect(self.on_search_results)
        self.search_thread.start()

    def on_search_results(self, results, query):
        """Handle search results from thread"""
        self.btn_search.setEnabled(True)
        self.results_text.clear()
        self.results_text.setPlainText("⏳ Обработка результатов...")
        QApplication.processEvents()
        
        self.display_results(results, query)
        self.update_status()

    def display_results(self, results: List[Tuple[str, int, str, List[int]]], query: str):
        self.results_text.clear()
        self.found_words_text.clear()
        self.result_doc_ids = []

        if not results:
            self.results_text.setPlainText("Ничего не найдено.")
            return

        query_lemmas = self.engine.get_query_lemmas(query)
        query_lemmas_set = set(query_lemmas)
        is_phrase_search = len(query_lemmas) > 1
        
        # Collect found phrases/words per document
        phrases_by_doc: Dict[str, List[str]] = {}  # filename -> list of found phrases
        
        cursor = self.results_text.textCursor()
        
        # Formats
        title_fmt = QTextCharFormat()
        title_fmt.setFontWeight(700)
        title_fmt.setForeground(QColor("#0066cc"))
        
        count_fmt = QTextCharFormat()
        count_fmt.setForeground(QColor("#666666"))
        
        para_num_fmt = QTextCharFormat()
        para_num_fmt.setForeground(QColor("#888888"))
        para_num_fmt.setFontWeight(600)
        
        # Soft highlight - light blue background, readable
        highlight_fmt = QTextCharFormat()
        highlight_fmt.setBackground(QColor("#b3e5fc"))  # Light blue
        highlight_fmt.setForeground(QColor("#01579b"))  # Dark blue text
        highlight_fmt.setFontWeight(600)
        
        normal_fmt = QTextCharFormat()
        
        total_results = len(results)

        for idx, (doc_id, count, filename, positions) in enumerate(results):
            # Update progress
            self.status_label.setText(f"Обработка документа {idx + 1} из {total_results}...")
            QApplication.processEvents()
            
            self.result_doc_ids.append(doc_id)
            phrases_by_doc[filename] = []
            
            full_text = self.engine.documents[doc_id]['text']
            
            if is_phrase_search:
                # For phrase search, extract the actual phrases found
                for pos in positions:
                    phrase = self._extract_phrase_at_position(full_text, pos, len(query_lemmas))
                    if phrase:
                        phrases_by_doc[filename].append(phrase)
                total_count = count  # count already contains phrase matches
            else:
                # For single word search, count all occurrences
                total_count = 0
                for match in re.finditer(r'[а-яёА-ЯЁa-zA-Z]+(?:-[а-яёА-ЯЁa-zA-Z]+)*', full_text):
                    word = match.group()
                    lemma = self.engine._lemmatize_word(word)
                    if lemma in query_lemmas_set:
                        phrases_by_doc[filename].append(word.lower())
                        total_count += 1
            
            cursor.insertText(f"📄 {filename}", title_fmt)
            cursor.insertText(f"  [найдено: {total_count}]\n\n", count_fmt)
            
            # Get sentences with matches (1 before, match, 1 after)
            sentence_contexts = self.engine.get_sentences_with_matches(doc_id, positions)
            
            for match_num, context_text in sentence_contexts:
                cursor.insertText(f"[{match_num}] ", para_num_fmt)
                
                # Highlight matches in context
                if is_phrase_search:
                    self._highlight_phrases_in_context(cursor, context_text, query_lemmas, highlight_fmt, normal_fmt)
                else:
                    last_end = 0
                    for match in re.finditer(r'[а-яёА-ЯЁa-zA-Z]+(?:-[а-яёА-ЯЁa-zA-Z]+)*', context_text):
                        word = match.group()
                        start = match.start()
                        
                        if start > last_end:
                            cursor.insertText(context_text[last_end:start], normal_fmt)
                        
                        lemma = self.engine._lemmatize_word(word)
                        if lemma in query_lemmas_set:
                            cursor.insertText(word, highlight_fmt)
                        else:
                            cursor.insertText(word, normal_fmt)
                        
                        last_end = match.end()
                    
                    if last_end < len(context_text):
                        cursor.insertText(context_text[last_end:], normal_fmt)
                
                cursor.insertText("\n\n", normal_fmt)
            
            cursor.insertText("─" * 70 + "\n\n", normal_fmt)

        cursor.insertText(f"Найдено документов: {len(results)}", normal_fmt)
        
        # Populate found phrases panel grouped by document
        words_cursor = self.found_words_text.textCursor()
        
        doc_title_fmt = QTextCharFormat()
        doc_title_fmt.setFontWeight(700)
        doc_title_fmt.setForeground(QColor("#0066cc"))
        
        total_fmt = QTextCharFormat()
        total_fmt.setFontWeight(700)
        total_fmt.setForeground(QColor("#2e7d32"))
        
        stats_fmt = QTextCharFormat()
        stats_fmt.setForeground(QColor("#555555"))
        
        variation_fmt = QTextCharFormat()
        variation_fmt.setForeground(QColor("#7b1fa2"))  # Purple for variations count
        variation_fmt.setFontWeight(600)
        
        word_fmt = QTextCharFormat()
        
        grand_total = 0
        grand_variations = 0
        
        for filename, phrases in phrases_by_doc.items():
            if phrases:
                # Count unique phrases
                phrase_counts = defaultdict(int)
                for phrase in phrases:
                    phrase_counts[phrase.lower()] += 1
                
                doc_total = sum(phrase_counts.values())
                doc_variations = len(phrase_counts)
                grand_total += doc_total
                grand_variations += doc_variations
                
                words_cursor.insertText(f"📄 {filename}\n", doc_title_fmt)
                words_cursor.insertText(f"   Всего: {doc_total} | ", stats_fmt)
                words_cursor.insertText(f"Вариаций: {doc_variations}\n", variation_fmt)
                
                sorted_phrases = sorted(phrase_counts.items(), key=lambda x: (-x[1], x[0]))
                for phrase, cnt in sorted_phrases:
                    words_cursor.insertText(f"  • {phrase} ({cnt})\n", word_fmt)
                words_cursor.insertText("\n", word_fmt)
        
        # Add grand total
        words_cursor.insertText("═" * 25 + "\n", word_fmt)
        words_cursor.insertText(f"📊 ИТОГО\n", total_fmt)
        words_cursor.insertText(f"   Совпадений: {grand_total}\n", stats_fmt)
        words_cursor.insertText(f"   Вариаций: {grand_variations}\n", variation_fmt)
        words_cursor.insertText(f"   Документов: {len(results)}", stats_fmt)
        
        self.status_label.setText(f"Найдено {grand_total} совпадений ({grand_variations} вариаций) в {len(results)} документах")

    def _extract_phrase_at_position(self, text: str, start_pos: int, word_count: int) -> str:
        """Extract a phrase starting at given position with given number of words"""
        words = []
        for match in re.finditer(r'[а-яёА-ЯЁa-zA-Z]+(?:-[а-яёА-ЯЁa-zA-Z]+)*', text[start_pos:]):
            words.append(match.group())
            if len(words) >= word_count:
                break
        return ' '.join(words) if len(words) == word_count else ''

    def _highlight_phrases_in_context(self, cursor, context: str, query_lemmas: List[str], 
                                       highlight_fmt, normal_fmt):
        """Highlight phrase matches in context text"""
        # Find all word positions in context
        word_matches = list(re.finditer(r'[а-яёА-ЯЁa-zA-Z]+(?:-[а-яёА-ЯЁa-zA-Z]+)*', context))
        
        if not word_matches:
            cursor.insertText(context, normal_fmt)
            return
        
        # Find phrase matches (sequences of words matching query lemmas)
        highlight_ranges = []
        i = 0
        while i <= len(word_matches) - len(query_lemmas):
            matched = True
            for j, query_lemma in enumerate(query_lemmas):
                word = word_matches[i + j].group()
                word_lemma = self.engine._lemmatize_word(word)
                if word_lemma != query_lemma:
                    matched = False
                    break
            
            if matched:
                start = word_matches[i].start()
                end = word_matches[i + len(query_lemmas) - 1].end()
                highlight_ranges.append((start, end))
                i += len(query_lemmas)  # Skip past matched phrase
            else:
                i += 1
        
        # Output text with highlights
        last_end = 0
        for start, end in highlight_ranges:
            if start > last_end:
                cursor.insertText(context[last_end:start], normal_fmt)
            cursor.insertText(context[start:end], highlight_fmt)
            last_end = end
        
        if last_end < len(context):
            cursor.insertText(context[last_end:], normal_fmt)

    def open_file(self, filepath: str):
        try:
            if platform.system() == 'Darwin':
                subprocess.run(['open', filepath], check=True)
            elif platform.system() == 'Windows':
                os.startfile(filepath)
            else:
                subprocess.run(['xdg-open', filepath], check=True)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть файл: {e}")

    def copy_found_words(self):
        """Copy all found words to clipboard"""
        text = self.found_words_text.toPlainText()
        
        if text.strip():
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            lines = [l for l in text.split('\n') if l.strip()]
            self.status_label.setText(f"Скопировано {len(lines)} строк в буфер обмена")


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Dark palette (optional, comment out for light theme)
    # from PyQt6.QtGui import QPalette
    # palette = QPalette()
    # palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
    # palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    # app.setPalette(palette)
    
    window = LemmaCheckApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
