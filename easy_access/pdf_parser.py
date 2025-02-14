"""
This module reads/parsers/transforms pdfs in SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
to prepare them for NLP analysis.
"""

import hashlib
from _collections_abc import dict_keys
from easy_access.downloader import Downloader
from easy_access.utils import File, Directory, cool, warn, info
from easy_access.settings import SETTINGS, DirSetting
import polars as pl
from pathlib import Path
import torch
from torch.utils.data import Dataset, DataLoader
from transformers import AutoTokenizer, AutoModelForSequenceClassification, Trainer, TrainingArguments
import spacy
from spacy_layout import spaCyLayout
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder
from typing import Any
from difflib import SequenceMatcher

from dataclasses import dataclass

def retrieve_files() -> list[File] | list[None]:
    """
    Retrieves all pdf files from the pdf downloads dir.
    """
    pdf_dir = Directory(path=SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].full)
    if not pdf_dir:
        return []
    if not pdf_dir.files:
        return []
    return [file for file in pdf_dir.files if file.extension == '.pdf']

def retrieve_manually_classified() -> dict[str, dict[str,str]]:
    """
    Retrieves all currently finalized manual classifications from SETTINGS.dirs[DirSetting.SCRIPT_DATA] / full_data.parquet.
    """
    print(SETTINGS.classification_options, type(SETTINGS.classification_options))
    possible_classifications: list[str] = SETTINGS.classification_options
    if isinstance(possible_classifications[0], list):
        possible_classifications = possible_classifications[0]

    manual_classifications_file = File(path=SETTINGS.dirs[DirSetting.SCRIPT_DATA].full / 'full_data.parquet')
    if not manual_classifications_file.exists:
        return {}
    data = pl.read_parquet(source=manual_classifications_file.path, columns=['material_id', 'manual_classification', 'filename', 'workflow_status'])
    data = data.filter(data['workflow_status'] == 'Done')
    data = data.with_columns(pl.col('manual_classification').str.to_lowercase().replace("-",new="").str.strip_chars().str.replace("",None)).filter(data['manual_classification'].is_not_null())

    data = data.filter(data['manual_classification'].is_in(possible_classifications))
    info(f"Found {len(data)} manual classifications.")
    data_dict = data.to_dicts()
    return {item['material_id']: item for item in data_dict}


def get_manual_classification_for_files() -> dict[str, dict[str, str|Path]] | dict[None]:
    """
    Retrieves all manual classifications for all files in SETTINGS.dirs[DirSetting.PDF_DOWNLOADS].
    """
    manual_classifications = retrieve_manually_classified()
    downloader = Downloader()
    subset: dict_keys[str, dict[str, str]] = manual_classifications.keys()
    info(f'Checking {len(subset)} files for a pdf, or downloading.')
    downloader.download_pdfs(subset=subset)
    files = retrieve_files()
    if not files or not manual_classifications:
        return {}
    manual_classifications_for_files = {}
    for file in files:
        try:
            mat_id = file.name.split('_')[0]
        except Exception as e:
            warn(f'Could not extract material_id from {file.name}')
            continue
        if mat_id in manual_classifications:
            manual_classifications_for_files[mat_id] ={"classification":manual_classifications[mat_id].get('manual_classification'), "file_path": file.path}
    info(f'Found {len(manual_classifications_for_files)} manual classifications for {len(files)} files.')
    return manual_classifications_for_files

@dataclass
class PDFDataset(Dataset):
    """Dataset for training the model"""
    texts: list[str]
    labels: list[int]
    tokenizer: Any
    max_length: int = 512

    def __len__(self):
        return len(self.texts)

    def __getitem__(self, idx):
        text = self.texts[idx]
        label = self.labels[idx]

        encoding = self.tokenizer(
            text,
            truncation=True,
            max_length=self.max_length,
            padding='max_length',
            return_tensors='pt'
        )

        return {
            'input_ids': encoding['input_ids'][0],
            'attention_mask': encoding['attention_mask'][0],
            'labels': torch.tensor(label)
        }

def extract_text_from_pdfs_batch(file_paths: list[Path]) -> dict[Path, str]:
    """
    Extract text content from multiple PDF files using spacy_layout's batch processing
    First checks for cached text files, processes only uncached PDFs
    Returns dictionary mapping file paths to their extracted text
    """
    results = {}
    paths_to_process = []

    # First check for cached versions
    for file_path in file_paths:
        text_path = file_path.with_suffix('.txt')
        if text_path.exists():
            try:
                with open(text_path, 'r', encoding='utf-8') as f:
                    results[file_path] = f.read()
            except Exception as e:
                warn(f"Error reading cached text file {text_path}: {e}")
                paths_to_process.append(file_path)
        else:
            paths_to_process.append(file_path)

    if paths_to_process:
        try:
            # Initialize spacy with transformer model
            nlp = spacy.load("en_core_web_trf")
            layout = spaCyLayout(nlp)

            # Process PDFs in batch
            for doc in layout.pipe([str(p) for p in paths_to_process]):
                doc = nlp(doc)  # Apply NLP pipeline
                file_path = Path(doc._.pdf_path)

                # Extract text while preserving section structure
                sections = []
                for span in doc.spans["layout"]:
                    if span.label_ in ["text", "title", "section_header"]:
                        section_text = f"{span.label_}: {span.text}"
                        sections.append(section_text)

                extracted_text = "\n\n".join(sections)
                results[file_path] = extracted_text

                # Cache the extracted text
                text_path = file_path.with_suffix('.txt')
                try:
                    with open(text_path, 'w', encoding='utf-8') as f:
                        f.write(extracted_text)
                except Exception as e:
                    warn(f"Error caching extracted text to {text_path}: {e}")

        except Exception as e:
            warn(f"Error in batch PDF processing: {e}")

    return results

def extract_text_from_pdf(file_path: Path) -> str:
    """
    Extract text content from a single PDF file
    Uses batch processing internally for consistency
    """
    results = extract_text_from_pdfs_batch([file_path])
    return results.get(file_path, "")


def deduplicate_pdfs() -> None:
    """
    Deduplicates PDFs in SETTINGS.dirs[DirSetting.PDF_DOWNLOADS] based on their hash or embedding or something
    make sure to store this somehow that it can be incorporated into the main data

    e.g. if a duplicate is found, link the material ids and show this in the sheets

    """
    pdf_dir = SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
    if not pdf_dir or not pdf_dir.files:
        info("No PDF files found for deduplication.")
        return

    file_hashes = {}
    pdf_texts = {}
    for file in pdf_dir.files:
        if file.extension.lower() == ".pdf":
            pdf_texts[file.path] = extract_text_from_pdf(file.path) or ""

    # print exact duplicates (same hash)
    for pdf_hash, paths in file_hashes.items():
        if len(paths) > 1:
            info(f"Duplicate PDF files found with hash {pdf_hash}: {paths}")
            # do something with the duplicate files

    # find near-duplicates (different hash but highly similar text)
    paths_list = list(pdf_texts.keys())
    n = len(paths_list)
    checked_pairs = set()
    for i in range(n):
        for j in range(i+1, n):
            pair = tuple(sorted([paths_list[i], paths_list[j]]))
            if pair in checked_pairs:
                continue
            checked_pairs.add(pair)

            text_a = pdf_texts[paths_list[i]]
            text_b = pdf_texts[paths_list[j]]
            similarity = SequenceMatcher(None, text_a, text_b).ratio()

            # adjust threshold as needed
            if similarity > 0.95:
                info(f"Near-duplicate PDF files (text similarity > 95%): {pair}")
                # do something with near-duplicate files
