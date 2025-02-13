"""
This module reads/parsers/transforms pdfs in SETTINGS.dirs[DirSetting.PDF_DOWNLOADS]
to prepare them for NLP analysis.
"""


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
from typing import Dict, List, Tuple, Any
import numpy as np
from dataclasses import dataclass
import evaluate

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
    texts: List[str]
    labels: List[int]
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

def extract_text_from_pdfs_batch(file_paths: List[Path]) -> Dict[Path, str]:
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

def prepare_dataset(data_dict: Dict[str, Dict[str, Any]]) -> Tuple[List[str], List[str]]:
    """Prepare texts and labels from the data dictionary"""
    # Get all file paths
    file_paths = [item['file_path'] for item in data_dict.values()]

    # Extract text from all PDFs in batch
    extracted_texts = extract_text_from_pdfs_batch(file_paths)

    texts = []
    labels = []
    for material_id, item in data_dict.items():
        text = extracted_texts.get(item['file_path'])
        if text:  # Only include if we successfully extracted text
            texts.append(text)
            labels.append(item['classification'])

    return texts, labels

def compute_metrics(eval_pred):
    """Compute metrics for model evaluation"""
    metric = evaluate.load("accuracy")
    predictions, labels = eval_pred
    predictions = np.argmax(predictions, axis=1)
    return metric.compute(predictions=predictions, references=labels)

def train_model(training_data: Dict[str, Dict[str, Any]], model_save_path: str = "pdf_classifier") -> Tuple[Any, Any, LabelEncoder]:
    """
    Train a ModernBERT model on the PDF dataset

    Args:
        training_data: Dictionary with material_ids as keys and dicts containing 'classification' and 'file_path' as values
        model_save_path: Where to save the trained model

    Returns:
        Tuple of (trained model, tokenizer, label encoder)
    """
    info("Preparing dataset for training...")
    texts, labels = prepare_dataset(training_data)

    if not texts:
        raise ValueError("No valid text data extracted from PDFs")

    # Encode labels
    label_encoder = LabelEncoder()
    encoded_labels = label_encoder.fit_transform(labels)

    # Split dataset
    train_texts, val_texts, train_labels, val_labels = train_test_split(
        texts, encoded_labels, test_size=0.2, random_state=42
    )

    # Initialize tokenizer and model
    model_id = "answerdotai/ModernBERT-large"
    tokenizer = AutoTokenizer.from_pretrained(model_id)
    model = AutoModelForSequenceClassification.from_pretrained(
        model_id,
        num_labels=len(label_encoder.classes_)
    )

    # Create datasets
    train_dataset = PDFDataset(train_texts, train_labels, tokenizer)
    val_dataset = PDFDataset(val_texts, val_labels, tokenizer)

    # Training arguments
    training_args = TrainingArguments(
        output_dir=model_save_path,
        evaluation_strategy="epoch",
        save_strategy="epoch",
        learning_rate=2e-5,
        per_device_train_batch_size=4,
        per_device_eval_batch_size=4,
        num_train_epochs=3,
        weight_decay=0.01,
        load_best_model_at_end=True,
    )

    # Initialize trainer
    trainer = Trainer(
        model=model,
        args=training_args,
        train_dataset=train_dataset,
        eval_dataset=val_dataset,
        compute_metrics=compute_metrics,
    )

    # Train the model
    info("Starting model training...")
    trainer.train()

    # Save the model
    trainer.save_model(model_save_path)
    tokenizer.save_pretrained(model_save_path)

    # Evaluate the model
    eval_results = trainer.evaluate()
    info(f"Evaluation results: {eval_results}")

    return model, tokenizer, label_encoder

def predict_classifications(
    data_dict: Dict[str, Dict[str, Any]],
    model_path: str = "pdf_classifier"
) -> Dict[str, Dict[str, Any]]:
    """
    Predict classifications for new PDFs

    Args:
        data_dict: Dictionary with material_ids as keys and dicts containing 'file_path' as values
        model_path: Path to the saved model

    Returns:
        Dictionary with predictions and confidence scores added
    """
    # Load model, tokenizer, and label encoder
    model = AutoModelForSequenceClassification.from_pretrained(model_path)
    tokenizer = AutoTokenizer.from_pretrained(model_path)

    # Process each PDF
    result_dict = {}
    for material_id, item in data_dict.items():
        # Extract text from PDF
        text = extract_text_from_pdf(item['file_path'])
        if not text:
            warn(f"Could not extract text from PDF for material_id {material_id}")
            continue

        # Tokenize
        inputs = tokenizer(
            text,
            truncation=True,
            max_length=512,
            padding='max_length',
            return_tensors='pt'
        )

        # Get prediction
        with torch.no_grad():
            outputs = model(**inputs)
            probabilities = torch.nn.functional.softmax(outputs.logits, dim=-1)

        # Get predicted class and confidence
        predicted_class_idx = torch.argmax(probabilities).item()
        confidence = probabilities[0][predicted_class_idx].item()
        predicted_class = model.config.id2label[predicted_class_idx]

        # Store results
        result_dict[material_id] = {
            **item,
            'predicted_classification': predicted_class,
            'confidence': confidence
        }

    return result_dict

def main():
    """Example usage of the training and prediction functions"""
    # Get training data
    training_data = get_manual_classification_for_files()
    if not training_data:
        warn("No training data available")
        return

    # Train model
    try:
        model, tokenizer, label_encoder = train_model(training_data)
        cool("Model training completed successfully")

        # Example prediction
        new_data = {
            'test_id': {
                'file_path': Path('path/to/test.pdf')
            }
        }
        predictions = predict_classifications(new_data)
        info(f"Predictions: {predictions}")

    except Exception as e:
        warn(f"Error during model training: {e}")

if __name__ == "__main__":
    main()
