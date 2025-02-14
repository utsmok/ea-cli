from easy_access.pdf_parser import extract_text_from_pdfs_batch, PDFDataset, extract_text_from_pdf, get_manual_classification_for_files
from pathlib import Path
from typing import Any
import numpy as np
from easy_access.utils import info, warn, cool
import evaluate
from sklearn.model_selection import train_test_split
from transformers import AutoTokenizer, AutoModelForSequenceClassification, Trainer, TrainingArguments
import torch
from sklearn.preprocessing import LabelEncoder

def prepare_dataset(data_dict: dict[str, dict[str, Any]]) -> tuple[list[str], list[str]]:
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

def train_model(training_data: dict[str, dict[str, Any]], model_save_path: str = "pdf_classifier") -> tuple[Any, Any, LabelEncoder]:
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
    data_dict: dict[str, dict[str, Any]],
    model_path: str = "pdf_classifier"
) -> dict[str, dict[str, Any]]:
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



def run_classifier():
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
