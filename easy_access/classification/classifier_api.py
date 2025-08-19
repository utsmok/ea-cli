"""
This module provides functionalities to classify PDF documents using an external
generative AI API (specifically Google's Gemini). It handles:
- Sending PDF content (either extracted text or full PDF bytes) to the API.
- Constructing appropriate prompts for classification.
- Parsing the structured JSON response from the API into Pydantic models.
- Managing API client initialization and concurrent requests with rate limiting.
- Storing classification results.
"""

import asyncio
import io
import logging  # Added
import os  # For os.path.exists, to be replaced
import time
from functools import partial
from typing import Any

import pikepdf  # For PDF manipulation (splitting pages)
from aiometer import amap  # For rate-limited concurrent async calls
from google import genai  # External dependency for Gemini API

# from rich.console import Console # Removed, using logging
from easy_access.classification.classifier_models import (
    Classification,  # Pydantic model for response
)
from easy_access.db.base import init as init_tortoise_orm
from easy_access.db.ingest import load_llm_classifications  # To save results to DB
from easy_access.db.models import PDF, CopyrightItem  # ORM Models
from easy_access.settings import SETTINGS, DirSetting

# from easy_access.utils import File, warn # File not used directly, warn replaced by logger

logger = logging.getLogger(__name__)

# --- Configuration Candidates (Critical to move to settings.yaml or env vars) ---
GEMINI_API_KEY_CONFIG_NAME: str = "GEMINI_API_KEY"  # Name of env var or settings key
GEMINI_MODEL_NAME: str = "models/gemini-1.5-flash-latest"  # Updated model name based on common Gemini names (verify actual)

# Prompt - this is very large and central to the logic.
# Ideally, load from a template file or a more structured config.
CLASSIFICATION_PROMPT: str = """From the included document, first extract and determine a list of metadata, then determine the copyright status and item type for this item.
Finally determine the most important classification: if the item is allowed to be shared with students in the context of the University of Twente learning environment.
Use all available (meta)data in the file or that you extracted earlier (e.g. the author name, publisher name, and copyright holder name, license statements, etc.) to help determine these statuses.
The copyright status should be focused on the overall document. You can ignore any possible copyrighted elements included inside the work (like images from other works).
For determining allowed use, take into account that the works are being shared internally at the University of Twente, a Dutch public institute, for educational purposes only, in a closed environment.
This means that clearly copyrighted commercial works cannot be used, except if educational use is explicitly allowed for instance.
There will never be any commercial use in this context. Assume attribution is always given.
If the detected 'publisher' or 'author' is the University of Twente or is employed by the University of Twente, the work should be classified as OWN_MATERIAL.
Include reasoning for the classification in the response in the corresponding fields.

The requested output format is replicated here as a set of Python classes, including additional details, hints, and suggestions.
    class ItemType(str, Enum):
        Possible item types of the item.
        PRESENTATION = "presentation" # a powerpoint in pdf format for example. By definition, this should have CopyrightStatus.OWN_MATERIAL.
        READER = "reader" # often self-written information by teachers for students for this specific course. By definition, this should have CopyrightStatus.OWN_MATERIAL.
        BOOK = "book" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
        ARTICLE = "article" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
        REPORT = "report" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
        ASSIGNMENT = "assignment" # an assignment description for this course. By definition, this should have CopyrightStatus.OWN_MATERIAL
        THESIS = "thesis" # Often COPYRIGHTED_MATERIAL or OPEN_ACCESS.
        MANUAL = "manual" # e.g. for a measuring device. Often COPYRIGHTED_MATERIAL or OPEN_ACCESS, but can be OWN_MATERIAL.
        UNKNOWN = "unknown" # if not possible to determine.
    class CopyrightStatus(str, Enum):
        The possible copyright classifications
        OPEN_ACCESS = "open access" # free to use
        OWN_MATERIAL = "own material" # made for or by an employee of the university of Twente
        COPYRIGHTED_MATERIAL = "copyrighted material" # not free to use, owned by a publisher for instance
        OTHER = "other" # should not be used? maybe if unable to classify otherwise.
    class AllowedUsageByUT(str, Enum):
        These possible classifications denote if the item is allowed to be shared with students in the context of the University of Twente learning environment.
        ALLOWED = "allowed" # the item can be shared with students without further limitations, e.g. it is open access or own material by a UT employee.
        RESTRICTED = "restricted" # the item has limitations on sharing; e.g. only this year, only if the uploader is the author, or only with specific permissions/acknowledgements etc.
        NOT_ALLOWED = "not allowed" # the item cannot be shared without further permissions, e.g. it is fully copyrighted without any other routes to obtain permission
        UNDETERMINED = "undetermined" # the item cannot be classified as allowed or not allowed, e.g. if the classification is not possible due to missing or conflicting information.
    class Classification(BaseModel):
        allowed_usage: AllowedUsageByUT = AllowedUsageByUT.UNDETERMINED
        allowed_usage_reasoning: str # add a 1 to 2 sentence explanation on why this allowed usage was chosen
        copyright_status: CopyrightStatus = CopyrightStatus.OTHER
        copyright_classification_reason: str  # add a 1 to 2 sentence explanation on why this copyright status was chosen
        item_type: ItemType = ItemType.UNKNOWN
        item_type_classification_reason: str # add a 1 to 2 sentence explanation on why this item type was chosen
        pdf_name: str

        # metadata fields -- if not possible to determine from the pdf store an empty string instead
        author_name: list[str] # the name of the author(s) that created the item, if possible to determine
        publisher_name: str # who published the item, if possible to determine
        copyright_holder: str # who holds the copyright, if possible to determine
        item_title: str # the title of the item, if possible to determine
        doi: list[str] # the DOI(s) for the item if included in the document itself
        isbn: list[str]  # the ISBN(s) for the item if included in the document itself
        source_url: list[str] # the source URL(s) for the item if included in the document itself
        license: list[str] # the license(s) for the item if included in the document itself, or any license-related statement like 'all rights reserved', 'creative commons', 'Reproduction is allowed with acknowledgement'.
        topic: str # the topic of the item, what it covers
        pdf_page_count: int # the amount of pages in the pdf
        remarks: str # any additional remarks on the item relevant to copyright status, metadata, and item type
"""

# API Client - initialized by activate_client()
gemini_client: genai.GenerativeModel | None = (
    None  # Changed from genai.Client to GenerativeModel
)

# --- End Configuration Candidates ---


def activate_client() -> bool:
    """
    Initializes the Google Gemini API client using an API key.
    The API key is expected to be stored in the environment variable GEMINI_API_KEY.
    If the key is not found or client initialization fails, an error is logged.

    Returns:
        bool: True if client was successfully initialized, False otherwise.
    """
    global gemini_client
    api_key = os.environ.get(GEMINI_API_KEY_CONFIG_NAME)
    if not api_key:
        # Fallback to settings file if not in env (less secure, for dev only)
        # api_key = SETTINGS.get("GEMINI_API_KEY") # Assuming SETTINGS can store it
        # For now, strictly from env:
        logger.error(
            f"API key for Gemini not found in environment variable '{GEMINI_API_KEY_CONFIG_NAME}'."
        )
        logger.error(
            "Please set this environment variable to use the classification API."
        )
        return False

    try:
        # genai.configure(api_key=api_key) # General configuration
        # self.client = genai.GenerativeModel(model_name="gemini-pro-vision")
        gemini_client = genai.GenerativeModel(
            model_name=GEMINI_MODEL_NAME
        )  # Initialize with specific model
        # Test with a simple call if possible, or assume success if object created.
        # For now, assume client object creation means success.
        logger.info(
            f"Google Gemini API client initialized successfully with model {GEMINI_MODEL_NAME}."
        )
        return True
    except Exception as e:
        logger.error(f"Failed to initialize Google Gemini API client: {e}")
        gemini_client = None
        return False


async def _send_pdf_to_gemini_get_file_object(
    pdf: PDF, max_pages: int = 20
) -> genai.types.File | None:
    """
    Uploads the first `max_pages` of a PDF file to the Gemini API's temporary file store
    and returns a `genai.types.File` object representing the uploaded file.

    Args:
        pdf (PDF): The PDF ORM object containing the path to the PDF file.
        max_pages (int): The maximum number of pages from the PDF to upload.

    Returns:
        Optional[genai.types.File]: A Gemini File object if upload is successful, else None.
    """
    if not gemini_client:
        logger.error("Gemini client not activated. Cannot send PDF.")
        return None
    if not pdf.path.exists():
        logger.warning(f"PDF file does not exist, cannot send to Gemini: {pdf.path}")
        return None

    pdf_bytes_to_upload: bytes | None = None
    try:
        with pikepdf.open(pdf.path) as opened_pdf_doc:
            # Create a new PDF in memory with the first 'max_pages'
            temp_pdf_in_memory = pikepdf.Pdf.new()
            for i in range(min(max_pages, len(opened_pdf_doc.pages))):
                temp_pdf_in_memory.pages.append(opened_pdf_doc.pages[i])

            temp_stream = io.BytesIO()
            temp_pdf_in_memory.save(temp_stream)
            pdf_bytes_to_upload = temp_stream.getvalue()

    except Exception as e_pikepdf:
        logger.error(
            f"Error processing PDF {pdf.path} with pikepdf before upload: {e_pikepdf}"
        )
        return None

    if not pdf_bytes_to_upload:
        logger.warning(f"No bytes generated from PDF {pdf.path} for upload.")
        return None

    try:
        logger.debug(
            f"Uploading {pdf.current_file_name} (first {max_pages} pages) to Gemini file store."
        )
        # Use a unique name for the file on Gemini, e.g., material_id
        gemini_file_name = f"pdf_{pdf.material_id}_{int(time.time())}"
        # The client.files.upload returns a File object
        uploaded_file_response = (
            gemini_client.client.files.upload(  # Access underlying client for files API
                file=io.BytesIO(pdf_bytes_to_upload),
                config={"mime_type": "application/pdf", "name": gemini_file_name},
            )
        )
        logger.info(
            f"PDF {pdf.current_file_name} uploaded to Gemini as '{uploaded_file_response.name}'."
        )
        return uploaded_file_response
    except Exception as e_upload:
        logger.error(
            f"Error uploading PDF {pdf.current_file_name} to Gemini: {e_upload}"
        )
        return None


async def classify_pdf(
    pdf: PDF, use_full_pdf_upload: bool = False, max_pages_for_upload: int = 20
) -> tuple[PDF, Classification | None]:
    """
    Classifies a single PDF document using the Gemini API.

    It can either send extracted text (if available and not `use_full_pdf_upload`)
    or upload the first few pages of the PDF directly to the API.

    Args:
        pdf (PDF): The PDF ORM object to classify.
        use_full_pdf_upload (bool): If True, uploads the PDF file directly.
                                    Otherwise, sends extracted text.
        max_pages_for_upload (int): Max pages to upload if `use_full_pdf_upload` is True.

    Returns:
        Tuple[PDF, Optional[Classification]]: The original PDF object and the parsed
                                              Classification object, or None if classification fails.
    """
    if not gemini_client:
        logger.error("Gemini client not activated. Cannot classify PDF.")
        return pdf, None  # Return original PDF and None for classification

    contents_for_api: Any = None  # Can be str (text) or list (file + text prompt)
    material_id_str = str(pdf.material_id)  # For logging and potential use as file ID

    # Determine content to send: extracted text or full PDF
    if (
        use_full_pdf_upload or not pdf.extracted_text or len(pdf.extracted_text) < 100
    ):  # Min length for useful text
        if not use_full_pdf_upload:
            logger.info(
                f"Text for {pdf.current_file_name} is too short or missing. Attempting full PDF upload."
            )

        gemini_file_obj = await _send_pdf_to_gemini_get_file_object(
            pdf, max_pages=max_pages_for_upload
        )
        if not gemini_file_obj:
            logger.warning(
                f"Failed to upload PDF {pdf.current_file_name} to Gemini. Cannot classify."
            )
            return pdf, None
        # Construct content list: [File object, prompt text]
        contents_for_api = [
            gemini_file_obj,  # The File object from Gemini API
            f"You received the pdf file {pdf.current_file_name}.\n{CLASSIFICATION_PROMPT}",
        ]
    else:  # Use extracted text
        pdf_text_to_send = pdf.extracted_text
        # Truncate text if too long for API limits (adjust as needed)
        # This limit should ideally be based on token count, not char count.
        # TODO: Add token counting and chunking if text often exceeds limits.
        max_text_chars = (
            150000  # Example limit, Gemini might have higher/lower or token based
        )
        if len(pdf_text_to_send) > max_text_chars:
            pdf_text_to_send = pdf_text_to_send[:max_text_chars]
            logger.debug(
                f"Truncated extracted text for {pdf.current_file_name} to {max_text_chars} chars."
            )

        contents_for_api = f"{CLASSIFICATION_PROMPT}\n\n| text content of pdf file {pdf.current_file_name} is as follows: |\n{pdf_text_to_send}"

    logger.info(
        f"Sending classification request for PDF {material_id_str} ({'full PDF' if isinstance(contents_for_api, list) else 'extracted text'})..."
    )

    parsed_classification: Classification | None = None
    api_error_reason: str | None = None
    try:
        # Assuming gemini_client is GenerativeModel instance
        response = await asyncio.to_thread(  # Run blocking SDK call in thread
            gemini_client.generate_content,
            contents=contents_for_api,
            generation_config=genai.types.GenerationConfig(
                response_mime_type="application/json",
                response_schema=Classification,  # Pass Pydantic model as schema
            ),
        )

        # Accessing parsed content depends on how the Gemini SDK structures it with schema
        # This might be response.candidates[0].content.parts[0].text (if schema is not auto-parsed)
        # Or, if SDK parses it directly into the Pydantic model:
        if (
            response.candidates
            and response.candidates[0].content
            and response.candidates[0].content.parts
        ):
            # Assuming the first part contains the JSON string if not auto-parsed by SDK
            json_text_from_response = response.candidates[0].content.parts[0].text
            # Manually parse into Pydantic model
            parsed_classification = Classification.model_validate_json(
                json_text_from_response
            )

        # If the SDK has a direct way to get the Pydantic model (e.g. response.parsed_as(Classification)) use that.
        # The example `response.parsed` is not standard for this SDK. Let's assume manual parsing for now.

        if parsed_classification:
            parsed_classification.pdf_name = (
                pdf.current_file_name
            )  # Add pdf_name from our context
            logger.info(
                f"Successfully parsed classification response for {material_id_str}."
            )
        else:  # No valid classification parsed
            api_error_reason = "Response parsing failed or no content."
            if response.candidates and response.candidates[0].finish_reason:
                api_error_reason = response.candidates[
                    0
                ].finish_reason.name  # Get enum name string
            logger.warning(
                f"No classification result for {material_id_str}. Finish reason: {api_error_reason}. Response: {response.text if not response.candidates else 'See prompt feedback'}"
            )
            if (
                response.prompt_feedback and response.prompt_feedback.block_reason
            ):  # Check for safety blocks
                logger.error(
                    f"Classification for {material_id_str} blocked. Reason: {response.prompt_feedback.block_reason_message}"
                )

    except Exception as e_classify:
        logger.error(
            f"Error during Gemini API call for {pdf.current_file_name}: {e_classify}"
        )
        logger.debug(traceback.format_exc())
        api_error_reason = str(e_classify)  # Store error as reason
    finally:
        # Clean up uploaded file from Gemini if it was uploaded
        if isinstance(contents_for_api, list) and isinstance(
            contents_for_api[0], genai.types.File
        ):
            try:
                # Access underlying client for files API
                gemini_client.client.files.delete(name=contents_for_api[0].name)
                logger.info(
                    f"Deleted uploaded file {contents_for_api[0].name} from Gemini file store."
                )
            except Exception as e_delete:
                logger.warning(
                    f"Failed to delete uploaded file {contents_for_api[0].name} from Gemini: {e_delete}"
                )

    return pdf, parsed_classification


async def _delete_all_gemini_files() -> (
    None
):  # Made private as it's a utility for this module
    """Deletes all files from the Gemini API file store. Use with caution."""
    if not gemini_client:
        logger.error("Gemini client not activated. Cannot delete files.")
        return

    logger.info("Attempting to delete all files from Gemini file store...")
    deleted_count = 0
    try:
        # Access underlying client for files API
        for f_gemini in gemini_client.client.files.list():  # type: ignore
            logger.debug(f"Deleting Gemini file: {f_gemini.name}")
            gemini_client.client.files.delete(name=f_gemini.name)  # type: ignore
            deleted_count += 1
        logger.info(f"Deleted {deleted_count} files from Gemini storage.")
    except Exception as e:
        logger.error(f"Error deleting files from Gemini storage: {e}")


async def classify_items_in_batch(
    pdf_list: list[PDF],
    use_full_pdf_upload_for_all: bool = False,
    max_pages_for_full_upload: int = 20,
) -> int:
    """
    Classifies a list of PDF items concurrently using the Gemini API.

    Manages concurrency and rate limiting using `aiometer.amap`.
    Results (Classification objects) are stored as JSON files.

    Args:
        pdf_list (List[PDF]): A list of PDF ORM objects to classify.
        use_full_pdf_upload_for_all (bool): If True, forces all PDFs to be uploaded fully
                                           instead of sending extracted text.
        max_pages_for_full_upload (int): Max pages to upload if full PDF upload is used.

    Returns:
        int: The number of items successfully classified and stored.
    """
    if not pdf_list:
        logger.info("No PDFs provided to classify_items_in_batch.")
        return 0

    # Rate limits - consider making these configurable
    max_concurrent_tasks: int = SETTINGS.get(
        "CLASSIFIER_MAX_CONCURRENT", 5
    )  # Example: Get from settings or default
    max_requests_per_second: int = SETTINGS.get("CLASSIFIER_MAX_RPS", 1)  # Example

    successful_classifications_count: int = 0

    # Use functools.partial to pass fixed arguments to classify_pdf
    classify_func_partial = partial(
        classify_pdf,
        use_full_pdf_upload=use_full_pdf_upload_for_all,
        max_pages_for_upload=max_pages_for_full_upload,
    )

    # aiometer.amap processes items concurrently with rate limiting
    async with amap(
        classify_func_partial,  # type: ignore # amap expects a coroutine function
        pdf_list,
        max_at_once=max_concurrent_tasks,
        max_per_second=max_requests_per_second,
    ) as classification_results:
        async for result_tuple in classification_results:
            # result_tuple should be (PDF, Optional[Classification])
            if isinstance(result_tuple, tuple) and len(result_tuple) == 2:
                pdf_obj, classification_obj = result_tuple
                if classification_obj:  # If classification was successful
                    logger.info(
                        f"Received classification for PDF: {pdf_obj.material_id} - {classification_obj.allowed_usage.value}"
                    )

                    # Store classification result to JSON file
                    classifications_dir = SETTINGS.dirs.get(DirSetting.CLASSIFICATIONS)
                    if not classifications_dir or not classifications_dir.exists:
                        logger.error(
                            "Classifications directory not configured or found. Cannot save result."
                        )
                        continue  # Skip saving this result if dir is bad

                    json_filename = (
                        f"{pdf_obj.material_id}.json"  # Use material_id for filename
                    )
                    json_filepath = classifications_dir.full / json_filename

                    # Handle existing file (e.g., rename to _old)
                    if json_filepath.exists():
                        old_json_filepath = (
                            classifications_dir.full / f"{pdf_obj.material_id}_old.json"
                        )
                        try:
                            json_filepath.rename(old_json_filepath)
                            logger.debug(
                                f"Renamed existing classification {json_filepath.name} to {old_json_filepath.name}"
                            )
                        except OSError as e_rename:
                            logger.warning(
                                f"Could not rename existing {json_filepath.name}: {e_rename}"
                            )

                    try:
                        with open(json_filepath, "w", encoding="utf-8") as f_json:
                            # Pydantic's model_dump_json is preferred for serialization
                            f_json.write(classification_obj.model_dump_json(indent=2))
                        logger.info(
                            f"Stored classification for {pdf_obj.material_id} to {json_filename}"
                        )
                        successful_classifications_count += 1
                    except OSError as e_save:
                        logger.error(
                            f"Could not store classification JSON for {pdf_obj.material_id} to {json_filename}: {e_save}"
                        )
                    except Exception as e_dump:  # Catch Pydantic errors or others
                        logger.error(
                            f"Error serializing classification for {pdf_obj.material_id}: {e_dump}"
                        )

                else:  # Classification was None (failed)
                    logger.warning(
                        f"Classification failed for PDF: {pdf_obj.material_id} ({pdf_obj.current_file_name})"
                    )
            else:  # Should not happen if classify_pdf returns correctly
                logger.error(
                    f"Unexpected result format from classify_pdf: {result_tuple}"
                )

    return successful_classifications_count


async def main_classification_flow(
    material_ids_subset: list[int] | None = None,
    force_reclassify_all_in_subset: bool = False,  # If True, reclassifies even if LLM data exists
    default_to_full_pdf_upload: bool = False,
    max_pages_for_upload: int = 20,
) -> None:
    """
    Main orchestration function for classifying PDF documents.

    Steps:
    1. Initializes database and Gemini API client.
    2. Fetches PDF records from the database:
        - If `material_ids_subset` is provided, only those PDFs.
        - Otherwise, fetches all PDFs that don't yet have an LLM classification linked.
    3. If `force_reclassify_all_in_subset` is True, all PDFs in the subset are processed
       regardless of existing classification.
    4. Batches PDFs and calls `classify_items_in_batch` to get classifications from Gemini.
    5. After classification, calls `load_llm_classifications` to load new JSON results into the DB
       and link them to CopyrightItems.
    6. Cleans up any temporary files uploaded to Gemini API storage.

    Args:
        material_ids_subset (Optional[List[int]]): A specific list of material IDs to process.
                                                If None, processes items needing classification.
        force_reclassify_all_in_subset (bool): If True and `material_ids_subset` is given,
                                             re-classifies these items even if they already have one.
        default_to_full_pdf_upload (bool): Forces all classifications in this run to use full PDF upload.
        max_pages_for_upload (int): Max pages to use if full PDF upload is chosen.
    """
    logger.info(
        f"Starting main classification flow. Subset: {'Provided' if material_ids_subset else 'All needing classification'}. Force reclassify: {force_reclassify_all_in_subset}"
    )

    if not activate_client():  # Ensure client is active before proceeding
        logger.error("Failed to activate Gemini client. Classification cannot proceed.")
        return

    await init_tortoise_orm()  # Ensure DB is ready

    pdfs_to_process_list: list[PDF]

    if material_ids_subset:
        logger.info(
            f"Processing specified subset of {len(material_ids_subset)} material IDs."
        )
        # Fetch all PDFs for the subset first
        all_pdfs_in_subset = await PDF.filter(material_id__in=material_ids_subset).all()
        if force_reclassify_all_in_subset:
            pdfs_to_process_list = all_pdfs_in_subset
            logger.info(
                f"Forcing re-classification for all {len(pdfs_to_process_list)} PDFs in the subset."
            )
        else:
            # Filter out those that already have an LLM classification linked via CopyrightItem
            # This requires checking CopyrightItem's llm_classification_id field.
            items_in_subset_with_class = await CopyrightItem.filter(
                material_id__in=material_ids_subset, llm_classification_id__isnull=False
            ).values_list("material_id", flat=True)

            ids_already_classified: Set[int] = set(items_in_subset_with_class)  # type: ignore
            pdfs_to_process_list = [
                pdf
                for pdf in all_pdfs_in_subset
                if pdf.material_id not in ids_already_classified
            ]
            logger.info(
                f"{len(pdfs_to_process_list)} PDFs from subset require classification (others already have one)."
            )
    else:  # No subset, process all items that don't have a classification yet
        items_needing_classification_ids = await CopyrightItem.filter(
            llm_classification_id__isnull=True
        ).values_list("material_id", flat=True)

        pdfs_to_process_list = await PDF.filter(
            material_id__in=list(items_needing_classification_ids)
        ).all()  # type: ignore
        logger.info(
            f"Found {len(pdfs_to_process_list)} PDFs in total that require LLM classification."
        )

    if not pdfs_to_process_list:
        logger.info(
            "No PDF files found requiring classification based on the criteria."
        )
    else:
        # Batch processing logic (example from original `main` function)
        batch_size_for_api: int = SETTINGS.get(
            "CLASSIFIER_API_BATCH_SIZE", 10
        )  # Example: get from settings or default
        api_call_delay_seconds: int = SETTINGS.get(
            "CLASSIFIER_API_DELAY_SECONDS", 120
        )  # Example

        current_batch: list[PDF] = []
        total_processed_count: int = 0

        for pdf_item in pdfs_to_process_list:
            current_batch.append(pdf_item)
            if len(current_batch) >= batch_size_for_api:
                batch_start_time = time.time()
                logger.info(
                    f"Processing batch of {len(current_batch)} PDFs for classification..."
                )
                await classify_items_in_batch(
                    current_batch,
                    use_full_pdf_upload_for_all=default_to_full_pdf_upload,
                    max_pages_for_full_upload=max_pages_for_upload,
                )
                total_processed_count += len(current_batch)
                logger.info(
                    f"Batch processed. Total processed so far: {total_processed_count}/{len(pdfs_to_process_list)}."
                )
                current_batch = []  # Reset for next batch

                # Rate limiting delay
                elapsed_time = time.time() - batch_start_time
                if elapsed_time < api_call_delay_seconds:
                    sleep_time = api_call_delay_seconds - elapsed_time
                    logger.info(
                        f"Sleeping for {sleep_time:.1f} seconds to respect rate limits."
                    )
                    await asyncio.sleep(sleep_time)

        if current_batch:  # Process any remaining items in the last batch
            logger.info(f"Processing final batch of {len(current_batch)} PDFs...")
            await classify_items_in_batch(
                current_batch,
                use_full_pdf_upload_for_all=default_to_full_pdf_upload,
                max_pages_for_full_upload=max_pages_for_upload,
            )
            total_processed_count += len(current_batch)
            logger.info(
                f"Final batch processed. Total processed: {total_processed_count}/{len(pdfs_to_process_list)}."
            )

        # After all classifications are done (JSON files created), load them into the DB
        logger.info(
            "Loading all new/updated LLM classification JSON files into the database..."
        )
        await (
            load_llm_classifications()
        )  # This function should handle its own Tortoise init/close

    # Clean up any files uploaded to Gemini, regardless of whether classification ran for all items
    await _delete_all_gemini_files()
    await Tortoise.close_connections()
    logger.info("Classification main flow finished.")


if __name__ == "__main__":  # pragma: no cover
    # Setup basic logging for direct script execution
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )

    # Example of how to run the main classification flow
    # Ensure DB is populated with PDFs first.
    # asyncio.run(main_classification_flow())

    # Example for a subset:
    # asyncio.run(main_classification_flow(material_ids_subset=[12345, 67890]))

    logger.info("classifier_api.py executed directly (likely for testing).")
