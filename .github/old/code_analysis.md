# Code Analysis and Refactoring Plan

## Current Data Flow Analysis

The core of the application is a data processing pipeline that ingests data from various sources, enriches it, and generates reports. The current data flow is complex and involves multiple, sometimes circular, steps.

### High-Level Data Flow

The process is orchestrated by the `EasyAccessTool` class in `easy_access/main.py`. A typical run involves the following stages:

1.  **Initial Ingestion**: A raw data export from an external tool (in Excel format) is read and ingested into a SQLite database.
2.  **Synchronization from Sheets**: The application reads multiple Excel workbooks that have been manually edited by users. It then attempts to merge these changes into the database.
3.  **Enrichment**: The data is enriched with information scraped from external web sources (e.g., Osiris). This process involves creating intermediate JSON files.
4.  **Processing**: The application performs calculations on the data (e.g., calculating potential fines).
5.  **Report Generation**: The application generates a new set of Excel sheets (overviews, weekly reports) based on the processed data.

###  Visualization of Current Data Flow

```mermaid
graph TD
    A[Raw Excel Export] --> B{Ingest into DB};
    C[Faculty Excel Sheets] --> D{Merge into DB};
    B --> E[Load from DB to DataFrame];
    D --> E;
    E --> F{Enrichment (via JSON)};
    F --> G{Update DB};
    G --> H[Load from DB to DataFrame];
    H --> I{Calculations};
    I --> J{Update DB};
    J --> K[Load from DB to DataFrame];
    K --> L[Generate Excel Reports];
```

### Key Issues

-   **No Single Source of Truth**: The application treats the raw Excel export, the faculty sheets, and the database as sources of truth at different times. This creates ambiguity and requires complex, error-prone logic to resolve conflicts.
-   **Circular and Inefficient Data Flow**: Data is repeatedly read from and written to the database and file system. For example, data is loaded from the DB into a DataFrame, processed, and then immediately written back. This is inefficient.
-   **Implicit Side Effects**: Functions often have side effects that are not obvious from their names. For instance, a function named `create_overviews` also modifies the database.
-   **Complex and Brittle Update Logic**: The logic for merging data from different sources is spread across multiple files (`main.py`, `db/update.py`) and is difficult to understand and maintain.

## Proposed Refactoring Plan

The proposed refactoring aims to establish a clear, unidirectional data flow with the database as the single source of truth.

### Guiding Principles

-   **The Database is the Single Source of Truth**: All data is stored in the SQLite database. Excel files are treated as either inputs for ingestion or outputs for reporting.
-   **Unidirectional Data Flow**: Data flows in one direction: `Ingestion -> Processing -> Export`.
-   **Clear Separation of Concerns**: The logic for ingestion, processing, and exporting data will be separated into distinct, well-defined modules.

### Proposed New Data Flow

```mermaid
graph TD
    # Code Analysis (moved)

    The detailed analysis and plan were split into three files to improve maintainability:

    - `.github/analysis.md` — background, design, and recommendations.
    - `.github/todo.md` — canonical to-do checklist and prioritized tasks.
    - `.github/changelog.md` — high-level change log for the analysis files.

    Please edit the files above for further updates.
## Review summary
