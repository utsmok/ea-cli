# Instructions for LLM-Powered Django Application Development

### General Python coding instructions

# Use uv!
Use uv for all python env management:
- uv run code.py to run a file instead of python code.py
- uv add library to install libraries instead of pip install library

# libraries
- use libraries as described in this document
- add ones you think are required, prefer libraries mentioned in this document


# Docstrings and type hints
- Write clear and concise comments for each function.
- Ensure functions have descriptive names and include type hints.
- Provide docstrings following PEP 257 conventions.
- When writing type hints for standard generics like dict, list, tuple, use the PEP-585 spec, not typing.Dict, typing.List, etc.
- When generating union types, use the union operator, | , not the typing.Union type
- Where possible, prefer duck-typing tests than isinstance, e.g. hasattr(x, attr) not isinstance(x, SpecificClass)

Example function with docstrings and type hints:

```python
def calculate_area(radius: float) -> float:
    """
    Calculate the area of a circle given the radius.

    Parameters:
    radius (float): The radius of the circle.

    Returns:
    float: The area of the circle, calculated as π * radius^2.
    """
    import math
    return math.pi * radius ** 2
```

# Usage of standard libraries
- When merging dictionaries, use the union operator
- Prefer pathlib over os.path for operations like path joining
- When using open() in text-mode, explicitly set encoding to utf-8
- Use the builtin methods in the itertools module for common tasks on iterables rather than creating code to achieve the same result
- Prefer f-strings for formatting strings rather than .format or % formatting
- Use modern Python 3.12+ syntax

# Explicitly required libraries for this project
- Use polars for dataframes.
    - do not use pandas
- use sqlalchemy2 for sql operations
- use msgspec for data serialization & validation
- use alembic for migrations
- use pymupdf / pymupdf4llm for pdf text & metadata extraction
    - as fallback (if ocr is required) use Kreuzberg
    - for rendering (if applicable) maybe use pypdfium2
- use xlsxwriter for excel files (in combination with polars where possible)
- Use httpx over requests or other http libs.
- use selenium and/or bs4 for additional web scraping operations
- use loguru for logging.
- use rich for printing.
- use xxhash for hashing
- use qdrant w/ fastembed for embeddings
- use gemini 2.0 through google api as main llm


# Overview of recommended python libraries/modules



[terminal/cli/print]
rich 			 rich text printing/display
				 https://rich.readthedocs.io/
plotext			 terminal plots
				 https://github.com/piccolomo/plotext
				 compatible with rich -> https://github.com/piccolomo/plotext/blob/master/readme/environments.md#rich
textual			 tui / webapp
				 https://textual.textualize.io/
tabulate		 quick & easy table printer
				 https://pypi.org/project/tabulate/
typer			 extensive cli
				https://typer.tiangolo.com/
fire			 very simple cli
				 https://github.com/google/python-fire
questionary		 ask for input
				 https://questionary.readthedocs.io/en/stable/
				 alt: typer prompts -> https://typer.tiangolo.com/tutorial/options/prompt/
				 alt: rich.prompt -> https://rich.readthedocs.io/en/stable/prompt.html
[web/api/scrape/get]
httpx			 http request/get/post/etc, also async, replacese requests
				 https://www.python-httpx.org/
bs4				 webscrape tool
				 https://www.crummy.com/software/BeautifulSoup/bs4/doc/#
litestar		 lightweight asgi api, replace fastapi or django
				 works with orm, htmx, sqlalchemy, and more
				 https://litestar.dev/
litestar-fullstack		 example/boilerplate setup to get litestar running
				 https://docs.fullstack.litestar.dev/latest/
				 https://github.com/litestar-org/litestar-fullstack
granian				 asgi/wsgi server
				https://github.com/emmett-framework/granian/tree/master
[non-text files]
pypdfium2			 fast & high quality pdf text extraction
				 can also render pdfs as imgs!
				 https://pypi.org/project/pypdfium2/
pymupdf				 fast and extensive pdf read/write lib
				 https://pymupdf.readthedocs.io/en/latest/index.html
pymupdf4llm			 pdf extraction specifically for llm/rag
				 https://pymupdf.readthedocs.io/en/latest/pymupdf4llm/index.html
pikepdf				 pdf manipulation
				 https://pikepdf.readthedocs.io/en/latest/
Kreuzberg			 parse all kinds of files using ocr if necessary -- read the docs!
				 https://github.com/Goldziher/kreuzberg
python-fsutil 			 file utils
				 https://github.com/fabiocaccamo/python-fsutil
xlsxwriter			 fast excel parsing/create
				 https://xlsxwriter.readthedocs.io/
				 works nicely with polars -> https://xlsxwriter.readthedocs.io/working_with_polars.html
				 conditional formatting -> https://xlsxwriter.readthedocs.io/example_conditional_format.html
openpyxl			 slower excel parse/create with more options (?)
				 http://openpyxl.readthedocs.io/
pyexcel				 simple api for quick reading of csv, xls(x), etc
				 https://github.com/pyexcel/pyexcel
[data-frames & -bases]
polars				 dataframes, faster and better api than pandas
				 https://docs.pola.rs/
dataset				 super-simplified wrapper around sqlalchemy for quick db storage (no relations!)
				 https://dataset.readthedocs.io/en/latest/quickstart.html
redis-dict			 simplify working with redis as dictstore / cache
				 https://pypi.org/project/redis-dict/
pycaret				 low code machine learning toolkit
				 https://pycaret.gitbook.io/docs
				 - example for classification https://nbviewer.org/github/pycaret/pycaret/blob/master/tutorials/Tutorial%20-%20Multiclass%20Classification.ipynb
alembic				 db migrations w/ sqlalchemy
				 https://alembic.sqlalchemy.org/en/latest/
sqlalchemy			 db engine + orm
				 https://www.sqlalchemy.org/
pygwalker			 explore dataframes and/or sql dbs interactively
				 can integrate with django, flask, streamlit, any html framework really
				 https://docs.kanaries.net/pygwalker/api-reference/dataset

[serialize/validate/parse]
srsly				 performant serialize/unserialize library
				 handles yaml, json, jsonl, msgpack, pickle4
			 https://github.com/explosion/srsly
msgspec				 -extremely- fast and simple validation and serialization
				 https://jcristharif.com/msgspec
lxml				 xml parsing
				 https://lxml.de/
nameparser			 parse strs into human name parts / initials / etc
				 https://github.com/derek73/python-nameparser
[graphics]
mss				 screenshotting library - very fast
				 https://python-mss.readthedocs.io/examples.html
vtracer			 bitmap to vector tool
				 https://github.com/visioncortex/vtracer
pychartjs		 create chart.js html code in python
				 https://github.com/tavallaie/pychartjs
[dicts & utils]
python-benedict[all]		 better dicts
				 https://github.com/fabiocaccamo/python-benedict
glom			 declarative deep-depth data restructing using native python dict
				 https://glom.readthedocs.io/en/latest/index.html
boltons			 grab-bag of bolt-on python utils
				 https://boltons.readthedocs.io/en/latest/index.html
dpath			 access dicts by path (/)
				 https://pypi.org/project/dpath/
[functional / types]
toolz			 functional programming toolkit
				 https://toolz.readthedocs.io/
returns			type safe return values replacing None (e.g. Maybe, Future, etc)
				 https://returns.readthedocs.io/en/latest/pages/quickstart.html
pipe			 functions with piped syntax (fun1 | fun2 instead of fun2(fun1))
				 https://github.com/JulienPalard/Pipe
piped functions example		blog post with extended example of piped funcs in python
				 https://pybit.es/articles/a-practical-example-of-the-pipeline-pattern-in-python/
[machine learning / llm / embedding]
GLiClass		 zero-shot classifier based on GLiNER (https://github.com/urchade/GLiNER/tree/main)
				https://github.com/Knowledgator/GLiClass
modernbert info			https://medium.com/@simeon.emanuilov/modernbert-a-modernized-bert-for-nlp-tasks-unfoldai-0ca2c66bc669
sqlite-vec			vector store extension for sqlite
				https://github.com/asg017/sqlite-vec
fastembed			lightweight embedding gen
			 https://qdrant.github.io/fastembed/
qdrant				 performant local vector store
				 https://github.com/qdrant/qdrant


# General coding guidelines
- When creating dummy data, don't use "Foo" and "Bar", be more creative
- When asked to create a function, class, or other piece of standalone code, don't append example calls unless otherwise told to
- Break down complex functions into smaller, more manageable functions.
- Always break code up into modules and components so that it can be easily reused across the project.
- Always prioritize readability and clarity.
- For algorithm-related code, include explanations of the approach used.
- Write code with good maintainability practices, including comments on why c certain design decisions were made.
- Handle edge cases and write clear exception handling.
- For libraries or external dependencies, mention their usage and purpose in comments.
- Use consistent naming conventions and follow language-specific best practices.
- Write concise, efficient, and idiomatic code that is also easily understandable.






## Model Design

- **Use abstract base models**: Inherit from shared base classes for common fields
- **Keep models focused**: Each model should represent a concept with clear responsibilities
- **Use descriptive field names**: Names should clearly indicate the purpose of each field
- **Set appropriate field constraints**: Add validators, choices, and help text
- **Implement proper indexes**: Add database indexes for frequently queried fields
- **Use meaningful relationships**: Define appropriate ForeignKey/ManyToMany relationships with descriptive related_name
- **Add string representation**: Always implement `__str__` methods
- **Group model choices as classes**: Use enums for enum-like fields
- **Keep meta options organized**
- **Document model purpose**: Add docstrings explaining the model's purpose and key fields

## Existing Application Analysis

### Core Functionality
The existing `easy_access` application is a Python CLI tool that processes and manages copyright data for educational materials. Its primary functions include:

- **Copyright Data Management**: Processing raw copyright exports, creating organized sheets by faculty and program
- **Document Classification**: Classification of materials using LLM APIs
- **PDF Processing**: Handling and analysis of PDF documents
- **Data Enrichment**: Integration with Osiris (academic system) to enrich copyright data
- **Reporting**: Creating overview sheets, export sheets, and faculty-specific reports
- **Database Operations**: Storage, retrieval, and updating of copyright information

### Key Components

1. **EasyAccessTool**: Main class orchestrating all operations with methods for:
 - Processing copyright data
 - Creating faculty sheets, program sheets, and overview sheets
 - Updating database from faculty sheets
 - Data validation and cleaning
 - Export sheet generation

2. **File System**: Custom `Directory` and `File` classes managing file operations:
 - File manipulation (copy, delete, move)
 - Directory creation and management
 - Path handling with consistent interfaces

3. **Classification System**:
 - Integration with external classification APIs
 - Models for different classifications (open access, copyrighted, etc.)
 - PDF handling and processing

4. **Database Layer**:
 - Models for storing copyright data
 - Functions for data retrieval, updates, and ingestion
 - Asynchronous database operations

### Data Models and Relations

This is a description of the current models in tortoiseorm.
All models have a mix-in that stores the creation datetime & last modified datetime
Use enums where possible instead of bare charfields. See current app for details.

**CopyrightItem**
- the primary class of the app
- stores data for a file found on the teaching environment
- initially directly imported from another system
- then enriched and linked to other models in this db
- also stores workflow status (todo/inprogress/done), manual classification data, remarks, is_duplicate, etc
- numerous relations: courses (manytomany), faculty (organization) (fk), llmclassification (1to1), replace_with (fk), itemupdates (manyto1)

**ItemUpdate**
- stores changes to CopyrightItems
- reverse relation to copyrightitem
- single field with a json storing the changes. Use this format: {changed_field_1: {'old': old_value, 'new':new_value}, changed_field_2: ... }

**Course**:
- data for a course: name, ids, ecs, category, remarks, etc
- has m2m to person using a through-table that also includes the role of the person (teacher /examinator/contact/...)
- fk to faculty (primary faculty that hosts this course)
- m2m to programme (programmes that include this course in their curriculum)

**Person**:
- data for an employee
- fk to a faculty (organization), the primary faculty that employs this person
- m2m to organizations for all orgs related to this person (multiple hierarchy levels)
- reverse m2m relation to courses (where this employee is a teacher /examinator/contact/...)

**LLMClassification**:
- classification data from LLM for a PDF
- one-to-one relation to a PDF file
    - which in turn should have a 1-to-1 to a copyright item
    - maybe do a direct 1-to-1 to the copyright item instead?
- stores data from a structured response

**Organization**:
- Organizational units, like faculties, research groups, departments, etc.
- contains metadata like abbreviation, name
- hierarchical, has fk to the model itself (parent_org)
- hierachy_level (int) stores how much parent levels are above this one, starting at 0 for the university itself
- to be decided: make inheriting models to specifiy specific org types like faculties etc?

**Programme**:
- Describes a programme (educational programme; 'department' in raw copyrightdata)
- fk to faculty (parent faculty)
- contains metadata for the programme
- each CopyrightItem should have a fk to a programme (based on the 'department' field)

 **PDF**:
- Linked to a CopyrightItem (nullable OneToOne)
- File itself downloaded from Canvas using the url in CopyrightItem
- Model stores metadata & extracted text for each PDF file.

### Input/Output Files

**Input**:
- Raw copyright data exports (Excel files)
- Faculty sheets with manual updates
- Osiris data for enrichment
- PDF files for classification

**Output**:
- Faculty-specific Excel sheets
- Program-specific Excel sheets
- All-items Excel sheets
- Overview reports and analyses
- Export sheets for processed items
- Backup files

### Data Flow
1. Import raw copyright data
2. Clean and validate data
3. Update with faculty sheet information
4. Enrich with Osiris data
5. Process classifications
6. Generate output sheets by faculty/program
7. Create overview reports
8. Manage backups

### Key Features
- Operations for database and external APIs
- Structured logging system with different levels
- File backup and restoration capabilities
- Data validation and cleaning
- Configurable settings via YAML files
- PDF processing and classification
- Excel sheet generation with formatting
