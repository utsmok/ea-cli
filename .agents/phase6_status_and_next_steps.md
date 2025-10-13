# Phase 6 Status: Tortoise → SQLAlchemy Migration

**Date**: October 13, 2025
**Status**: 4 of 14 tasks completed (~29%)
**Remaining**: Core modules (relations, update, retrieve) + 5 smaller modules + cleanup

---

## ✅ COMPLETED MODULES

### 1. `easy_access/db/compat.py` - Compatibility Layer (✅ DONE)

**Purpose**: Provides drop-in replacements for Tortoise ORM methods using SQLAlchemy.

**Functions Implemented**:
- `bulk_create()` - Batch inserts with PostgreSQL upsert (ON CONFLICT)
- `bulk_update()` - Batch updates with primary key matching
- `get_or_create()` - Atomic get-or-create with race condition handling
- `update_or_create()` - Update existing or create new
- `get_by_id()`, `get_one()`, `get_or_none()` - Single object retrieval
- `filter_all(**kwargs)` - Query and return all matching objects
- `filter_values(*fields, **kwargs)` - Return specific fields as dicts
- `all_values(*fields)` - Get all objects as dicts
- `count(**kwargs)`, `exists(**kwargs)` - Aggregation helpers
- `delete_where(**kwargs)` - Bulk delete
- `create_instance(**kwargs)`, `save_instance(obj)` - CRUD helpers
- `transaction()` - Async context manager for transactions
- `_batched(iterable, n)` - Internal batching utility

**Key Patterns**:
```python
# Get or Create
obj, created = await get_or_create(
    Model,
    field=value,
    defaults={"other": "value"}
)

# Bulk Operations
await bulk_create(
    Model,
    rows=[{"id": 1, "name": "test"}],
    on_conflict=["id"],
    update_fields=["name"]
)

# Transactions
async with transaction() as session:
    # All operations here are atomic
    pass
```

**Status**: ✅ No further work needed. All lint errors resolved.

---

### 2. `easy_access/db/session.py` - Session Management (✅ DONE)

**Changes**:
- Modified `init_db()` to accept both `Settings` object and URL string
- Extracts PostgreSQL connection from `DATABASE_URL` environment variable
- Maintains async session factory pattern with `expire_on_commit=False`

**Usage**:
```python
from easy_access.db.session import init_db, get_session

# Initialize at app startup
init_db(settings)  # or init_db("postgresql+asyncpg://...")

# Use in code
async for session in get_session():
    async with session.begin():
        await session.execute(...)
```

**Status**: ✅ Complete and tested.

---

### 3. `easy_access/db/base.py` - Database Initialization (✅ MOSTLY DONE)

**Changes Made**:
- ❌ Removed: `Tortoise.init()`, `Tortoise.generate_schemas()`, `Tortoise.close_connections()`
- ✅ Added: `init_db(settings)`, `Base.metadata.create_all()`, `shutdown_db()`
- ✅ Updated: `init_faculties()` to use `get_or_create()` compat function
- ✅ Updated: `init()` to use `Base.metadata.create_all()` via engine.begin()

**Remaining Work**:
- ⚠️ `copyright_item_from_dict()` function still references Tortoise `Model` type
- ⚠️ Minor type errors in that function (None-checking for datetime conversions)

**Note**: This function may not be actively used anymore - check call sites before investing time.

**Status**: ✅ Core functionality complete. Minor cleanup needed.

---

### 4. `easy_access/db/ingest.py` - Data Ingestion (✅ DONE)

**Conversions**:
- ❌ `Model.get_or_create()` → ✅ `get_or_create(Model, ...)`
- ❌ `Model.bulk_create()` → ✅ `bulk_create(Model, rows=...)`
- ❌ `Model.all().values()` → ✅ `all_values(Model, *fields)`
- ❌ `Model.all().count()` → ✅ `count(Model)`
- ❌ `Model.get_or_none()` → ✅ `get_or_none(Model, ...)`
- ❌ `Tortoise.close_connections()` → ✅ `shutdown_db()`

**Functions Converted**:
- ✅ `load_org_data_from_settings()`
- ✅ `load_base_data()`
- ✅ `load_pdfs()`
- ✅ `load_raw_copyright_data_to_staging()`
- ✅ `load_faculty_updates_to_staging()`

**Remaining Lint Errors**:
- 3 errors about `programme.name.lower()` where `name` could be None - not critical

**Status**: ✅ Functionally complete.

---

## 🚧 IN-PROGRESS / NOT STARTED

### 5. `easy_access/db/relations.py` - Relationship Management (⏳ NOT STARTED)

**Size**: ~500 lines
**Complexity**: HIGH - Complex M2M operations, queryset resolution

**Key Functions to Convert**:
1. `link_courses(settings)` - Links CopyrightItems to Courses based on codes
   - Uses `CopyrightItem.filter()` with `prefetch_related("courses")`
   - Batch fetches Course objects
   - Uses `.add()` for M2M relationships
   - Has test-aware queryset resolution (`_resolve_queryset_candidate`)

2. `link_persons_to_courses(settings, mapping)` - Creates CourseEmployee M2M relations
   - Batch fetches Person and Course objects
   - Uses `CourseEmployee.create()` for through-table records

3. `match_v1_to_copyright_items(settings)` - Matches legacy items
   - Complex polars DataFrame operations
   - Uses `Model.filter().values()` and `Model.filter().all()`

**Conversion Strategy**:
```python
# Before: M2M .add()
await item.courses.add(*course_objects)

# After: Option A - Use SQLAlchemy relationship
async for session in get_session():
    async with session.begin():
        # Refresh item if needed
        await session.refresh(item, ["courses"])
        item.courses.extend(course_objects)
        await session.flush()

# After: Option B - Direct association table insert (faster for batch)
from sqlalchemy import insert
stmt = insert(copyright_data_course).values([
    {"copyright_data_id": item.id, "course_cursuscode": c.cursuscode}
    for c in course_objects
])
async for session in get_session():
    async with session.begin():
        await session.execute(stmt)

# Before: prefetch_related
items = await CopyrightItem.filter().prefetch_related("courses")

# After: selectinload
from sqlalchemy.orm import selectinload
from sqlalchemy import select

async for session in get_session():
    stmt = select(CopyrightItem).options(selectinload(CopyrightItem.courses))
    result = await session.execute(stmt)
    items = list(result.scalars().all())
```

**Specific Conversions Needed**:
- `in_transaction` → `async with transaction():`
- `Model.filter()` → `filter_all(Model, ...)` or raw `select(Model).where(...)`
- `prefetch_related()` → `selectinload()` or `joinedload()`
- `M2M .add()` → SQLAlchemy relationship or association table insert
- `_resolve_queryset_candidate()` → Simplify or remove (test helper)

**Status**: ⏳ Ready to start. High priority due to complexity.

---

### 6. `easy_access/db/update.py` - Complex Update Logic (⏳ NOT STARTED)

**Size**: ~1,734 lines
**Complexity**: VERY HIGH - Strategy pattern, merge rules, bulk operations

**Key Sections**:
1. **Custom Exceptions** (lines 1-60) - No changes needed
2. **Field Comparison Strategies** (lines 60-220) - No changes needed
3. **Data Preprocessing** (lines 220-300) - Uses `CopyrightItem.all().values()`, needs conversion
4. **Merge/Update Logic** (lines 300-1000+) - Complex, uses many Tortoise methods
5. **Staged Processing** (lines 1000-1400) - Uses `in_transaction`, bulk operations
6. **Helper Functions** (lines 1400-1734) - Various CRUD operations

**Key Conversions Needed**:
- `in_transaction` → `async with transaction():`
- `bulk_create()`, `bulk_update()` → compat layer versions
- `Model.save()` → `session.add()` + `session.flush()`
- `Model.filter().all()` → `filter_all(Model, ...)`
- `Model.all().values()` → `all_values(Model, *fields)`

**Recommended Approach**:
1. Convert in phases (preprocessing → merge → staged processing)
2. Extract business logic from DB operations where possible
3. Keep strategy pattern intact (it's well-designed)
4. Add comprehensive error handling with transactions

**Status**: ⏳ NOT STARTED. Plan to tackle after relations.py.

---

### 7. `easy_access/db/retrieve.py` - Data Retrieval (⏳ NOT STARTED)

**Size**: ~827 lines
**Complexity**: MEDIUM - Mixed SQLAlchemy/Tortoise usage

**Current State**:
- ✅ Already uses SQLAlchemy engine + polars for main queries
- ❌ Has some Tortoise imports and Q expressions
- ❌ Some functions use Tortoise ORM

**Key Conversions**:
- Remove Tortoise imports (`from tortoise import Tortoise`)
- Remove Q expression imports (`from tortoise.expressions import Q`)
- Replace any remaining Tortoise model queries with compat layer

**Status**: ⏳ LOW PRIORITY. Most code is already SQLAlchemy-based.

---

### 8-12. Application Modules (⏳ NOT STARTED)

#### `pdf/download.py` & `pdf/parse.py`
**Patterns**: Simple CRUD - `get_or_none()`, `create()`, `filter()`, `save()`
**Strategy**: Straightforward compat layer replacement
**Estimated Effort**: 1 hour each

#### `enrichment/osiris.py`
**Patterns**: `all()`, `filter()`, `distinct()`, M2M `.add()`, `delete()`
**Strategy**: Similar to relations.py, focus on M2M operations
**Estimated Effort**: 2 hours

#### `maintenance/file_existence.py` & `maintenance/v1_items.py`
**Patterns**: `filter()`, `get()`, `update()`, `update_or_create()`
**Strategy**: Straightforward compat layer replacement
**Estimated Effort**: 1 hour each

---

## 📋 FINAL CLEANUP TASKS

### 13. Remove Tortoise Dependencies

**Files to Clean**:
- `pyproject.toml` - Remove `tortoise-orm` from dependencies
- Search all files for remaining Tortoise imports
- Remove `easy_access/db/models.py` (old Tortoise models)
- Verify no code references Tortoise classes

**Command**:
```powershell
# Search for remaining Tortoise references
Get-ChildItem -Path e:\ea-cli\easy_access -Recurse -Filter *.py |
  Select-String -Pattern "from tortoise|import tortoise"
```

---

### 14. Update Documentation

**Files to Update**:
- `README.md` - Add PostgreSQL setup instructions
- `README.md` - Update database initialization steps
- `README.md` - Add migration notes from SQLite
- `.agents/memory.instruction.md` - Update DB section

**Key Topics**:
- PostgreSQL 18 installation (Docker recommended)
- Environment variables (`DATABASE_URL`)
- Alembic migration commands
- Data migration from SQLite (pgloader)

---

## 🎯 RECOMMENDED NEXT STEPS

### Immediate (High Priority)

1. **Fix remaining lint errors in base.py**
   - Fix `copyright_item_from_dict()` or mark as deprecated

2. **Convert `db/relations.py`**
   - Critical module with M2M operations
   - Needs careful handling of association tables
   - Test thoroughly after conversion

3. **Convert `db/update.py`** (in phases)
   - Phase A: Preprocessing functions
   - Phase B: Merge logic (keep strategy pattern)
   - Phase C: Staged processing
   - Phase D: Helper functions

### Medium Priority

4. **Convert application modules**
   - `pdf/download.py`, `pdf/parse.py` - Straightforward
   - `enrichment/osiris.py` - Similar to relations.py
   - `maintenance/file_existence.py`, `maintenance/v1_items.py` - Straightforward

5. **Clean up `db/retrieve.py`**
   - Remove Tortoise imports
   - Unify query patterns

### Final Steps

6. **Remove Tortoise dependencies**
   - Clean up imports
   - Remove old files
   - Update pyproject.toml

7. **Update documentation**
   - README with PostgreSQL setup
   - Migration guide
   - Update memory.instruction.md

---

## 🧪 TESTING STRATEGY

After each module conversion:

1. **Run the application** with converted module
2. **Test key workflows**:
   - `uv run run.py process` (full pipeline)
   - Ingest → Enrichment → Export
3. **Verify data integrity**:
   - Check row counts match expected
   - Validate relationships (M2M tables)
4. **Check for N+1 queries**:
   - Enable SQL logging: `init_db(settings, echo=True)`
   - Watch for repeated similar queries
5. **Validate transactions**:
   - Test rollback scenarios
   - Verify atomic operations

---

## 📊 PROGRESS METRICS

| Category | Completed | Total | % Done |
|----------|-----------|-------|--------|
| **Foundation** | 3 | 3 | 100% |
| **Core DB Modules** | 1 | 4 | 25% |
| **App Modules** | 0 | 5 | 0% |
| **Cleanup** | 0 | 2 | 0% |
| **TOTAL** | 4 | 14 | **29%** |

**Estimated Remaining Effort**: 12-16 hours

---

## 🚀 QUICK REFERENCE: Common Conversions

```python
# 1. Simple Query
# Before: items = await Model.filter(field=value).all()
# After:  items = await filter_all(Model, field=value)

# 2. Get Single Object
# Before: obj = await Model.get_or_none(field=value)
# After:  obj = await get_or_none(Model, field=value)

# 3. Create with Get or Create
# Before: obj, created = await Model.get_or_create(field=value, defaults={...})
# After:  obj, created = await get_or_create(Model, field=value, defaults={...})

# 4. Bulk Insert with Upsert
# Before: await Model.bulk_create(objects=[...], on_conflict=[...], update_fields=[...])
# After:  await bulk_create(Model, rows=[...], on_conflict=[...], update_fields=[...])

# 5. Transaction
# Before: async with in_transaction():
# After:  async with transaction() as session:

# 6. Count/Exists
# Before: count = await Model.all().count()
# After:  count = await count(Model)

# 7. M2M Relationship
# Before: await item.related.add(*objects)
# After:  async for session in get_session():
#             async with session.begin():
#                 item.related.extend(objects)
#                 await session.flush()
```

---

## 📝 NOTES

- **Database URL**: Currently hardcoded to read from `DATABASE_URL` env var in `session.py`
- **Naming Convention**: SQLAlchemy MetaData uses naming conventions for constraints - ensure migrations respect this
- **Type Safety**: Added type hints to compat layer, maintain consistency
- **Performance**: Bulk operations use batching (batch_size=500 by default)
- **Error Handling**: Compat layer includes race condition handling for `get_or_create`

---

**Last Updated**: October 13, 2025
**Next Review**: After relations.py conversion
