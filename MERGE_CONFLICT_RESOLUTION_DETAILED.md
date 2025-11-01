# Merge Conflict Resolution - Detailed Documentation

## Context
Resolved merge conflicts when merging `main` (commit 25d55d1) into `feat/db/postgres-sqlalchemy-migration` (commit 7b242a1), incorporating performance optimizations from PR #12.

## Files with Conflicts

### 1. uv.lock
**Resolution:** Accepted version from `main`  
**Rationale:** As requested, will regenerate lockfile later with correct dependencies

### 2. easy_access/db/update.py

#### Change 1: Imports (lines 23-38)
**Main branch had:**
```python
from easy_access.db.enums import (
    CLASSIFICATION_MAPPING_V1_TO_V2,
    Classification,
    ClassificationMapping,
    ClassificationV2,
)
from easy_access.db.models import (  # Tortoise models
```

**Postgres refactor had:**
```python
from easy_access.db.enums import (
    Classification,
    Infringement,
    Status,
    WorkflowStatus,
)
from easy_access.db.sa_models import (  # SQLAlchemy models
```

**Resolution:**
```python
from easy_access.db.enums import (
    CLASSIFICATION_MAPPING_V1_TO_V2,  # Added from main
    Classification,
    ClassificationMapping,            # Added from main
    ClassificationV2,                 # Added from main
    Infringement,
    Status,
    WorkflowStatus,
)
from easy_access.db.sa_models import (  # Kept SQLAlchemy models
```

#### Change 2: Module-level constants (after line 93)
**Added from PR #12 for performance optimization:**
```python
# Pre-calculated lookup dictionaries for Classification enum normalization
# Used in map_v1_to_v2_classifications to avoid recreating sets for each item
_CLASSIFICATION_NORMALIZE_PATTERN = re.compile(r'[\s_-]')
LOWER_TO_CLASSIFICATION = {e.value.lower(): e for e in Classification}
NORMALIZED_TO_CLASSIFICATION = {
    _CLASSIFICATION_NORMALIZE_PATTERN.sub('', e.value.lower()): e for e in Classification
}
```

#### Change 3: map_v1_to_v2_classifications function

**Main branch implementation (Tortoise ORM):**
```python
selected_items = await CopyrightItem.filter(
    Q(v2_manual_classification__isnull=True)
    | Q(v2_manual_classification=ClassificationV2.ONBEKEND)
).all().prefetch_related('v1_items', 'faculty')

async with in_transaction():
    for item in selected_items:
        # ... process item ...
        await item.save(update_fields=[...])
```

**Postgres refactor stub (raw SQL):**
```python
result = await session.execute(text("""
    SELECT material_id, period, department, ... FROM copyright_data
    WHERE v2_manual_classification IS NULL OR v2_manual_classification = 'Onbekend'
"""))
# Manually convert rows to objects
```

**Resolution (SQLAlchemy with PR #12 optimizations):**
```python
async for session in get_session():
    # Query using SQLAlchemy select
    result = await session.execute(
        select(CopyrightItem).where(
            (CopyrightItem.v2_manual_classification.is_(None))
            | (CopyrightItem.v2_manual_classification == ClassificationV2.ONBEKEND.value)
        )
    )
    selected_items = result.scalars().all()
    
    for item in selected_items:
        # ... normalization logic ...
        
        # PR #12 optimization: Replace match-case with dictionary lookups
        key = LOWER_TO_CLASSIFICATION.get(current)
        if not key:
            normalized = _CLASSIFICATION_NORMALIZE_PATTERN.sub('', current)
            key = NORMALIZED_TO_CLASSIFICATION.get(normalized)
        if not key:
            key = Classification.ONBEKEND
        
        # ... mapping logic ...
        
        # Update items directly (no .save() method in SQLAlchemy)
        if item.v2_manual_classification != mapped.classification.value:
            item.v2_manual_classification = mapped.classification.value
            item.v2_lengte = mapped.length.value
            item.v2_overnamestatus = mapped.overname_status.value
            modified_count += 1
    
    # Single commit at the end
    await session.commit()
```

## Key Adaptations

### From Tortoise to SQLAlchemy:
- `CopyrightItem.filter(Q(...))` → `session.execute(select(CopyrightItem).where(...))`
- `in_transaction()` → `async for session in get_session()`
- `prefetch_related('v1_items', 'faculty')` → Removed (relationships not yet in SA models)
- `await item.save(update_fields=[...])` → Direct attribute assignment + `await session.commit()`
- Simplified detail dict (removed faculty/v1_items tracking)

### PR #12 Performance Optimization:

**Before (match statement - O(n) complexity per iteration):**
```python
match current:
    case val if val in {e.value for e in Classification}:  # Creates set every iteration
        key = Classification(val)
    case val if val.lower() in {e.value.lower() for e in Classification}:  # Redundant
        key = Classification(next(e.value for e in Classification if e.value.lower() == val.lower()))
    case val if re.sub(r"[\s_-]", "", val.lower()) in {re.sub(r"[\s_-]", "", e.value.lower()) for e in Classification}:
        # Creates set + compiles regex every iteration
        key = Classification(next(...))
    case _:
        key = Classification.ONBEKEND
```

**After (if/elif with lookups - O(1) complexity):**
```python
# At module level (one-time cost):
_CLASSIFICATION_NORMALIZE_PATTERN = re.compile(r'[\s_-]')
LOWER_TO_CLASSIFICATION = {e.value.lower(): e for e in Classification}
NORMALIZED_TO_CLASSIFICATION = {
    _CLASSIFICATION_NORMALIZE_PATTERN.sub('', e.value.lower()): e for e in Classification
}

# In loop (constant time):
key = LOWER_TO_CLASSIFICATION.get(current)  # O(1) dictionary lookup
if not key:
    normalized = _CLASSIFICATION_NORMALIZE_PATTERN.sub('', current)  # Pre-compiled regex
    key = NORMALIZED_TO_CLASSIFICATION.get(normalized)  # O(1) dictionary lookup
if not key:
    key = Classification.ONBEKEND
```

**Benefits:**
- **Performance**: O(1) dictionary lookups vs O(n) set comprehensions on each iteration
- **Clarity**: Simpler if/elif structure is easier to understand than nested match cases
- **Correctness**: Removed redundant case (checking lowercase on already-lowercase value)
- **Efficiency**: Pre-compiled regex pattern avoids repeated compilation

## Result
The final resolution successfully combines:
- All postgres refactor changes (SQLAlchemy migration)
- New v2 classification enums and mapping from main
- Adapted v2 classification mapping function using SQLAlchemy
- Performance optimizations from PR #12

All other files auto-merged without conflicts.
