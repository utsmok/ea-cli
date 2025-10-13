# Relations.py Conversion Guide

## File: `easy_access/db/relations.py`

**Status**: 🔄 In Progress
**Complexity**: High (M2M operations, prefetch patterns, test compatibility)
**Lines**: 500+

---

## Current Tortoise Patterns

### 1. M2M Add Pattern (Tortoise)
```python
# Old code
async with in_transaction():
    await item.courses.add(*course_objects)
```

### 2. Prefetch Pattern (Tortoise)
```python
# Old code
items = await CopyrightItem.filter().prefetch_related("courses")
```

### 3. Through Model Create (Tortoise)
```python
# Old code
await CourseEmployee.create(course=course_obj, person=person_obj, role=role)
```

---

## New SQLAlchemy Patterns

### 1. M2M Add → Bulk Association Insert
```python
# New code - using association table
from sqlalchemy import insert
from easy_access.db.sa_models import copyright_item_course_association
from easy_access.db.compat import transaction

links_to_add = [
    {"copyright_item_id": item.material_id, "course_id": course.cursuscode}
    for course in courses_to_link
]

async with transaction() as session:
    await session.execute(
        insert(copyright_item_course_association).values(links_to_add)
    )
```

### 2. Prefetch → selectinload
```python
# New code - using selectinload for eager loading
from sqlalchemy import select
from sqlalchemy.orm import selectinload
from easy_access.db.sa_models import CopyrightItem as SACopyrightItem
from easy_access.db.session import get_session

async for session in get_session():
    stmt = select(SACopyrightItem).options(
        selectinload(SACopyrightItem.courses)
    )
    result = await session.execute(stmt)
    items = list(result.scalars().all())

    # Now items have .courses populated (no N+1 queries)
    for item in items:
        for course in item.courses:  # Already loaded
            print(course.cursuscode)
```

### 3. Through Model → Direct Insert
```python
# New code - using insert with table reference
from sqlalchemy import insert
from easy_access.db.sa_models import CourseEmployee as SACourseEmployee
from easy_access.db.compat import transaction

rows = [
    {"course_id": course_pk, "person_id": person_pk, "role": role}
    for course_pk, person_pk, role in desired_links
]

async with transaction() as session:
    await session.execute(
        insert(SACourseEmployee.__table__).values(rows)
    )
```

---

## Conversion Steps for `link_courses()`

### Step 1: Replace Query with selectinload
```python
# OLD
all_items_candidate = CopyrightItem.filter()
all_items = await _resolve_queryset_candidate(all_items_candidate, "courses")

# NEW
async for session in get_session():
    stmt = select(SACopyrightItem).options(
        selectinload(SACopyrightItem.courses)
    )
    result = await session.execute(stmt)
    items = list(result.scalars().all())
    break
```

### Step 2: Query Courses
```python
# OLD
courses = await Course.filter(code__in=list(valid_course_codes))

# NEW
async for session in get_session():
    stmt = select(SACourse).where(
        SACourse.cursuscode.in_(list(valid_course_codes))
    )
    result = await session.execute(stmt)
    courses = list(result.scalars().all())
    break
```

### Step 3: Build M2M Links
```python
# NEW - collect links instead of calling .add()
links_to_add: list[dict[str, int]] = []

for item in items:
    # Get existing course IDs (already prefetched)
    existing_course_ids = {c.cursuscode for c in item.courses}

    # Determine new links needed
    for course_code_str in item_course_map[item.material_id]:
        int_code = safe_int(course_code_str)
        if int_code in existing_course_ids:
            continue  # Already linked
        if int_code not in course_map:
            continue  # Course doesn't exist

        links_to_add.append({
            "copyright_item_id": item.material_id,
            "course_id": int_code,
        })
```

### Step 4: Bulk Insert M2M Links
```python
# NEW - single bulk insert
if links_to_add:
    async with transaction() as session:
        await session.execute(
            insert(copyright_item_course_association).values(links_to_add)
        )
    logger.success(f"Added {len(links_to_add)} course links")
```

---

## Conversion Steps for `link_persons_to_courses()`

### Step 1: Query Courses and Persons
```python
async for session in get_session():
    # Courses
    stmt = select(SACourse).where(
        SACourse.cursuscode.in_(list(all_course_codes))
    )
    result = await session.execute(stmt)
    courses = list(result.scalars().all())

    # Persons
    stmt = select(SAPerson).where(
        SAPerson.people_page_url.in_(list(all_people_page_urls))
    )
    result = await session.execute(stmt)
    persons = list(result.scalars().all())
```

### Step 2: Fetch Existing Links
```python
# Query existing CourseEmployee records
from sqlalchemy import and_

stmt = select(SACourseEmployee).where(
    and_(
        SACourseEmployee.course_id.in_(list(course_pks)),
        SACourseEmployee.person_id.in_(list(person_pks)),
    )
)
result = await session.execute(stmt)
existing = list(result.scalars().all())
existing_pairs = {(e.course_id, e.person_id) for e in existing}
```

### Step 3: Bulk Insert New Links
```python
to_create = [
    {"course_id": cpk, "person_id": ppk, "role": role}
    for cpk, ppk, role in desired
    if (cpk, ppk) not in existing_pairs
]

if to_create:
    async with transaction() as session:
        await session.execute(
            insert(SACourseEmployee.__table__).values(to_create)
        )
```

---

## Conversion Steps for `match_v1_to_copyright_items()`

### Step 1: Fetch All Items
```python
async for session in get_session():
    stmt_v1 = select(SAv1_CopyrightItem)
    result_v1 = await session.execute(stmt_v1)
    v1_items = list(result_v1.scalars().all())

    stmt_current = select(SACopyrightItem)
    result_current = await session.execute(stmt_current)
    current_items = list(result_current.scalars().all())
```

### Step 2: Convert to Polars DataFrames
```python
# Extract column values into dicts
v1_item_dicts = [
    {col.name: getattr(item, col.name)
     for col in SAv1_CopyrightItem.__table__.columns}
    for item in v1_items
]

# Same for current_items...
# Then create DataFrames as before
v1_df = pl.DataFrame(v1_item_dicts, ...)
```

### Step 3: Update in Transaction
```python
# After matching logic...
async with transaction() as session:
    for row in all_matched_df.to_dicts():
        v1_item = v1_item_dict[row["material_id"]]
        current_item = current_item_dict[row["material_id_right"]]

        # Update v1_item
        v1_item.matching_copyright_item_id = current_item.material_id
        session.add(v1_item)

        # Update current_item
        if should_update_classification:
            current_item.manual_classification = v1_item.manual_classification
            session.add(current_item)
```

---

## Required Imports

```python
# Remove
from tortoise.transactions import in_transaction

# Add
from sqlalchemy import and_, insert, select
from sqlalchemy.orm import selectinload
from easy_access.db.session import get_session
from easy_access.db.compat import transaction
from easy_access.db.sa_models import (
    CopyrightItem as SACopyrightItem,
    Course as SACourse,
    CourseEmployee as SACourseEmployee,
    Person as SAPerson,
    v1_CopyrightItem as SAv1_CopyrightItem,
    copyright_item_course_association,
)
```

---

## Test Compatibility Notes

The old code had extensive test compatibility logic:
- `_resolve_queryset_candidate()` - Handled mock querysets
- Bulk update mock shortcuts
- Try/except around transactions

**New approach**: Simpler, cleaner code. Tests should:
1. Mock `get_session()` to return test sessions
2. Use real SQLAlchemy models in tests (in-memory SQLite)
3. Remove complex mock resolution logic

---

## Association Table Reference

Make sure `copyright_item_course_association` is imported from `sa_models.py`:

```python
# In sa_models.py
copyright_item_course_association = Table(
    "copyright_item_course",
    Base.metadata,
    Column("copyright_item_id", Integer, ForeignKey("copyright_item.material_id")),
    Column("course_id", Integer, ForeignKey("course.cursuscode")),
)
```

---

## Validation Checklist

After conversion:
- [ ] No `tortoise` imports remain
- [ ] All async functions use `get_session()` or `transaction()`
- [ ] M2M operations use association table inserts
- [ ] Prefetch replaced with `selectinload()`
- [ ] Lint errors resolved
- [ ] Test file updated (if tests exist)
- [ ] Logged operations remain descriptive

---

## Estimated Effort

- **link_courses()**: 45 minutes (complex M2M logic)
- **link_persons_to_courses()**: 30 minutes (similar pattern)
- **match_v1_to_copyright_items()**: 20 minutes (mostly DataFrame work)
- **Remove test compatibility code**: 15 minutes
- **Testing and validation**: 30 minutes

**Total**: ~2.5 hours for complete conversion
