# Phase 6 Progress: Tortoise to SQLAlchemy Migration

## Completed Work

### 1. Enhanced compat.py (✅ COMPLETE)
Created comprehensive compatibility layer with SQLAlchemy implementations:
- `bulk_create()` - Batch inserts with upsert support (PostgreSQL ON CONFLICT)
- `bulk_update()` - Batch updates
- `get_or_create()` - Get or create with race condition handling
- `update_or_create()` - Update or create pattern
- `get_by_id()`, `get_one()`, `get_or_none()` - Query helpers
- `filter_all()` - Filter and return all results
- `filter_values()`, `all_values()` - Return specific fields as dicts
- `count()`, `exists()` - Aggregation helpers
- `delete_where()` - Bulk delete
- `create_instance()`, `save_instance()` - CRUD helpers
- `transaction()` - Async context manager for transactions
- `_batched()` - Internal batching utility

### 2. Updated session.py (✅ COMPLETE)
- Modified `init_db()` to accept both Settings object and URL string
- Extracts PostgreSQL connection from environment variable (DATABASE_URL)
- Maintains async session factory pattern

### 3. Converted db/base.py (✅ MOSTLY COMPLETE)
- Replaced `Tortoise.init()` with `init_db(settings)`
- Replaced `Tortoise.generate_schemas()` with `Base.metadata.create_all()`
- Replaced `Tortoise.close_connections()` with `shutdown_db()`
- Updated `init_faculties()` to use `get_or_create()` compat function
- Note: `copyright_item_from_dict()` still needs conversion

### 4. Converted db/ingest.py (✅ MOSTLY COMPLETE)
- Replaced `Model.get_or_create()` with `get_or_create()` compat
- Replaced `Model.bulk_create()` with `bulk_create()` compat
- Replaced `Model.all().values()` with `all_values()` compat
- Replaced `Model.all().count()` with `count()` compat
- Replaced `Model.get_or_none()` with `get_or_none()` compat
- Replaced `Tortoise.close_connections()` with `shutdown_db()`

## Remaining Work

### 5. db/relations.py (IN PROGRESS)
**Key Patterns to Convert:**
- `in_transaction` → `async with transaction():`
- `Model.filter()` → `filter_all(Model, **kwargs)`
- `prefetch_related()` → SQLAlchemy `selectinload()` or `joinedload()`
- `M2M .add()` → Insert into association table or use SQLAlchemy relationship helpers
- `_resolve_queryset_candidate()` → Simplify with direct SQLAlchemy queries

**Strategy:**
- Create dedicated functions for M2M operations
- Use SQLAlchemy's relationship loading strategies
- Batch M2M inserts for performance

### 6. db/update.py (NOT STARTED - COMPLEX)
**Key Patterns to Convert:**
- `in_transaction` → `async with transaction():`
- Complex merge strategies already use strategy pattern - keep intact
- `bulk_create`, `bulk_update` → Use compat layer
- Field-by-field updates → Use SQLAlchemy session tracking

**Strategy:**
- Break down into smaller functions
- Extract business logic from DB operations
- Create service layer for complex operations
- Add comprehensive error handling

### 7. db/retrieve.py (NOT STARTED)
**Key Patterns to Convert:**
- Already uses SQLAlchemy engine for polars reads - keep that pattern
- Remove Tortoise imports and Q expressions
- Replace Tortoise ORM queries with raw SQL or compat layer

**Strategy:**
- Keep polars + SQLAlchemy Core pattern for read-heavy operations
- Unify query patterns across the module

### 8. pdf/download.py (NOT STARTED)
**Patterns:** `get_or_none()`, `create()`, `filter()`
**Strategy:** Straightforward compat layer usage

### 9. pdf/parse.py (NOT STARTED)
**Patterns:** `all()`, `create()`, `save()`
**Strategy:** Straightforward compat layer usage

### 10. enrichment/osiris.py (NOT STARTED - MEDIUM COMPLEXITY)
**Patterns:** `all()`, `filter()`, `distinct()`, `get_or_none()`, M2M `.add()`, `delete()`
**Strategy:** Similar to relations.py, focus on M2M operations

### 11. maintenance/file_existence.py (NOT STARTED)
**Patterns:** `filter()`, `get()`, `update()`
**Strategy:** Straightforward compat layer usage

### 12. maintenance/v1_items.py (NOT STARTED)
**Patterns:** `update_or_create()`, `all()`, `values_list()`, `filter()`
**Strategy:** Straightforward compat layer usage

## Implementation Guidelines

### Common Conversion Patterns

#### 1. Simple Queries
```python
# Before (Tortoise)
items = await Model.filter(field=value).all()

# After (SQLAlchemy via compat)
items = await filter_all(Model, field=value)
```

#### 2. Get or Create
```python
# Before
obj, created = await Model.get_or_create(
    field=value,
    defaults={"other": "value"}
)

# After
obj, created = await get_or_create(
    Model,
    field=value,
    defaults={"other": "value"}
)
```

#### 3. Bulk Operations
```python
# Before
await Model.bulk_create(
    objects=[Model(**d) for d in data],
    on_conflict=["id"],
    update_fields=["name"]
)

# After
await bulk_create(
    Model,
    rows=data,
    on_conflict=["id"],
    update_fields=["name"]
)
```

#### 4. Transactions
```python
# Before
async with in_transaction():
    # operations

# After
async with transaction() as session:
    # operations
```

#### 5. M2M Relationships
```python
# Before (Tortoise)
await item.courses.add(*course_objects)

# After (SQLAlchemy) - Option A: Use relationship
async for session in get_session():
    async with session.begin():
        item.courses.extend(course_objects)
        await session.flush()

# After (SQLAlchemy) - Option B: Direct association table insert
from sqlalchemy import insert
stmt = insert(association_table).values([
    {"item_id": item.id, "course_id": c.id}
    for c in course_objects
])
async for session in get_session():
    async with session.begin():
        await session.execute(stmt)
```

#### 6. Prefetch Related
```python
# Before (Tortoise)
items = await Model.filter().prefetch_related("related_field")

# After (SQLAlchemy)
from sqlalchemy.orm import selectinload

async for session in get_session():
    stmt = select(Model).options(selectinload(Model.related_field))
    result = await session.execute(stmt)
    items = list(result.scalars().all())
```

### Q Expressions Conversion

Tortoise Q expressions need to be converted to SQLAlchemy filter expressions:

```python
# Before
from tortoise.expressions import Q
items = await Model.filter(Q(field1=value1) | Q(field2=value2))

# After
from sqlalchemy import or_

async for session in get_session():
    stmt = select(Model).where(
        or_(Model.field1 == value1, Model.field2 == value2)
    )
    result = await session.execute(stmt)
    items = list(result.scalars().all())
```

## Testing Strategy

After each module conversion:
1. Run the application with converted module
2. Test key workflows (ingest → enrichment → export)
3. Verify data integrity
4. Check for N+1 query issues
5. Validate transaction boundaries

## Performance Considerations

- Use `bulk_create` and `bulk_update` for batch operations
- Add proper indexes to SQLAlchemy models
- Use `selectinload` or `joinedload` to prevent N+1 queries
- Consider using SQLAlchemy Core for read-heavy operations
- Monitor query performance with `echo=True` during development

## Next Steps

1. ✅ Complete db/ingest.py cleanup (remove unused imports)
2. Start db/relations.py conversion
3. Tackle db/update.py (most complex)
4. Convert remaining simpler modules
5. Remove Tortoise dependencies from pyproject.toml
6. Update documentation

## Notes

- Keep the compat layer minimal and focused
- Add new compat functions only when needed
- Document any SQLAlchemy-specific patterns that differ significantly
- Maintain backward compatibility during transition
- Use type hints consistently
