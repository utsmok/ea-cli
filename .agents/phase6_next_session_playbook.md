# Phase 6 - Next Session Playbook

**For**: Next developer/session continuing the migration
**Status**: 4/14 modules complete (29%)
**Priority**: Complete relations.py, then update.py

---

## 🎯 Immediate Task: Complete relations.py

**File**: `easy_access/db/relations.py`
**Status**: Conversion guide exists, implementation pending
**Reference**: `.agents/relations_conversion_guide.md`
**Estimated Time**: 2.5 hours

### Quick Start
1. Open `.agents/relations_conversion_guide.md`
2. Follow step-by-step instructions for each function
3. Use code examples provided (copy-paste and adapt)
4. Test each function after conversion

### Key Imports to Add
```python
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

### Imports to Remove
```python
from tortoise.transactions import in_transaction
# Remove _resolve_queryset_candidate function (test compatibility no longer needed)
```

### Validation Checklist
After converting relations.py:
- [ ] No `tortoise` imports remain
- [ ] All functions use `get_session()` or `transaction()`
- [ ] M2M operations use association table inserts
- [ ] Prefetch replaced with `selectinload()`
- [ ] Run linter: `uv run ruff check easy_access/db/relations.py`
- [ ] Check for errors: Look for any compilation errors
- [ ] Update TODO: Mark task #5 as complete

---

## 📋 Task Priority Order

### Phase 1: Core Database Operations (High Priority)
1. ✅ compat.py (DONE)
2. ✅ session.py (DONE)
3. ✅ base.py (DONE)
4. ✅ ingest.py (DONE)
5. **⏳ relations.py (NEXT - use guide)**
6. **update.py** (1,734 lines - break into sub-tasks)
7. retrieve.py (mostly cleanup)

### Phase 2: Application Features (Medium Priority)
8. enrichment/osiris.py (similar to relations.py)
9. pdf/download.py (straightforward)
10. pdf/parse.py (straightforward)

### Phase 3: Maintenance & Cleanup (Low Priority)
11. maintenance/file_existence.py (straightforward)
12. maintenance/v1_items.py (straightforward)
13. Remove tortoise-orm from pyproject.toml
14. Update README.md

---

## 🔧 Common Patterns Reference

### Pattern 1: Simple Query
```python
from easy_access.db.compat import filter_all

# Get all items matching criteria
items = await filter_all(Model, field=value)
```

### Pattern 2: Get or Create
```python
from easy_access.db.compat import get_or_create

obj, created = await get_or_create(
    Model,
    lookup_field=lookup_value,
    defaults={"field1": "value1"}
)
```

### Pattern 3: Bulk Insert with Upsert
```python
from easy_access.db.compat import bulk_create

await bulk_create(
    Model,
    rows=[{"id": 1, "name": "test"}],
    on_conflict=["id"],
    update_fields=["name"]
)
```

### Pattern 4: Transaction
```python
from easy_access.db.compat import transaction

async with transaction() as session:
    # All operations here are atomic
    await session.execute(...)
```

### Pattern 5: M2M Insert (Direct Association Table)
```python
from sqlalchemy import insert
from easy_access.db.sa_models import association_table
from easy_access.db.compat import transaction

links = [
    {"left_id": 1, "right_id": 2},
    {"left_id": 1, "right_id": 3},
]

async with transaction() as session:
    await session.execute(
        insert(association_table).values(links)
    )
```

### Pattern 6: Eager Loading (Prefetch)
```python
from sqlalchemy import select
from sqlalchemy.orm import selectinload
from easy_access.db.session import get_session

async for session in get_session():
    stmt = select(Model).options(
        selectinload(Model.related_field)
    )
    result = await session.execute(stmt)
    items = list(result.scalars().all())

    # Now items have related_field populated
    for item in items:
        for related in item.related_field:
            # No additional queries
            pass
    break
```

---

## 📚 Essential Files to Reference

### Conversion Guides
- `.agents/relations_conversion_guide.md` - Detailed guide for relations.py
- `.agents/phase6_progress.md` - Pattern examples and quick reference
- `.agents/phase6_status_and_next_steps.md` - Full status with code examples

### Implementation References
- `easy_access/db/compat.py` - All Tortoise replacement functions
- `easy_access/db/ingest.py` - Example of complete conversion
- `easy_access/db/base.py` - Example of initialization conversion
- `easy_access/db/session.py` - Session management pattern

### Project Context
- `.agents/memory.instruction.md` - Full project context and preferences
- `migration_to_postgres_sqlalchemy.md` - Original migration plan
- `settings.yaml` - Database configuration

---

## 🧪 Testing Strategy

### After Each Module Conversion
1. **Lint Check**: `uv run ruff check <file_path>`
2. **Type Check**: `uv run pyright <file_path>` (if installed)
3. **Import Test**: `uv run python -c "from easy_access.db.<module> import *"`
4. **Visual Inspection**: Look for any obvious errors

### Before Committing
1. Run full linter: `uv run ruff check easy_access/`
2. Check git status: Ensure only intended files changed
3. Review diffs: Make sure no unintended changes
4. Update TODO list: Mark completed tasks

---

## 🚨 Common Pitfalls to Avoid

### 1. Session Management
❌ **Don't**: Create session without proper cleanup
```python
session = AsyncSession(engine)  # No cleanup!
```

✅ **Do**: Use get_session() generator or transaction()
```python
async for session in get_session():
    # Auto cleanup
    pass
```

### 2. M2M Operations
❌ **Don't**: Try to use ORM .add() method
```python
await item.courses.add(course)  # Tortoise pattern!
```

✅ **Do**: Use association table insert
```python
async with transaction() as session:
    await session.execute(
        insert(association_table).values([{...}])
    )
```

### 3. Prefetch Operations
❌ **Don't**: Query related objects in loop (N+1)
```python
for item in items:
    courses = await session.execute(
        select(Course).where(...)
    )  # N+1 query!
```

✅ **Do**: Use selectinload()
```python
stmt = select(Item).options(selectinload(Item.courses))
items = await session.execute(stmt)
# Now items have courses prefetched
```

### 4. Import Organization
❌ **Don't**: Mix Tortoise and SQLAlchemy
```python
from tortoise import Model  # Old!
from easy_access.db.sa_models import Model  # New!
```

✅ **Do**: Only use SQLAlchemy models
```python
from easy_access.db.sa_models import Model as SAModel
# Or use compat layer that handles it
```

---

## 📊 Progress Tracking

### After Each Module
1. Update TODO list via `manage_todo_list` tool
2. Note completion in `.agents/memory.instruction.md`
3. Add any new patterns discovered to guides
4. Document any issues encountered

### Current Progress
```
✅ compat.py (Foundation)
✅ session.py (Foundation)
✅ base.py (Foundation)
✅ ingest.py (Foundation)
⏳ relations.py (In Progress) ← YOU ARE HERE
⏳ update.py (Next)
⏳ retrieve.py
... 7 more modules
```

---

## 🎯 Success Metrics

Track these as you work:
- [ ] Modules converted: 4/14 (29%)
- [ ] Lines converted: ~500 / ~4,000 (13%)
- [ ] Lint errors: Trending down
- [ ] Import errors: None after each module
- [ ] Test compatibility: Maintained or improved

---

## 💡 Pro Tips

1. **Copy-Paste from Guides**: The conversion guides have working code - use it!
2. **One Function at a Time**: Convert and test incrementally
3. **Use Compat Layer**: Don't reinvent - compat.py has what you need
4. **Check Imports First**: Fix imports before logic
5. **Backup Before Major Changes**: `Copy-Item` to create .bak files
6. **Reference ingest.py**: It's a complete example of conversion
7. **Ask for Clarification**: If pattern unclear, check guides or memory.instruction.md

---

## 🔄 If You Get Stuck

### Issue: Conversion pattern unclear
**Solution**: Check these in order:
1. `.agents/relations_conversion_guide.md` (for relations.py)
2. `.agents/phase6_progress.md` (for general patterns)
3. `easy_access/db/ingest.py` (working example)
4. `easy_access/db/compat.py` (available functions)

### Issue: Association table not found
**Solution**: Check `easy_access/db/sa_models.py` for table definition
```python
# Should exist:
copyright_item_course_association = Table(...)
```

### Issue: Import errors
**Solution**: Verify all imports from conversion guide
- Remove: `from tortoise...`
- Add: `from sqlalchemy...`, `from easy_access.db.sa_models...`

### Issue: Test failures (if tests exist)
**Solution**:
1. Focus on production code first
2. Update tests after conversion complete
3. Use real SQLAlchemy models in tests (simpler than mocks)

---

## ✅ Definition of Done (Per Module)

A module is complete when:
- [ ] No Tortoise imports remain
- [ ] All functions use SQLAlchemy patterns
- [ ] Lint check passes (or only acceptable warnings)
- [ ] No import errors when testing
- [ ] TODO list updated
- [ ] Memory/docs updated if needed
- [ ] Git diff reviewed

---

## 🚀 Ready to Start?

1. Open `.agents/relations_conversion_guide.md`
2. Open `easy_access/db/relations.py`
3. Follow guide step-by-step
4. Test after each function conversion
5. Update TODO when done
6. Move to update.py (next highest priority)

**You've got this!** The foundation is complete, patterns are clear, and guides are comprehensive. Just follow the established patterns and you'll make quick progress.

---

**Last Updated**: October 13, 2025
**Next Reviewer**: Continue from relations.py using detailed conversion guide
