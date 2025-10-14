"""Small compatibility layer exposing a tiny subset of the old Tortoise API.

This file provides helper adapters so we can incrementally replace Tortoise
call sites with SQLAlchemy implementations. Keep this file intentionally
small — add functions as needed when migrating modules.
"""

from __future__ import annotations

from collections.abc import Iterable
from contextlib import asynccontextmanager
from typing import Any

from loguru import logger
from sqlalchemy import and_, delete, insert, select, update
from sqlalchemy.dialects.postgresql import insert as pg_insert
from sqlalchemy.exc import IntegrityError

from .session import get_session


async def bulk_create(
    model,
    objects: Iterable[Any] | None = None,
    rows: Iterable[dict[str, Any]] | None = None,
    batch_size: int = 500,
    on_conflict: list[str] | None = None,
    update_fields: list[str] | None = None,
) -> None:
    """Insert rows in batches using SQLAlchemy, optionally with upsert logic.

    Args:
        model: SQLAlchemy ORM model class
        objects: Iterable of model instances (alternative to rows)
        rows: Iterable of dicts with column values (alternative to objects)
        batch_size: Number of rows per batch
        on_conflict: List of column names that define conflicts (for upsert)
        update_fields: List of column names to update on conflict

    Example:
        await bulk_create(MyModel, rows=[{"id": 1, "name": "test"}])
        await bulk_create(MyModel, objects=[MyModel(id=1, name="test")])
        await bulk_create(
            MyModel,
            rows=[{"id": 1, "name": "updated"}],
            on_conflict=["id"],
            update_fields=["name"]
        )
    """
    # Convert objects to dicts if provided
    if objects is not None:
        data = []
        for obj in objects:
            if hasattr(obj, "__dict__"):
                # Extract only column attributes, not relationship attributes
                row = {
                    k: v
                    for k, v in obj.__dict__.items()
                    if not k.startswith("_") and not callable(v)
                }
                data.append(row)
            else:
                raise ValueError(f"Object {obj} does not have __dict__")
    elif rows is not None:
        data = list(rows)
    else:
        raise ValueError("Either objects or rows must be provided")

    if not data:
        return

    # For upsert operations, use much smaller batch size to avoid PostgreSQL parameter limits
    # Each row can have dozens of parameters, so limit to very small batches
    effective_batch_size = batch_size
    if on_conflict and update_fields:
        # PostgreSQL has a hard limit of ~65k parameters per prepared statement
        # For upsert operations, each row contributes parameters to both INSERT and UPDATE clauses
        # Be extremely conservative - use a fixed small batch size
        effective_batch_size = min(batch_size, 50)  # Fixed small batch size for upsert
        logger.info(
            f"Using batch size {effective_batch_size} for upsert operation (fixed conservative size)"
        )

    async for session in get_session():
        async with session.begin():
            # Process in batches
            for batch in _batched(data, effective_batch_size):
                # Normalize batch values: coerce booleans for integer-backed fields
                try:
                    # Batch can be a list of dicts
                    if isinstance(batch, list):
                        for row in batch:
                            if isinstance(row, dict):
                                if "in_collection" in row and isinstance(
                                    row["in_collection"], bool
                                ):
                                    row["in_collection"] = (
                                        1 if row["in_collection"] else 0
                                    )
                                if "file_exists" in row and isinstance(
                                    row["file_exists"], bool
                                ):
                                    row["file_exists"] = 1 if row["file_exists"] else 0
                    elif isinstance(batch, dict):
                        if "in_collection" in batch and isinstance(
                            batch["in_collection"], bool
                        ):
                            batch["in_collection"] = 1 if batch["in_collection"] else 0
                        if "file_exists" in batch and isinstance(
                            batch["file_exists"], bool
                        ):
                            batch["file_exists"] = 1 if batch["file_exists"] else 0
                except Exception:
                    # Never fail the whole batch normalization; log and continue
                    logger.debug(
                        "Failed to normalize boolean fields in batch; continuing without coercion"
                    )

                if on_conflict and update_fields:
                    # PostgreSQL upsert
                    stmt = pg_insert(model.__table__).values(batch)

                    # Only include fields that are actually present in the batch data
                    available_fields = set()
                    if isinstance(batch, list) and batch:
                        available_fields = set(batch[0].keys()) if batch[0] else set()
                    elif isinstance(batch, dict):
                        available_fields = set(batch.keys())

                    # Filter update_fields to only include fields present in the data
                    valid_update_fields = [
                        field for field in update_fields if field in available_fields
                    ]

                    if valid_update_fields:
                        stmt = stmt.on_conflict_do_update(
                            index_elements=on_conflict,
                            set_={
                                field: stmt.excluded[field]
                                for field in valid_update_fields
                            },
                        )
                    await session.execute(stmt)
                else:
                    # Simple insert
                    await session.execute(insert(model.__table__), batch)


async def bulk_update(
    model,
    objects: Iterable[Any] | None = None,
    updates: Iterable[dict[str, Any]] | None = None,
    fields: list[str] | None = None,
) -> int:
    """Bulk update rows using SQLAlchemy.

    Args:
        model: SQLAlchemy ORM model class
        objects: Iterable of model instances with updated values
        updates: Iterable of dicts with primary key + updated values
        fields: List of field names to update (optional, updates all if None)

    Returns:
        Number of rows updated

    Example:
        await bulk_update(MyModel, updates=[{"id": 1, "name": "updated"}])
        await bulk_update(MyModel, objects=[my_instance], fields=["name"])
    """
    if objects is not None:
        data = []
        for obj in objects:
            if hasattr(obj, "__dict__"):
                # Get primary key name
                pk_cols = [col.name for col in model.__table__.primary_key.columns]
                row = {pk: getattr(obj, pk, None) for pk in pk_cols}

                # Add specified fields or all non-pk fields
                if fields:
                    for field in fields:
                        row[field] = getattr(obj, field, None)
                else:
                    for k, v in obj.__dict__.items():
                        if not k.startswith("_") and k not in pk_cols:
                            row[k] = v
                data.append(row)
            else:
                raise ValueError(f"Object {obj} does not have __dict__")
    elif updates is not None:
        data = list(updates)
    else:
        raise ValueError("Either objects or updates must be provided")

    if not data:
        return 0

    count = 0
    async for session in get_session():
        async with session.begin():
            # SQLAlchemy bulk_update_mappings for performance
            await session.execute(update(model.__table__), data)
            count = len(data)

    return count


async def get_or_create(
    model,
    defaults: dict[str, Any] | None = None,
    **filter_kwargs,
) -> tuple[Any, bool]:
    """Get an object or create it if it doesn't exist.

    Args:
        model: SQLAlchemy ORM model class
        defaults: Dict of default values for creation
        **filter_kwargs: Filter criteria

    Returns:
        Tuple of (instance, created) where created is True if newly created

    Example:
        user, created = await get_or_create(
            User,
            defaults={"email": "user@example.com"},
            username="john"
        )
    """
    async for session in get_session():
        # Try to get existing
        filters = [getattr(model, k) == v for k, v in filter_kwargs.items()]
        stmt = select(model).where(and_(*filters))
        result = await session.execute(stmt)
        instance = result.scalar_one_or_none()

        if instance:
            return instance, False

        # Create new instance
        create_kwargs = {**filter_kwargs, **(defaults or {})}
        instance = model(**create_kwargs)

        async with session.begin():
            session.add(instance)
            try:
                await session.flush()
            except IntegrityError:
                # Race condition: another process created it
                await session.rollback()
                result = await session.execute(stmt)
                instance = result.scalar_one_or_none()
                if instance:
                    return instance, False
                raise

        return instance, True

    # Should never reach here, but satisfy type checker
    raise RuntimeError("Session generator did not yield")


async def update_or_create(
    model,
    defaults: dict[str, Any] | None = None,
    **filter_kwargs,
) -> tuple[Any, bool]:
    """Update an object or create it if it doesn't exist.

    Args:
        model: SQLAlchemy ORM model class
        defaults: Dict of values to set (used for both update and create)
        **filter_kwargs: Filter criteria

    Returns:
        Tuple of (instance, created) where created is True if newly created

    Example:
        user, created = await update_or_create(
            User,
            defaults={"email": "new@example.com"},
            username="john"
        )
    """
    async for session in get_session():
        # Try to get existing
        filters = [getattr(model, k) == v for k, v in filter_kwargs.items()]
        stmt = select(model).where(and_(*filters))
        result = await session.execute(stmt)
        instance = result.scalar_one_or_none()

        if instance:
            # Update existing
            if defaults:
                for key, value in defaults.items():
                    setattr(instance, key, value)
                async with session.begin():
                    await session.flush()
            return instance, False

        # Create new
        create_kwargs = {**filter_kwargs, **(defaults or {})}
        instance = model(**create_kwargs)

        async with session.begin():
            session.add(instance)
            await session.flush()

        return instance, True

    # Should never reach here, but satisfy type checker
    raise RuntimeError("Session generator did not yield")


async def get_by_id(model, id_value) -> Any | None:
    """Get an object by primary key.

    Args:
        model: SQLAlchemy ORM model class
        id_value: Primary key value

    Returns:
        Model instance or None
    """
    async for session in get_session():
        return await session.get(model, id_value)


async def get_one(model, **filter_kwargs) -> Any:
    """Get exactly one object matching the filter, raise if not found or multiple.

    Args:
        model: SQLAlchemy ORM model class
        **filter_kwargs: Filter criteria

    Returns:
        Model instance

    Raises:
        NoResultFound: If no matching object
        MultipleResultsFound: If multiple matching objects
    """
    async for session in get_session():
        filters = [getattr(model, k) == v for k, v in filter_kwargs.items()]
        stmt = select(model).where(and_(*filters))
        result = await session.execute(stmt)
        return result.scalar_one()


async def get_or_none(model, **filter_kwargs) -> Any | None:
    """Get an object matching the filter or None if not found.

    Args:
        model: SQLAlchemy ORM model class
        **filter_kwargs: Filter criteria

    Returns:
        Model instance or None
    """
    async for session in get_session():
        filters = [getattr(model, k) == v for k, v in filter_kwargs.items()]
        stmt = select(model).where(and_(*filters))
        result = await session.execute(stmt)
        return result.scalar_one_or_none()


async def filter_all(model, **filter_kwargs) -> list[Any]:
    """Get all objects matching the filter.

    Args:
        model: SQLAlchemy ORM model class
        **filter_kwargs: Filter criteria

    Returns:
        List of model instances
    """
    async for session in get_session():
        filters = [getattr(model, k) == v for k, v in filter_kwargs.items()]
        stmt = select(model).where(and_(*filters)) if filters else select(model)
        result = await session.execute(stmt)
        return list(result.scalars().all())

    # Should never reach here, but satisfy type checker
    return []


async def filter_values(model, *fields, **filter_kwargs) -> list[dict[str, Any]]:
    """Get specific fields as dicts for objects matching the filter.

    Args:
        model: SQLAlchemy ORM model class
        *fields: Field names to retrieve
        **filter_kwargs: Filter criteria

    Returns:
        List of dicts with requested fields

    Example:
        results = await filter_values(User, "id", "name", is_active=True)
        # Returns: [{"id": 1, "name": "John"}, ...]
    """
    async for session in get_session():
        # Build column selection
        cols = [getattr(model, f) for f in fields] if fields else [model]

        # Build filters
        filters = [getattr(model, k) == v for k, v in filter_kwargs.items()]
        stmt = select(*cols).where(and_(*filters)) if filters else select(*cols)

        result = await session.execute(stmt)

        if fields:
            # Return as list of dicts
            return [dict(zip(fields, row, strict=False)) for row in result.all()]
        else:
            # Return full objects as dicts
            return [obj.__dict__ for obj in result.scalars().all()]

    # Should never reach here, but satisfy type checker
    return []


async def all_values(model, *fields) -> list[dict[str, Any]]:
    """Get all objects as dicts with specific fields.

    Args:
        model: SQLAlchemy ORM model class
        *fields: Field names to retrieve (empty = all fields)

    Returns:
        List of dicts

    Example:
        results = await all_values(User, "id", "name")
    """
    return await filter_values(model, *fields)


async def count(model, **filter_kwargs) -> int:
    """Count objects matching the filter.

    Args:
        model: SQLAlchemy ORM model class
        **filter_kwargs: Filter criteria

    Returns:
        Count of matching objects
    """
    async for session in get_session():
        from sqlalchemy import func

        filters = [getattr(model, k) == v for k, v in filter_kwargs.items()]
        stmt = (
            select(func.count()).select_from(model).where(and_(*filters))
            if filters
            else select(func.count()).select_from(model)
        )
        result = await session.execute(stmt)
        return result.scalar_one()

    # Should never reach here, but satisfy type checker
    return 0


async def exists(model, **filter_kwargs) -> bool:
    """Check if any objects match the filter.

    Args:
        model: SQLAlchemy ORM model class
        **filter_kwargs: Filter criteria

    Returns:
        True if at least one matching object exists
    """
    return await count(model, **filter_kwargs) > 0


async def delete_where(model, **filter_kwargs) -> int:
    """Delete objects matching the filter.

    Args:
        model: SQLAlchemy ORM model class
        **filter_kwargs: Filter criteria

    Returns:
        Number of rows deleted
    """
    async for session in get_session():
        async with session.begin():
            filters = [getattr(model, k) == v for k, v in filter_kwargs.items()]
            stmt = delete(model).where(and_(*filters)) if filters else delete(model)
            result = await session.execute(stmt)
            return result.rowcount  # type: ignore[return-value]

    # Should never reach here, but satisfy type checker
    return 0


async def create_instance(model, **kwargs) -> Any:
    """Create and persist a new instance.

    Args:
        model: SQLAlchemy ORM model class
        **kwargs: Field values

    Returns:
        Created model instance
    """
    async for session in get_session():
        instance = model(**kwargs)
        async with session.begin():
            session.add(instance)
            await session.flush()
            await session.refresh(instance)
        return instance


async def save_instance(instance: Any) -> None:
    """Save changes to an existing instance.

    Args:
        instance: Model instance to save
    """
    async for session in get_session():
        async with session.begin():
            session.add(instance)
            await session.flush()


@asynccontextmanager
async def transaction():
    """Context manager for database transactions.

    Example:
        async with transaction():
            await create_instance(User, name="John")
            await create_instance(User, name="Jane")
            # Both or neither will be committed
    """
    async for session in get_session():
        async with session.begin():
            yield session


def _batched(iterable: Iterable[Any], n: int) -> Iterable[list[Any]]:
    """Batch an iterable into chunks of size n."""
    from itertools import islice

    it = iter(iterable)
    while True:
        batch = list(islice(it, n))
        if not batch:
            return
        yield batch
