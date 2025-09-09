#!/usr/bin/env python3
"""
Failure Inspection and Retry Helper for Staged Processing Failures

This script provides utilities to:
1. Inspect StagedProcessingFailure records
2. Analyze failure patterns
3. Retry failed processing with improved error handling
4. Clean up old failure records

Usage:
    python scripts/failure_helper.py inspect [--limit N] [--material-id ID]
    python scripts/failure_helper.py retry [--material-id ID] [--dry-run]
    python scripts/failure_helper.py cleanup [--days-old N] [--dry-run]
    python scripts/failure_helper.py stats
"""

import argparse
import asyncio
import json
import sys
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

from loguru import logger
from tortoise import Tortoise

# Add the project root to the Python path
sys.path.insert(0, str(Path(__file__).parent.parent))

from easy_access.db.models import (
    StagedCopyrightItem,
    StagedFacultyUpdate,
    StagedProcessingFailure,
)
from easy_access.db.update import (
    process_staged_faculty_updates,
    process_staged_raw_data,
)
from easy_access.settings import Settings


class FailureInspector:
    """Helper class for inspecting and managing StagedProcessingFailure records."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    async def inspect_failures(
        self,
        limit: int = 50,
        material_id: int | None = None,
        show_payload: bool = False,
    ) -> list[dict[str, Any]]:
        """Inspect recent failure records."""
        logger.info(
            f"Inspecting failure records (limit: {limit}, material_id: {material_id})"
        )

        query = StagedProcessingFailure.all().order_by("-created_at")

        if material_id:
            query = query.filter(material_id=material_id)

        failures = await query.limit(limit)

        results = []
        for failure in failures:
            result = {
                "id": failure.id,
                "material_id": failure.material_id,
                "error_message": failure.error_message,
                "created_at": failure.created_at,
                "staged_payload": failure.staged_payload if show_payload else None,
            }
            results.append(result)

        return results

    async def get_failure_stats(self) -> dict[str, Any]:
        """Get statistics about failure records."""
        logger.info("Generating failure statistics")

        total_failures = await StagedProcessingFailure.all().count()

        # Group by error patterns
        failures = await StagedProcessingFailure.all()
        error_patterns = {}

        for failure in failures:
            if failure.error_message:
                # Extract common error patterns
                error_key = self._categorize_error(failure.error_message)
                error_patterns[error_key] = error_patterns.get(error_key, 0) + 1

        # Get failures by material_id
        material_failures = await StagedProcessingFailure.filter(
            material_id__not_isnull=True
        ).count()
        unknown_material_failures = await StagedProcessingFailure.filter(
            material_id__isnull=True
        ).count()

        # Get recent failures (last 24 hours)
        yesterday = datetime.now(UTC) - timedelta(days=1)
        recent_failures = await StagedProcessingFailure.filter(
            created_at__gte=yesterday
        ).count()

        return {
            "total_failures": total_failures,
            "error_patterns": error_patterns,
            "material_failures": material_failures,
            "unknown_material_failures": unknown_material_failures,
            "recent_failures": recent_failures,
        }

    def _categorize_error(self, error_message: str) -> str:
        """Categorize error messages into common patterns."""
        error_lower = error_message.lower()

        if "faculty" in error_lower and (
            "not found" in error_lower or "does not exist" in error_lower
        ):
            return "Faculty Lookup Error"
        elif "material_id" in error_lower and (
            "invalid" in error_lower or "missing" in error_lower
        ):
            return "Invalid Material ID"
        elif "classification" in error_lower:
            return "Classification Error"
        elif "database" in error_lower or "connection" in error_lower:
            return "Database Error"
        elif "permission" in error_lower or "access" in error_lower:
            return "Permission Error"
        elif "timeout" in error_lower:
            return "Timeout Error"
        elif "validation" in error_lower:
            return "Validation Error"
        else:
            return "Other Error"

    async def retry_failures(
        self, material_id: int | None = None, dry_run: bool = True
    ) -> dict[str, Any]:
        """Retry processing of failed records."""
        logger.info(
            f"Retrying failures (material_id: {material_id}, dry_run: {dry_run})"
        )

        query = StagedProcessingFailure.all()

        if material_id:
            query = query.filter(material_id=material_id)

        failures = await query

        results = {
            "total_attempted": len(failures),
            "successful_retries": 0,
            "failed_retries": 0,
            "errors": [],
        }

        for failure in failures:
            try:
                if failure.staged_payload and failure.material_id:
                    # Try to reprocess the staged data
                    success = await self._retry_single_failure(failure, dry_run)
                    if success:
                        results["successful_retries"] += 1
                        if not dry_run:
                            await failure.delete()  # Remove successful retry
                    else:
                        results["failed_retries"] += 1
                        results["errors"].append(
                            f"Failed to retry material_id {failure.material_id}"
                        )
                else:
                    results["failed_retries"] += 1
                    results["errors"].append(
                        f"Missing payload or material_id for failure {failure.id}"
                    )

            except Exception as e:
                results["failed_retries"] += 1
                results["errors"].append(
                    f"Error retrying failure {failure.id}: {str(e)}"
                )

        return results

    async def _retry_single_failure(
        self, failure: StagedProcessingFailure, dry_run: bool
    ) -> bool:
        """Retry processing a single failure."""
        try:
            material_id = failure.material_id
            payload = failure.staged_payload

            if not payload or not material_id:
                return False

            # Check if the original staged record still exists
            staged_raw = await StagedCopyrightItem.filter(
                material_id=material_id
            ).first()
            staged_faculty = await StagedFacultyUpdate.filter(
                material_id=material_id
            ).first()

            if staged_raw:
                # Retry raw data processing
                if not dry_run:
                    await process_staged_raw_data(self.settings)
                return True
            elif staged_faculty:
                # Retry faculty update processing
                if not dry_run:
                    await process_staged_faculty_updates(self.settings)
                return True
            else:
                logger.warning(f"No staged record found for material_id {material_id}")
                return False

        except Exception as e:
            logger.error(
                f"Error retrying failure for material_id {failure.material_id}: {str(e)}"
            )
            return False

    async def cleanup_old_failures(
        self, days_old: int = 30, dry_run: bool = True
    ) -> dict[str, Any]:
        """Clean up old failure records."""
        logger.info(
            f"Cleaning up failures older than {days_old} days (dry_run: {dry_run})"
        )

        cutoff_date = datetime.now(UTC) - timedelta(days=days_old)

        query = StagedProcessingFailure.filter(created_at__lt=cutoff_date)
        old_failures = await query

        results = {
            "cutoff_date": cutoff_date,
            "failures_to_delete": len(old_failures),
            "deleted_count": 0,
        }

        if not dry_run and old_failures:
            deleted_count = await query.delete()
            results["deleted_count"] = deleted_count

        return results


async def main():
    """Main entry point for the failure helper script."""
    parser = argparse.ArgumentParser(description="Staged Processing Failure Helper")
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # Inspect command
    inspect_parser = subparsers.add_parser("inspect", help="Inspect failure records")
    inspect_parser.add_argument(
        "--limit", type=int, default=50, help="Limit number of records"
    )
    inspect_parser.add_argument("--material-id", type=int, help="Filter by material ID")
    inspect_parser.add_argument(
        "--show-payload", action="store_true", help="Show staged payload"
    )

    # Stats command
    subparsers.add_parser("stats", help="Show failure statistics")

    # Retry command
    retry_parser = subparsers.add_parser("retry", help="Retry failed processing")
    retry_parser.add_argument(
        "--material-id", type=int, help="Retry specific material ID"
    )
    retry_parser.add_argument(
        "--dry-run", action="store_true", help="Show what would be done"
    )

    # Cleanup command
    cleanup_parser = subparsers.add_parser(
        "cleanup", help="Clean up old failure records"
    )
    cleanup_parser.add_argument(
        "--days-old", type=int, default=30, help="Delete records older than N days"
    )
    cleanup_parser.add_argument(
        "--dry-run", action="store_true", help="Show what would be done"
    )

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return

    # Initialize database
    settings = Settings()
    await Tortoise.init(
        db_url=f"sqlite://{settings.db_path}",
        modules={"models": ["easy_access.db.models"]},
    )

    try:
        inspector = FailureInspector(settings)

        if args.command == "inspect":
            failures = await inspector.inspect_failures(
                limit=args.limit,
                material_id=args.material_id,
                show_payload=args.show_payload,
            )

            if failures:
                print(f"\nFound {len(failures)} failure records:")
                print("-" * 80)

                for failure in failures:
                    print(f"ID: {failure['id']}")
                    print(f"Material ID: {failure['material_id']}")
                    print(f"Created: {failure['created_at']}")
                    print(f"Error: {failure['error_message']}")

                    if args.show_payload and failure["staged_payload"]:
                        print(
                            f"Payload: {json.dumps(failure['staged_payload'], indent=2)}"
                        )

                    print("-" * 80)
            else:
                print("No failure records found.")

        elif args.command == "stats":
            stats = await inspector.get_failure_stats()

            print("\nFailure Statistics:")
            print("-" * 40)
            print(f"Total failures: {stats['total_failures']}")
            print(f"Failures with material_id: {stats['material_failures']}")
            print(f"Failures without material_id: {stats['unknown_material_failures']}")
            print(f"Recent failures (24h): {stats['recent_failures']}")

            if stats["error_patterns"]:
                print("\nError Patterns:")
                for pattern, count in stats["error_patterns"].items():
                    print(f"  {pattern}: {count}")

        elif args.command == "retry":
            results = await inspector.retry_failures(
                material_id=args.material_id, dry_run=args.dry_run
            )

            print(f"\nRetry Results ({'DRY RUN' if args.dry_run else 'LIVE'}):")
            print("-" * 40)
            print(f"Total attempted: {results['total_attempted']}")
            print(f"Successful retries: {results['successful_retries']}")
            print(f"Failed retries: {results['failed_retries']}")

            if results["errors"]:
                print("\nErrors:")
                for error in results["errors"][:10]:  # Show first 10 errors
                    print(f"  {error}")
                if len(results["errors"]) > 10:
                    print(f"  ... and {len(results['errors']) - 10} more errors")

        elif args.command == "cleanup":
            results = await inspector.cleanup_old_failures(
                days_old=args.days_old, dry_run=args.dry_run
            )

            print(f"\nCleanup Results ({'DRY RUN' if args.dry_run else 'LIVE'}):")
            print("-" * 40)
            print(f"Cutoff date: {results['cutoff_date']}")
            print(f"Failures to delete: {results['failures_to_delete']}")
            print(f"Actually deleted: {results['deleted_count']}")

    finally:
        await Tortoise.close_connections()


if __name__ == "__main__":
    asyncio.run(main())
