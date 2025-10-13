"""
Add missing fields to staged_copyright_item table

Revision ID: d06fa1051710
Revises: f2355a540318
Create Date: 2025-10-13 22:08:15.377703
"""

# revision identifiers, used by Alembic.
revision = "d06fa1051710"
down_revision = "f2355a540318"
branch_labels = None
depends_on = None

import sqlalchemy as sa
from alembic import op


def upgrade() -> None:
    # Add missing fields to staged_copyright_item table
    op.add_column(
        "staged_copyright_item",
        sa.Column("id_course", sa.String(length=255), nullable=True),
    )
    op.add_column(
        "staged_copyright_item",
        sa.Column("id_material", sa.String(length=255), nullable=True),
    )
    op.add_column(
        "staged_copyright_item",
        sa.Column("last_scan_date_university", sa.String(length=255), nullable=True),
    )
    op.add_column(
        "staged_copyright_item",
        sa.Column("filehash", sa.String(length=255), nullable=True),
    )
    op.add_column(
        "staged_copyright_item",
        sa.Column(
            "manual_classification_report", sa.String(length=2048), nullable=True
        ),
    )
    op.add_column(
        "staged_copyright_item",
        sa.Column("count_downloads_material", sa.String(length=255), nullable=True),
    )
    op.add_column(
        "staged_copyright_item",
        sa.Column("last_scan_date_course", sa.String(length=255), nullable=True),
    )


def downgrade() -> None:
    # Remove the added fields from staged_copyright_item table
    op.drop_column("staged_copyright_item", "last_scan_date_course")
    op.drop_column("staged_copyright_item", "count_downloads_material")
    op.drop_column("staged_copyright_item", "manual_classification_report")
    op.drop_column("staged_copyright_item", "filehash")
    op.drop_column("staged_copyright_item", "last_scan_date_university")
    op.drop_column("staged_copyright_item", "id_material")
    op.drop_column("staged_copyright_item", "id_course")
