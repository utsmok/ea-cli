"""
${message}

Revision ID: ${up_revision}
Revises: ${down_revision}
Create Date: ${create_date}
"""

# revision identifiers, used by Alembic.
revision = "${up_revision}"
down_revision = ${down_revision}
branch_labels = ${branch_labels}
depends_on = ${depends_on}

from alembic import op
import sqlalchemy as sa


def upgrade() -> None:
    ${upgrades if upgrades else 'pass'}


def downgrade() -> None:
    ${downgrades if downgrades else 'pass'}
