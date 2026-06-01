"""add_gdpr_consent_to_candidates

Revision ID: 929b9bab15de
Revises: bc44b021d1f9
Create Date: 2026-06-01 19:21:28.796999

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = '929b9bab15de'
down_revision: Union[str, Sequence[str], None] = 'bc44b021d1f9'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    with op.batch_alter_table('candidates', schema=None) as batch_op:
        batch_op.add_column(sa.Column('gdpr_consent', sa.Integer(), nullable=True, server_default='0'))
        batch_op.add_column(sa.Column('gdpr_consent_at', sa.DateTime(), nullable=True))


def downgrade() -> None:
    with op.batch_alter_table('candidates', schema=None) as batch_op:
        batch_op.drop_column('gdpr_consent_at')
        batch_op.drop_column('gdpr_consent')
