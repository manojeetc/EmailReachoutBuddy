"""One-time, idempotent transfer of local SQLite records into Supabase PostgreSQL."""

import sqlite3
from pathlib import Path

from storage import connect_database


ROOT = Path(__file__).resolve().parent
TABLES = ('professors', 'drafts', 'outbound', 'inbound', 'settings')
IDENTITY_TABLES = ('professors', 'outbound', 'inbound')


def main():
    source_path = ROOT / 'outreach.sqlite3'
    if not source_path.exists():
        raise SystemExit(f'Local SQLite database not found: {source_path.name}')

    source = sqlite3.connect(source_path)
    source.row_factory = sqlite3.Row
    totals = {}
    try:
        with connect_database() as target:
            for table in TABLES:
                rows = source.execute(f'SELECT * FROM {table}').fetchall()
                if not rows:
                    totals[table] = (0, 0)
                    continue
                columns = [column['name'] for column in source.execute(f'PRAGMA table_info({table})')]
                column_sql = ', '.join(f'"{column}"' for column in columns)
                placeholders = ', '.join('?' for _ in columns)
                query = f'INSERT INTO public.{table} ({column_sql}) VALUES ({placeholders}) ON CONFLICT DO NOTHING'
                inserted = 0
                for row in rows:
                    cursor = target.execute(query, tuple(row[column] for column in columns))
                    inserted += cursor.rowcount
                totals[table] = (len(rows), inserted)

            for table in IDENTITY_TABLES:
                target.execute(
                    f"SELECT setval(pg_get_serial_sequence('public.{table}', 'id'), "
                    f"COALESCE((SELECT MAX(id) FROM public.{table}), 1), "
                    f"EXISTS (SELECT 1 FROM public.{table}))"
                )
    finally:
        source.close()

    for table, (read, inserted) in totals.items():
        print(f'{table}: {inserted} inserted; {read - inserted} already present; {read} local total')
    print('Migration completed. The local SQLite database was left unchanged.')


if __name__ == '__main__':
    main()
