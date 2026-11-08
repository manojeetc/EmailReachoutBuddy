"""Apply the checked-in Supabase schema using the private .env connection settings."""

from pathlib import Path

from storage import connect_database


ROOT = Path(__file__).resolve().parent


def main():
    schema_path = ROOT / 'supabase' / 'schema.sql'
    statements = [part.strip() for part in schema_path.read_text(encoding='utf-8').split(';') if part.strip()]
    with connect_database() as connection:
        for statement in statements:
            connection.execute(statement)
    print(f'Applied {len(statements)} schema statements. The Supabase tables are ready.')


if __name__ == '__main__':
    main()
