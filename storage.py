"""PostgreSQL connection setup for the Supabase-backed Flask app."""

import os
from pathlib import Path
from urllib.parse import unquote, urlsplit

import psycopg
from psycopg.conninfo import conninfo_to_dict
from psycopg.rows import dict_row
from dotenv import load_dotenv


load_dotenv(Path(__file__).resolve().parent / '.env')


class QmarkConnection(psycopg.Connection):
    """Allow the app's existing parameterized SQL to use qmark placeholders."""

    def execute(self, query, params=None, *, prepare=None, binary=False):
        if isinstance(query, str):
            query = query.replace('?', '%s')
        return super().execute(query, params, prepare=prepare, binary=binary)


def connect_database():
    connection_string = os.getenv('SUPABASE_DB_CON_STRING', '').strip()
    if connection_string:
        if connection_string.startswith(('postgres://', 'postgresql://')):
            _, remainder = connection_string.split('://', 1)
            authority, separator, path_and_query = remainder.partition('/')
            userinfo, at_separator, host_and_port = authority.rpartition('@')
            username, colon_separator, password = userinfo.partition(':')
            parsed_host = urlsplit('//' + host_and_port)
            database = path_and_query.partition('?')[0]
            database = unquote(database)
            if not (separator and at_separator and colon_separator and username and parsed_host.hostname and database):
                raise RuntimeError('SUPABASE_DB_CON_STRING must be a PostgreSQL URI with username, password, host, and database.')
            return QmarkConnection.connect(
                host=parsed_host.hostname,
                port=parsed_host.port or 5432,
                dbname=database,
                user=unquote(username),
                password=unquote(password),
                sslmode='require',
                row_factory=dict_row,
            )

        try:
            parameters = conninfo_to_dict(connection_string)
        except Exception as exc:
            raise RuntimeError('SUPABASE_DB_CON_STRING is not a valid PostgreSQL connection string.') from exc
        if not parameters.get('host') or '@' in parameters['host']:
            raise RuntimeError('The host in SUPABASE_DB_CON_STRING is malformed. URL-encode reserved characters in the password.')
        return QmarkConnection.connect(
            connection_string,
            sslmode='require',
            row_factory=dict_row,
        )

    required = {
        'SUPABASE_DB_HOST': os.getenv('SUPABASE_DB_HOST', ''),
        'SUPABASE_DB_NAME': os.getenv('SUPABASE_DB_NAME', ''),
        'SUPABASE_DB_USER': os.getenv('SUPABASE_DB_USER', ''),
        'SUPABASE_DB_PASSWORD': os.getenv('SUPABASE_DB_PASSWORD', ''),
    }
    missing = [name for name, value in required.items() if not value or value.startswith('REPLACE_')]
    if missing:
        raise RuntimeError('Set ' + ', '.join(missing) + ' in .env using the newly rotated Supabase database credentials.')
    try:
        port = int(os.getenv('SUPABASE_DB_PORT', '5432'))
    except ValueError as exc:
        raise RuntimeError('SUPABASE_DB_PORT must be a number.') from exc
    return QmarkConnection.connect(
        host=required['SUPABASE_DB_HOST'],
        port=port,
        dbname=required['SUPABASE_DB_NAME'],
        user=required['SUPABASE_DB_USER'],
        password=required['SUPABASE_DB_PASSWORD'],
        sslmode='require',
        row_factory=dict_row,
    )
