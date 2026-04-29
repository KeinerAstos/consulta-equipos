import os
import psycopg2
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ruta_sistem = os.path.join(BASE_DIR, 'datos', 'SISTEM.xlsx')
ruta_re = os.path.join(BASE_DIR, 'datos', 'reasignacion.xlsx')

def get_connection():
    return psycopg2.connect(
        dbname="neondb",
        user="neondb_owner",
        password="npg_B0ZyzNDGFb3k",
        host="ep-cool-snow-ad0pqcmu-pooler.c-2.us-east-1.aws.neon.tech",
        port="5432",
        sslmode="require"
    )