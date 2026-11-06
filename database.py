
import os
from pathlib import Path
from dotenv import load_dotenv

from sqlalchemy import create_engine, event
from sqlalchemy.orm import declarative_base, sessionmaker

BASE_DIR = Path(__file__).resolve().parent
load_dotenv()

TURSO_URL = os.environ.get("TURSO_DATABASE_URL")
TURSO_TOKEN = os.environ.get("TURSO_AUTH_TOKEN")
DATABASE_URL = os.environ.get("DATABASE_URL")

connect_args = {}

if TURSO_URL:
    # Conexión remota a Turso
    host = TURSO_URL.replace("libsql://", "").replace("https://", "").rstrip("/")

    SQLALCHEMY_DATABASE_URL = f"sqlite+libsql://{host}?secure=true"

    connect_args = {
        "auth_token": TURSO_TOKEN
    }

    _es_sqlite = True

elif DATABASE_URL:
    SQLALCHEMY_DATABASE_URL = DATABASE_URL

    if SQLALCHEMY_DATABASE_URL.startswith("postgres://"):
        SQLALCHEMY_DATABASE_URL = SQLALCHEMY_DATABASE_URL.replace(
            "postgres://", "postgresql+psycopg://", 1
        )
    elif SQLALCHEMY_DATABASE_URL.startswith("postgresql://"):
        SQLALCHEMY_DATABASE_URL = SQLALCHEMY_DATABASE_URL.replace(
            "postgresql://", "postgresql+psycopg://", 1
        )

    _es_sqlite = SQLALCHEMY_DATABASE_URL.startswith("sqlite")

    if _es_sqlite:
        connect_args = {"check_same_thread": False}

else:
    # SQLite para desarrollo local
    INSTANCE_DIR = BASE_DIR / "instance"
    INSTANCE_DIR.mkdir(exist_ok=True)

    SQLALCHEMY_DATABASE_URL = f"sqlite:///{INSTANCE_DIR / 'recetario.db'}"

    connect_args = {"check_same_thread": False}
    _es_sqlite = True

engine = create_engine(
    SQLALCHEMY_DATABASE_URL,
    connect_args=connect_args,
    pool_pre_ping=not _es_sqlite,
)

if _es_sqlite:
    @event.listens_for(engine, "connect")
    def _activar_llaves_foraneas(dbapi_connection, _record):
        cursor = dbapi_connection.cursor()
        cursor.execute("PRAGMA foreign_keys=ON")
        cursor.close()

LocalSession = sessionmaker(
    autocommit=False,
    autoflush=False,
    bind=engine
)

Base = declarative_base()
