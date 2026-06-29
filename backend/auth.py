import os
import bcrypt
import psycopg
from datetime import datetime, timedelta
from typing import Optional
from jose import JWTError, jwt

SECRET_KEY = os.environ.get("SECRET_KEY", "infotel-cv-generator-secret-key-2024")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_HOURS = 8
DATABASE_URL = os.environ.get("DATABASE_URL", "")


def get_conn():
    return psycopg.connect(DATABASE_URL, row_factory=psycopg.rows.dict_row)


def init_db():
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user',
                created_at TIMESTAMP DEFAULT NOW(),
                last_login TIMESTAMP,
                total_generated INTEGER DEFAULT 0,
                monthly_generated INTEGER DEFAULT 0
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS history (
                id SERIAL PRIMARY KEY,
                username TEXT NOT NULL,
                filename TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT NOW()
            )
        """)
        row = conn.execute("SELECT 1 FROM users WHERE username = 'admin'").fetchone()
        if not row:
            hashed = bcrypt.hashpw("admin123".encode(), bcrypt.gensalt()).decode()
            conn.execute(
                "INSERT INTO users (username, password, role) VALUES (%s, %s, %s)",
                ("admin", hashed, "admin")
            )
        conn.commit()


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()


def verify_password(plain: str, hashed: str) -> bool:
    return bcrypt.checkpw(plain.encode(), hashed.encode())


def get_user(username: str):
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM users WHERE username = %s", (username,)).fetchone()
        return dict(row) if row else None


def authenticate_user(username: str, password: str):
    user = get_user(username)
    if not user or not verify_password(password, user["password"]):
        return None
    with get_conn() as conn:
        conn.execute("UPDATE users SET last_login = NOW() WHERE username = %s", (username,))
        conn.commit()
    return user


def create_token(username: str):
    expire = datetime.utcnow() + timedelta(hours=ACCESS_TOKEN_EXPIRE_HOURS)
    return jwt.encode({"sub": username, "exp": expire}, SECRET_KEY, algorithm=ALGORITHM)


def decode_token(token: str) -> Optional[str]:
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        return payload.get("sub")
    except JWTError:
        return None


def log_generation(username: str, filename: str):
    with get_conn() as conn:
        conn.execute("INSERT INTO history (username, filename) VALUES (%s, %s)", (username, filename))
        conn.execute(
            "UPDATE users SET total_generated = total_generated + 1, monthly_generated = monthly_generated + 1 WHERE username = %s",
            (username,)
        )
        conn.commit()


def add_user(username: str, password: str, role: str = "user"):
    try:
        with get_conn() as conn:
            conn.execute(
                "INSERT INTO users (username, password, role) VALUES (%s, %s, %s)",
                (username, hash_password(password), role)
            )
            conn.commit()
        return True, "OK"
    except Exception:
        return False, "Utilisateur déjà existant"


def delete_user(username: str):
    if username == "admin":
        return False, "Impossible de supprimer le compte admin principal"
    with get_conn() as conn:
        cur = conn.execute("DELETE FROM users WHERE username = %s", (username,))
        conn.commit()
        if cur.rowcount == 0:
            return False, "Utilisateur introuvable"
    return True, "OK"


def get_all_users():
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT username, role, total_generated, monthly_generated,
                   last_login::text as last_login, created_at::text as created_at
            FROM users ORDER BY created_at
        """).fetchall()
        return [dict(r) for r in rows]


def get_stats():
    with get_conn() as conn:
        stats = dict(conn.execute("""
            SELECT COUNT(*) as total_users,
                   COALESCE(SUM(total_generated),0) as total_generated,
                   COALESCE(SUM(monthly_generated),0) as monthly_generated
            FROM users
        """).fetchone())
        history = [dict(r) for r in conn.execute("""
            SELECT username, filename, created_at::text as date,
                   TO_CHAR(created_at, 'DD/MM/YYYY HH24:MI') as date_display
            FROM history ORDER BY created_at DESC LIMIT 20
        """).fetchall()]
    return {
        "total_users": stats["total_users"],
        "total_generated": stats["total_generated"],
        "monthly_generated": stats["monthly_generated"],
        "recent_history": history
    }


try:
    init_db()
    print("[DB] Base de données initialisée avec succès")
except Exception as e:
    print(f"[DB] Erreur d'initialisation: {e}")
