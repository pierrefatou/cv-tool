import json
import os
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional
from jose import JWTError, jwt
from passlib.context import CryptContext

SECRET_KEY = os.environ.get("SECRET_KEY", "infotel-cv-generator-secret-key-2024")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_HOURS = 8

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
DB_PATH = Path(__file__).parent.parent / "data" / "users.json"


def init_db():
    DB_PATH.parent.mkdir(exist_ok=True)
    if not DB_PATH.exists():
        default = {
            "users": [
                {
                    "username": "admin",
                    "password": pwd_context.hash("admin123"),
                    "role": "admin",
                    "created_at": datetime.now().isoformat(),
                    "total_generated": 0,
                    "monthly_generated": 0,
                    "last_login": None,
                    "history": []
                }
            ]
        }
        DB_PATH.write_text(json.dumps(default, indent=2, ensure_ascii=False))


def load_db():
    init_db()
    return json.loads(DB_PATH.read_text(encoding="utf-8"))


def save_db(data):
    DB_PATH.write_text(json.dumps(data, indent=2, ensure_ascii=False))


def get_user(username: str):
    db = load_db()
    return next((u for u in db["users"] if u["username"] == username), None)


def verify_password(plain, hashed):
    return pwd_context.verify(plain, hashed)


def authenticate_user(username: str, password: str):
    user = get_user(username)
    if not user or not verify_password(password, user["password"]):
        return None
    db = load_db()
    for u in db["users"]:
        if u["username"] == username:
            u["last_login"] = datetime.now().isoformat()
    save_db(db)
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
    db = load_db()
    now = datetime.now()
    for u in db["users"]:
        if u["username"] == username:
            u["total_generated"] = u.get("total_generated", 0) + 1
            u["monthly_generated"] = u.get("monthly_generated", 0) + 1
            if "history" not in u:
                u["history"] = []
            u["history"].insert(0, {
                "filename": filename,
                "date": now.isoformat(),
                "date_display": now.strftime("%d/%m/%Y %H:%M")
            })
            u["history"] = u["history"][:50]
    save_db(db)


def add_user(username: str, password: str, role: str = "user"):
    db = load_db()
    if any(u["username"] == username for u in db["users"]):
        return False, "Utilisateur déjà existant"
    db["users"].append({
        "username": username,
        "password": pwd_context.hash(password),
        "role": role,
        "created_at": datetime.now().isoformat(),
        "total_generated": 0,
        "monthly_generated": 0,
        "last_login": None,
        "history": []
    })
    save_db(db)
    return True, "OK"


def delete_user(username: str):
    if username == "admin":
        return False, "Impossible de supprimer l'admin"
    db = load_db()
    db["users"] = [u for u in db["users"] if u["username"] != username]
    save_db(db)
    return True, "OK"


def get_all_users():
    db = load_db()
    return db["users"]


def get_stats():
    db = load_db()
    users = db["users"]
    total = sum(u.get("total_generated", 0) for u in users)
    monthly = sum(u.get("monthly_generated", 0) for u in users)
    history = []
    for u in users:
        for h in u.get("history", []):
            history.append({**h, "username": u["username"]})
    history.sort(key=lambda x: x["date"], reverse=True)
    return {
        "total_users": len(users),
        "total_generated": total,
        "monthly_generated": monthly,
        "recent_history": history[:20]
    }
