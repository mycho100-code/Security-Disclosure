"""
DB Helper - SQLite 데이터베이스 초기화 및 CRUD 헬퍼 모듈
(실제 엑셀 컬럼 구조 반영)
"""
import sqlite3
import pandas as pd
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "it_asset.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    # ── 기준정보-자산 ──
    cur.execute("""
        CREATE TABLE IF NOT EXISTS master_asset (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            asset_number    TEXT,
            description     TEXT,
            cost_center     TEXT,
            serial_number   TEXT,
            cost_center2    TEXT,
            asset_class     TEXT,
            depreciation    REAL DEFAULT 0,
            exclude_yn      TEXT DEFAULT '',
            it_yn           TEXT DEFAULT '',
            sec_yn          TEXT DEFAULT ''
        )
    """)

    # ── 기준정보-비용 ──
    cur.execute("""
        CREATE TABLE IF NOT EXISTS master_cost (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            gl_date         TEXT,
            profit_center   TEXT,
            cost_center     TEXT,
            account         TEXT,
            account_name    TEXT,
            business_area   TEXT,
            doc_type        TEXT,
            description     TEXT,
            amount          REAL DEFAULT 0,
            exclude_yn      TEXT DEFAULT '',
            it_yn           TEXT DEFAULT '',
            sec_yn          TEXT DEFAULT ''
        )
    """)

    # ── 분석용 임시-자산 ──
    cur.execute("""
        CREATE TABLE IF NOT EXISTS target_asset (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            asset_number    TEXT,
            description     TEXT,
            cost_center     TEXT,
            serial_number   TEXT,
            cost_center2    TEXT,
            asset_class     TEXT,
            depreciation    REAL DEFAULT 0,
            exclude_yn      TEXT DEFAULT '',
            it_yn           TEXT DEFAULT '',
            sec_yn          TEXT DEFAULT '',
            match_type      TEXT DEFAULT '',
            match_score     REAL DEFAULT 0,
            matched_desc    TEXT DEFAULT ''
        )
    """)

    # ── 분석용 임시-비용 ──
    cur.execute("""
        CREATE TABLE IF NOT EXISTS target_cost (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            gl_date         TEXT,
            profit_center   TEXT,
            cost_center     TEXT,
            account         TEXT,
            account_name    TEXT,
            business_area   TEXT,
            doc_type        TEXT,
            description     TEXT,
            amount          REAL DEFAULT 0,
            exclude_yn      TEXT DEFAULT '',
            it_yn           TEXT DEFAULT '',
            sec_yn          TEXT DEFAULT '',
            match_type      TEXT DEFAULT '',
            match_score     REAL DEFAULT 0,
            matched_desc    TEXT DEFAULT ''
        )
    """)

    conn.commit()
    conn.close()

    # ── 사용자 테이블 & 기본 Admin 계정 ──
    _init_users()


def _init_users():
    """사용자 테이블 생성 및 기본 Admin 계정 삽입"""
    import hashlib
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role     TEXT NOT NULL DEFAULT 'user',
            name     TEXT DEFAULT ''
        )
    """)
    # Admin 계정이 없으면 생성
    cur.execute("SELECT COUNT(*) FROM users WHERE username = 'Admin'")
    if cur.fetchone()[0] == 0:
        pw_hash = hashlib.sha256("Admin".encode()).hexdigest()
        cur.execute(
            "INSERT INTO users (username, password, role, name) VALUES (?, ?, ?, ?)",
            ("Admin", pw_hash, "admin", "관리자")
        )
    conn.commit()
    conn.close()


def verify_user(username: str, password: str) -> dict | None:
    """로그인 인증. 성공 시 {id, username, role, name} 반환, 실패 시 None"""
    import hashlib
    pw_hash = hashlib.sha256(password.encode()).hexdigest()
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, username, role, name FROM users WHERE username = ? AND password = ?",
                (username, pw_hash))
    row = cur.fetchone()
    conn.close()
    if row:
        return {"id": row[0], "username": row[1], "role": row[2], "name": row[3]}
    return None


def get_all_users() -> pd.DataFrame:
    """전체 사용자 목록 조회 (비밀번호 제외)"""
    conn = get_conn()
    df = pd.read_sql("SELECT id, username, role, name FROM users", conn)
    conn.close()
    return df


def add_user(username: str, password: str, role: str = "user", name: str = "") -> bool:
    """사용자 추가. 성공 True, 중복 시 False"""
    import hashlib
    pw_hash = hashlib.sha256(password.encode()).hexdigest()
    conn = get_conn()
    try:
        conn.execute(
            "INSERT INTO users (username, password, role, name) VALUES (?, ?, ?, ?)",
            (username, pw_hash, role, name)
        )
        conn.commit()
        conn.close()
        return True
    except sqlite3.IntegrityError:
        conn.close()
        return False


def delete_user(user_id: int):
    """사용자 삭제 (Admin 제외)"""
    conn = get_conn()
    conn.execute("DELETE FROM users WHERE id = ? AND username != 'Admin'", (int(user_id),))
    conn.commit()
    conn.close()


def reset_password(user_id: int, new_password: str):
    """비밀번호 초기화"""
    import hashlib
    pw_hash = hashlib.sha256(new_password.encode()).hexdigest()
    conn = get_conn()
    conn.execute("UPDATE users SET password = ? WHERE id = ?", (pw_hash, int(user_id)))
    conn.commit()
    conn.close()


def load_table(table_name: str) -> pd.DataFrame:
    conn = get_conn()
    df = pd.read_sql(f"SELECT * FROM {table_name}", conn)
    conn.close()
    return df


def insert_row(table_name: str, data: dict):
    conn = get_conn()
    cols = ", ".join(data.keys())
    placeholders = ", ".join(["?"] * len(data))
    conn.execute(f"INSERT INTO {table_name} ({cols}) VALUES ({placeholders})", list(data.values()))
    conn.commit()
    conn.close()


def update_row(table_name: str, row_id: int, data: dict):
    conn = get_conn()
    set_clause = ", ".join([f"{k} = ?" for k in data.keys()])
    conn.execute(f"UPDATE {table_name} SET {set_clause} WHERE id = ?", list(data.values()) + [row_id])
    conn.commit()
    conn.close()


def delete_rows(table_name: str, ids: list):
    conn = get_conn()
    placeholders = ",".join(["?"] * len(ids))
    conn.execute(f"DELETE FROM {table_name} WHERE id IN ({placeholders})", ids)
    conn.commit()
    conn.close()


def replace_all(table_name: str, df: pd.DataFrame):
    """테이블 전체 삭제 후 DataFrame 일괄 삽입 (ID 자동증가 초기화 포함)"""
    conn = get_conn()
    conn.execute(f"DELETE FROM {table_name}")
    # AUTOINCREMENT 카운터 초기화 (id가 1부터 다시 시작)
    conn.execute(f"DELETE FROM sqlite_sequence WHERE name=?", (table_name,))
    df.to_sql(table_name, conn, if_exists="append", index=False)
    conn.commit()
    conn.close()


def truncate_table(table_name: str):
    conn = get_conn()
    conn.execute(f"DELETE FROM {table_name}")
    conn.commit()
    conn.close()


def bulk_update_classifications(table_name: str, updates: list[dict]):
    conn = get_conn()
    cur = conn.cursor()
    for u in updates:
        cur.execute(f"""
            UPDATE {table_name}
            SET it_yn = ?, sec_yn = ?, exclude_yn = ?,
                match_type = ?, match_score = ?, matched_desc = ?
            WHERE id = ?
        """, (u["it_yn"], u["sec_yn"], u["exclude_yn"],
              u["match_type"], u["match_score"], u["matched_desc"], u["id"]))
    conn.commit()
    conn.close()


def reset_classifications(table_name: str):
    """분석 실행 전 모든 분류값·매칭 결과를 초기화 (이전 분석 잔존 방지)"""
    conn = get_conn()
    conn.execute(f"""
        UPDATE {table_name}
        SET exclude_yn = '', it_yn = '', sec_yn = '',
            match_type = '', match_score = 0, matched_desc = ''
    """)
    conn.commit()
    conn.close()
