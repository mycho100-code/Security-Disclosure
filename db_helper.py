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


# ═══════════════════════════════════
# 백업 / 복원
# ═══════════════════════════════════

def export_backup() -> bytes:
    """전체 DB를 JSON으로 백업. 반환: JSON bytes"""
    import json
    conn = get_conn()
    tables = ["master_asset", "master_cost", "target_asset", "target_cost", "users"]
    backup = {}
    for t in tables:
        try:
            df = pd.read_sql(f"SELECT * FROM {t}", conn)
            backup[t] = df.to_dict(orient="records")
        except:
            backup[t] = []
    conn.close()
    backup["_meta"] = {
        "version": "4.0",
        "exported_at": pd.Timestamp.now().isoformat(),
        "tables": list(backup.keys()),
    }
    return json.dumps(backup, ensure_ascii=False, default=str).encode("utf-8")


def import_backup(data: bytes) -> dict:
    """JSON 백업 파일에서 전체 DB 복원. 반환: {테이블명: 건수} 딕셔너리"""
    import json
    backup = json.loads(data.decode("utf-8"))
    result = {}

    conn = get_conn()

    table_map = {
        "master_asset": ["asset_number","description","cost_center","serial_number",
                         "cost_center2","asset_class","depreciation","exclude_yn","it_yn","sec_yn"],
        "master_cost": ["gl_date","profit_center","cost_center","account","account_name",
                        "business_area","doc_type","description","amount","exclude_yn","it_yn","sec_yn"],
        "target_asset": ["asset_number","description","cost_center","serial_number",
                         "cost_center2","asset_class","depreciation","exclude_yn","it_yn","sec_yn",
                         "match_type","match_score","matched_desc"],
        "target_cost": ["gl_date","profit_center","cost_center","account","account_name",
                        "business_area","doc_type","description","amount","exclude_yn","it_yn","sec_yn",
                        "match_type","match_score","matched_desc"],
    }

    for table_name, cols in table_map.items():
        if table_name in backup and backup[table_name]:
            conn.execute(f"DELETE FROM {table_name}")
            df = pd.DataFrame(backup[table_name])
            # DB 컬럼만 추출 (id 제외)
            valid_cols = [c for c in cols if c in df.columns]
            if valid_cols:
                df[valid_cols].to_sql(table_name, conn, if_exists="append", index=False)
                result[table_name] = len(df)
            else:
                result[table_name] = 0
        else:
            result[table_name] = 0

    # 사용자 복원 (기존 Admin 유지하면서 추가 사용자 복원)
    if "users" in backup and backup["users"]:
        import hashlib
        for user in backup["users"]:
            username = user.get("username", "")
            if username == "Admin":
                # Admin은 비밀번호만 업데이트
                if "password" in user and user["password"]:
                    conn.execute("UPDATE users SET password = ?, name = ? WHERE username = 'Admin'",
                                 (user["password"], user.get("name", "관리자")))
            else:
                # 다른 사용자는 없으면 추가
                try:
                    conn.execute(
                        "INSERT INTO users (username, password, role, name) VALUES (?, ?, ?, ?)",
                        (username, user.get("password", ""), user.get("role", "user"), user.get("name", ""))
                    )
                except:
                    pass  # 이미 존재하면 무시
        result["users"] = len(backup["users"])

    conn.commit()
    conn.close()
    return result


def is_db_empty() -> bool:
    """DB가 비어있는지 확인 (Sleep 후 깨어남 감지용)"""
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM master_asset")
        master_a = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM master_cost")
        master_c = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM users WHERE username != 'Admin'")
        extra_users = cur.fetchone()[0]
        conn.close()
        # 마스터 데이터 없고 추가 사용자도 없으면 = 초기화된 상태
        return master_a == 0 and master_c == 0 and extra_users == 0
    except:
        conn.close()
        return True
