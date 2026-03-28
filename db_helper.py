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
