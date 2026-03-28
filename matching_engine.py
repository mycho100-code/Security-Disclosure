"""
Matching Engine - Description 기반 유사도 매칭 로직
분류값: "O" = 해당, "" = 비해당
"""
import re
import pandas as pd
from rapidfuzz import fuzz


def preprocess(text: str) -> str:
    """공백·특수문자 제거 후 대문자 변환"""
    if not isinstance(text, str):
        return ""
    text = re.sub(r"[^A-Za-z0-9가-힣]", "", text)
    return text.upper()


def run_matching(target_df: pd.DataFrame, master_df: pd.DataFrame,
                 threshold: int = 85) -> tuple[pd.DataFrame, dict]:
    """
    target_df의 description을 master_df와 비교하여 분류값(제외/정보기술/정보보호) 매핑.
    """
    target_df = target_df.copy()
    master_df = master_df.copy()

    target_df["_key"] = target_df["description"].apply(preprocess)
    master_df["_key"] = master_df["description"].apply(preprocess)

    # 결과 컬럼 초기화
    target_df["match_type"] = ""
    target_df["match_score"] = 0.0
    target_df["matched_desc"] = ""

    # 마스터 딕셔너리 (1차 exact match)
    master_map = {}
    for _, row in master_df.iterrows():
        key = row["_key"]
        if key and key not in master_map:
            master_map[key] = {
                "exclude_yn": str(row.get("exclude_yn", "")) if pd.notna(row.get("exclude_yn", "")) else "",
                "it_yn": str(row.get("it_yn", "")) if pd.notna(row.get("it_yn", "")) else "",
                "sec_yn": str(row.get("sec_yn", "")) if pd.notna(row.get("sec_yn", "")) else "",
                "description": str(row.get("description", "")),
            }

    # 마스터 리스트 (2차 fuzzy match)
    master_list = []
    for _, row in master_df.iterrows():
        master_list.append({
            "key": row["_key"],
            "exclude_yn": str(row.get("exclude_yn", "")) if pd.notna(row.get("exclude_yn", "")) else "",
            "it_yn": str(row.get("it_yn", "")) if pd.notna(row.get("it_yn", "")) else "",
            "sec_yn": str(row.get("sec_yn", "")) if pd.notna(row.get("sec_yn", "")) else "",
            "description": str(row.get("description", "")),
        })

    exact_count = 0
    fuzzy_count = 0
    unmatched_count = 0

    for idx, row in target_df.iterrows():
        tkey = row["_key"]
        if not tkey:
            unmatched_count += 1
            target_df.at[idx, "match_type"] = "미매칭"
            continue

        # ── 1차: 완전 일치 ──
        if tkey in master_map:
            m = master_map[tkey]
            target_df.at[idx, "exclude_yn"] = m["exclude_yn"]
            target_df.at[idx, "it_yn"] = m["it_yn"]
            target_df.at[idx, "sec_yn"] = m["sec_yn"]
            target_df.at[idx, "match_type"] = "완전일치"
            target_df.at[idx, "match_score"] = 100.0
            target_df.at[idx, "matched_desc"] = m["description"]
            exact_count += 1
            continue

        # ── 2차: 유사도 매칭 ──
        best_score = 0
        best_match = None
        for m in master_list:
            if not m["key"]:
                continue
            score = fuzz.token_sort_ratio(tkey, m["key"])
            if score > best_score:
                best_score = score
                best_match = m

        if best_score >= threshold and best_match:
            target_df.at[idx, "exclude_yn"] = best_match["exclude_yn"]
            target_df.at[idx, "it_yn"] = best_match["it_yn"]
            target_df.at[idx, "sec_yn"] = best_match["sec_yn"]
            target_df.at[idx, "match_type"] = "유사매칭"
            target_df.at[idx, "match_score"] = round(best_score, 1)
            target_df.at[idx, "matched_desc"] = best_match["description"]
            fuzzy_count += 1
        else:
            target_df.at[idx, "match_type"] = "미매칭"
            target_df.at[idx, "match_score"] = round(best_score, 1) if best_match else 0
            unmatched_count += 1

    target_df.drop(columns=["_key"], inplace=True)

    stats = {
        "total": len(target_df),
        "exact": exact_count,
        "fuzzy": fuzzy_count,
        "unmatched": unmatched_count,
    }
    return target_df, stats
