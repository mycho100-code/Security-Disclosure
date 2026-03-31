"""
GitHub Backup - GitHub 저장소에 백업 파일 자동 저장/복원
Streamlit secrets에 GITHUB_TOKEN, GITHUB_REPO 설정 필요
"""
import requests
import base64
import json

BACKUP_PATH = "backup/data.json"


def _get_config():
    """Streamlit secrets에서 GitHub 설정 읽기"""
    try:
        import streamlit as st
        token = st.secrets.get("GITHUB_TOKEN", "")
        repo = st.secrets.get("GITHUB_REPO", "")  # "owner/repo-name" 형식
        return token, repo
    except:
        return "", ""


def _headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
    }


def is_configured() -> bool:
    """GitHub 백업이 설정되어 있는지 확인"""
    token, repo = _get_config()
    return bool(token) and bool(repo)


def push_backup(backup_bytes: bytes) -> dict:
    """
    GitHub 저장소에 백업 파일 push (생성 또는 업데이트).
    반환: {"success": True/False, "message": "..."}
    """
    token, repo = _get_config()
    if not token or not repo:
        return {"success": False, "message": "GitHub 설정이 없습니다. secrets에 GITHUB_TOKEN, GITHUB_REPO를 설정하세요."}

    url = f"https://api.github.com/repos/{repo}/contents/{BACKUP_PATH}"
    headers = _headers(token)

    # 기존 파일의 SHA 확인 (업데이트 시 필요)
    sha = None
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        if resp.status_code == 200:
            sha = resp.json().get("sha")
    except:
        pass

    # 파일 생성/업데이트
    import datetime
    content_b64 = base64.b64encode(backup_bytes).decode("utf-8")
    payload = {
        "message": f"Auto backup - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "content": content_b64,
        "branch": "main",
    }
    if sha:
        payload["sha"] = sha

    try:
        resp = requests.put(url, headers=headers, json=payload, timeout=30)
        if resp.status_code in (200, 201):
            return {"success": True, "message": "GitHub 백업 완료"}
        else:
            return {"success": False, "message": f"GitHub API 오류: {resp.status_code} - {resp.text[:200]}"}
    except Exception as e:
        return {"success": False, "message": f"네트워크 오류: {e}"}


def pull_backup() -> dict:
    """
    GitHub 저장소에서 백업 파일 가져오기.
    반환: {"success": True/False, "data": bytes or None, "message": "..."}
    """
    token, repo = _get_config()
    if not token or not repo:
        return {"success": False, "data": None, "message": "GitHub 설정이 없습니다."}

    url = f"https://api.github.com/repos/{repo}/contents/{BACKUP_PATH}"
    headers = _headers(token)

    try:
        resp = requests.get(url, headers=headers, timeout=15)
        if resp.status_code == 200:
            content_b64 = resp.json().get("content", "")
            data = base64.b64decode(content_b64)
            return {"success": True, "data": data, "message": "GitHub 백업 로드 완료"}
        elif resp.status_code == 404:
            return {"success": False, "data": None, "message": "GitHub에 백업 파일이 없습니다."}
        else:
            return {"success": False, "data": None, "message": f"GitHub API 오류: {resp.status_code}"}
    except Exception as e:
        return {"success": False, "data": None, "message": f"네트워크 오류: {e}"}
