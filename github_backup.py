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
        # st.secrets는 딕셔너리 접근만 지원하는 버전이 있음
        token = ""
        repo = ""
        try:
            token = st.secrets["GITHUB_TOKEN"]
        except (KeyError, FileNotFoundError):
            pass
        try:
            repo = st.secrets["GITHUB_REPO"]
        except (KeyError, FileNotFoundError):
            pass
        return str(token).strip(), str(repo).strip()
    except Exception:
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


def _get_default_branch(token, repo):
    """저장소의 기본 브랜치명 확인 (main 또는 master 등)"""
    try:
        resp = requests.get(
            f"https://api.github.com/repos/{repo}",
            headers=_headers(token), timeout=10
        )
        if resp.status_code == 200:
            return resp.json().get("default_branch", "main")
    except:
        pass
    return "main"


def push_backup(backup_bytes: bytes) -> dict:
    """
    GitHub 저장소에 백업 파일 push (생성 또는 업데이트).
    반환: {"success": True/False, "message": "..."}
    """
    token, repo = _get_config()
    if not token or not repo:
        return {"success": False, "message": "GitHub 설정이 없습니다."}

    branch = _get_default_branch(token, repo)
    url = f"https://api.github.com/repos/{repo}/contents/{BACKUP_PATH}"
    headers = _headers(token)

    # 기존 파일의 SHA 확인 (업데이트 시 필요)
    sha = None
    try:
        resp = requests.get(url, headers=headers, params={"ref": branch}, timeout=15)
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
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha

    try:
        resp = requests.put(url, headers=headers, json=payload, timeout=30)
        if resp.status_code in (200, 201):
            return {"success": True, "message": f"GitHub 백업 완료 (branch: {branch})"}
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

    branch = _get_default_branch(token, repo)
    url = f"https://api.github.com/repos/{repo}/contents/{BACKUP_PATH}"
    headers = _headers(token)

    try:
        resp = requests.get(url, headers=headers, params={"ref": branch}, timeout=15)
        if resp.status_code == 200:
            resp_json = resp.json()

            # 방법 1: base64 인코딩된 content (소용량 파일)
            content_b64 = resp_json.get("content", "")
            if content_b64:
                # ★ GitHub API는 base64에 줄바꿈(\n)을 포함시킴 → 제거 필요
                content_b64 = content_b64.replace("\n", "").replace("\r", "").strip()
                data = base64.b64decode(content_b64)
                return {"success": True, "data": data,
                        "message": f"GitHub 백업 로드 완료 (branch: {branch})"}

            # 방법 2: download_url로 직접 다운로드 (대용량 파일, >1MB)
            download_url = resp_json.get("download_url", "")
            if download_url:
                dl_resp = requests.get(download_url, headers=headers, timeout=30)
                if dl_resp.status_code == 200:
                    return {"success": True, "data": dl_resp.content,
                            "message": f"GitHub 백업 로드 완료 - download_url (branch: {branch})"}

            return {"success": False, "data": None, "message": "백업 파일 내용을 읽을 수 없습니다."}
        elif resp.status_code == 404:
            return {"success": False, "data": None, "message": "GitHub에 백업 파일이 없습니다."}
        else:
            return {"success": False, "data": None, "message": f"GitHub API 오류: {resp.status_code}"}
    except Exception as e:
        return {"success": False, "data": None, "message": f"네트워크 오류: {e}"}
