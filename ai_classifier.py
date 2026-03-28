"""
AI Classifier - OpenAI API를 사용하여 미매칭 Description을 분류
분류 기준: 정보기술(IT) / 정보보호(Security) / 제외(Exclude)
"""
import json
import requests


SYSTEM_PROMPT = """당신은 IT 자산 및 비용 분류 전문가입니다.
주어진 Description(설명)을 분석하여 아래 3가지 분류 중 해당하는 항목을 판단해 주세요.

[분류 기준]
1. 정보기술 (IT): 컴퓨터, 서버, 네트워크 장비, 소프트웨어, 클라우드 서비스, 통신, IT 유지보수, 시스템 개발, IT 인프라 등 정보기술과 관련된 자산이나 비용. 일반 IT와 관련된 제품이나 서비스도 포함
2. 정보보호 (Information Security): 방화벽, IPS/IDS, 백신, 보안 솔루션, 보안 컨설팅, 취약점 점검, 보안 관제, 암호화 장비, 접근제어 등 정보보호와 관련된 자산이나 비용. 일반 정보보호와 관련된 제품이나 서비스도 포함
3. 제외 (Exclude): 사무용품, 가구, 일반 시설, 차량, 복리후생, 기타 IT/정보보호/보안/정보보안과 무관한 항목

[중요 규칙]
- 정보기술과 정보보호 모두 해당될 수 있습니다 (예: VPN 장비 → IT=O, 보안=O)
- 확실하지 않으면 confidence를 낮게 설정하세요

[응답 형식 - 반드시 JSON만 출력]
{
  "it_yn": "O" 또는 "",
  "sec_yn": "O" 또는 "",
  "exclude_yn": "O" 또는 "",
  "confidence": 0.0~1.0,
  "reason": "판단 근거를 한 줄로 설명"
}
"""


def classify_single(description: str, api_key: str, model: str = "gpt-4o-mini") -> dict:
    """단일 Description을 AI로 분류"""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"다음 Description을 분류해 주세요:\n\n\"{description}\""},
        ],
        "temperature": 0.1,
        "max_tokens": 300,
    }

    try:
        resp = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=30,
        )
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"]

        # JSON 파싱 (```json ... ``` 감싸진 경우도 처리)
        content = content.strip()
        if content.startswith("```"):
            content = content.split("\n", 1)[1]  # 첫 줄 제거
            content = content.rsplit("```", 1)[0]  # 마지막 ``` 제거
        content = content.strip()

        result = json.loads(content)
        return {
            "it_yn": result.get("it_yn", ""),
            "sec_yn": result.get("sec_yn", ""),
            "exclude_yn": result.get("exclude_yn", ""),
            "confidence": float(result.get("confidence", 0)),
            "reason": result.get("reason", ""),
            "error": None,
        }
    except requests.exceptions.HTTPError as e:
        return {"it_yn": "", "sec_yn": "", "exclude_yn": "",
                "confidence": 0, "reason": "", "error": f"API 오류: {e}"}
    except json.JSONDecodeError:
        return {"it_yn": "", "sec_yn": "", "exclude_yn": "",
                "confidence": 0, "reason": content if 'content' in dir() else "",
                "error": "응답 파싱 실패"}
    except Exception as e:
        return {"it_yn": "", "sec_yn": "", "exclude_yn": "",
                "confidence": 0, "reason": "", "error": str(e)}


def classify_batch(descriptions: list[dict], api_key: str,
                   model: str = "gpt-4o-mini",
                   progress_callback=None) -> list[dict]:
    """
    여러 Description을 일괄 분류.
    descriptions: [{"id": ..., "description": ...}, ...]
    반환: [{"id": ..., "description": ..., "it_yn": ..., ...}, ...]
    """
    results = []
    total = len(descriptions)

    for i, item in enumerate(descriptions):
        desc = item.get("description", "")
        result = classify_single(desc, api_key, model)
        result["id"] = item["id"]
        result["description"] = desc
        results.append(result)

        if progress_callback:
            progress_callback(i + 1, total)

    return results
