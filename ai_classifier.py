"""
AI Classifier - OpenAI API를 사용하여 미매칭 Description을 분류
분류 기준: 정보기술(IT) / 정보보호(Security) / 제외(Exclude)
"""
import json
import requests


SYSTEM_PROMPT = """당신은 KISA(한국인터넷진흥원) 「정보보호 공시 가이드라인(2024.03 개정본)」에 따른 IT 자산 및 비용 분류 전문가입니다.
주어진 Description(설명)을 분석하여 '정보기술부문', '정보보호부문', '제외' 중 해당하는 항목을 판단해 주세요.

※ 참조: KISA 정보보호 공시 가이드라인
  - https://isds.kisa.or.kr/kr/bbs/view.do?bbsId=B0000011&menuNo=204948&nttId=38
  - 근거법령: 「정보보호산업의 진흥에 관한 법률」 제13조

[분류 기준 - KISA 가이드라인 자산분류표 기반]

1. 정보기술부문 (IT) — 포괄적 관점에서 작성
   정보기술이란 전기통신설비를 이용하고, 컴퓨터 등 정보처리능력을 가진 하드웨어·소프트웨어 및 이를 사용하기 위한 간접설비를 포함하는 개념입니다.
   
   자산 예시:
   - 서버/스토리지: 서버, NAS, SAN, 스토리지, 백업장비
   - 네트워크: 라우터, 스위치(L2/L3/L4), 허브, AP(무선랜), 로드밸런서
   - 단말/PC: 데스크톱, 노트북, 태블릿, 모니터, 씬클라이언트
   - 출력장비: 프린터, 복합기, 스캐너
   - 통신장비: 전화교환기(PBX), IP-Phone, 화상회의장비, 전용회선장비
   - 소프트웨어: OS, DB, 미들웨어, 오피스SW, ERP, CRM, 그룹웨어, 개발도구
   - 클라우드: IaaS, PaaS, SaaS 이용료
   - 전산시설: UPS, 항온항습기, 전산실 시설, 배선(케이블링)
   
   비용 예시:
   - IT 유지보수, SW 라이선스, 통신비(인터넷/전용회선), 호스팅비
   - IT 컨설팅, 시스템 개발/구축, IT 외주용역, 클라우드 이용료

2. 정보보호부문 (Security) — 보수적 관점에서 작성
   정보보호란 정보의 수집·가공·저장·검색·송수신 중 발생하는 정보의 훼손·변조·유출 등을 방지하기 위한 관리적·기술적 수단을 말합니다.
   
   자산 예시:
   - 네트워크 보안: 방화벽(Firewall), IPS/IDS, WAF, DDoS 대응장비, VPN장비
   - 시스템 보안: NAC(네트워크접근제어), 서버보안, DB보안/접근제어, 보안USB
   - 애플리케이션 보안: 웹방화벽, 소스코드 취약점분석도구, 시큐어코딩도구
   - 데이터 보안: DLP(정보유출방지), 문서보안(DRM), 암호화솔루션, 개인정보탐지
   - 인증/접근제어: OTP, SSO, IAM/PAM, 생체인증, 인증서관리
   - 보안관제: SIEM, ESM, SOAR, 로그분석, 통합관제
   - 단말보안: 백신(Antivirus), EDR, MDM(모바일단말관리), 패치관리(PMS)
   - 물리보안(IT관련): CCTV, 출입통제(카드리더/지문인식), 영상분석
   
   비용 예시:
   - 보안솔루션 유지보수, 보안컨설팅, 취약점점검/모의해킹, 보안관제 용역
   - 보안인증(ISMS/ISMS-P/ISO27001) 심사비, 보안교육, 침해사고 대응

   ★ 정보보호 자산/비용은 동시에 정보기술에도 해당됩니다 (IT=O, 보안=O 모두 표시)

3. 제외 (Exclude) — IT/보안과 무관한 항목
   - 사무용품(A4, 토너 등), 가구(책상, 의자), 일반 시설(건물, 인테리어)
   - 차량, 복리후생, 일반 경비, 광고비, 여비교통비
   - 기타 정보기술·정보보호와 직접적 관련이 없는 항목

[중요 규칙]
- 정보보호부문에 해당하면 정보기술부문에도 해당됩니다 (it_yn="O", sec_yn="O")
- 정보기술부문만 해당될 수 있습니다 (it_yn="O", sec_yn="")
- 제외 항목은 IT와 보안 모두 해당하지 않습니다 (exclude_yn="O")
- 확실하지 않으면 confidence를 낮게 설정하세요
- KISA 가이드라인의 자산분류표(p.72~77)를 기준으로 판단하세요

[응답 형식 - 반드시 JSON만 출력]
{
  "it_yn": "O" 또는 "",
  "sec_yn": "O" 또는 "",
  "exclude_yn": "O" 또는 "",
  "confidence": 0.0~1.0,
  "reason": "KISA 가이드라인 기준 판단 근거를 한 줄로 설명"
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
            {"role": "user", "content": f"KISA 정보보호 공시 가이드라인(2024.03)의 자산분류표를 기준으로 다음 Description을 분류해 주세요:\n\n\"{description}\""},
        ],
        "temperature": 0.1,
        "max_tokens": 500,
    }

    try:
        resp = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=60,
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
