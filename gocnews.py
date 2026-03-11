# -*- coding: utf-8 -*-
"""
2단계:
POP3로 최근 메일들을 읽어서
전체를 한 번에 LLM에 넣고
'신문 편집본' 형태의 HTML 생성

출력:
- news_output_step2.html

환경변수 예시 (Windows CMD)
set POP3_HOST=pop3.samsung.net
set POP3_PORT=995
set POP3_USE_SSL=1
set POP3_TIMEOUT_SEC=20
set POP3_USER=your_id
set POP3_PASSWORD=your_password

set LLM_API_BASE_URL=http://apigw.samsungds.net:8000/gpt-oss/1/gpt-oss-120b/v1/chat/completions
set LLM_CREDENTIAL_KEY=...
set LLM_USER_ID=sungmook.cho
set LLM_USER_TYPE=AD_ID
set LLM_SEND_SYSTEM_NAME=GOC_MAIL_RAG_PIPELINE
"""

import os
import re
import json
import uuid
import html
import poplib
import requests
from dataclasses import dataclass
from datetime import datetime, timedelta
from email import policy
from email.parser import BytesParser
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime
from typing import List, Optional, Dict, Any


# =========================================================
# 설정
# =========================================================
API_BASE_URL = os.getenv(
    "LLM_API_BASE_URL",
    "http://apigw.samsungds.net:8000/gpt-oss/1/gpt-oss-120b/v1/chat/completions"
)
CREDENTIAL_KEY = os.getenv("LLM_CREDENTIAL_KEY", "").strip()
USER_ID = os.getenv("LLM_USER_ID", "sungmook.cho").strip()
USER_TYPE = os.getenv("LLM_USER_TYPE", "AD_ID").strip()
SEND_SYSTEM_NAME = os.getenv("LLM_SEND_SYSTEM_NAME", "GOC_MAIL_RAG_PIPELINE").strip()
EXCLUDE_KEYWORDS = ["[공통]", "[공급망운영 그룹]", "EDP 파트 주요 이슈"]
DEFAULT_LOOKBACK_DAYS = 7
FILTER_KEYWORDS = ["[HBM]", "[FLASH]", "[물류]", "[MOBILE]", "[EDP]", "[DO]", "[운영관리]", "[운영기획]"]


# =========================================================
# 데이터 모델
# =========================================================
@dataclass
class MailQueryParams:
    user: str
    password: str
    max_count: int = 12
    lookback_days: int = DEFAULT_LOOKBACK_DAYS


@dataclass
class MailItem:
    subject: str
    sender: str
    date_str: str
    date_obj: Optional[datetime]
    body: str


# =========================================================
# LLM 호출
# =========================================================
def call_gpt_oss(prompt: str, system_prompt: Optional[str] = None,
                 temperature: float = 0.3, max_tokens: int = 1800) -> Dict[str, Any]:
    if not CREDENTIAL_KEY:
        return {"error": "LLM_CREDENTIAL_KEY 환경변수가 비어 있습니다."}

    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    messages.append({"role": "user", "content": prompt})

    payload = json.dumps({
        "model": "openai/gpt-oss-120b",
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "stream": False
    })

    headers = {
        "x-dep-ticket": CREDENTIAL_KEY,
        "Send-System-Name": SEND_SYSTEM_NAME,
        "User-Id": USER_ID,
        "User-Type": USER_TYPE,
        "Prompt-Msg-Id": str(uuid.uuid4()),
        "Completion-Msg-Id": str(uuid.uuid4()),
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    try:
        resp = requests.post(API_BASE_URL, headers=headers, data=payload, timeout=90)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.RequestException as e:
        return {"error": str(e)}


def extract_json_block(text: str) -> str:
    text = text.strip()
    # 코드블록 감싸진 경우 제거
    text = re.sub(r"^```json\s*", "", text, flags=re.I)
    text = re.sub(r"^```\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()


# =========================================================
# POP3 메일 읽기
# =========================================================
def _pop3_connect(params: MailQueryParams):
    host = os.getenv("POP3_HOST", "pop3.samsung.net")
    use_ssl = os.getenv("POP3_USE_SSL", "1").lower() not in {"0", "false", "no"}
    port_env = os.getenv("POP3_PORT", "").strip()
    port = int(port_env) if port_env else (995 if use_ssl else 110)
    timeout = int(os.getenv("POP3_TIMEOUT_SEC", "20"))

    if use_ssl:
        client = poplib.POP3_SSL(host, port, timeout=timeout)
    else:
        client = poplib.POP3(host, port, timeout=timeout)

    client.user(params.user)
    client.pass_(params.password)
    return client


def decode_mime_header(value: Optional[str]) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\r", "")
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\u200b", "", text)
    return text.strip()


def html_to_text_basic(html_text: str) -> str:
    if not html_text:
        return ""
    text = re.sub(r"(?is)<script.*?>.*?</script>", " ", html_text)
    text = re.sub(r"(?is)<style.*?>.*?</style>", " ", text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p>", "\n", text)
    text = re.sub(r"(?is)<.*?>", " ", text)
    text = html.unescape(text)
    return clean_text(text)


def trim_mail_body(text: str, max_len: int = 1800) -> str:
    if not text:
        return ""
    # 회신/전달 흔적 이후 잘라내기
    split_patterns = [
        r"\n[-]{2,}\s*Original Message\s*[-]{2,}",
        r"\n보낸 사람\s*:",
        r"\nFrom\s*:",
        r"\n-----Original Message-----",
        r"\n발신\s*:",
    ]
    for p in split_patterns:
        m = re.search(p, text, flags=re.I)
        if m:
            text = text[:m.start()]
            break
    text = clean_text(text)
    if len(text) > max_len:
        text = text[:max_len] + " ..."
    return text


def normalize_compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "").lower()


def contains_any_keyword(text: str, keywords: List[str], ignore_spaces: bool = False) -> bool:
    if not text:
        return False
    target = normalize_compact(text) if ignore_spaces else text.lower()
    for keyword in keywords:
        candidate = normalize_compact(keyword) if ignore_spaces else keyword.lower()
        if candidate and candidate in target:
            return True
    return False


def should_include_mail(subject: str, body: str, date_obj: Optional[datetime], cutoff: datetime) -> bool:
    if date_obj is not None:
        if date_obj.tzinfo is not None:
            cutoff_cmp = cutoff.replace(tzinfo=date_obj.tzinfo)
        else:
            cutoff_cmp = cutoff
        if date_obj < cutoff_cmp:
            return False

    if contains_any_keyword(subject, FILTER_KEYWORDS, ignore_spaces=True) is False:
        return False

    if contains_any_keyword(subject, EXCLUDE_KEYWORDS) or contains_any_keyword(body, EXCLUDE_KEYWORDS):
        return False

    return True


def extract_body_from_message(msg) -> str:
    plain_parts = []
    html_parts = []

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition", "")).lower()
            if "attachment" in disposition:
                continue

            try:
                payload = part.get_payload(decode=True)
                if payload is None:
                    continue
                charset = part.get_content_charset() or "utf-8"
                text = payload.decode(charset, errors="replace")
            except Exception:
                try:
                    text = part.get_content()
                except Exception:
                    text = ""

            if content_type == "text/plain":
                plain_parts.append(text)
            elif content_type == "text/html":
                html_parts.append(text)
    else:
        content_type = msg.get_content_type()
        try:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                text = payload.decode(charset, errors="replace")
            else:
                text = msg.get_content()
        except Exception:
            text = ""

        if content_type == "text/plain":
            plain_parts.append(text)
        elif content_type == "text/html":
            html_parts.append(text)

    if plain_parts:
        return trim_mail_body("\n\n".join(plain_parts))
    if html_parts:
        return trim_mail_body(html_to_text_basic("\n\n".join(html_parts)))
    return ""


def fetch_recent_mails(params: MailQueryParams) -> List[MailItem]:
    client = _pop3_connect(params)
    items: List[MailItem] = []
    cutoff = datetime.now() - timedelta(days=params.lookback_days)

    try:
        count, _ = client.stat()

        for i in range(count, 0, -1):
            _, lines, _ = client.retr(i)
            raw_email = b"\n".join(lines)
            msg = BytesParser(policy=policy.default).parsebytes(raw_email)

            subject = decode_mime_header(msg.get("Subject"))
            sender = decode_mime_header(msg.get("From"))
            raw_date = msg.get("Date", "")
            date_obj = None
            try:
                date_obj = parsedate_to_datetime(raw_date)
            except Exception:
                pass

            body = extract_body_from_message(msg)
            if not body and not subject:
                continue
            if not should_include_mail(subject, body, date_obj, cutoff):
                continue

            items.append(MailItem(
                subject=subject or "(제목 없음)",
                sender=sender or "",
                date_str=raw_date or "",
                date_obj=date_obj,
                body=body
            ))
            if len(items) >= params.max_count:
                break
    finally:
        try:
            client.quit()
        except Exception:
            pass

    return items


# =========================================================
# 전체 편집본 생성
# =========================================================
def build_mail_bundle_for_llm(mails: List[MailItem]) -> str:
    blocks = []
    for idx, m in enumerate(mails, start=1):
        dt = m.date_obj.strftime("%Y-%m-%d %H:%M") if m.date_obj else m.date_str
        block = f"""
[메일 {idx}]
제목: {m.subject}
발신자: {m.sender}
일시: {dt}
본문:
{m.body}
"""
        blocks.append(block.strip())
    return "\n\n".join(blocks)


def generate_newspaper_plan(mails: List[MailItem]) -> Dict[str, Any]:
    bundle = build_mail_bundle_for_llm(mails)

    system_prompt = """
당신은 사내 메일 편집국 에디터입니다.
여러 개의 메일을 읽고, 중복되거나 비슷한 주제는 하나의 이슈로 묶어서
'사내 뉴스 신문 편집본' JSON으로 재구성하세요.

반드시 JSON만 출력하세요.
설명문, 마크다운, 코드블록 없이 JSON만 출력하세요.

출력 스키마:
{
  "paper_title": "신문 이름 또는 오늘자 헤드라인",
  "paper_subtitle": "오늘 메일 브리핑 한 줄 설명",
  "top_story": {
    "headline": "",
    "subheadline": "",
    "summary": "",
    "bullets": ["", "", ""],
    "related_mail_indexes": [1, 3]
  },
  "sections": [
    {
      "section_name": "주요 이슈",
      "articles": [
        {
          "headline": "",
          "summary": "",
          "bullets": ["", ""],
          "related_mail_indexes": [2, 5]
        }
      ]
    },
    {
      "section_name": "단신",
      "articles": [
        {
          "headline": "",
          "summary": "",
          "bullets": ["", ""],
          "related_mail_indexes": [4]
        }
      ]
    }
  ],
  "editor_note": "전체 흐름을 한두 문장으로 정리"
}

규칙:
- 비슷한 메일은 related_mail_indexes로 묶기
- top_story는 가장 중요한 이슈 1개
- sections는 최소 2개
- 기사 문체는 간결하고 신문형
- 과장 금지, 원문 기반
- 없는 내용 지어내지 말 것
"""

    user_prompt = f"""
아래 메일 묶음을 읽고, 오늘의 사내 신문 편집본 JSON을 생성하세요.

{bundle}
"""

    result = call_gpt_oss(
        prompt=user_prompt,
        system_prompt=system_prompt,
        temperature=0.2,
        max_tokens=2200
    )

    if "error" in result:
        return {
            "paper_title": "오늘의 사내 메일 브리핑",
            "paper_subtitle": f"LLM 편집 실패: {result['error']}",
            "top_story": {
                "headline": mails[0].subject if mails else "메일 없음",
                "subheadline": "자동 편집 실패",
                "summary": mails[0].body[:300] if mails else "",
                "bullets": [],
                "related_mail_indexes": [1] if mails else []
            },
            "sections": [],
            "editor_note": "LLM 응답을 받지 못해 기본 형태로 생성되었습니다."
        }

    try:
        content = result["choices"][0]["message"]["content"]
        content = extract_json_block(content)
        data = json.loads(content)
        return data
    except Exception as e:
        raw = result.get("choices", [{}])[0].get("message", {}).get("content", "")
        return {
            "paper_title": "오늘의 사내 메일 브리핑",
            "paper_subtitle": "LLM 응답 파싱 실패",
            "top_story": {
                "headline": mails[0].subject if mails else "메일 없음",
                "subheadline": "원문 기반 임시 기사",
                "summary": mails[0].body[:300] if mails else "",
                "bullets": [],
                "related_mail_indexes": [1] if mails else []
            },
            "sections": [
                {
                    "section_name": "원본 응답",
                    "articles": [
                        {
                            "headline": "LLM 원문",
                            "summary": raw[:1200],
                            "bullets": [],
                            "related_mail_indexes": []
                        }
                    ]
                }
            ],
            "editor_note": f"파싱 오류: {e}"
        }


# =========================================================
# HTML 렌더링
# =========================================================
def esc(v) -> str:
    return html.escape(str(v or ""))


def esc_br(v) -> str:
    return esc(v).replace("\n", "<br>")


def render_related_sources(indexes: List[int], mails: List[MailItem]) -> str:
    if not indexes:
        return ""
    rows = []
    for idx in indexes:
        if 1 <= idx <= len(mails):
            m = mails[idx - 1]
            dt = m.date_obj.strftime("%Y-%m-%d %H:%M") if m.date_obj else m.date_str
            rows.append(
                "<tr>"
                "<td style=\"padding:0 0 10px 0;font-size:13px;line-height:1.6;color:#5f5a50;\">"
                f"<strong>[메일 {idx}]</strong> {esc(m.subject)}"
                f"<br><span>{esc(m.sender)} | {esc(dt)}</span>"
                "</td>"
                "</tr>"
            )
    if not rows:
        return ""
    return (
        "<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" "
        "style=\"margin-top:12px;border-top:1px solid #ddd2b8;padding-top:12px;\">"
        f"{''.join(rows)}</table>"
    )


def render_bullet_block(items: List[str], limit: int) -> str:
    rows = []
    for item in (items or [])[:limit]:
        rows.append(
            "<tr>"
            "<td valign=\"top\" style=\"padding:0 8px 8px 0;font-size:15px;line-height:1.7;color:#1d1d1d;\">-</td>"
            f"<td style=\"padding:0 0 8px 0;font-size:15px;line-height:1.7;color:#1d1d1d;\">{esc_br(item)}</td>"
            "</tr>"
        )
    if not rows:
        return ""
    return (
        "<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" "
        "style=\"margin-top:12px;\">"
        f"{''.join(rows)}</table>"
    )


def render_article_card(article: Dict[str, Any], mails: List[MailItem]) -> str:
    bullets_html = render_bullet_block(article.get("bullets", []) or [], 4)
    source_html = render_related_sources(article.get("related_mail_indexes", []), mails)

    return f"""
    <tr>
        <td style="padding:0 0 16px 0;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #d9d2c1;background:#ffffff;">
                <tr>
                    <td style="padding:18px 20px 20px 20px;">
                        <div style="font-size:24px;line-height:1.4;font-weight:700;color:#1d1d1d;">{esc(article.get("headline"))}</div>
                        <div style="padding-top:10px;font-size:15px;line-height:1.8;color:#2f2b25;">{esc_br(article.get("summary"))}</div>
                        {bullets_html}
                        {source_html}
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    """


def render_newspaper_html_step2(plan: Dict[str, Any], mails: List[MailItem], output_path: str):
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    top = plan.get("top_story", {}) or {}
    top_bullets = render_bullet_block(top.get("bullets") or [], 5)
    top_sources = render_related_sources(top.get("related_mail_indexes", []), mails)

    sections_html = ""
    for sec in plan.get("sections", []) or []:
        articles_html = "".join(render_article_card(a, mails) for a in (sec.get("articles") or []))
        sections_html += f"""
        <tr>
            <td style="padding:0 24px 24px 24px;">
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-top:3px solid #222222;">
                    <tr>
                        <td style="padding:14px 0 16px 0;font-size:22px;line-height:1.3;font-weight:700;color:#1d1d1d;">
                            {esc(sec.get("section_name"))}
                        </td>
                    </tr>
                    {articles_html}
                </table>
            </td>
        </tr>
        """

    html_text = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>{esc(plan.get("paper_title", "사내 메일 신문"))}</title>
</head>
<body style="margin:0;padding:0;background-color:#efe9dc;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;margin:0;padding:24px 0;background-color:#efe9dc;">
        <tr>
            <td align="center">
                <table role="presentation" width="760" cellpadding="0" cellspacing="0" style="width:760px;max-width:760px;background-color:#fffdf8;border:1px solid #d4cdbf;font-family:'Malgun Gothic','Apple SD Gothic Neo',Arial,sans-serif;color:#1d1d1d;">
                    <tr>
                        <td style="padding:28px 30px 18px 30px;text-align:center;background-color:#fbf6ea;border-bottom:4px double #222222;">
                            <div style="font-size:42px;line-height:1.1;font-weight:700;letter-spacing:2px;">{esc(plan.get("paper_title", "GOC DAILY MAIL TIMES"))}</div>
                            <div style="padding-top:8px;font-size:15px;line-height:1.5;color:#666666;">{esc(plan.get("paper_subtitle", "사내 메일 자동 편집 신문"))}</div>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:10px 20px;border-bottom:1px solid #ddd7ca;background-color:#ffffff;font-size:13px;line-height:1.6;color:#666666;text-align:center;">
                            생성 시각: {esc(created_at)} | 원본 메일 수: {len(mails)}건 | 검색 기간: 최근 {DEFAULT_LOOKBACK_DAYS}일
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:16px 24px;text-align:center;background-color:#faf8f1;border-bottom:1px solid #ddd7ca;font-size:24px;line-height:1.5;font-weight:700;color:#1d1d1d;">
                            {esc((top.get("headline") or plan.get("paper_title") or "오늘의 주요 이슈"))}
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:24px;">
                            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border:1px solid #ddd2b8;background-color:#fffdfa;">
                                <tr>
                                    <td style="padding:24px;">
                                        <div style="display:inline-block;padding:5px 10px;background-color:#111111;color:#ffffff;font-size:12px;line-height:1.2;font-weight:700;">TOP STORY</div>
                                        <div style="padding-top:14px;font-size:36px;line-height:1.3;font-weight:700;color:#1d1d1d;">{esc(top.get("headline"))}</div>
                                        <div style="padding-top:10px;font-size:18px;line-height:1.6;color:#6b6b6b;">{esc_br(top.get("subheadline"))}</div>
                                        <div style="padding-top:16px;font-size:17px;line-height:1.8;color:#2f2b25;">{esc_br(top.get("summary"))}</div>
                                        {top_bullets}
                                        {top_sources}
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:0 24px 24px 24px;">
                            <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td style="padding:16px;border:1px solid #ddd2b8;background-color:#faf6ea;">
                                        <div style="font-size:18px;line-height:1.4;font-weight:700;color:#1d1d1d;">편집자 노트</div>
                                        <div style="padding-top:10px;font-size:14px;line-height:1.8;color:#3f3a34;">{esc_br(plan.get("editor_note", ""))}</div>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="height:12px;font-size:0;line-height:0;">&nbsp;</td>
                                </tr>
                                <tr>
                                    <td style="padding:16px;border:1px solid #ddd2b8;background-color:#faf6ea;">
                                        <div style="font-size:18px;line-height:1.4;font-weight:700;color:#1d1d1d;">오늘 편집 방향</div>
                                        <div style="padding-top:10px;font-size:14px;line-height:1.8;color:#3f3a34;">개별 메일 나열이 아니라, 유사 주제를 하나의 기사로 통합해 신문형 레이아웃으로 재편집했습니다.</div>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    {sections_html}
                    <tr>
                        <td style="padding:16px 20px;border-top:1px solid #ddd7ca;background-color:#fafafa;text-align:center;font-size:12px;line-height:1.6;color:#777777;">
                            본 HTML은 사내 메일을 기반으로 자동 생성된 신문형 초안입니다.
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_text)


# =========================================================
# 메인
# =========================================================
def main():
    user = os.getenv("POP3_USER", "").strip()
    password = os.getenv("POP3_PASSWORD", "").strip()

    if not user or not password:
        raise RuntimeError("POP3_USER / POP3_PASSWORD 환경변수를 설정해 주세요.")
    if not CREDENTIAL_KEY:
        raise RuntimeError("LLM_CREDENTIAL_KEY 환경변수를 설정해 주세요.")

    params = MailQueryParams(
        user=user,
        password=password,
        max_count=10,
        lookback_days=DEFAULT_LOOKBACK_DAYS
    )

    print("[1] 최근 메일 수집 중...")
    mails = fetch_recent_mails(params)
    print(f"[INFO] 수집된 메일 수: {len(mails)}")

    if not mails:
        raise RuntimeError("가져온 메일이 없습니다.")

    print("[2] 전체 메일 묶음을 신문 편집본으로 재구성 중...")
    plan = generate_newspaper_plan(mails)

    print("[3] HTML 생성 중...")
    output_path = "news_output_step2.html"
    render_newspaper_html_step2(plan, mails, output_path)

    print(f"[DONE] 완료: {output_path}")


if __name__ == "__main__":
    main()
