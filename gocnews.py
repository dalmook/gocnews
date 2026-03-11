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
from datetime import datetime
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


# =========================================================
# 데이터 모델
# =========================================================
@dataclass
class MailQueryParams:
    user: str
    password: str
    max_count: int = 12


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

    try:
        count, _ = client.stat()
        start_idx = max(1, count - params.max_count + 1)

        for i in range(count, start_idx - 1, -1):
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

            items.append(MailItem(
                subject=subject or "(제목 없음)",
                sender=sender or "",
                date_str=raw_date or "",
                date_obj=date_obj,
                body=body
            ))
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


def render_related_sources(indexes: List[int], mails: List[MailItem]) -> str:
    if not indexes:
        return ""
    rows = []
    for idx in indexes:
        if 1 <= idx <= len(mails):
            m = mails[idx - 1]
            dt = m.date_obj.strftime("%Y-%m-%d %H:%M") if m.date_obj else m.date_str
            rows.append(
                f"<li><strong>[메일 {idx}]</strong> {esc(m.subject)}"
                f"<br><span>{esc(m.sender)} · {esc(dt)}</span></li>"
            )
    if not rows:
        return ""
    return f"<ul class='source-list'>{''.join(rows)}</ul>"


def render_article_card(article: Dict[str, Any], mails: List[MailItem]) -> str:
    bullets = article.get("bullets", []) or []
    bullets_html = "".join(f"<li>{esc(b)}</li>" for b in bullets[:4])
    source_html = render_related_sources(article.get("related_mail_indexes", []), mails)

    return f"""
    <article class="article-card">
        <h4>{esc(article.get("headline"))}</h4>
        <p class="article-summary">{esc(article.get("summary"))}</p>
        {'<ul class="bullet-list">' + bullets_html + '</ul>' if bullets_html else ''}
        {source_html}
    </article>
    """


def render_newspaper_html_step2(plan: Dict[str, Any], mails: List[MailItem], output_path: str):
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    top = plan.get("top_story", {}) or {}
    top_bullets = "".join(f"<li>{esc(b)}</li>" for b in (top.get("bullets") or [])[:5])
    top_sources = render_related_sources(top.get("related_mail_indexes", []), mails)

    sections_html = ""
    for sec in plan.get("sections", []) or []:
        articles_html = "".join(render_article_card(a, mails) for a in (sec.get("articles") or []))
        sections_html += f"""
        <section class="section-block">
            <div class="section-title">{esc(sec.get("section_name"))}</div>
            <div class="section-grid">
                {articles_html}
            </div>
        </section>
        """

    html_text = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(plan.get("paper_title", "사내 메일 신문"))}</title>
<style>
    body {{
        margin: 0;
        background: #efe9dc;
        color: #1d1d1d;
        font-family: "Malgun Gothic", "Apple SD Gothic Neo", Arial, sans-serif;
    }}
    .page {{
        max-width: 1280px;
        margin: 20px auto;
        background: #fffdf8;
        border: 1px solid #d4cdbf;
        box-shadow: 0 6px 24px rgba(0,0,0,0.08);
    }}
    .masthead {{
        padding: 24px 30px 18px;
        border-bottom: 4px double #222;
        text-align: center;
        background: #fbf6ea;
    }}
    .masthead h1 {{
        margin: 0;
        font-size: 42px;
        letter-spacing: 2px;
    }}
    .masthead .subtitle {{
        margin-top: 8px;
        font-size: 15px;
        color: #666;
    }}
    .topbar {{
        display: flex;
        justify-content: space-between;
        gap: 12px;
        padding: 10px 20px;
        border-bottom: 1px solid #ddd;
        background: #fff;
        font-size: 13px;
        color: #666;
    }}
    .lead-banner {{
        padding: 16px 24px;
        text-align: center;
        font-size: 24px;
        font-weight: 700;
        border-bottom: 1px solid #ddd;
        background: #faf8f1;
    }}
    .main-grid {{
        display: grid;
        grid-template-columns: 2.1fr 1fr;
        gap: 0;
    }}
    .hero {{
        padding: 28px;
        border-right: 1px solid #ddd;
    }}
    .hero-label {{
        display: inline-block;
        padding: 5px 10px;
        background: #111;
        color: white;
        font-size: 12px;
        margin-bottom: 12px;
    }}
    .hero h2 {{
        margin: 0 0 10px;
        font-size: 36px;
        line-height: 1.3;
    }}
    .hero h3 {{
        margin: 0 0 14px;
        font-size: 18px;
        line-height: 1.5;
        color: #6b6b6b;
        font-weight: normal;
    }}
    .hero-summary {{
        font-size: 17px;
        line-height: 1.8;
        margin-bottom: 16px;
    }}
    .bullet-list {{
        margin: 0 0 18px 18px;
        line-height: 1.7;
    }}
    .side-note {{
        padding: 24px;
        background: #fffdfa;
    }}
    .editor-box {{
        border: 1px solid #ddd2b8;
        background: #faf6ea;
        padding: 16px;
        margin-bottom: 18px;
    }}
    .editor-box h4 {{
        margin: 0 0 10px;
        font-size: 18px;
    }}
    .source-list {{
        margin: 12px 0 0 18px;
        color: #666;
        line-height: 1.6;
        font-size: 13px;
    }}
    .source-list li {{
        margin-bottom: 8px;
    }}
    .sections {{
        padding: 0 24px 24px;
    }}
    .section-block {{
        margin-top: 24px;
        border-top: 3px solid #222;
        padding-top: 14px;
    }}
    .section-title {{
        font-size: 22px;
        font-weight: 700;
        margin-bottom: 16px;
    }}
    .section-grid {{
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 16px;
    }}
    .article-card {{
        border: 1px solid #ddd;
        background: #fff;
        padding: 18px;
    }}
    .article-card h4 {{
        margin: 0 0 10px;
        font-size: 23px;
        line-height: 1.4;
    }}
    .article-summary {{
        margin: 0 0 10px;
        font-size: 15px;
        line-height: 1.75;
    }}
    .footer {{
        border-top: 1px solid #ddd;
        padding: 16px 20px;
        text-align: center;
        color: #777;
        font-size: 12px;
        background: #fafafa;
    }}
    @media (max-width: 960px) {{
        .main-grid {{
            grid-template-columns: 1fr;
        }}
        .hero {{
            border-right: none;
            border-bottom: 1px solid #ddd;
        }}
        .section-grid {{
            grid-template-columns: 1fr;
        }}
        .masthead h1 {{
            font-size: 30px;
        }}
        .hero h2 {{
            font-size: 28px;
        }}
    }}
</style>
</head>
<body>
    <div class="page">
        <header class="masthead">
            <h1>{esc(plan.get("paper_title", "GOC DAILY MAIL TIMES"))}</h1>
            <div class="subtitle">{esc(plan.get("paper_subtitle", "사내 메일 자동 편집 신문"))}</div>
        </header>

        <div class="topbar">
            <div>생성 시각: {esc(created_at)}</div>
            <div>원본 메일 수: {len(mails)}건</div>
        </div>

        <div class="lead-banner">
            {esc((top.get("headline") or plan.get("paper_title") or "오늘의 주요 이슈"))}
        </div>

        <div class="main-grid">
            <section class="hero">
                <div class="hero-label">TOP STORY</div>
                <h2>{esc(top.get("headline"))}</h2>
                <h3>{esc(top.get("subheadline"))}</h3>
                <p class="hero-summary">{esc(top.get("summary"))}</p>
                {'<ul class="bullet-list">' + top_bullets + '</ul>' if top_bullets else ''}
                {top_sources}
            </section>

            <aside class="side-note">
                <div class="editor-box">
                    <h4>편집자 노트</h4>
                    <p>{esc(plan.get("editor_note", ""))}</p>
                </div>
                <div class="editor-box">
                    <h4>오늘 편집 방향</h4>
                    <p>개별 메일 나열이 아니라, 유사 주제를 하나의 기사로 통합해 신문형 레이아웃으로 재편집했습니다.</p>
                </div>
            </aside>
        </div>

        <div class="sections">
            {sections_html}
        </div>

        <footer class="footer">
            본 HTML은 사내 메일을 기반으로 자동 생성된 신문형 초안입니다.
        </footer>
    </div>
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
        max_count=10
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