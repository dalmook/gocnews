# -*- coding: utf-8 -*-
"""
메일 리포팅 시스템
- 주별 또는 지정 날짜 범위의 메일을 크롤링하여 요약 리포트 생성
- mail_sender_api.py를 통해 리포트 메일 발송
"""

import os
import re
import poplib
import email
from email.message import Message
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime
from datetime import datetime, timezone, timedelta
from typing import List, Dict, Optional, Tuple, Union
from mail_sender_api import send_mail_api
import html as html_module

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False
    print("경고: BeautifulSoup4가 설치되지 않았습니다. 'pip install beautifulsoup4'로 설치하세요.")

# ==================== 설정 ====================
POP3_HOST = "pop3.samsung.net"

USER = os.getenv("MAIL_USER", "issue.goc")
PW = os.getenv("MAIL_PASS", "qhsqn2063!")

# 리포트 수신자 (아이디만 입력, @samsung.com 자동 추가)
# 여러 사람 입력 가능: ["id1", "id2", "id3"]
REPORT_RECIPIENT_IDS = ["sungmook.cho","sung.w.jung","junsoo.jung","w2635.lee","jc2573.lee","sj82.han","cheon.kim","jh3.park","kyungchan.seong","suy.kim","sunok78.han","jjlive.kim"]  # 여기에 실제 수신자 이메일 아이디 입력
# REPORT_RECIPIENT_IDS = ["sungmook.cho"]  # 여기에 실제 수신자 이메일 아이디 입력
REPORT_DOMAIN = "samsung.com"
REPORT_SENDER = os.getenv("REPORT_SENDER", USER)

# 요약 길이 (문자 수)
SUMMARY_LENGTH = 200

# 필터링할 키워드 (제목에 포함된 메일만 수집, 공백 무시)
FILTER_KEYWORDS = ["[HBM]", "[FLASH]", "[물류]", "[MOBILE]", "[EDP]", "[DO]","[운영관리]","[운영기획]"]

# 카테고리별 이모지 및 파스텔 톤 색상
CATEGORY_STYLES = {
    "[HBM]": {"emoji": "💾", "color": "#FFB3BA"},      # 파스텔 핑크
    "[FLASH]": {"emoji": "⚡", "color": "#FFDFBA"},    # 파스텔 오렌지
    "[물류]": {"emoji": "🚚", "color": "#FFFFBA"},    # 파스텔 옐로우
    "[MOBILE]": {"emoji": "📱", "color": "#BAFFC9"},   # 파스텔 그린
    "[EDP]": {"emoji": "💻", "color": "#BAE1FF"},      # 파스텔 블루
    "[DO]": {"emoji": "📊", "color": "#E0BBE4"},      # 파스텔 퍼플
    "[운영관리]": {"emoji": "⚙️", "color": "#C7CEEA"}, # 파스텔 인디고
    "[운영기획]": {"emoji": "📋", "color": "#F0E68C"}, # 파스텔 카키
    "기타": {"emoji": "📌", "color": "#E8E8E8"},       # 파스텔 그레이
}

# 제외할 키워드 (제목 또는 본문에 포함된 메일 제외)
EXCLUDE_KEYWORDS = ["[공통]", "[공급망운영 그룹]","EDP 파트 주요 이슈"]

# 기본 검색 기간 (일)
DEFAULT_LOOKBACK_DAYS = 7

# 시간대 설정
KST = timezone(timedelta(hours=9))


# ==================== 유틸리티 함수 ====================
def _safe_filename(s: str) -> str:
    """파일명으로 사용할 수 없는 문자 제거"""
    s = (s or "").strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"\s+", " ", s)
    return (s[:120] or "no_subject").strip()


def _decode_str(value: str) -> str:
    """헤더 문자열 디코딩"""
    if not value:
        return ""
    return str(make_header(decode_header(value))).strip()


def _msg_date_kst(msg: Message) -> Optional[datetime]:
    """메일 날짜를 KST로 변환"""
    raw = msg.get("Date")
    if not raw:
        return None
    try:
        dt = parsedate_to_datetime(raw)
        if dt is None:
            return None
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(KST)
    except Exception:
        return None


def _get_week_range(target_date: Optional[datetime] = None) -> Tuple[datetime, datetime]:
    """
    해당 날짜가 속한 주의 월요일부터 일요일까지의 범위 반환
    """
    if target_date is None:
        target_date = datetime.now(KST)
    
    # 월요일(0)부터 일요일(6)까지, weekday()는 월요일=0
    weekday = target_date.weekday()
    monday = target_date - timedelta(days=weekday)
    sunday = monday + timedelta(days=6)
    
    # 시간 초기화
    monday = monday.replace(hour=0, minute=0, second=0, microsecond=0)
    sunday = sunday.replace(hour=23, minute=59, second=59, microsecond=999999)
    
    return monday, sunday


def _extract_text_from_msg(msg: Message) -> str:
    """
    메일 본문에서 텍스트 추출
    BeautifulSoup을 사용하여 HTML을 깔끔하게 처리
    """
    text = ""
    
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            
            # 첨부파일 제외
            if "attachment" in content_disposition:
                continue
            
            if content_type == "text/plain":
                try:
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or "utf-8"
                    text += payload.decode(charset, errors="ignore")
                except Exception:
                    pass
            elif content_type == "text/html":
                try:
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or "utf-8"
                    html_content = payload.decode(charset, errors="ignore")
                    # BeautifulSoup으로 HTML 파싱
                    if HAS_BS4:
                        soup = BeautifulSoup(html_content, 'html.parser')
                        # style, script, table, img 태그 제거
                        for tag in soup(['style', 'script', 'head', 'meta', 'link', 'table', 'img', 'figure', 'figcaption']):
                            tag.decompose()
                        # 텍스트 추출 (줄바꿈 유지)
                        for br in soup.find_all('br'):
                            br.replace_with('\n')
                        for p in soup.find_all('p'):
                            p.append('\n')
                        for div in soup.find_all('div'):
                            div.append('\n')
                        text += soup.get_text(separator='', strip=False)
                    else:
                        # BeautifulSoup이 없으면 정규식으로 처리
                        # style, script 태그 내용 제거
                        html_content = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
                        html_content = re.sub(r'<script[^>]*>.*?</script>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
                        # HTML 태그 제거
                        text += re.sub(r'<[^>]+>', ' ', html_content)
                except Exception:
                    pass
    else:
        content_type = msg.get_content_type()
        try:
            payload = msg.get_payload(decode=True)
            charset = msg.get_content_charset() or "utf-8"
            text = payload.decode(charset, errors="ignore")
            if content_type == "text/html":
                if HAS_BS4:
                    soup = BeautifulSoup(text, 'html.parser')
                    for tag in soup(['style', 'script', 'head', 'meta', 'link', 'table', 'img', 'figure', 'figcaption']):
                        tag.decompose()
                    # 텍스트 추출 (줄바꿈 유지)
                    for br in soup.find_all('br'):
                        br.replace_with('\n')
                    for p in soup.find_all('p'):
                        p.append('\n')
                    for div in soup.find_all('div'):
                        div.append('\n')
                    text = soup.get_text(separator='', strip=False)
                else:
                    text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
                    text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
                    text = re.sub(r'<[^>]+>', ' ', text)
        except Exception:
            pass
    
    # 연속된 공백은 단일 공백으로, 줄바꿈은 유지
    text = re.sub(r'[ \t]+', ' ', text)  # 탭과 공백만 단일 공백으로
    text = re.sub(r'\n\s*\n', '\n\n', text)  # 빈 줄 정리
    text = text.strip()
    return text


def _summarize_text(text: str, max_length: int = SUMMARY_LENGTH) -> str:
    """
    텍스트 요약 (단순 트렁케이션)
    향후 LLM 기반 요약으로 확장 가능
    """
    if not text:
        return ""
    
    # 의미 없는 패턴 제거 (CSS 코드, 이메일 주소 등)
    text = re.sub(r'@charset[^;]+;', '', text)
    text = re.sub(r'body,html\{[^}]+\}', '', text)
    text = re.sub(r'overflow[^;]+;', '', text)
    text = re.sub(r'Copyright[^}]+', '', text)
    text = re.sub(r'All Right Reserved', '', text)
    text = re.sub(r'[\w\.-]+@[\w\.-]+\.\w+', '[이메일]', text)  # 이메일 주소 마스킹
    
    # 연속된 공백 제거
    text = re.sub(r'\s+', ' ', text).strip()
    
    if not text:
        return ""
    
    # 최대 길이로 자르고, 문장 끝에서 자르기 시도
    if len(text) <= max_length:
        return text
    
    truncated = text[:max_length]
    # 마지막 문장 끝 찾기
    last_period = truncated.rfind('.')
    last_exclamation = truncated.rfind('!')
    last_question = truncated.rfind('?')
    
    last_sentence_end = max(last_period, last_exclamation, last_question)
    
    if last_sentence_end > max_length * 0.5:  # 너무 짧게 잘리지 않도록
        return truncated[:last_sentence_end + 1]
    
    return truncated + "..."


def _format_text_to_html(text: str) -> str:
    """
    텍스트를 HTML로 변환하여 가독성 향상
    - 헤더 정보(Sender, Date, Title)를 별도 박스로 분리
    - ■ 기호로 시작하는 섹션을 구분
    - 줄바꿈을 적절하게 처리
    """
    if not text:
        return ""
    
    lines = text.split('\n')
    html_parts = []
    
    # 헤더 정보 추출 (Sender, Date, Title 등)
    header_info = []
    body_lines = []
    in_header = True
    
    for line in lines:
        line = line.strip()
        if not line:
            # 빈 줄은 본문에 유지
            if not in_header:
                body_lines.append('')
            continue
        
        # 헤더 패턴 확인
        header_patterns = [
            r'^Sender\s*:',
            r'^From\s*:',
            r'^Date\s*:',
            r'^Sent\s*:',
            r'^Cc\s*:',
            r'^Subject\s*:',
            r'^Title\s*:',
        ]
        
        is_header = any(re.match(p, line, re.IGNORECASE) for p in header_patterns)
        
        if is_header and in_header:
            header_info.append(line)
        else:
            in_header = False
            body_lines.append(line)
    
    # 헤더 정보 HTML 생성
    if header_info:
        html_parts.append('<div class="mail-header">')
        for h in header_info:
            # 이메일 주소 강조
            h = re.sub(r'([\w\.-]+@[\w\.-]+\.\w+)', r'<span class="email">\1</span>', h)
            html_parts.append(f'<div class="header-item">{h}</div>')
        html_parts.append('</div>')
    
    # 본문 처리
    if body_lines:
        html_parts.append('<div class="mail-body">')
        
        current_section = None
        for line in body_lines:
            # ■ 기호로 시작하는 섹션 헤더
            if line.startswith('■'):
                if current_section is not None:
                    html_parts.append('</div>')
                section_title = line[1:].strip()
                html_parts.append(f'<div class="section"><div class="section-title">■ {section_title}</div><div class="section-content">')
                current_section = True
            # 인용문 구분선
            elif re.match(r'^-+Original Message-+$', line, re.IGNORECASE):
                html_parts.append('<div class="quote-divider"></div>')
                html_parts.append(f'<div class="quote-text">{line}</div>')
            # 일반 텍스트
            else:
                # 이메일 주소 강조
                line = re.sub(r'([\w\.-]+@[\w\.-]+\.\w+)', r'<span class="email">\1</span>', line)
                if line:  # 빈 줄이 아닌 경우만 추가
                    html_parts.append(f'<div class="body-line">{line}</div>')
                else:
                    html_parts.append('<div class="body-line">&nbsp;</div>')
        
        if current_section is not None:
            html_parts.append('</div></div>')
        html_parts.append('</div>')
    
    return ''.join(html_parts)


# ==================== 메일 크롤러 ====================
class MailCrawler:
    """메일 크롤링 클래스"""
    
    def __init__(self, host: str = POP3_HOST, user: str = USER, password: str = PW):
        self.host = host
        self.user = user
        self.password = password
        self.server = None
    
    def connect(self):
        """POP3 서버 연결"""
        self.server = poplib.POP3_SSL(self.host)
        self.server.user(self.user)
        self.server.pass_(self.password)
    
    def disconnect(self):
        """서버 연결 종료"""
        if self.server:
            self.server.quit()
            self.server = None
    
    def fetch_mails_in_range(
        self,
        start_date: datetime,
        end_date: datetime,
        max_mails: int = 1000,
        keywords: Optional[List[str]] = None
    ) -> List[Dict]:
        """
        지정된 날짜 범위의 메일 가져오기
        
        Args:
            start_date: 시작 날짜 (KST)
            end_date: 종료 날짜 (KST)
            max_mails: 최대 메일 수
            keywords: 필터링할 키워드 리스트 (None이면 필터링 안함)
        
        Returns:
            메일 정보 리스트 [{date, from, subject, body, summary}, ...]
        """
        if not self.server:
            self.connect()
        
        mail_count = self.server.stat()[0]
        mails = []
        
        # 최대 메일 수 제한
        start = min(mail_count, max_mails)
        end = max(1, mail_count - max_mails + 1)
        
        print(f"메일 크롤링 시작: {start_date.date()} ~ {end_date.date()}")
        print(f"총 메일 수: {mail_count}, 검색 범위: {start} ~ {end}")
        
        for i in range(start, end - 1, -1):
            try:
                raw_bytes = b"\n".join(self.server.retr(i)[1])
                msg = email.message_from_bytes(raw_bytes)
                
                dt_kst = _msg_date_kst(msg)
                if not dt_kst:
                    continue
                
                # 날짜 범위 체크
                if dt_kst < start_date or dt_kst > end_date:
                    continue
                
                subject = _decode_str(msg.get("Subject", ""))
                from_ = _decode_str(msg.get("From", ""))
                body = _extract_text_from_msg(msg)
                summary = _summarize_text(body)
                
                # 제외 키워드 필터링
                if EXCLUDE_KEYWORDS:
                    subject_lower = subject.lower()
                    body_lower = body.lower()
                    exclude_match = any(
                        kw.lower() in subject_lower or kw.lower() in body_lower
                        for kw in EXCLUDE_KEYWORDS
                    )
                    if exclude_match:
                        continue
                
                # 포함 키워드 필터링 (제목만, 공백 무시)
                if keywords:
                    subject_normalized = re.sub(r'\s+', '', subject.lower())
                    keyword_match = any(
                        re.sub(r'\s+', '', kw.lower()) in subject_normalized
                        for kw in keywords
                    )
                    if not keyword_match:
                        continue
                
                mail_info = {
                    "date": dt_kst,
                    "from": from_,
                    "subject": subject,
                    "body": body,
                    "summary": summary
                }
                mails.append(mail_info)
                
                print(f"[{len(mails)}] {dt_kst.strftime('%Y-%m-%d %H:%M')} | {from_} | {subject[:50]}")
                
            except Exception as e:
                print(f"메일 {i} 처리 중 오류: {e}")
                continue
        
        # 날짜순 정렬 (최신순)
        mails.sort(key=lambda x: x["date"], reverse=True)
        
        return mails


# ==================== 리포트 생성기 ====================
class ReportGenerator:
    """HTML 리포트 생성 클래스"""
    
    @staticmethod
    def generate_html_report(
        mails: List[Dict],
        start_date: datetime,
        end_date: datetime
    ) -> str:
        """
        메일 리스트를 HTML 리포트로 변환
        
        Args:
            mails: 메일 정보 리스트
            start_date: 시작 날짜
            end_date: 종료 날짜
        
        Returns:
            HTML 형식의 리포트 문자열
        """
        # HTML 이스케이프 함수
        def escape_html(text: str) -> str:
            if text is None:
                return ""
            return html_module.escape(str(text), quote=True)
        
        # 카테고리별 메일 그룹화
        def get_category(mail: Dict) -> str:
            """메일의 카테고리 반환 (제목만, 공백 무시)"""
            subject_normalized = re.sub(r'\s+', '', mail["subject"].lower())
            
            for keyword in FILTER_KEYWORDS:
                keyword_normalized = re.sub(r'\s+', '', keyword.lower())
                if keyword_normalized in subject_normalized:
                    return keyword
            return "기타"
        
        start_str = start_date.strftime("%Y-%m-%d")
        end_str = end_date.strftime("%Y-%m-%d")
        
        # HTML 템플릿
        html_template_raw = """<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <style>
        body {
            font-family: "Malgun Gothic", "맑은 고딕", Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            margin: 0;
            padding: 20px;
            background-color: #fafafa;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        }
        h1 {
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .report-link-btn {
            background-color: #5dade2;
            color: white;
            padding: 8px 16px;
            border-radius: 5px;
            text-decoration: none;
            font-size: 0.9em;
            font-weight: bold;
            transition: background-color 0.3s;
        }
        .report-link-btn:hover {
            background-color: #3498db;
        }
        .summary {
            background-color: #e3f2fd;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
            border-left: 4px solid #3498db;
        }
        .category-section {
            margin-bottom: 30px;
        }
        .category-header {
            padding: 12px 15px;
            border-radius: 5px 5px 0 0;
            font-weight: bold;
            font-size: 1.1em;
        }
        .category-emoji {
            margin-right: 8px;
            font-size: 1.2em;
        }
        .category-count {
            background-color: rgba(255, 255, 255, 0.7);
            padding: 2px 8px;
            border-radius: 10px;
            margin-left: 10px;
            font-size: 0.9em;
            color: #333;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 0;
        }
        th {
            background-color: #f8f9fa;
            color: #2c3e50;
            padding: 12px;
            text-align: left;
            font-weight: bold;
            border-bottom: 2px solid #e0e0e0;
        }
        td {
            padding: 12px;
            border-bottom: 1px solid #eee;
            vertical-align: top;
        }
        tr:hover {
            background-color: #f0f4f8;
        }
        .date {
            white-space: nowrap;
            color: #666;
            font-size: 0.9em;
        }
        .from {
            font-weight: bold;
            color: #2c3e50;
        }
        .subject {
            font-weight: bold;
            color: #2c3e50;
        }
        .summary-text {
            color: #555;
            font-size: 0.95em;
        }
        details {
            margin: 0;
        }
        summary {
            color: #3498db;
            font-weight: bold;
            cursor: pointer;
            text-decoration: underline;
        }
        summary:hover {
            color: #2980b9;
        }
        .mail-detail {
            background-color: #fafafa;
            padding: 15px;
            margin-top: 10px;
            border-radius: 5px;
            border-left: 3px solid #3498db;
        }
        .detail-header {
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 10px;
            font-size: 0.95em;
        }
        .detail-body {
            word-wrap: break-word;
            color: #333;
            line-height: 1.5;
            font-size: 0.95em;
        }
        .mail-header {
            background-color: #e8f4f8;
            padding: 10px 15px;
            border-radius: 5px;
            margin-bottom: 15px;
            border-left: 4px solid #3498db;
        }
        .header-item {
            padding: 2px 0;
            color: #2c3e50;
            font-size: 0.9em;
            line-height: 1.4;
        }
        .header-item .email {
            color: #2980b9;
            font-weight: 500;
        }
        .mail-body {
            padding: 10px 0;
        }
        .section {
            margin-top: 15px;
            background-color: #ffffff;
            border-radius: 5px;
            padding: 12px 15px;
            border: 1px solid #e8e8e8;
        }
        .section-title {
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 10px;
            padding-bottom: 5px;
            border-bottom: 1px solid #eee;
        }
        .section-content {
            padding-left: 5px;
        }
        .body-line {
            padding: 4px 0;
            color: #444;
            line-height: 1.5;
        }
        .body-line .email {
            color: #2980b9;
            font-weight: 500;
        }
        .quote-divider {
            border-top: 2px dashed #ddd;
            margin: 15px 0;
        }
        .quote-text {
            color: #999;
            font-style: italic;
            font-size: 0.85em;
            text-align: center;
            margin: 5px 0;
            padding: 5px;
        }
        .no-mail {
            text-align: center;
            padding: 40px;
            color: #999;
            font-style: italic;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>
            <span>이슈지 리포트</span>
            <a href="https://confluence.samsungds.net/spaces/GOCMEM/pages/3150978308/%EC%9D%B4%EC%8A%88%EC%A7%80" class="report-link-btn" target="_blank">이슈지 바로 가기</a>
        </h1>
        
        <div class="summary">
            <strong>🗓️ 기간:</strong> {start_str} ~ {end_str}<br>
            <strong>📧 총 메일 수:</strong> {mail_count}건<br>
            <strong>🕒 매주 월요일 오전 발송</strong>
        </div>
{content}
    </div>
</body>
</html>
"""

        # ✅ CSS/JS의 { } 를 format이 오해하지 않도록 전부 이스케이프하고,
        # ✅ 우리가 쓸 플레이스홀더만 다시 복구
        html_template = (
            html_template_raw
            .replace("{", "{{").replace("}", "}}")
            .replace("{{start_str}}", "{start_str}")
            .replace("{{end_str}}", "{end_str}")
            .replace("{{mail_count}}", "{mail_count}")
            .replace("{{content}}", "{content}")
        )

        
        # 메일이 없는 경우
        if not mails:
            content = """
        <div class="no-mail">
            해당 기간에 수신된 메일이 없습니다.
        </div>
"""
            html = html_template.format(
                start_str=start_str,
                end_str=end_str,
                mail_count=len(mails),
                content=content
            )
            return html
        
        # 메일을 카테고리별로 그룹화
        categorized_mails = {}
        for mail in mails:
            category = get_category(mail)
            if category not in categorized_mails:
                categorized_mails[category] = []
            categorized_mails[category].append(mail)
        
        # 카테고리 순서 (FILTER_KEYWORDS 순서 + 기타)
        category_order = [kw for kw in FILTER_KEYWORDS if kw in categorized_mails]
        if "기타" in categorized_mails:
            category_order.append("기타")
        
        # 카테고리별 HTML 생성
        content_parts = []
        mail_idx = 0
        
        for category in category_order:
            category_mails = categorized_mails[category]
            style = CATEGORY_STYLES.get(category, CATEGORY_STYLES["기타"])
            category_html = f"""
        <div class="category-section">
            <div class="category-header" style="background-color: {style['color']}; color: #333;">
                <span class="category-emoji">{style['emoji']}</span>
                {category} <span class="category-count">{len(category_mails)}건</span>
            </div>
            <table>
                <thead>
                    <tr>
                        <th style="width: 15%">날짜</th>
                        <th style="width: 20%">발신자</th>
                        <th style="width: 25%">제목</th>
                        <th style="width: 40%">요약</th>
                    </tr>
                </thead>
                <tbody>
"""
            for mail in category_mails:
                date_str = mail["date"].strftime("%m-%d %H:%M")
                from_ = mail["from"][:50] if len(mail["from"]) > 50 else mail["from"]
                subject = mail["subject"][:60] if len(mail["subject"]) > 60 else mail["subject"]
                summary = mail["summary"]
                full_body = _format_text_to_html(mail["body"])
                
                category_html += f"""
                    <tr>
                        <td class="date">{date_str}</td>
                        <td class="from">{from_}</td>
                        <td class="subject">
                            <details>
                                <summary>{subject}</summary>
                                <div class="mail-detail">
                                    <div class="detail-header">📧 전체 내용</div>
                                    <div class="detail-body">{full_body}</div>
                                </div>
                            </details>
                        </td>
                        <td class="summary-text">{summary}</td>
                    </tr>
"""
                mail_idx += 1
            
            category_html += """
                </tbody>
            </table>
        </div>
"""
            content_parts.append(category_html)
        
        content = "".join(content_parts)
        html = html_template.format(
            start_str=start_str,
            end_str=end_str,
            mail_count=len(mails),
            content=content
        )
        return html


# ==================== 메일 리포터 ====================
class MailReporter:
    """메일 리포팅 메인 클래스"""
    
    def __init__(self, recipient_ids: Union[List[str], str] = None, sender: str = REPORT_SENDER):
        if recipient_ids is None:
            recipient_ids = REPORT_RECIPIENT_IDS
        # 문자열인 경우 리스트로 변환
        if isinstance(recipient_ids, str):
            recipient_ids = [recipient_ids]
        # 아이디에 @samsung.com 도메인 추가
        self.recipients = [
            f"{rid}@{REPORT_DOMAIN}" if "@" not in rid else rid
            for rid in recipient_ids
        ]
        self.sender = sender
        self.crawler = MailCrawler()
    
    def generate_weekly_report(self, target_date: Optional[datetime] = None, keywords: Optional[List[str]] = None) -> bool:
        """
        주간 리포트 생성 및 발송
        
        Args:
            target_date: 기준 날짜 (None이면 현재 날짜)
            keywords: 필터링할 키워드 리스트 (None이면 기본 키워드 사용)
        
        Returns:
            성공 여부
        """
        start_date, end_date = _get_week_range(target_date)
        return self.generate_report(start_date, end_date, keywords)
    
    def generate_custom_report(self, start_date: datetime, end_date: datetime, keywords: Optional[List[str]] = None) -> bool:
        """
        사용자 지정 기간 리포트 생성 및 발송
        
        Args:
            start_date: 시작 날짜
            end_date: 종료 날짜
            keywords: 필터링할 키워드 리스트 (None이면 기본 키워드 사용)
        
        Returns:
            성공 여부
        """
        return self.generate_report(start_date, end_date, keywords)
    
    def generate_report(self, start_date: datetime, end_date: datetime, keywords: Optional[List[str]] = None) -> bool:
        """
        리포트 생성 및 발송 공통 메서드
        
        Args:
            start_date: 시작 날짜
            end_date: 종료 날짜
            keywords: 필터링할 키워드 리스트 (None이면 기본 키워드 사용)
        
        Returns:
            성공 여부
        """
        if keywords is None:
            keywords = FILTER_KEYWORDS
        
        try:
            # 1. 메일 크롤링
            self.crawler.connect()
            mails = self.crawler.fetch_mails_in_range(start_date, end_date, keywords=keywords)
            self.crawler.disconnect()
            
            keyword_str = ", ".join(keywords) if keywords else "없음"
            exclude_str = ", ".join(EXCLUDE_KEYWORDS) if EXCLUDE_KEYWORDS else "없음"
            print(f"\n포함 키워드: {keyword_str}")
            print(f"제외 키워드: {exclude_str}")
            print(f"총 {len(mails)}건의 메일을 찾았습니다.")
            
            # 2. HTML 리포트 생성
            report_html = ReportGenerator.generate_html_report(mails, start_date, end_date)
            
            # 3. 메일 발송
            subject = f"[이슈지 리포트] 주간 이슈 정리 ({start_date.strftime('%m-%d')} ~ {end_date.strftime('%m-%d')})"
            
            recipients = [
                {"emailAddress": recipient, "recipientType": "TO"}
                for recipient in self.recipients
            ]
            
            print(f"\n리포트 발송 중... 수신자: {', '.join(self.recipients)}")
            
            result = send_mail_api(
                sender_id=self.sender,
                subject=subject,
                contents=report_html,
                content_type="HTML",
                recipients=recipients
            )
            
            print(f"발송 결과: {result}")
            print("리포트 발송 완료!")
            
            return True
            
        except Exception as e:
            print(f"리포트 생성/발송 중 오류 발생: {e}")
            return False


# ==================== 메인 ====================
def main():
    """메인 실행 함수"""
    global PW
    global EXCLUDE_KEYWORDS
    import argparse
    
    parser = argparse.ArgumentParser(description="메일 리포팅 시스템")
    parser.add_argument(
        "--mode",
        choices=["weekly", "custom", "recent"],
        default="recent",
        help="리포트 모드: weekly(주별), custom(지정 날짜), recent(최근 N일)"
    )
    parser.add_argument(
        "--start",
        help="시작 날짜 (YYYY-MM-DD), custom 모드에서만 사용"
    )
    parser.add_argument(
        "--end",
        help="종료 날짜 (YYYY-MM-DD), custom 모드에서만 사용"
    )
    parser.add_argument(
        "--days",
        type=int,
        default=DEFAULT_LOOKBACK_DAYS,
        help=f"최근 N일 (recent 모드, 기본값: {DEFAULT_LOOKBACK_DAYS})"
    )
    parser.add_argument(
        "--keywords",
        nargs="+",
        help=f"포함할 키워드 (공백으로 구분, 기본값: {', '.join(FILTER_KEYWORDS)})"
    )
    parser.add_argument(
        "--exclude",
        nargs="+",
        help=f"제외할 키워드 (공백으로 구분, 기본값: {', '.join(EXCLUDE_KEYWORDS)})"
    )
    parser.add_argument(
        "--no-filter",
        action="store_true",
        help="키워드 필터링 사용 안함"
    )
    parser.add_argument(
        "--recipient",
        nargs="*",
        default=REPORT_RECIPIENT_IDS,
        help=f"리포트 수신자 아이디 (공백으로 구분, @samsung.com 자동 추가, 기본값: {', '.join(REPORT_RECIPIENT_IDS)})"
    )
    
    args = parser.parse_args()
    
    # 비밀번호 입력 확인
    if not PW:
        PW = input("MAIL_PASS 입력: ").strip()
    
    # 키워드 설정
    keywords = None
    if args.no_filter:
        keywords = None
    elif args.keywords:
        keywords = args.keywords
    else:
        keywords = FILTER_KEYWORDS
    
    # 제외 키워드 설정
    if args.exclude:
        EXCLUDE_KEYWORDS = args.exclude
    
    # 리포터 생성
    reporter = MailReporter(recipient_ids=args.recipient)
    
    if args.mode == "weekly":
        print("=== 주간 리포트 모드 ===")
        success = reporter.generate_weekly_report(keywords=keywords)
    elif args.mode == "recent":
        print(f"=== 최근 {args.days}일 리포트 모드 ===")
        end_date = datetime.now(KST)
        start_date = end_date - timedelta(days=args.days)
        start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
        success = reporter.generate_custom_report(start_date, end_date, keywords=keywords)
    else:
        print("=== 지정 날짜 리포트 모드 ===")
        if not args.start or not args.end:
            print("오류: custom 모드에서는 --start와 --end가 필요합니다.")
            print("예: python mail_reporter.py --mode custom --start 2025-01-20 --end 2025-01-26")
            return
        
        try:
            start_date = datetime.strptime(args.start, "%Y-%m-%d").replace(tzinfo=KST)
            end_date = datetime.strptime(args.end, "%Y-%m-%d").replace(
                hour=23, minute=59, second=59, microsecond=999999, tzinfo=KST
            )
            success = reporter.generate_custom_report(start_date, end_date, keywords=keywords)
        except ValueError as e:
            print(f"날짜 형식 오류: {e}")
            print("날짜는 YYYY-MM-DD 형식으로 입력해주세요.")
            return
    
    if success:
        print("\n✅ 리포트 생성 및 발송 완료!")
    else:
        print("\n❌ 리포트 생성 및 발송 실패!")


if __name__ == "__main__":
    main()
