# -*- coding: utf-8 -*-
import os, json, mimetypes, requests
from typing import List, Dict, Optional, Any

CONFIG = {
    "HOST": os.getenv("MAIL_API_HOST", "https://openapi.samsung.net"),
    "TOKEN": os.getenv("MAIL_API_TOKEN", "Bearer 931e0fcb-31b8-33cf-8699-0d0ef752c85b"),
    "SYSTEM_ID": os.getenv("MAIL_API_SYSTEM_ID", "KCC10REST00621"),
}

def send_mail_api(
    *, sender_id: str, subject: str, contents: str,
    content_type: str = "HTML", doc_secu_type: str = "PERSONAL",
    recipients: List[Dict[str, str]] = None,
    reserved_time: Optional[str] = None,
    attachments: Optional[List[str]] = None,
    proxies: Optional[Dict[str, str]] = None,
    verify_ssl: bool = False, timeout: int = 30,
) -> dict:
    headers_common = {
        "Authorization": CONFIG["TOKEN"],
        "System-ID": CONFIG["SYSTEM_ID"],
    }
    mail_json: Dict[str, Any] = {
        "subject": subject,
        "contents": contents,
        "contentType": content_type,     # "TEXT" | "HTML" | "MIME"
        "docSecuType": doc_secu_type,    # "PERSONAL" | "OFFICIAL"
        "sender": {"emailAddress": f"{sender_id}@samsung.com"},
        "recipients": recipients or [],
    }
    if reserved_time:
        mail_json["reservedTime"] = reserved_time  # "yyyy-MM-dd HH:mm"

    url = f'{CONFIG["HOST"].rstrip("/")}/mail/api/v2.0/mails/send?userId={sender_id}'
    s = requests.Session()
    if proxies:
        s.proxies.update(proxies)

    attach_list = attachments or []
    if not attach_list:
        headers = dict(headers_common)
        headers["Content-Type"] = "application/json"
        r = s.post(url, data=json.dumps(mail_json, ensure_ascii=False), headers=headers,
                   verify=verify_ssl, timeout=timeout)
    else:
        files = [("mail", (None, json.dumps(mail_json, ensure_ascii=False), "application/json"))]
        for path in attach_list:
            fname = os.path.basename(path)
            ctype = mimetypes.guess_type(fname)[0] or "application/octet-stream"
            files.append(("attachments", (fname, open(path, "rb"), ctype)))
        r = s.post(url, headers=headers_common, files=files, verify=verify_ssl, timeout=timeout)
        for _, ft in files[1:]:
            try: ft[1].close()
            except Exception: pass

    r.raise_for_status()
    return r.json() if (r.text or "").strip() else {"ok": True}
