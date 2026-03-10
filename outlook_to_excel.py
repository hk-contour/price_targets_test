import os
import re
import requests
import pandas as pd
import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()

TENANT = os.getenv("TENANT_ID")
CLIENT = os.getenv("CLIENT_ID")
SECRET = os.getenv("CLIENT_SECRET")
MAILBOX = os.getenv("MAILBOX")
IMAP_USER = os.getenv("IMAP_USER")
IMAP_PASS = os.getenv("IMAP_PASS")
IMAP_SERVER = os.getenv("IMAP_SERVER", "outlook.office365.com")

FILE_ID = "YOUR_SHAREPOINT_FILE_ID"
GRAPH = "https://graph.microsoft.com/v1.0"


# -----------------------
# AUTH
# -----------------------
def token():
    # Keep MSAL token helper available if Graph access is still desired
    from msal import ConfidentialClientApplication

    app = ConfidentialClientApplication(
        CLIENT,
        authority=f"https://login.microsoftonline.com/{TENANT}",
        client_credential=SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    return result["access_token"]


# -----------------------
# EMAIL FETCH
# -----------------------
def get_emails(token):
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"{GRAPH}/users/{MAILBOX}/mailFolders/inbox/messages"
        f"?$top=25&$select=subject,receivedDateTime,body"
    )
    r = requests.get(url, headers=headers)
    return r.json()["value"]


def _decode_mime_words(value):
    if not value:
        return ""
    parts = decode_header(value)
    out = []
    for part, enc in parts:
        if isinstance(part, bytes):
            out.append(part.decode(enc or "utf-8", errors="ignore"))
        else:
            out.append(part)
    return "".join(out)


def get_emails_imap(user, password, server=IMAP_SERVER, mailbox="INBOX", limit=25):
    m = imaplib.IMAP4_SSL(server)
    m.login(user, password)
    m.select(mailbox)
    status, data = m.search(None, "ALL")
    if status != "OK":
        m.logout()
        return []

    ids = data[0].split()
    if not ids:
        m.logout()
        return []

    # take the last `limit` messages
    last_ids = ids[-limit:]
    emails = []

    for num in reversed(last_ids):
        status, msg_data = m.fetch(num, "(RFC822)")
        if status != "OK":
            continue
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)
        subject = _decode_mime_words(msg.get("Subject"))

        # parse date
        date_hdr = msg.get("Date")
        try:
            dt = parsedate_to_datetime(date_hdr)
        except Exception:
            dt = datetime.utcnow()

        # get body (prefer plain text)
        body_text = ""
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                disp = str(part.get("Content-Disposition"))
                if ctype == "text/plain" and "attachment" not in disp:
                    try:
                        body_text = part.get_payload(decode=True).decode(
                            part.get_content_charset() or "utf-8", errors="ignore"
                        )
                        break
                    except Exception:
                        continue
            if not body_text:
                # fallback to first text/html
                for part in msg.walk():
                    if part.get_content_type() == "text/html":
                        try:
                            body_text = part.get_payload(decode=True).decode(
                                part.get_content_charset() or "utf-8", errors="ignore"
                            )
                            break
                        except Exception:
                            continue
        else:
            try:
                body_text = msg.get_payload(decode=True).decode(
                    msg.get_content_charset() or "utf-8", errors="ignore"
                )
            except Exception:
                body_text = str(msg.get_payload())

        # strip simple HTML tags if present
        if "<" in body_text and ">" in body_text:
            body_text = re.sub(r"<[^>]+>", "", body_text)

        emails.append(
            {
                "subject": subject,
                "receivedDateTime": dt.isoformat(),
                "body": {"content": body_text},
            }
        )

    m.logout()
    return emails


# -----------------------
# PARSERS
# -----------------------
def extract_ticker(subject):
    m = re.search(r"\(([A-Z]+)", subject)
    if m:
        return m.group(1)
    return None


def extract_targets(text):
    upside = re.search(r"Upside[:=]\s*\$?([0-9\.]+)", text, re.I)
    downside = re.search(r"Downside[:=]\s*\$?([0-9\.]+)", text, re.I)
    u = float(upside.group(1)) if upside else None
    d = float(downside.group(1)) if downside else None
    return u, d


# -----------------------
# EXCEL
# -----------------------
def download_excel(token):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH}/me/drive/items/{FILE_ID}/content"
    r = requests.get(url, headers=headers)
    with open("targets.xlsx", "wb") as f:
        f.write(r.content)


def upload_excel(token):
    headers = {"Authorization": f"Bearer {token}"}
    with open("targets.xlsx", "rb") as f:
        data = f.read()
    url = f"{GRAPH}/me/drive/items/{FILE_ID}/content"
    requests.put(url, headers=headers, data=data)


# -----------------------
# UPDATE TABLE
# -----------------------
def update_excel(row):
    df = pd.read_excel("targets.xlsx")
    issuer = row["Issuer"]
    today = row["BeginDate"]

    # update previous EndDate
    prev = df[df["Issuer"] == issuer]
    if not prev.empty:
        idx = prev.index[0]
        df.loc[idx, "EndDate"] = today - timedelta(days=1)

    # insert new row at top
    df = pd.concat([pd.DataFrame([row]), df]).reset_index(drop=True)
    df.to_excel("targets.xlsx", index=False)


# -----------------------
# MAIN
# -----------------------
def run():
    # Determine email source: prefer IMAP if credentials provided
    emails = []
    t = None

    if IMAP_USER and IMAP_PASS:
        emails = get_emails_imap(IMAP_USER, IMAP_PASS)
        use_graph_for_files = False
    else:
        t = token()
        emails = get_emails(t)
        use_graph_for_files = True

    # Excel file: if FILE_ID left as placeholder, operate on local targets.xlsx
    use_graph_for_files = use_graph_for_files and (
        FILE_ID and FILE_ID != "YOUR_SHAREPOINT_FILE_ID"
    )

    if use_graph_for_files:
        download_excel(t)
    else:
        # ensure a local file exists
        if not os.path.exists("targets.xlsx"):
            df = pd.DataFrame(
                columns=[
                    "BeginDate",
                    "EndDate",
                    "Issuer",
                    "Upside Price Target",
                    "Downside Price Target",
                ]
            )
            df.to_excel("targets.xlsx", index=False)

    for mail in emails:
        subject = mail["subject"]
        ticker = extract_ticker(subject)
        if not ticker:
            continue
        body = mail["body"]["content"]
        upside, downside = extract_targets(body)
        if upside is None or downside is None:
            continue
        date = datetime.fromisoformat(
            mail["receivedDateTime"].replace("Z", "")
        )
        row = {
            "BeginDate": date.date(),
            "EndDate": "",
            "Issuer": ticker,
            "Upside Price Target": upside,
            "Downside Price Target": downside,
        }
        update_excel(row)

    if use_graph_for_files:
        upload_excel(t)


if __name__ == "__main__":
    run()
