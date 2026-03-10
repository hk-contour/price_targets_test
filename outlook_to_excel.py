import os
import re
import pandas as pd
import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()

IMAP_USER = os.getenv("IMAP_USER")  # your email: hari.kumar@contourasset.com
IMAP_PASS = os.getenv("IMAP_PASS")  # app password or account password
IMAP_SERVER = os.getenv("IMAP_SERVER", "outlook.office365.com")
SHARED_MAILBOX = os.getenv("SHARED_MAILBOX", "")  # e.g. pricetargets@contourasset.com


# -----------------------
# EMAIL FETCH
# -----------------------
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
    if not IMAP_USER or not IMAP_PASS:
        raise RuntimeError("IMAP_USER and IMAP_PASS must be set in .env")

    # For shared mailboxes in Office 365, login as user\shared_mailbox
    login_user = IMAP_USER
    if SHARED_MAILBOX:
        login_user = f"{IMAP_USER}\\{SHARED_MAILBOX}"

    emails = get_emails_imap(login_user, IMAP_PASS)

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


if __name__ == "__main__":
    run()
