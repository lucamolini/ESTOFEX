# estofex_mailer.py (robusto)

import os
import re
import ssl
import smtplib
from email.message import EmailMessage
from datetime import datetime
from zoneinfo import ZoneInfo
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

import requests
from bs4 import BeautifulSoup

LIST_URL = os.getenv("LIST_URL", "https://www.estofex.org/cgi-bin/polygon/showforecast.cgi?list=yes")
ALT_LIST_URL = "https://www.estofex.org/cgi-bin/polygon/showforecast.cgi?all=yes&list=yes"
FILENAME_BASE = os.getenv("FILENAME_BASE", "estofex_latest")

TO_EMAIL = os.getenv("TO_EMAIL", "luca.molini@cimafoundation.org")
EMAIL_SUBJECT = os.getenv("EMAIL_SUBJECT", "ESTOFEX — mappa più recente")
EMAIL_BODY = os.getenv("EMAIL_BODY", "In allegato la mappa ESTOFEX (https://www.estofex.org) più recente (storm forecast).")

DEBUG_SMTP = os.getenv("DEBUG_SMTP", "0").strip() == "1"
CC_EMAILS = os.getenv("CC_EMAILS", "").strip()
BCC_EMAILS = os.getenv("BCC_EMAILS", "").strip()

ROME_HOUR_GATE = os.getenv("ROME_HOUR_GATE", "").strip()  # es. "17"; vuoto = no gate
FORCE_SEND = os.getenv("FORCE_SEND", "false").strip().lower() == "true"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9,it;q=0.8",
}


def guard_by_rome_hour() -> bool:
    if FORCE_SEND:
        print("[INFO] FORCE_SEND=true -> bypass controllo orario Europe/Rome.")
        return True
    if not ROME_HOUR_GATE:
        return True
    now_rome = datetime.now(ZoneInfo("Europe/Rome"))
    if str(now_rome.hour) != str(ROME_HOUR_GATE):
        print(f"[INFO] Europe/Rome ora {now_rome:%Y-%m-%d %H:%M:%S}; gate={ROME_HOUR_GATE} -> skip invio email.")
        return False
    return True


def extract_fcst_links_from_html(html: str, base_url: str):
    """Estrae TUTTI i link a showforecast.cgi con fcstfile=... (DOM + regex fallback)."""
    links = []

    # 1) DOM parsing
    try:
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if re.search(r"showforecast\.cgi.*fcstfile=", href, re.I):
                links.append(urljoin(base_url, href))
    except Exception:
        pass

    # 2) Regex fallback (se DOM non ha trovato nulla)
    if not links:
        for m in re.finditer(r'(?:href=["\']?)?(/?cgi-bin/\S*showforecast\.cgi\?[^"\'\s>]*fcstfile=[^"\'\s>]+)',
                             html, flags=re.I):
            href = m.group(1)
            links.append(urljoin(base_url, href))

    # Dedup preservando l'ordine
    seen = set()
    uniq = []
    for u in links:
        if u not in seen:
            uniq.append(u)
            seen.add(u)
    return uniq


def find_latest_fcst_url(list_url: str) -> str:
    """Trova il forecast più recente dalla lista; preferisce 'stormforecast.xml'."""
    for attempt, url in enumerate([list_url, ALT_LIST_URL], start=1):
        print(f"[INFO] ({attempt}/2) Apertura lista: {url}")
        r = requests.get(url, headers=HEADERS, timeout=30, allow_redirects=True)
        r.raise_for_status()
        html = r.text
        links = extract_fcst_links_from_html(html, url)
        if not links:
            print("[WARN] Nessun link trovato in questa lista (potrebbe essere markup minimale). "
                  f"Primi 300 char: {html[:300]!r}")
            continue

        # preferisci stormforecast.xml, altrimenti prendi il primo
        def is_stormforecast(u: str) -> bool:
            return "stormforecast.xml" in u.lower()

        best = next((u for u in links if is_stormforecast(u)), links[0])
        print(f"[INFO] Forecast più recente: {best}")
        return best

    raise RuntimeError("Nessun link di forecast trovato nelle liste ESTOFEX.")


def to_map_image_url(forecast_url: str) -> str:
    """...&text=yes -> ...&lightningmap=yes"""
    u = urlparse(forecast_url)
    qs = parse_qs(u.query, keep_blank_values=True)
    qs.pop("text", None)
    qs["lightningmap"] = ["yes"]
    new_query = urlencode({k: v[0] if isinstance(v, list) else v for k, v in qs.items()})
    return urlunparse((u.scheme, u.netloc, u.path, u.params, new_query, u.fragment))


def download_map_image(img_url: str, base_name: str) -> str:
    r = requests.get(img_url, headers=HEADERS, timeout=60)
    r.raise_for_status()
    ctype = (r.headers.get("Content-Type") or "").lower()

    if "image/" not in ctype:
        # fallback: prova da estensione nell'URL
        print(f"[WARN] Content-Type inatteso: {ctype!r}. Provo a inferire dall'URL.")
    if "png" in ctype or img_url.lower().endswith(".png"):
        ext = "png"
    elif "gif" in ctype or img_url.lower().endswith(".gif"):
        ext = "gif"
    elif "jpeg" in ctype or "jpg" in ctype or img_url.lower().endswith((".jpg", ".jpeg")):
        ext = "jpg"
    else:
        ext = "png"

    today = datetime.now(ZoneInfo("Europe/Rome")).strftime("%Y%m%d")
    filename = f"{base_name}_{today}.{ext}"
    with open(filename, "wb") as f:
        f.write(r.content)

    # copia anche il nome fisso
    with open(f"{base_name}.{ext}", "wb") as f:
        f.write(r.content)

    print(f"[OK] Mappa salvata: {filename} (e {base_name}.{ext})")
    return filename


def send_email_with_attachment(path: str) -> bool:
    SMTP_HOST = os.getenv("SMTP_HOST")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))
    SMTP_USER = os.getenv("SMTP_USER")
    SMTP_PASS = os.getenv("SMTP_PASS")
    FROM_EMAIL = os.getenv("FROM_EMAIL")

    missing = [k for k, v in {
        "SMTP_HOST": SMTP_HOST, "SMTP_USER": SMTP_USER, "SMTP_PASS": SMTP_PASS, "FROM_EMAIL": FROM_EMAIL
    }.items() if not v]
    if missing:
        print(f"[ERROR] Config SMTP mancante: {', '.join(missing)}. Email NON inviata.")
        return False

    cc_list = [x.strip() for x in CC_EMAILS.split(",") if x.strip()]
    bcc_list = [x.strip() for x in BCC_EMAILS.split(",") if x.strip()]
    all_rcpts = [TO_EMAIL] + cc_list + bcc_list

    msg = EmailMessage()
    msg["From"] = FROM_EMAIL
    msg["To"] = TO_EMAIL
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = EMAIL_SUBJECT
    msg["Date"] = datetime.now(ZoneInfo("Europe/Rome")).strftime("%a, %d %b %Y %H:%M:%S %z")
    msg.set_content(EMAIL_BODY)

    with open(path, "rb") as f:
        data = f.read()
    ext = os.path.splitext(path)[1].lower().strip(".") or "png"
    subtype = "png" if ext == "png" else "gif" if ext == "gif" else "jpeg" if ext in ("jpg", "jpeg") else "octet-stream"
    msg.add_attachment(data, maintype="image", subtype=subtype, filename=os.path.basename(path))

    context = ssl.create_default_context()
    print(f"[INFO] Invio email: from={FROM_EMAIL} to={all_rcpts} via {SMTP_HOST}:{SMTP_PORT}")

    try:
        if SMTP_PORT == 465:
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context, timeout=60) as s:
                if DEBUG_SMTP:
                    s.set_debuglevel(1)
                s.login(SMTP_USER, SMTP_PASS)
                refused = s.sendmail(FROM_EMAIL, all_rcpts, msg.as_string())
        else:
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as s:
                if DEBUG_SMTP:
                    s.set_debuglevel(1)
                s.starttls(context=context)
                s.login(SMTP_USER, SMTP_PASS)
                refused = s.sendmail(FROM_EMAIL, all_rcpts, msg.as_string())

        if refused:
            print(f"[WARN] Alcuni destinatari rifiutati: {refused}")
        else:
            print("[OK] Il server SMTP ha accettato tutti i destinatari.")
        return True
    except Exception as e:
        print(f"[ERROR] Invio email fallito: {e}")
        return False


def main():
    # 1) trova l’ultimo forecast e costruisci l’URL della mappa
    fcst_url = find_latest_fcst_url(LIST_URL)
    img_url = to_map_image_url(fcst_url)
    print(f"[INFO] URL immagine mappa: {img_url}")

    # 2) scarica sempre l’immagine (così puoi caricarla come artifact)
    out_path = download_map_image(img_url, FILENAME_BASE)

    # 3) invia email solo se (gate orario ok) o forzato
    if guard_by_rome_hour():
        sent = send_email_with_attachment(out_path)
        if sent:
            print(f"[OK] Email inviata a {TO_EMAIL} con allegato {out_path}")
        else:
            print("[WARN] Email non inviata (vedi log).")
    else:
        print("[INFO] Fuori orario: email non inviata. File scaricato correttamente.")


if __name__ == "__main__":
    main()
