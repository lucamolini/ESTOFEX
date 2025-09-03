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
FILENAME_BASE = os.getenv("FILENAME_BASE", "estofex_latest")

TO_EMAIL = os.getenv("TO_EMAIL", "luca.molini@cimafoundation.org")
EMAIL_SUBJECT = os.getenv("EMAIL_SUBJECT", "ESTOFEX — mappa più recente")
EMAIL_BODY = os.getenv("EMAIL_BODY", "In allegato la mappa ESTOFEX più recente (storm forecast).")

DEBUG_SMTP = os.getenv("DEBUG_SMTP", "0").strip() == "1"
CC_EMAILS = os.getenv("CC_EMAILS", "").strip()
BCC_EMAILS = os.getenv("BCC_EMAILS", "").strip()

ROME_HOUR_GATE = os.getenv("ROME_HOUR_GATE", "").strip()    # es. "17"; vuoto = no gate
FORCE_SEND = os.getenv("FORCE_SEND", "false").strip().lower() == "true"


def guard_by_rome_hour() -> bool:
    """Se FORCE_SEND è attivo, bypassa il gate orario; altrimenti rispetta ROME_HOUR_GATE (se definito)."""
    if FORCE_SEND:
        print("[INFO] FORCE_SEND=true -> bypass controllo orario Europe/Rome.")
        return True
    if not ROME_HOUR_GATE:
        return True
    now_rome = datetime.now(ZoneInfo("Europe/Rome"))
    if str(now_rome.hour) != str(ROME_HOUR_GATE):
        print(f"[INFO] Europe/Rome ora {now_rome:%Y-%m-%d %H:%M:%S}; gate={ROMЕ_HOUR_GATE} -> skip invio email.")
        return False
    return True


def find_latest_fcst_url(list_url: str) -> str:
    """
    Apre la pagina 'list=yes' e ricava il link al forecast più recente
    (ancora valido). Preferisce '...stormforecast.xml'; se non trovato, usa il primo fcstfile disponibile.
    """
    print(f"[INFO] Apertura lista: {list_url}")
    r = requests.get(list_url, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    # Trova tutti i link con fcstfile=...
    anchors = soup.find_all("a", href=re.compile(r"showforecast\.cgi\?fcstfile=", re.I))
    if not anchors:
        raise RuntimeError("Nessun link di forecast trovato nella lista ESTOFEX.")

    # Ordine pagina = più recente in alto → prendi il primo 'stormforecast', altrimenti il primo in assoluto
    def is_stormforecast(href: str) -> bool:
        return bool(re.search(r"stormforecast\.xml", href))

    first_storm = next((a for a in anchors if is_stormforecast(a.get("href", ""))), None)
    target = first_storm or anchors[0]
    href = target.get("href")
    full = urljoin(list_url, href)
    print(f"[INFO] Forecast più recente: {full}")
    return full


def to_map_image_url(forecast_url: str) -> str:
    """
    Converte un link tipo ...showforecast.cgi?fcstfile=...&text=yes
    in ...showforecast.cgi?fcstfile=...&lightningmap=yes (endpoint immagine).
    """
    u = urlparse(forecast_url)
    qs = parse_qs(u.query, keep_blank_values=True)
    # mantieni fcstfile, sostituisci text=yes -> lightningmap=yes
    qs.pop("text", None)
    qs["lightningmap"] = ["yes"]
    new_query = urlencode({k: v[0] if isinstance(v, list) else v for k, v in qs.items()})
    new_url = urlunparse((u.scheme, u.netloc, u.path, u.params, new_query, u.fragment))
    print(f"[INFO] URL immagine mappa: {new_url}")
    return new_url


def download_map_image(img_url: str, base_name: str) -> str:
    """Scarica l'immagine della mappa e salva come <base_name>_YYYYMMDD.png (o .gif)."""
    r = requests.get(img_url, timeout=60)
    r.raise_for_status()
    ctype = r.headers.get("Content-Type", "").lower()
    if "png" in ctype:
        ext = "png"
    elif "gif" in ctype:
        ext = "gif"
    elif "jpeg" in ctype or "jpg" in ctype:
        ext = "jpg"
    else:
        # fallback: prova a inferire da url
        ext = "png" if img_url.lower().endswith(".png") else "gif" if img_url.lower().endswith(".gif") else "png"

    today = datetime.now(ZoneInfo("Europe/Rome")).strftime("%Y%m%d")
    filename = f"{base_name}_{today}.{ext}"
    with open(filename, "wb") as f:
        f.write(r.content)

    # Copia anche un nome fisso per eventuale artifact
    with open(f"{base_name}.{'png' if ext=='png' else ext}", "wb") as f:
        f.write(r.content)

    print(f"[OK] Mappa salvata: {filename}")
    return filename


def send_email_with_attachment(path: str) -> bool:
    """Invia la mail con allegato; ritorna True/False."""
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

    # Destinatari
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

    # Allegato
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
    # 1) Prendi l’ultimo forecast e costruisci l’URL mappa
    fcst_url = find_latest_fcst_url(LIST_URL)
    img_url = to_map_image_url(fcst_url)

    # 2) Scarica l’immagine sempre (così puoi caricarla come artifact se vuoi)
    out_path = download_map_image(img_url, FILENAME_BASE)

    # 3) Invia email solo se gate orario ok (o se forzato)
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
