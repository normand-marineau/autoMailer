from __future__ import annotations

import base64
import csv
import html as _html
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from turtle import mode
from typing import Callable, Dict, List, Set, Tuple

from ulaval_mailer.core.text_utils import EMAIL_RE, now_stamp, render_template

# Google libs (installed via pip)
try:
    from google.oauth2.credentials import Credentials  # type: ignore
    from google_auth_oauthlib.flow import InstalledAppFlow  # type: ignore
    from google.auth.transport.requests import Request  # type: ignore
    from googleapiclient.discovery import build  # type: ignore
except Exception:
    Credentials = None
    InstalledAppFlow = None
    Request = None
    build = None

LogFn = Callable[[str], None]
ProgressFn = Callable[[int, int], None]

# Send-only scope — draft mode removed (not required for Google OAuth verification)
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
]

def gmail_available() -> bool:
    return all(x is not None for x in (Credentials, InstalledAppFlow, Request, build))

def _secrets_dir(root: Path | None = None) -> Path:
    base = root or Path.cwd()
    d = base / "secrets"
    d.mkdir(parents=True, exist_ok=True)
    return d

def credentials_path(root: Path | None = None) -> Path:
    return _secrets_dir(root) / "credentials.json"

def token_path(root: Path | None = None) -> Path:
    return _secrets_dir(root) / "token.json"

def ensure_gmail_service(*, log: LogFn, root: Path | None = None):
    """
    Returns an authenticated Gmail API service.
    - First time: opens browser OAuth consent, stores token.json.
    - Later: silent refresh if expired.
    """
    if not gmail_available():
        raise RuntimeError(
            "Dépendances Gmail manquantes. Installe:\n"
            "  pip install google-api-python-client google-auth google-auth-oauthlib google-auth-httplib2"
        )

    cred_p = credentials_path(root)
    tok_p = token_path(root)

    if not cred_p.exists():
        raise RuntimeError(
            "Fichier OAuth manquant:\n"
            f"  {cred_p}\n\n"
            "Télécharge un 'OAuth client ID' de type Desktop depuis Google Cloud Console, "
            "puis renomme-le en credentials.json et place-le dans secrets/."
        )

    creds = None
    if tok_p.exists():
        creds = Credentials.from_authorized_user_file(str(tok_p), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(cred_p), SCOPES)
            creds = flow.run_local_server(port=0)
        tok_p.write_text(creds.to_json(), encoding="utf-8")

    svc = build("gmail", "v1", credentials=creds)
    # lightweight sanity call (no heavy listing)
    _ = svc.users().getProfile(userId="me").execute()
    log("[OK] Gmail OAuth prêt (token.json présent).")
    return svc

def _build_raw_email(to_addr: str, subject: str, text_body: str, html_body: str) -> str:
    """
    Build multipart/alternative (plain + html) and return base64url raw string.
    """
    msg = MIMEMultipart("alternative")
    msg["To"] = to_addr
    msg["Subject"] = subject

    text_body = text_body or ""
    html_body = html_body or ""

    # If HTML empty, auto-derive a safe HTML version from text
    if html_body.strip() == "":
        safe = _html.escape(text_body).replace("\n", "<br>\n")
        html_body = f"<html><body>{safe}</body></html>"

    msg.attach(MIMEText(text_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    raw_bytes = msg.as_bytes()
    return base64.urlsafe_b64encode(raw_bytes).decode("utf-8")

def run_gmail_batch(
    *,
    mode: str,  # "send_now"
    run_rows: List[Tuple[int, Dict[str, str]]],
    recipient_key: str,
    subject_tpl: str,
    text_tpl: str,
    html_tpl: str,
    required_keys: Set[str],
    throttle_per_min: int,
    logs_dir: Path,
    stop_event,
    log: LogFn,
    progress: ProgressFn,
) -> str:
    """
    Executes Gmail Drafts / Send Now for a batch.
    Writes logs/send_log_*.csv and logs/skipped_rows_*.csv under logs_dir.
    """
    svc = ensure_gmail_service(log=log)

    if mode not in ("send_now", "send"):
        raise NotImplementedError(f"Mode non supporté pour Gmail: {mode!r}. Seul 'send_now' est disponible.")
    if mode == "send":
        mode = "send_now"

    stamp = now_stamp()
    send_log_path = logs_dir / f"send_log_{stamp}.csv"
    skip_path = logs_dir / f"skipped_rows_{stamp}.csv"

    total = len(run_rows)
    log(f"— Gmail {mode}: démarrage… ({total} lignes)")
    log(f"[LOG] {send_log_path}")
    log(f"[LOG] {skip_path}")

    per_min = max(1, int(throttle_per_min))
    delay = 60.0 / per_min

    with send_log_path.open("w", encoding="utf-8", newline="") as fsend, skip_path.open("w", encoding="utf-8", newline="") as fskip:
        send_w = csv.writer(fsend)
        skip_w = csv.writer(fskip)

        # Keep same schema as Outlook logs (plus we place ids inside the reason column)
        send_w.writerow(["timestamp", "row_index", "recipient", "provider", "mode", "status", "reason"])
        skip_w.writerow(["row_index", "recipient", "reason"])

        sent = 0
        skipped = 0
        errored = 0

        for n, (i, row) in enumerate(run_rows, start=1):
            if stop_event.is_set():
                log(f"[STOP] Arrêt demandé. Fin au prochain point sûr (ligne {i}).")
                break

            recip = (row.get(recipient_key, "") or "").strip()
            reason = ""

            if not recip:
                reason = "Message non envoyé: champ destinataire vide."
            elif not EMAIL_RE.match(recip):
                reason = "Message non envoyé: destinataire invalide (format courriel)."
            else:
                for req in required_keys:
                    if (row.get(req, "") or "").strip() == "":
                        reason = f"Message non envoyé: champ requis vide ({req})."
                        break

            if reason:
                skipped += 1
                send_w.writerow([datetime.now().isoformat(timespec="seconds"), i, recip, "gmail", mode, "skipped", reason])
                skip_w.writerow([i, recip, reason])
                log(f"[SKIP] Ligne {i}: {recip} — {reason}")
                progress(n, total)
                continue

            subj = render_template(subject_tpl, row)
            txt = render_template(text_tpl, row)
            htm = render_template(html_tpl, row)

            try:
                raw = _build_raw_email(recip, subj, txt, htm)

                res = svc.users().messages().send(
                    userId="me",
                    body={"raw": raw},
                ).execute()
                sent += 1
                mid = res.get("id", "")
                send_w.writerow([datetime.now().isoformat(timespec="seconds"), i, recip, "gmail", mode, "sent", f"OK (msg_id={mid})"])
                log(f"[SENT]  Ligne {i}: {recip} (msg_id={mid})")

            except Exception as e:
                errored += 1
                err = str(e)
                send_w.writerow([datetime.now().isoformat(timespec="seconds"), i, recip, "gmail", mode, "error", err])
                log(f"[ERROR] Ligne {i}: {recip} — {err}")

            progress(n, total)
            time.sleep(min(0.25, delay))

    summary = f"Gmail terminé. sent={sent}, skipped={skipped}, error={errored}"
    if stop_event.is_set():
        summary += " (arrêt demandé)"
    log("— " + summary)
    return summary
