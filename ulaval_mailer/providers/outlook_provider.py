from __future__ import annotations

import csv
import time
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Set, Tuple

from ulaval_mailer.core.text_utils import EMAIL_RE, now_stamp, render_template

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except Exception:
    win32com = None
    pythoncom = None

LogFn = Callable[[str], None]
ProgressFn = Callable[[int, int], None]

def outlook_available() -> bool:
    return win32com is not None and pythoncom is not None

def _warmup_outlook(outlook, stop_event, log: LogFn) -> None:
    # ---- Outlook cold-start guardrail ----
    ns = outlook.GetNamespace("MAPI")
    try:
        ns.Logon("", "", False, False)
    except Exception:
        pass

    ready = False
    for _ in range(30):  # ~9 seconds max
        if stop_event.is_set():
            break
        try:
            _ = ns.Folders.Count  # triggers store init
            _warm = outlook.CreateItem(0)
            _ = _warm.Subject
            ready = True
            break
        except Exception:
            time.sleep(0.3)

    if not ready:
        log("[WARN] Outlook took too long to become ready; continuing anyway.")

def run_outlook_batch(
    *,
    mode: str,  # "draft" | "send_now"
    run_rows: List[Tuple[int, Dict[str, str]]],
    recipient_key: str,
    subject_tpl: str,
    body_tpl: str,
    required_keys: Set[str],
    throttle_per_min: int,
    logs_dir: Path,
    stop_event,
    log: LogFn,
    progress: ProgressFn,
) -> str:
    """
    Executes Outlook Draft / Send Now for a batch of rows.
    Writes logs/send_log_*.csv and logs/skipped_rows_*.csv under logs_dir.
    Returns a summary string.
    """
    if not outlook_available():
        raise RuntimeError("pywin32 (win32com/pythoncom) is not available.")

    if mode not in ("draft", "send_now", "send"):
        raise NotImplementedError(f"Mode non supporté pour Outlook: {mode!r} (Phase C désactivée)")
    if mode == "send":
        mode = "send_now"

    stamp = now_stamp()
    send_log_path = logs_dir / f"send_log_{stamp}.csv"
    skip_path = logs_dir / f"skipped_rows_{stamp}.csv"

    total = len(run_rows)
    log(f"— Outlook {mode}: démarrage… ({total} lignes)")
    log(f"[LOG] {send_log_path}")
    log(f"[LOG] {skip_path}")

    per_min = max(1, int(throttle_per_min))
    delay = 60.0 / per_min

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        _warmup_outlook(outlook, stop_event, log)

        with send_log_path.open("w", encoding="utf-8", newline="") as fsend, skip_path.open("w", encoding="utf-8", newline="") as fskip:
            send_w = csv.writer(fsend)
            skip_w = csv.writer(fskip)

            send_w.writerow(["timestamp", "row_index", "recipient", "provider", "mode", "status", "reason"])
            skip_w.writerow(["row_index", "recipient", "reason"])

            sent = 0
            drafted = 0
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
                    send_w.writerow([datetime.now().isoformat(timespec="seconds"), i, recip, "outlook", mode, "skipped", reason])
                    skip_w.writerow([i, recip, reason])
                    log(f"[SKIP] Ligne {i}: {recip} — {reason}")
                    progress(n, total)
                    continue

                subj = render_template(subject_tpl, row)
                body = render_template(body_tpl, row)

                try:
                    mail = outlook.CreateItem(0)  # 0 = olMailItem
                    mail.To = recip
                    mail.Subject = subj
                    mail.Body = body

                    if mode == "draft":
                        mail.Save()
                        drafted += 1
                        status = "drafted"
                        log(f"[DRAFT] Ligne {i}: {recip}")
                    else:
                        mail.Send()
                        sent += 1
                        status = "sent"
                        log(f"[SENT]  Ligne {i}: {recip}")

                    send_w.writerow([datetime.now().isoformat(timespec="seconds"), i, recip, "outlook", mode, status, "OK"])

                except Exception as e:
                    errored += 1
                    err = str(e)
                    send_w.writerow([datetime.now().isoformat(timespec="seconds"), i, recip, "outlook", mode, "error", err])
                    log(f"[ERROR] Ligne {i}: {recip} — {err}")

                progress(n, total)
                time.sleep(min(0.25, delay))

        summary = f"Outlook terminé. drafted={drafted}, sent={sent}, skipped={skipped}, error={errored}"
        if stop_event.is_set():
            summary += " (arrêt demandé)"
        log("— " + summary)
        return summary

    finally:
        pythoncom.CoUninitialize()
