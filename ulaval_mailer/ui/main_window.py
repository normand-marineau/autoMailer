from __future__ import annotations

import csv
import queue
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk

from ulaval_mailer.core.data_io import (
    LoadedData,
    build_key_maps,
    detect_recipient_key,
    read_csv,
    read_xlsx,
    rows_to_dicts,
)
from ulaval_mailer.core.paths import ensure_logs_dir
from ulaval_mailer.core.text_utils import EMAIL_RE, now_stamp, placeholders_used, render_template

from ulaval_mailer.providers.outlook_provider import outlook_available, run_outlook_batch
from ulaval_mailer.providers.gmail_provider import gmail_available, ensure_gmail_service, run_gmail_batch


FEATURE_PHASE_C = False  # hide "Envoyer plus tard" until implemented

class MailerV2App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self.title("ULaval Mailer v2")
        self.geometry("1100x760")
        self.minsize(980, 680)

        # State
        self.loaded: Optional[LoadedData] = None
        self.recipient_key: tk.StringVar = tk.StringVar(value="")
        self.provider: tk.StringVar = tk.StringVar(value="outlook")  # outlook | gmail
        self.mode: tk.StringVar = tk.StringVar(value="draft")        # draft | send_now
        self.throttle_per_min: tk.IntVar = tk.IntVar(value=30)
        self.test_run_enabled: tk.BooleanVar = tk.BooleanVar(value=True)
        self.test_run_n: tk.IntVar = tk.IntVar(value=3)

        self.subject_tpl: tk.StringVar = tk.StringVar(value="")
        self.file_path_var: tk.StringVar = tk.StringVar(value="")

        # Run control
        self._stop_event = threading.Event()
        self._worker: Optional[threading.Thread] = None
        self._uiq: "queue.Queue[tuple]" = queue.Queue()

        self._build_style()
        self._build_ui()

        self.after(100, self._poll_ui_queue)

    def _build_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TNotebook.Tab", padding=(14, 8))
        style.configure("TLabelframe", padding=10)
        style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"))

    def _build_ui(self) -> None:
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        self.tab_guide = ttk.Frame(self.nb)
        self.tab_message = ttk.Frame(self.nb)
        self.tab_params = ttk.Frame(self.nb)

        self.nb.add(self.tab_guide, text="Guide")
        self.nb.add(self.tab_message, text="Message")
        self.nb.add(self.tab_params, text="Paramétrisation")

        self._build_tab_guide()
        self._build_tab_message()
        self._build_tab_params()

        self._refresh_mode_availability()

    # -------------------------
    # Guide tab
    # -------------------------
    def _build_tab_guide(self) -> None:
        frm = ttk.Frame(self.tab_guide, padding=14)
        frm.pack(fill="both", expand=True)

        title = ttk.Label(frm, text="ULaval Mailer v2 — Guide rapide", font=("Segoe UI", 16, "bold"))
        title.pack(anchor="w")

        body = tk.Text(frm, wrap="word", height=24)
        body.pack(fill="both", expand=True, pady=(12, 0))

        guide = (
            "Workflow (recommandé)\n"
            "1) Paramétrisation → Charger un fichier CSV/XLSX\n"
            "2) Message → Rédiger sujet + corps (Texte + HTML) avec variables {cle}\n"
            "3) Paramétrisation → Vérification (prévol)\n"
            "4) Dry-run (sans envoi) → vérifier logs\n"
            "5) Brouillons (Outlook ou Gmail) — recommandé\n"
            "6) Envoyer maintenant (Outlook ou Gmail) — confirmation requise\n\n"
            "Gmail\n"
            "- Nécessite OAuth (bouton Connexion Gmail) + secrets/credentials.json.\n"
            "- Crée soit des Brouillons Gmail, soit des envois immédiats.\n\n"
            "Règles\n"
            "- 1 destinataire par ligne.\n"
            "- Si un champ requis est vide/invalide: ligne ignorée + log.\n"
        )
        body.insert("1.0", guide)
        body.configure(state="disabled")

    # -------------------------
    # Message tab
    # -------------------------
    def _build_tab_message(self) -> None:
        root = ttk.Frame(self.tab_message, padding=14)
        root.pack(fill="both", expand=True)

        subj_row = ttk.Frame(root)
        subj_row.pack(fill="x")
        ttk.Label(subj_row, text="Sujet:").pack(side="left")
        self.ent_subject = ttk.Entry(subj_row, textvariable=self.subject_tpl)
        self.ent_subject.pack(side="left", fill="x", expand=True, padx=(10, 0))

        pan = ttk.Panedwindow(root, orient="horizontal")
        pan.pack(fill="both", expand=True, pady=(12, 0))

        left = ttk.Frame(pan, padding=(0, 0, 10, 0))
        right = ttk.Frame(pan)
        pan.add(left, weight=1)
        pan.add(right, weight=3)

        ttk.Label(left, text="Variables disponibles", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.var_tree = ttk.Treeview(left, columns=("key", "hdr"), show="headings", height=18)
        self.var_tree.heading("key", text="Clé")
        self.var_tree.heading("hdr", text="Colonne")
        self.var_tree.column("key", width=160, anchor="w")
        self.var_tree.column("hdr", width=260, anchor="w")
        self.var_tree.pack(fill="both", expand=True, pady=(6, 0))
        ttk.Label(left, text="Double-clic: insère {clé} dans le champ actif.").pack(anchor="w", pady=(6, 0))
        self.var_tree.bind("<Double-1>", self._on_var_double_click)

        editor_box = ttk.Labelframe(right, text="Corps du message")
        editor_box.pack(fill="both", expand=True)

        # Sub-tabs inside Message: Texte / HTML
        self.body_nb = ttk.Notebook(editor_box)
        self.body_nb.pack(fill="both", expand=True)

        frm_txt = ttk.Frame(self.body_nb)
        frm_html = ttk.Frame(self.body_nb)
        self.body_nb.add(frm_txt, text="Texte")
        self.body_nb.add(frm_html, text="HTML")

        self.txt_body_text = tk.Text(frm_txt, wrap="word", height=12)
        ysb_t = ttk.Scrollbar(frm_txt, orient="vertical", command=self.txt_body_text.yview)
        self.txt_body_text.configure(yscrollcommand=ysb_t.set)
        self.txt_body_text.pack(side="left", fill="both", expand=True)
        ysb_t.pack(side="right", fill="y")

        self.txt_body_html = tk.Text(frm_html, wrap="word", height=12)
        ysb_h = ttk.Scrollbar(frm_html, orient="vertical", command=self.txt_body_html.yview)
        self.txt_body_html.configure(yscrollcommand=ysb_h.set)
        self.txt_body_html.pack(side="left", fill="both", expand=True)
        ysb_h.pack(side="right", fill="y")

        preview_box = ttk.Labelframe(right, text="Aperçu (ligne)")
        preview_box.pack(fill="both", expand=True, pady=(10, 0))

        top = ttk.Frame(preview_box)
        top.pack(fill="x")
        ttk.Label(top, text="Ligne:").pack(side="left")
        self.spin_row = ttk.Spinbox(top, from_=1, to=1, width=8, command=self._update_preview)
        self.spin_row.pack(side="left", padx=(8, 0))
        ttk.Button(top, text="Mettre à jour l'aperçu", command=self._update_preview).pack(side="left", padx=(10, 0))

        self.txt_preview = tk.Text(preview_box, wrap="word", height=12, state="disabled")
        ysb2 = ttk.Scrollbar(preview_box, orient="vertical", command=self.txt_preview.yview)
        self.txt_preview.configure(yscrollcommand=ysb2.set)
        self.txt_preview.pack(side="left", fill="both", expand=True, pady=(8, 0))
        ysb2.pack(side="right", fill="y", pady=(8, 0))

    def _focused_text_widget(self) -> Optional[tk.Text]:
        w = self.focus_get()
        return w if isinstance(w, tk.Text) else None

    def _on_var_double_click(self, _evt=None) -> None:
        sel = self.var_tree.selection()
        if not sel:
            return
        item = self.var_tree.item(sel[0])
        key = item["values"][0]
        token = "{" + str(key) + "}"

        w = self.focus_get()
        if w == self.ent_subject:
            pos = self.ent_subject.index(tk.INSERT)
            cur = self.subject_tpl.get()
            self.subject_tpl.set(cur[:pos] + token + cur[pos:])
            self.ent_subject.icursor(pos + len(token))
        else:
            tw = self._focused_text_widget()
            if tw is None:
                tw = self.txt_body_text
            tw.insert(tk.INSERT, token)

        self._update_preview()

    def _get_templates(self) -> tuple[str, str, str]:
        subject = self.subject_tpl.get()
        text_body = self.txt_body_text.get("1.0", "end-1c")
        html_body = self.txt_body_html.get("1.0", "end-1c")
        return subject, text_body, html_body

    def _update_preview(self) -> None:
        if not self.loaded or self.loaded.nrows == 0:
            self._set_preview("Aucune donnée chargée.")
            return
        try:
            idx = int(self.spin_row.get()) - 1
        except Exception:
            idx = 0
        idx = max(0, min(idx, self.loaded.nrows - 1))

        row = self.loaded.rows[idx]
        subject, text_body, html_body = self._get_templates()

        subj = render_template(subject, row)
        txt = render_template(text_body, row)
        htm = render_template(html_body, row)

        rk = self.recipient_key.get().strip()
        recip = row.get(rk, "") if rk else ""

        txtp = (
            f"--- Ligne {idx+1} / {self.loaded.nrows} ---\n"
            f"Destinataire ({rk or 'non sélectionné'}): {recip}\n\n"
            f"SUJET:\n{subj}\n\n"
            f"TEXTE:\n{txt}\n\n"
            f"HTML:\n{htm if htm.strip() else '(vide → HTML auto-généré à partir du texte pour Gmail)'}\n"
        )
        self._set_preview(txtp)

    def _set_preview(self, text: str) -> None:
        self.txt_preview.configure(state="normal")
        self.txt_preview.delete("1.0", "end")
        self.txt_preview.insert("1.0", text)
        self.txt_preview.configure(state="disabled")

    # -------------------------
    # Parametrisation tab
    # -------------------------
    def _build_tab_params(self) -> None:
        root = ttk.Frame(self.tab_params, padding=14)
        root.pack(fill="both", expand=True)

        self.lbl_status = ttk.Label(root, text="Statut: Aucun fichier chargé.")
        self.lbl_status.pack(anchor="w")

        f_file = ttk.Labelframe(root, text="Fichier de données")
        f_file.pack(fill="x", pady=(10, 0))

        row = ttk.Frame(f_file)
        row.pack(fill="x")
        ttk.Entry(row, textvariable=self.file_path_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Parcourir…", command=self._pick_file).pack(side="left", padx=(10, 0))

        f_schema = ttk.Labelframe(root, text="Colonnes & destinataire")
        f_schema.pack(fill="x", pady=(10, 0))

        row2 = ttk.Frame(f_schema)
        row2.pack(fill="x")
        ttk.Label(row2, text="Colonne destinataire:").pack(side="left")
        self.opt_recipient = ttk.Combobox(row2, textvariable=self.recipient_key, state="readonly", values=[])
        self.opt_recipient.pack(side="left", padx=(10, 0), fill="x", expand=True)

        f_provider = ttk.Labelframe(root, text="Fournisseur")
        f_provider.pack(fill="x", pady=(10, 0))

        row3 = ttk.Frame(f_provider)
        row3.pack(fill="x")
        ttk.Radiobutton(row3, text="Outlook (desktop)", value="outlook", variable=self.provider, command=self._refresh_mode_availability).pack(side="left")
        ttk.Radiobutton(row3, text="Gmail (API)", value="gmail", variable=self.provider, command=self._refresh_mode_availability).pack(side="left", padx=(16, 0))

        self.btn_gmail_auth = ttk.Button(row3, text="Connexion Gmail (OAuth)…", command=self._gmail_authorize)
        self.btn_gmail_auth.pack(side="left", padx=(16, 0))

        f_mode = ttk.Labelframe(root, text="Mode")
        f_mode.pack(fill="x", pady=(10, 0))

        row4 = ttk.Frame(f_mode)
        row4.pack(fill="x")
        self.rb_draft = ttk.Radiobutton(row4, text="Brouillons", value="draft", variable=self.mode)
        self.rb_send = ttk.Radiobutton(row4, text="Envoyer maintenant", value="send_now", variable=self.mode)

        self.rb_draft.pack(side="left")
        self.rb_send.pack(side="left", padx=(16, 0))

        if FEATURE_PHASE_C:
            self.rb_later = ttk.Radiobutton(row4, text="Envoyer plus tard", value="send_later", variable=self.mode)
            self.rb_later.pack(side="left", padx=(16, 0))

        f_ctrl = ttk.Labelframe(root, text="Contrôles (v2)")
        f_ctrl.pack(fill="x", pady=(10, 0))

        row5 = ttk.Frame(f_ctrl)
        row5.pack(fill="x")

        ttk.Label(row5, text="Throttle (emails/min):").pack(side="left")
        ttk.Spinbox(row5, from_=1, to=600, width=7, textvariable=self.throttle_per_min).pack(side="left", padx=(8, 0))

        ttk.Checkbutton(row5, text="Test run (premières N lignes)", variable=self.test_run_enabled).pack(side="left", padx=(18, 0))
        ttk.Spinbox(row5, from_=1, to=500, width=6, textvariable=self.test_run_n).pack(side="left", padx=(8, 0))

        row6 = ttk.Frame(f_ctrl)
        row6.pack(fill="x", pady=(8, 0))

        ttk.Button(row6, text="Vérification (prévol)", command=self._preflight).pack(side="left")
        ttk.Button(row6, text="Dry-run (sans envoi) → logs", command=self._dry_run).pack(side="left", padx=(10, 0))

        self.btn_start = ttk.Button(row6, text="Démarrer", command=self._start_run)
        self.btn_start.pack(side="left", padx=(10, 0))

        self.btn_stop = ttk.Button(row6, text="Stop", command=self._request_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(10, 0))

        f_logs = ttk.Labelframe(root, text="Journal")
        f_logs.pack(fill="both", expand=True, pady=(10, 0))

        self.pbar = ttk.Progressbar(f_logs, mode="determinate")
        self.pbar.pack(fill="x")

        self.txt_log = tk.Text(f_logs, wrap="word", height=12, state="disabled")
        ysb = ttk.Scrollbar(f_logs, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=ysb.set)
        self.txt_log.pack(side="left", fill="both", expand=True, pady=(8, 0))
        ysb.pack(side="right", fill="y", pady=(8, 0))

    def log(self, line: str) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", line.rstrip() + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def set_status(self, text: str) -> None:
        self.lbl_status.configure(text=f"Statut: {text}")

    def _refresh_mode_availability(self) -> None:
        if self.provider.get() == "gmail":
            self.btn_start.configure(text="Démarrer (Gmail)")
            self.btn_gmail_auth.configure(state="normal")
        else:
            self.btn_start.configure(text="Démarrer (Outlook)")
            self.btn_gmail_auth.configure(state="disabled")

    def _gmail_authorize(self) -> None:
        if not gmail_available():
            messagebox.showerror(
                "Gmail",
                "Dépendances Gmail manquantes.\n"
                "Installe:\n"
                "  pip install google-api-python-client google-auth google-auth-oauthlib google-auth-httplib2"
            )
            return
        try:
            ensure_gmail_service(log=self.log)
            messagebox.showinfo("Gmail", "Connexion Gmail OK.\nToken enregistré dans secrets/token.json.")
        except Exception as e:
            messagebox.showerror("Gmail", str(e))

    def _pick_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Choisir un fichier CSV ou XLSX",
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx"), ("Tous", "*.*")]
        )
        if not path:
            return
        self.file_path_var.set(path)
        try:
            self._load_file(Path(path))
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
            self.log(f"[ERREUR] {e}")

    def _load_file(self, path: Path) -> None:
        ext = path.suffix.lower()
        if ext == ".csv":
            headers, data = read_csv(path)
        elif ext == ".xlsx":
            headers, data = read_xlsx(path)
        else:
            raise RuntimeError("Format non supporté. Utilise CSV ou XLSX.")

        if not headers:
            raise RuntimeError("Le fichier ne contient pas d'en-têtes (ligne 1).")

        key_to_header, header_to_key = build_key_maps(headers)
        rows = rows_to_dicts(headers, data, header_to_key)

        self.loaded = LoadedData(
            path=path,
            raw_headers=headers,
            key_to_header=key_to_header,
            header_to_key=header_to_key,
            rows=rows,
        )

        keys = list(self.loaded.key_to_header.keys())
        self.opt_recipient.configure(values=keys)

        suggested = detect_recipient_key(self.loaded.key_to_header, self.loaded.rows)
        if suggested:
            self.recipient_key.set(suggested)
        elif keys:
            self.recipient_key.set(keys[0])

        for it in self.var_tree.get_children():
            self.var_tree.delete(it)
        for k, h in self.loaded.key_to_header.items():
            self.var_tree.insert("", "end", values=(k, h))

        n = max(1, self.loaded.nrows)
        self.spin_row.configure(to=n)
        self.spin_row.delete(0, "end")
        self.spin_row.insert(0, "1")

        self.set_status(f"Fichier chargé: {path.name} — {self.loaded.nrows} lignes")
        self.log(f"[OK] Chargé: {path} ({self.loaded.nrows} lignes, {len(headers)} colonnes)")
        self._update_preview()

    def _preflight(self) -> bool:
        self.log("— Prévol: démarrage…")
        if not self.loaded:
            self.log("[BLOQUANT] Aucun fichier chargé.")
            messagebox.showwarning("Prévol", "Aucun fichier chargé.")
            return False

        rk = self.recipient_key.get().strip()
        if not rk:
            self.log("[BLOQUANT] Colonne destinataire non sélectionnée.")
            messagebox.showwarning("Prévol", "Colonne destinataire non sélectionnée.")
            return False

        if rk not in self.loaded.key_to_header:
            self.log("[BLOQUANT] Colonne destinataire invalide.")
            messagebox.showwarning("Prévol", "Colonne destinataire invalide.")
            return False

        subject, text_body, html_body = self._get_templates()

        used = set(placeholders_used(subject, text_body) + placeholders_used("", html_body))
        used = sorted(used)
        unknown = [k for k in used if k not in self.loaded.key_to_header]
        if unknown:
            self.log("[BLOQUANT] Variables inconnues dans le message: " + ", ".join(unknown))
            messagebox.showwarning("Prévol", "Variables inconnues:\n" + "\n".join(unknown))
            return False

        required = set(used + [rk])
        sample = self.loaded.rows[: min(200, self.loaded.nrows)]
        invalid_recips = 0
        missing_required = 0
        for row in sample:
            recip = (row.get(rk, "") or "").strip()
            if not recip or not EMAIL_RE.match(recip):
                invalid_recips += 1
                continue
            for req in required:
                if (row.get(req, "") or "").strip() == "":
                    missing_required += 1
                    break

        self.log(f"[OK] Prévol: variables utilisées: {', '.join(used) if used else '(aucune)'}")
        self.log(f"[INFO] Échantillon (max 200): destinataires invalides: {invalid_recips} | champs requis manquants: {missing_required}")
        self.log("— Prévol: terminé.")
        messagebox.showinfo("Prévol", "Prévol OK (voir le journal).")
        return True

    def _compute_run_rows(self) -> List[Tuple[int, Dict[str, str]]]:
        assert self.loaded is not None
        rows = self.loaded.rows
        if self.test_run_enabled.get():
            n = max(1, int(self.test_run_n.get()))
            rows = rows[: min(n, len(rows))]
        return list(enumerate(rows, start=1))

    def _dry_run(self) -> None:
        if not self._preflight():
            return
        assert self.loaded is not None

        logs_dir = ensure_logs_dir()
        stamp = now_stamp()
        send_log_path = logs_dir / f"send_log_{stamp}.csv"
        skip_path = logs_dir / f"skipped_rows_{stamp}.csv"

        rk = self.recipient_key.get().strip()
        subject, text_body, html_body = self._get_templates()

        used = set(placeholders_used(subject, text_body) + placeholders_used("", html_body))
        used = sorted(used)
        required = set(used + [rk])

        run_rows = self._compute_run_rows()
        total = len(run_rows)
        self.pbar.configure(maximum=max(1, total), value=0)
        self.log(f"— Dry-run: {total} lignes (sur {self.loaded.nrows}) → écriture logs…")

        with send_log_path.open("w", encoding="utf-8", newline="") as fsend, skip_path.open("w", encoding="utf-8", newline="") as fskip:
            send_w = csv.writer(fsend)
            skip_w = csv.writer(fskip)

            send_w.writerow(["timestamp", "row_index", "recipient", "provider", "mode", "status", "reason"])
            skip_w.writerow(["row_index", "recipient", "reason"])

            skipped = 0
            prepared = 0

            for n, (i, row) in enumerate(run_rows, start=1):
                recip = (row.get(rk, "") or "").strip()
                reason = ""

                if not recip:
                    reason = "Message non envoyé: champ destinataire vide."
                elif not EMAIL_RE.match(recip):
                    reason = "Message non envoyé: destinataire invalide (format courriel)."
                else:
                    for req in required:
                        if (row.get(req, "") or "").strip() == "":
                            reason = f"Message non envoyé: champ requis vide ({req})."
                            break

                if reason:
                    skipped += 1
                    send_w.writerow([datetime.now().isoformat(timespec="seconds"), i, recip, self.provider.get(), self.mode.get(), "skipped", reason])
                    skip_w.writerow([i, recip, reason])
                    self.log(f"[SKIP] Ligne {i}: {recip} — {reason}")
                else:
                    _ = render_template(subject, row)
                    _ = render_template(text_body, row)
                    _ = render_template(html_body, row)
                    prepared += 1
                    send_w.writerow([datetime.now().isoformat(timespec="seconds"), i, recip, self.provider.get(), self.mode.get(), "prepared", "Dry-run: rendu OK"])
                    self.log(f"[OK]   Ligne {i}: {recip} — prêt (dry-run)")

                self.pbar.configure(value=n)
                self.update_idletasks()

                per_min = max(1, int(self.throttle_per_min.get()))
                delay = 60.0 / per_min
                time.sleep(min(0.15, delay))

        self.log(f"— Dry-run terminé. prepared={prepared}, skipped={skipped}")
        self.log(f"[LOG] {send_log_path}")
        self.log(f"[LOG] {skip_path}")
        messagebox.showinfo("Dry-run", f"Terminé.\n\nLogs:\n{send_log_path}\n{skip_path}")

    # -------------------------
    # Run (Outlook or Gmail)
    # -------------------------
    def _request_stop(self) -> None:
        if self._worker and self._worker.is_alive():
            self._stop_event.set()
            self.log("[STOP] Demande d'arrêt envoyée. (arrêt entre deux lignes)")
            self.set_status("Arrêt demandé…")

    def _start_run(self) -> None:
        if not self._preflight():
            return

        if self._worker and self._worker.is_alive():
            messagebox.showwarning("En cours", "Un traitement est déjà en cours.")
            return

        mode = self.mode.get()
        if mode not in ("draft", "send_now"):
            messagebox.showinfo("À venir", "Envoyer plus tard (calendar + queue) arrive en phase C.\nPour l'instant: Brouillons ou Envoyer maintenant.")
            return

        run_rows = self._compute_run_rows()
        total = len(run_rows)
        if total == 0:
            messagebox.showwarning("Aucune ligne", "Le fichier ne contient aucune ligne.")
            return

        # Confirmation gate for real send
        if mode == "send_now":
            token_needed = f"ENVOYER {total}"
            typed = simpledialog.askstring(
                "Confirmation d'envoi",
                f"Tu es sur le point d'envoyer {total} message(s).\n\n"
                f"Pour confirmer, tape exactement:\n{token_needed}\n\n"
                "Sinon, annule.",
                parent=self
            )
            if typed != token_needed:
                self.log("[ANNULÉ] Envoi annulé (confirmation incorrecte).")
                return
        else:
            ok = messagebox.askokcancel(
                "Créer des brouillons",
                f"Créer {total} brouillon(s) ?",
                parent=self
            )
            if not ok:
                self.log("[ANNULÉ] Création de brouillons annulée.")
                return

        # Provider readiness checks
        if self.provider.get() == "outlook":
            if not outlook_available():
                messagebox.showerror("Dépendance manquante", "pywin32 n'est pas disponible.\nInstalle: pip install pywin32")
                return
        else:
            if not gmail_available():
                messagebox.showerror(
                    "Dépendance manquante",
                    "Dépendances Gmail manquantes.\nInstalle:\n  pip install google-api-python-client google-auth google-auth-oauthlib google-auth-httplib2"
                )
                return
            # Ensure OAuth token is present (will open browser if needed)
            try:
                ensure_gmail_service(log=self.log)
            except Exception as e:
                messagebox.showerror("Gmail", str(e))
                return

        # Start background worker
        self._stop_event.clear()
        self.btn_stop.configure(state="normal")
        self.btn_start.configure(state="disabled")
        self.set_status(f"Traitement {self.provider.get()} en cours…")

        self._worker = threading.Thread(
            target=self._worker_run,
            args=(self.provider.get(), mode, run_rows),
            daemon=True
        )
        self._worker.start()

    def _poll_ui_queue(self) -> None:
        try:
            while True:
                kind, payload = self._uiq.get_nowait()
                if kind == "log":
                    self.log(payload)
                elif kind == "status":
                    self.set_status(payload)
                elif kind == "progress":
                    cur, total = payload
                    self.pbar.configure(maximum=max(1, total), value=cur)
                    self.update_idletasks()
                elif kind == "done":
                    self.btn_start.configure(state="normal")
                    self.btn_stop.configure(state="disabled")
                    self.set_status(payload)
                else:
                    pass
        except queue.Empty:
            pass
        self.after(120, self._poll_ui_queue)

    def _worker_run(self, provider: str, mode: str, run_rows: List[Tuple[int, Dict[str, str]]]) -> None:
        assert self.loaded is not None

        rk = self.recipient_key.get().strip()
        subject_tpl, text_tpl, html_tpl = self._get_templates()

        used = set(placeholders_used(subject_tpl, text_tpl) + placeholders_used("", html_tpl))
        used = sorted(used)
        required = set(used + [rk])

        total = len(run_rows)
        self._uiq.put(("progress", (0, total)))

        logs_dir = ensure_logs_dir()

        def _log(msg: str) -> None:
            self._uiq.put(("log", msg))

        def _progress(cur: int, tot: int) -> None:
            self._uiq.put(("progress", (cur, tot)))

        try:
            if provider == "outlook":
                summary = run_outlook_batch(
                    mode=mode,
                    run_rows=run_rows,
                    recipient_key=rk,
                    subject_tpl=subject_tpl,
                    body_tpl=text_tpl,  # Outlook uses plain text body
                    required_keys=required,
                    throttle_per_min=int(self.throttle_per_min.get()),
                    logs_dir=logs_dir,
                    stop_event=self._stop_event,
                    log=_log,
                    progress=_progress,
                )
            else:
                summary = run_gmail_batch(
                    mode=mode,
                    run_rows=run_rows,
                    recipient_key=rk,
                    subject_tpl=subject_tpl,
                    text_tpl=text_tpl,
                    html_tpl=html_tpl,
                    required_keys=required,
                    throttle_per_min=int(self.throttle_per_min.get()),
                    logs_dir=logs_dir,
                    stop_event=self._stop_event,
                    log=_log,
                    progress=_progress,
                )

            self._uiq.put(("done", summary))
        except Exception as e:
            self._uiq.put(("log", f"[FATAL] {provider}: {e}"))
            self._uiq.put(("done", f"Erreur {provider} (voir journal)"))


def main() -> int:
    app = MailerV2App()
    app.mainloop()
    return 0
