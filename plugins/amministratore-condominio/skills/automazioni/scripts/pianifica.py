#!/usr/bin/env python3
"""
pianifica.py — prepara il riepilogo settimanale su Linux: un servizio e un timer di systemd per l'utente.

  pianifica.py --dir "<cartella del condominio>" [--giorno lun] [--ora 08:00] [--modello haiku] [--sostituisci]
  pianifica.py --dir "<cartella del condominio>" --rimuovi

Scrive tre file: le istruzioni (in ~/.config/amministratore-condominio/) e il servizio e il timer
(in ~/.config/systemd/user/). Il servizio lancia `claude -p` nella cartella del condominio, con le
istruzioni come input e due soli permessi (lo script di stato e l'invio Gmail). NON attiva nulla:
stampa i comandi `systemctl --user …` che l'utente esegue. Con --rimuovi stampa i comandi per
disattivare e cancellare.

Nome, destinatario e dati vengono dal registro: `Nome` ed `Email amministratore` del foglio Condominio.
Il PATH del servizio è quello della shell da cui si lancia lo script, così `claude` e `python3`
(con openpyxl) sono gli stessi del terminale.
"""
import argparse
import json
import os
import re
import shlex
import shutil
import sys
import unicodedata

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "condominio-base", "scripts"))
try:
    import openpyxl
    import registro as reg
except ImportError as e:  # pragma: no cover
    if e.name == "openpyxl":
        sys.exit("openpyxl non installato. Installarlo con: python3 -m pip install openpyxl (Linux con errore "
                 "'externally-managed-environment': sudo apt install python3-openpyxl)")
    sys.exit(f"registro.py non trovato ({e}): il plugin è installato in modo incompleto")

ISTRUZIONI = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "references", "istruzioni-riepilogo.md")
GIORNI = {"lun": "Mon", "mar": "Tue", "mer": "Wed", "gio": "Thu", "ven": "Fri", "sab": "Sat", "dom": "Sun"}
# I soli permessi dell'automazione: caricare la skill di stato, eseguire il suo script, inviare con Gmail.
PERMESSI = ["Skill(amministratore-condominio:stato)", "Bash(python3 *stato.py*)", "mcp__claude_ai_Gmail__send_message"]


def slug(testo):
    t = unicodedata.normalize("NFKD", str(testo)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "-", t).strip("-") or "condominio"


def path_servizio(path):
    """PATH della shell senza le cartelle che Claude Code aggiunge per i plugin della sessione
    (~/.claude/plugins/…/<versione>/bin): cambiano a ogni aggiornamento. Senza doppioni."""
    voci = []
    for voce in path.split(os.pathsep):
        if voce and "/.claude/plugins/" not in voce and voce not in voci:
            voci.append(voce)
    return os.pathsep.join(voci)


def systemd_escape(valore):
    """Nei file di unità '%' introduce uno specificatore: va raddoppiato."""
    return str(valore).replace("%", "%%")


def arg(valore):
    """Argomento di ExecStart tra virgolette doppie (systemd le interpreta come una shell semplice)."""
    return '"' + systemd_escape(valore).replace("\\", "\\\\").replace('"', '\\"') + '"'


def contenuti(a, nome, email, claude, path_env, istruzioni_path):
    cartella = os.path.abspath(a.dir)
    servizio = "\n".join([
        "[Unit]",
        f"Description=Riepilogo settimanale del {systemd_escape(nome)} (amministratore-condominio)",
        "",
        "[Service]",
        "Type=oneshot",
        f"WorkingDirectory={systemd_escape(cartella)}",
        f"Environment={arg('PATH=' + path_env)}",
        f"StandardInput=file:{systemd_escape(istruzioni_path)}",
        "ExecStart=" + " ".join([arg(claude), "-p", "--model", arg(a.modello)]
                                + (["--plugin-dir", arg(os.path.abspath(a.plugin_dir))] if a.plugin_dir else [])
                                + ["--allowedTools"] + [arg(p) for p in PERMESSI]),
        "",
    ])
    timer = "\n".join([
        "[Unit]",
        f"Description=Ogni {a.giorno} alle {a.ora}: riepilogo del {systemd_escape(nome)}",
        "",
        "[Timer]",
        f"OnCalendar={GIORNI[a.giorno]} *-*-* {a.ora}:00",
        "Persistent=true",  # se il computer era spento, recupera una sola esecuzione all'avvio
        "",
        "[Install]",
        "WantedBy=timers.target",
        "",
    ])
    with open(ISTRUZIONI, encoding="utf-8") as f:
        istruzioni = f.read().replace("{nome}", str(nome)).replace("{email}", str(email))
    return servizio, timer, istruzioni


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--dir", required=True, help="cartella del condominio")
    p.add_argument("--giorno", choices=list(GIORNI), default="lun")
    p.add_argument("--ora", default="08:00")
    p.add_argument("--modello", default="haiku")
    p.add_argument("--sostituisci", action="store_true", help="sovrascrive un'automazione esistente diversa")
    p.add_argument("--rimuovi", action="store_true", help="stampa i comandi per disattivare e cancellare")
    p.add_argument("--plugin-dir", help="solo per provare una copia di sviluppo del plugin al posto di quella installata")
    p.add_argument("--unita-dir", default=os.path.expanduser("~/.config/systemd/user"))
    p.add_argument("--config-dir", default=os.path.expanduser("~/.config/amministratore-condominio"))
    a = p.parse_args()

    if not re.fullmatch(r"([01]\d|2[0-3]):[0-5]\d", a.ora):
        sys.exit(f"Ora non valida: {a.ora} (formato HH:MM, per esempio 08:00)")
    path_c = os.path.join(a.dir, reg.DEFAULT_FILE)
    if not os.path.exists(path_c):
        sys.exit("registro-condominio.xlsx non trovato: --dir deve essere la cartella del condominio")
    kv = reg._kv(openpyxl.load_workbook(path_c)["Condominio"])
    nome, email = kv.get("Nome"), kv.get("Email amministratore")
    nome_unita = f"riepilogo-{slug(nome or 'condominio')}"
    servizio_path = os.path.join(a.unita_dir, f"{nome_unita}.service")
    timer_path = os.path.join(a.unita_dir, f"{nome_unita}.timer")
    istruzioni_path = os.path.join(a.config_dir, f"{nome_unita}.md")

    if a.rimuovi:
        print(json.dumps({"ok": True, "nome": nome_unita, "comandi": [
            f"systemctl --user disable --now {nome_unita}.timer",
            "rm " + " ".join(shlex.quote(x) for x in (servizio_path, timer_path, istruzioni_path)),
            "systemctl --user daemon-reload",
        ]}, ensure_ascii=False, indent=2))
        return

    errori = []
    if reg._blank(email):
        errori.append("'Email amministratore' vuota nel foglio Condominio: è il destinatario del riepilogo")
    claude = shutil.which("claude")
    if not claude:
        errori.append("comando 'claude' non trovato nel PATH: il servizio non potrebbe partire")
    if errori:
        print(json.dumps({"ok": False, "errori": errori}, ensure_ascii=False, indent=2))
        sys.exit(2)

    servizio, timer, istruzioni = contenuti(a, nome or "Condominio", email, claude,
                                            path_servizio(os.environ.get("PATH", "")), istruzioni_path)
    nuovi = {servizio_path: servizio, timer_path: timer, istruzioni_path: istruzioni}
    diversi = []
    for path, testo in nuovi.items():
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                if f.read() != testo:
                    diversi.append(path)
    if diversi and not a.sostituisci:
        print(json.dumps({"ok": False, "errori": [
            "Esiste già un'automazione diversa per questo condominio: " + ", ".join(diversi)
            + ". Rilanciare con --sostituisci per aggiornarla."]}, ensure_ascii=False, indent=2))
        sys.exit(2)
    for path, testo in nuovi.items():
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(testo)

    print(json.dumps({
        "ok": True, "nome": nome_unita, "destinatario": email, "quando": f"ogni {a.giorno} alle {a.ora}",
        "file": list(nuovi), "aggiornati": diversi,
        "comandi": [
            "systemctl --user daemon-reload",
            f"systemctl --user enable --now {nome_unita}.timer",
        ],
        "prova_subito": f"systemctl --user start {nome_unita}.service",
        "registro_esecuzioni": f"journalctl --user -u {nome_unita}.service",
        "senza_sessione_aperta": "loginctl enable-linger $USER  (serve perché giri anche senza aver fatto l'accesso)",
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
