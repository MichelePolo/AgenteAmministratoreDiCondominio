#!/usr/bin/env python3
"""
situazione_personale.py — compone l'email "situazione personale" di ogni unità (una per condomino).

  situazione_personale.py --dir "<cartella>" [--unita U01,U03] [--esercizio 2026] [--oggi 2026-09-26] [--salva]

Non invia nulla. Restituisce un JSON con, per ogni unità che ha una riga in Situazione per
l'esercizio: destinatario (Email dell'intestatario, mai il conduttore), oggetto, corpo del messaggio
e i numeri per la tabella di conferma (dovuto, versato, saldo, scaduto). Le unità senza email sono
elencate a parte. I numeri vengono da `registro.py situazione`: il testo non va ricalcolato né riscritto.

--salva scrive una copia di ogni testo in <cartella riservata>/comunicazioni/ (mai nell'archivio
condiviso): da usare dopo l'invio, o per le unità senza email da consegnare a mano.
"""
import argparse
import datetime as dt
import json
import os
import sys
from decimal import Decimal

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "condominio-base", "scripts"))
try:
    import openpyxl
    import registro as reg
except ImportError as e:  # pragma: no cover
    if e.name == "openpyxl":
        sys.exit("openpyxl non installato. Installarlo con: python3 -m pip install openpyxl (Windows: py -m pip install openpyxl; Linux con errore 'externally-managed-environment': sudo apt install python3-openpyxl)")
    sys.exit(f"registro.py non trovato ({e}): il plugin è installato in modo incompleto")

STATI = {"pagata": "pagata", "da pagare": "da pagare", "parziale": "pagata in parte", "scaduta": "scaduta"}


def riga_rata(r):
    imp, pagato = Decimal(str(r["Importo"])), Decimal(str(r["Pagato"]))
    quando = reg.data_it(r["Scadenza"]) if r["Scadenza"] else "quote precedenti"
    nome = f"Rata {r['Rata']}" if r["Rata"] else "Dovuto precedente"
    stato = STATI.get(r["Stato"], r["Stato"])
    if r["Stato"] in ("parziale", "scaduta") and pagato > 0:
        stato += f" (versati {reg.euro_it(pagato)})"
    # niente colonne allineate con spazi: nei client di posta il carattere è proporzionale
    return f"- {nome} · {quando} · {reg.euro_it(imp)} · {stato}"


def componi(s, anag, kv, oggi):
    nome_cond = kv.get("Nome") or "Condominio"
    interno = anag.get("Interno") or s["ID unità"]
    saldo = Decimal(str(s["Saldo"]))
    scaduto = Decimal(str(s["Scaduto"])) if s["Scaduto"] is not None else None
    righe = [
        f"Gentile {s['Intestatario'] or anag.get('Intestatario') or ''},".replace(" ,", ","),
        "",
        f"le invio il riepilogo delle quote condominiali {s['Esercizio']} per l'unità interno {interno},",
        f"aggiornato al {reg.data_it(oggi)}.",
        "",
        f"Dovuto per l'esercizio: {reg.euro_it(s['Dovuto'])}",
        f"Versato finora: {reg.euro_it(s['Versato'])}",
    ]
    if saldo > 0:
        righe.append(f"Resta da versare: {reg.euro_it(saldo)}")
    elif saldo < 0:
        righe.append(f"Credito a Suo favore: {reg.euro_it(-saldo)}")
    else:
        righe.append("Le quote dell'esercizio risultano interamente versate.")
    if s["rate"]:
        righe += ["", "Rate:"] + [riga_rata(r) for r in s["rate"]]
    if scaduto:
        righe += ["", f"Di quanto resta da versare, {reg.euro_it(scaduto)} riguardano rate con scadenza già passata."]
    if s["versamenti"]:
        righe += ["", "Versamenti registrati:"]
        righe += [f"- {reg.data_it(v['Data'])} · {reg.euro_it(v['Importo'])}" for v in s["versamenti"]]
    if saldo > 0 and not reg._blank(kv.get("Conto corrente")):
        righe += ["", "Coordinate per il versamento:",
                  f"IBAN: {kv['Conto corrente']}",
                  f"intestato a: {nome_cond}",
                  f"causale: quote {s['Esercizio']} interno {interno}"]
    righe += ["", "Se nota un errore, o un versamento che non compare, mi risponda a questa email:",
              "lo verifichiamo insieme.", "", "Cordiali saluti,", str(kv.get("Amministratore") or "L'amministratore"),
              f"amministratore del {nome_cond}"]
    return {
        "unita": s["ID unità"], "intestatario": s["Intestatario"] or anag.get("Intestatario"),
        "a": str(anag.get("Email") or "").strip() or None,
        "oggetto": f"{nome_cond} — situazione quote {s['Esercizio']} — interno {interno}",
        "corpo": "\n".join(righe),
        "dovuto": s["Dovuto"], "versato": s["Versato"], "saldo": s["Saldo"], "scaduto": s["Scaduto"],
    }


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--dir", default=".")
    p.add_argument("--unita", help="ID unità separati da virgola (default: tutte)")
    p.add_argument("--esercizio", type=int)
    p.add_argument("--oggi", type=dt.date.fromisoformat, default=dt.date.today())
    p.add_argument("--salva", action="store_true", help="salva una copia dei testi nella cartella riservata")
    a = p.parse_args()

    path_c = os.path.join(a.dir, reg.DEFAULT_FILE)
    if not os.path.exists(path_c):
        sys.exit("registro-condominio.xlsx non trovato: eseguire condominio-setup")
    wb_c = openpyxl.load_workbook(path_c)
    path_r = reg.percorso_riservato(a.dir, wb_c)
    if not os.path.exists(path_r):
        sys.exit(f"Registro riservato non trovato: {path_r} (controllare 'Percorso registro riservato')")
    kv = reg._kv(wb_c["Condominio"])
    esercizio = a.esercizio or int(kv.get("Esercizio corrente") or a.oggi.year)
    anag = {str(r["ID unità"]).strip(): r for r in reg._rows(wb_c["Anagrafica"]) if not reg._blank(r.get("ID unità"))}
    richieste = [u.strip() for u in a.unita.split(",")] if a.unita else None

    sit = {s["ID unità"]: s for s in reg.situazione(openpyxl.load_workbook(path_r), a.oggi, esercizio)}
    unita = richieste or [u for u in anag if u in sit]
    messaggi, senza_situazione, avvisi = [], [], []
    for uid in unita:
        if uid not in sit:
            senza_situazione.append(uid)
            continue
        messaggi.append(componi(sit[uid], anag.get(uid, {}), kv, a.oggi))
    if reg._blank(kv.get("Conto corrente")):
        avvisi.append("'Conto corrente' vuoto nel foglio Condominio: le email non riportano l'IBAN")

    salvati = []
    if a.salva:
        cartella = os.path.join(os.path.dirname(path_r), "comunicazioni")
        os.makedirs(cartella, exist_ok=True)
        for m in messaggi:
            base = os.path.join(cartella, f"{a.oggi.isoformat()}_situazione_{m['unita']}")
            path, k = base + ".md", 1
            while os.path.exists(path):
                k += 1
                path = f"{base}_{k}.md"
            with open(path, "w", encoding="utf-8") as f:
                f.write(f"A: {m['a'] or '(nessuna email: consegna a mano)'}\nOggetto: {m['oggetto']}\n\n{m['corpo']}\n")
            salvati.append(path)

    collaudo = str(kv.get("Modalità collaudo") or "").strip().upper() == "SI"
    print(json.dumps({
        "ok": True, "esercizio": esercizio, "oggi": a.oggi.isoformat(),
        "collaudo": collaudo, "email_collaudo": kv.get("Email collaudo") if collaudo else None,
        "messaggi": [m for m in messaggi if m["a"]],
        "senza_email": [m for m in messaggi if not m["a"]],
        "senza_situazione": senza_situazione, "avvisi": avvisi, "salvati": salvati,
    }, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
