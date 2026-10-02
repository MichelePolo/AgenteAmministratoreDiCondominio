#!/usr/bin/env python3
"""
registro.py — accesso deterministico ai registri del condominio.

Uso (dalla cartella del condominio, oppure con --dir <cartella>):
  registro.py info
  registro.py read <Foglio> [--where Col=Val] [--where Col!=Val] [--where Col=] [--csv]
  registro.py append <Foglio> '<json>'                  # una riga; assegna ID progressivo se c'è la colonna ID
  registro.py update <Foglio> --where Col=Val '<json>'  # aggiorna le righe che corrispondono
  registro.py set Condominio "<Chiave>" "<Valore>"      # foglio chiave/valore
  registro.py diario "<Operazione>" [--dettaglio "..."] [--eseguito IA] [--approvato "Nome"]
  registro.py verifica                                  # millesimi a 1000, anagrafica, spese senza tabella…
                                                        # (riscrive anche TOTALE/Versato/Saldo dopo modifiche a mano)
  registro.py schema                                    # aggiunge fogli/colonne/chiavi mancanti (aggiornamento registri)
  registro.py --file registro-riservato.xlsx situazione [--unita U01] [--esercizio 2026]
                                                        # per unità: dovuto, versato, saldo, scaduto, stato delle rate

Opzioni comuni: --file registro-riservato.xlsx (default registro-condominio.xlsx), --dir <cartella>,
--oggi AAAA-MM-GG (data di riferimento per rate e scaduto; default oggi, utile per le prove)

I registri NON contengono formule: i valori derivati (riga TOTALE dei millesimi, Dovuto, Versato,
Saldo e Scaduto del foglio Situazione) vengono ricalcolati e scritti da questo strumento a ogni lettura e a ogni
scrittura. Così il file si legge correttamente anche dall'anteprima di Google Drive su cellulare,
che mostra solo i valori salvati nel file.
Richiede openpyxl (python3 -m pip install openpyxl).
"""
import argparse
import datetime as dt
import json
import os
import sys
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("openpyxl non installato. Installarlo con: python3 -m pip install openpyxl (Windows: py -m pip install openpyxl; Linux con errore 'externally-managed-environment': sudo apt install python3-openpyxl)")

DEFAULT_FILE = "registro-condominio.xlsx"
KV_SHEETS = {"Condominio"}
IN_BREVE = "In breve"
GENERATED_SHEETS = {IN_BREVE}  # rigenerati a ogni salvataggio: mai letti né scritti come dati
INBOX = "da analizzare"
INBOX_SKIP_EXT = {".gdoc", ".gsheet", ".gslides", ".gform", ".gdraw", ".gmap", ".tmp", ".crdownload", ".drivedownload", ".driveupload"}
FONT = "Arial"
FMT_EURO = "#,##0.00 €"
FMT_DATE = "yyyy-mm-dd"
CENT = Decimal("0.01")
COL_IMPORTI = {"importo", "quota", "dovuto", "versato", "saldo", "scaduto"}
COL_DATE = {"data", "data pagamento", "ultimo sollecito", "data elaborazione", "scadenza", "dal"}
COL_INT = {"esercizio", "piano", "id spesa", "livello", "livello sollecito", "id", "rata",
           "esercizio corrente", "numero unità"}
COL_PCT = {"quota conduttore %"}
HDR_FILL = PatternFill("solid", fgColor="DDE7F0")

# Schema di riferimento: usato da `schema` per aggiungere ciò che manca a registri creati da versioni precedenti.
SCHEMA = {
    "registro-condominio.xlsx": {
        "Condominio": ["Nome", "Indirizzo", "Codice fiscale", "Profilo", "Amministratore", "Email amministratore",
                       "Modalità collaudo", "Email collaudo", "Esercizio corrente", "Inizio esercizio", "Fine esercizio",
                       "Numero unità", "Tabelle millesimali", "Conto corrente", "Percorso registro riservato", "Ultima elaborazione"],
        "Anagrafica": ["ID unità", "Interno", "Piano", "Intestatario", "Email", "Telefono", "Conduttore", "Email conduttore",
                       "Dati catastali", "Note"],
        "Millesimi": ["ID unità", "Intestatario", "Tab A", "Tab B", "Tab C"],
        "Spese": ["ID", "Data", "Fornitore", "Descrizione", "Importo", "Tabella", "Quota conduttore %", "Esercizio", "File",
                  "Pagata", "Data pagamento", "Note", "ID movimento"],
        "Riparti manuali": ["ID spesa", "ID unità", "Quota", "Note"],
        "Scadenze": ["ID", "Data", "Ora", "Tipo", "Descrizione", "Ricorrenza", "ID evento", "Note"],
        "Diario": ["Data-ora", "Operazione", "Dettaglio", "Eseguito da", "Approvato da"],
        "Indice": ["Hash", "Nome originale", "Nome archivio", "Percorso", "Data elaborazione", "ID spesa"],
    },
    "registro-riservato.xlsx": {
        "Situazione": ["ID unità", "Intestatario", "Esercizio", "Dovuto", "Versato", "Saldo", "Ultimo sollecito",
                       "Livello sollecito", "Note", "Scaduto"],
        "Versamenti": ["ID", "Data", "ID unità", "Importo", "Esercizio", "Riferimento", "Rata", "Note", "ID movimento"],
        "Solleciti": ["ID", "Data", "ID unità", "Livello", "Inviato a", "Canale", "Esito", "Note"],
        "Rate": ["ID unità", "Esercizio", "Prospetto", "Rata", "Scadenza", "Importo", "Valida", "Note"],
        "Ordinanti": ["Ordinante", "ID unità", "Dal", "Attiva", "Note"],
    },
}


# ---------------------------------------------------------------- utilità

def _wb(path):
    if not os.path.exists(path):
        sys.exit(f"File non trovato: {path}. Eseguire condominio-setup.")
    return openpyxl.load_workbook(path)


def _blank(v):
    return v is None or (isinstance(v, str) and not v.strip())


def _headers(ws):
    """Intestazioni della riga 1, fino alla prima cella vuota (le note oltre non contano)."""
    out = []
    for c in ws[1]:
        if _blank(c.value):
            break
        out.append(c.value)
    return out


def _rows(ws):
    hdr = _headers(ws)
    if not hdr:
        return []
    out = []
    for r in ws.iter_rows(min_row=2, max_col=len(hdr), values_only=True):
        if all(_blank(v) for v in r):
            continue
        out.append(dict(zip(hdr, r)))
    return out


def _num(v):
    if v is None or v == "" or isinstance(v, bool):
        return None
    if isinstance(v, (int, float, Decimal)):
        return Decimal(str(v))
    try:
        return Decimal(str(v).strip().replace(",", "."))
    except InvalidOperation:
        return None


def _plain(d):
    """Decimal → int se intero, altrimenti float (Excel e JSON li gestiscono bene)."""
    d = Decimal(d)
    return int(d) if d == d.to_integral_value() else float(d)


def _norm(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat() if v.time() == dt.time(0, 0) else v.isoformat(timespec="seconds")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, Decimal):
        return _plain(v)
    return v


def _coerce(header, value):
    """Numeri e date come tipi nativi, così Excel/Sheets li trattano correttamente."""
    if _blank(value):
        return None
    h = str(header or "").lower()
    if h in COL_IMPORTI:
        d = _num(value)
        return float(d.quantize(CENT)) if d is not None else str(value).strip()
    if h.startswith("tab "):
        d = _num(value)
        return _plain(d) if d is not None else str(value).strip()
    if h in COL_INT:
        d = _num(value)
        return int(d) if d is not None and d == d.to_integral_value() else str(value).strip()
    if h in COL_PCT:
        d = _num(value)
        return _plain(d) if d is not None else str(value).strip()
    if h in COL_DATE:
        if isinstance(value, (dt.date, dt.datetime)):
            return value
        try:
            return dt.date.fromisoformat(str(value).strip()[:10])
        except ValueError:
            return str(value).strip()
    if isinstance(value, (int, float)):
        return value
    return str(value).strip()


def _put(ws, r, c, header, value):
    cell = ws.cell(r, c)
    cell.value = _coerce(header, value)
    cell.font = Font(name=FONT)
    h = str(header or "").lower()
    if h in COL_IMPORTI:
        cell.number_format = FMT_EURO
    elif h in COL_DATE:
        cell.number_format = FMT_DATE
    return cell


def _totale_row(ws):
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip().upper() == "TOTALE":
            return r
    return None


def _next_free_row(ws, ncol):
    r = ws.max_row
    while r >= 2 and all(_blank(ws.cell(r, c).value) for c in range(1, ncol + 1)):
        r -= 1
    return r + 1


def _touch(wb):
    if "Condominio" in wb.sheetnames:
        ws = wb["Condominio"]
        for row in ws.iter_rows(min_row=1, max_col=2):
            if row[0].value == "Ultima elaborazione":
                row[1].value = dt.datetime.now().isoformat(timespec="seconds")
                return


def _esercizio_di(vers):
    e = _num(vers.get("Esercizio"))
    if e is not None:
        return int(e)
    d = vers.get("Data")
    if isinstance(d, (dt.date, dt.datetime)):
        return d.year
    if isinstance(d, str) and d[:4].isdigit():
        return int(d[:4])
    return None


def _as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    try:
        return dt.date.fromisoformat(str(v).strip()[:10])
    except ValueError:
        return None


def _stesso_esercizio(es_riga, es):
    return es is None or es_riga is None or es_riga == es


def _conti(wb, uid, es, dovuto_scritto, oggi):
    """Conti di un'unità per un esercizio (es None = tutti), dal registro riservato.

    Dovuto: somma delle rate valide se l'unità ha righe in `Rate` per l'esercizio, altrimenti il
    valore scritto (registri 0.3). Versamenti imputati alle rate in ordine di scadenza; una rata senza
    scadenza (dovuto riportato da un registro 0.3) viene coperta per prima e non risulta mai scaduta. Scaduto:
    None senza rate (non si sa quanto sia in ritardo), altrimenti rate scadute non coperte."""
    uid = str(uid).strip()
    vers = [v for v in (_rows(wb["Versamenti"]) if "Versamenti" in wb.sheetnames else [])
            if str(v.get("ID unità") or "").strip() == uid and _stesso_esercizio(_esercizio_di(v), es)]
    versato = sum((_num(v.get("Importo")) or Decimal(0) for v in vers), Decimal(0))
    righe_rate = []
    if "Rate" in wb.sheetnames:
        for r in _rows(wb["Rate"]):
            e = _num(r.get("Esercizio"))
            if str(r.get("ID unità") or "").strip() == uid and _stesso_esercizio(int(e) if e is not None else None, es):
                righe_rate.append(r)
    valide = [r for r in righe_rate if str(r.get("Valida") or "SI").strip().upper() != "NO"]
    valide.sort(key=lambda r: (_as_date(r.get("Scadenza")) or dt.date.min, str(r.get("Prospetto") or ""), _num(r.get("Rata")) or 0))
    if righe_rate:
        dovuto = sum((_num(r.get("Importo")) or Decimal(0) for r in valide), Decimal(0))
    else:
        dovuto = _num(dovuto_scritto) or Decimal(0)
    rate, residuo, scoperto = [], versato, Decimal(0)
    for r in valide:
        imp = _num(r.get("Importo")) or Decimal(0)
        pagato = max(Decimal(0), min(imp, residuo))
        residuo -= pagato
        scad = _as_date(r.get("Scadenza"))
        if scad is not None and scad < oggi:
            scoperto += imp - pagato  # dall'imputazione, non da "scadute − versato": la rata 0 senza data assorbe i primi versamenti
        if pagato >= imp:
            stato = "pagata"
        elif scad is not None and scad < oggi:
            stato = "scaduta"
        elif pagato > 0:
            stato = "parziale"
        else:
            stato = "da pagare"
        rate.append({"Prospetto": r.get("Prospetto"), "Rata": _norm(r.get("Rata")), "Scadenza": _norm(r.get("Scadenza")),
                     "Importo": _plain(imp), "Pagato": _plain(pagato), "Stato": stato})
    scaduto = max(Decimal(0), scoperto) if righe_rate else None
    return {"Dovuto": dovuto, "Versato": versato, "Saldo": dovuto - versato, "Scaduto": scaduto,
            "rate": rate, "versamenti": vers}


def _refresh(wb, oggi=None):
    """Ricalcola i valori derivati e li scrive come numeri (mai formule)."""
    oggi = oggi or dt.date.today()
    for ws in wb.worksheets:
        hdr = _headers(ws)
        tab_cols = [i for i, h in enumerate(hdr, start=1) if isinstance(h, str) and h.startswith("Tab ")]
        tot = _totale_row(ws) if tab_cols else None
        if tot:
            for c in tab_cols:
                s = sum((_num(ws.cell(r, c).value) or Decimal(0)) for r in range(2, tot))
                cell = ws.cell(tot, c)
                cell.value = _plain(s)
                cell.font = Font(name=FONT, bold=True)
    if "Situazione" in wb.sheetnames and "Versamenti" in wb.sheetnames:
        ws = wb["Situazione"]
        hdr = _headers(ws)
        if {"ID unità", "Dovuto", "Versato", "Saldo"} <= set(hdr):
            ci = {h: i for i, h in enumerate(hdr, start=1)}
            for r in range(2, ws.max_row + 1):
                uid = ws.cell(r, ci["ID unità"]).value
                if _blank(uid):
                    continue
                es = _num(ws.cell(r, ci["Esercizio"]).value) if "Esercizio" in ci else None
                c = _conti(wb, uid, int(es) if es is not None else None, ws.cell(r, ci["Dovuto"]).value, oggi)
                for h in ("Dovuto", "Versato", "Saldo", "Scaduto"):
                    if h in ci:
                        _put(ws, r, ci[h], h, c[h])


def _save(wb, path, oggi=None):
    """Unico punto di salvataggio: ricalcola i valori derivati, aggiorna la data di elaborazione,
    rigenera il foglio In breve (solo registro condiviso), salva."""
    _refresh(wb, oggi)
    _touch(wb)
    if os.path.basename(path) == DEFAULT_FILE:
        _in_breve(wb, os.path.dirname(os.path.abspath(path)), oggi or dt.date.today())
    wb.save(path)


# ---------------------------------------------------------------- foglio "In breve"
# Primo foglio del registro condiviso, pensato per l'anteprima di Google Drive su telefono: due colonne,
# una sezione sotto l'altra, testo già formattato in italiano. Nessun dato per unità (nomi, ID, importi).

MESI = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre",
        "ottobre", "novembre", "dicembre"]


def data_it(d):
    """2026-11-30 → '30 novembre 2026'."""
    d = _as_date(d)
    return f"{d.day} {MESI[d.month - 1]} {d.year}" if d else ""


def euro_it(x):
    """1250.5 → '1.250,50 €'."""
    d = (_num(x) or Decimal(0)).quantize(CENT, rounding=ROUND_HALF_UP)
    testo = f"{abs(d):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if d < 0 else ''}{testo} €"


def da_ignorare(nome):
    """File dell'inbox che non sono documenti: nascosti, temporanei, scorciatoie di Google Drive."""
    return (nome.startswith(".") or nome.startswith("~$") or nome.lower() in ("desktop.ini", "thumbs.db")
            or os.path.splitext(nome)[1].lower() in INBOX_SKIP_EXT)


def file_in_attesa(cartella):
    inbox = os.path.join(cartella, INBOX)
    if not os.path.isdir(inbox):
        return 0
    return sum(1 for _, _, nomi in os.walk(inbox) for n in nomi if not da_ignorare(n))


def dati_in_breve(wb, cartella, oggi):
    """Raccoglie i dati pubblici del foglio In breve. Nessun dato per unità: le spese a carico di una sola
    unità sono sommate insieme, senza dire quale."""
    kv = _kv(wb["Condominio"]) if "Condominio" in wb.sheetnames else {}
    es = _num(kv.get("Esercizio corrente"))
    es = int(es) if es is not None else oggi.year
    etichette = {}
    for parte in str(kv.get("Tabelle millesimali") or "").split(";"):
        codice, _, nome = parte.partition(":")
        if codice.strip() and nome.strip():
            etichette[codice.strip().upper()] = nome.strip()

    scadenze = []
    for r in _rows(wb["Scadenze"]) if "Scadenze" in wb.sheetnames else []:
        d = _as_date(r.get("Data"))
        if d and d >= oggi:
            scadenze.append({"data": d, "ora": r.get("Ora"), "tipo": str(r.get("Tipo") or "").strip().lower(),
                             "descrizione": r.get("Descrizione") or ""})
    scadenze.sort(key=lambda x: (x["data"], str(x["ora"] or "")))

    per_tab, totale, numero, da_pagare = {}, Decimal(0), 0, 0
    for r in _rows(wb["Spese"]) if "Spese" in wb.sheetnames else []:
        imp = _num(r.get("Importo"))
        if imp is None:
            continue
        e = _num(r.get("Esercizio"))
        anno = int(e) if e is not None else (_as_date(r.get("Data")).year if _as_date(r.get("Data")) else None)
        if anno != es:
            continue
        tab = str(r.get("Tabella") or "").strip().upper()
        if tab.startswith("UNITA:") or tab.startswith("UNITÀ:"):
            chiave, nome = "UNITA", "a carico di singole unità"
        elif tab == "MANUALE":
            chiave, nome = "MANUALE", "riparto concordato"
        elif tab:
            chiave, nome = tab, etichette.get(tab, f"tabella {tab}")
        else:
            chiave, nome = "?", "tabella da assegnare"
        voce = per_tab.setdefault(chiave, {"codice": chiave, "nome": nome, "totale": Decimal(0), "numero": 0})
        voce["totale"] += imp
        voce["numero"] += 1
        totale += imp
        numero += 1
        if str(r.get("Pagata") or "").strip().upper() != "SI":
            da_pagare += 1

    archiviati = []
    for r in _rows(wb["Indice"]) if "Indice" in wb.sheetnames else []:
        nome = str(r.get("Nome archivio") or "")
        if nome and not nome.startswith("DUPLICATO"):
            archiviati.append({"data": _as_date(r.get("Data elaborazione")), "nome": nome})
    archiviati.sort(key=lambda x: x["data"] or dt.date.min, reverse=True)

    return {
        "nome": kv.get("Nome") or "Condominio",
        "oggi": oggi,
        "esercizio": es,
        "prossima_assemblea": next((x for x in scadenze if x["tipo"] == "assemblea"), None),
        "prossime_rate": [x for x in scadenze if x["tipo"] == "rata"],
        "prossime_scadenze": scadenze,
        "spese": {"totale": totale, "numero": numero, "non_pagate": da_pagare,
                  "per_tabella": sorted(per_tab.values(), key=lambda v: v["codice"])},
        "ultimi_archiviati": archiviati,
        "in_attesa": file_in_attesa(cartella),
    }


def descrivi_file(nome):
    """Nome d'archivio → (data del documento, descrizione leggibile).
    '2026-03-14_fattura_enel_luce-scale_412.50.pdf' → (2026-03-14, 'Fattura Enel, luce scale · 412,50 €').
    Un nome che non segue le convenzioni resta com'è, senza data."""
    base = os.path.splitext(os.path.basename(nome))[0]
    parti = base.split("_")
    data = _as_date(parti[0]) if parti else None
    if data is None or len(parti) < 3:
        return None, nome
    tipo, controparte = parti[1].capitalize(), parti[2].replace("-", " ").title()
    resto = parti[3:]
    importo = resto.pop() if resto and _num(resto[-1]) is not None else None
    testo = f"{tipo} {controparte}"
    if resto:
        testo += ", " + " ".join(resto).replace("-", " ")
    if importo is not None:
        testo += f" · {euro_it(importo)}"
    return data, testo


def sezioni_in_breve(dati):
    """Sceglie cosa vede per primo un condomino dal telefono: [(titolo sezione, [(voce, valore), ...]), ...].

    In cima ciò che si cerca (assemblea, prossima rata), poi poche altre scadenze, le spese dell'anno
    e gli ultimi documenti archiviati; in fondo, sempre, i documenti caricati e non ancora archiviati
    ("l'hanno visto?"). Liste corte: sul telefono in verticale il foglio deve restare breve."""
    def quando(x):
        return data_it(x["data"]) + (f", ore {x['ora']}" if x.get("ora") else "")

    sezioni, gia_mostrate = [], []
    in_cima = []
    if dati["prossima_assemblea"]:
        a = dati["prossima_assemblea"]
        in_cima.append(("Prossima assemblea", f"{quando(a)} — {a['descrizione']}"))
        gia_mostrate.append(a)
    if dati["prossime_rate"]:
        r = dati["prossime_rate"][0]
        in_cima.append(("Prossima rata", f"{quando(r)} — {r['descrizione']}"))
        gia_mostrate.append(r)
    if in_cima:
        sezioni.append(("Da ricordare", in_cima))

    altre = [x for x in dati["prossime_scadenze"] if not any(x is g for g in gia_mostrate)][:4]
    if altre:
        sezioni.append(("Altre scadenze", [(quando(x), x["descrizione"]) for x in altre]))

    sp = dati["spese"]
    if sp["numero"]:
        voci = [(f"Totale ({sp['numero']} {'spesa' if sp['numero'] == 1 else 'spese'})", euro_it(sp["totale"]))]
        voci += [(f"di cui {v['nome']}", euro_it(v["totale"])) for v in sp["per_tabella"]]
    else:
        voci = [("Nessuna spesa registrata", "")]
    sezioni.append((f"Spese registrate nel {dati['esercizio']}", voci))

    archiviati = []
    for x in dati["ultimi_archiviati"][:5]:
        data_doc, testo = descrivi_file(x["nome"])
        archiviati.append((data_it(data_doc or x["data"]), testo))
    if archiviati:
        sezioni.append(("Ultimi documenti archiviati", archiviati))

    sezioni.append(("Cartella «da analizzare»", [("Documenti caricati, in attesa di archiviazione", str(dati["in_attesa"]))]))
    return sezioni


def _in_breve(wb, cartella, oggi):
    """Rigenera il foglio In breve e lo mette per primo e attivo (è quello che l'anteprima Drive apre)."""
    dati = dati_in_breve(wb, cartella, oggi)
    if IN_BREVE in wb.sheetnames:
        del wb[IN_BREVE]
    ws = wb.create_sheet(IN_BREVE, 0)
    for altro in wb.worksheets:
        altro.sheet_view.tabSelected = False
    wb.active = 0
    ws.sheet_view.tabSelected = True
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 42
    a_capo = Alignment(wrap_text=True, vertical="top")

    ws.append([dati["nome"]])
    ws["A1"].font = Font(name=FONT, bold=True, size=14)
    ws.append([f"Aggiornato il {data_it(oggi)}"])
    ws["A2"].font = Font(name=FONT, italic=True, color="666666")
    for titolo, voci in sezioni_in_breve(dati):
        ws.append([])
        ws.append([titolo, None])
        for c in ws[ws.max_row]:
            c.font = Font(name=FONT, bold=True)
            c.fill = HDR_FILL
        for voce, valore in voci:
            ws.append([voce, valore])
            for c in ws[ws.max_row]:
                c.font = Font(name=FONT)
                c.alignment = a_capo
    ws.append([])
    for nota in ("I pagamenti delle singole unità non sono qui: chiedi all'amministratore.",
                 "Foglio aggiornato automaticamente: le modifiche fatte a mano vengono sovrascritte."):
        ws.append([nota])
        ws.cell(ws.max_row, 1).font = Font(name=FONT, italic=True, color="666666")


def _cond(cond):
    if "!=" in cond:
        k, v = cond.split("!=", 1)
        return k.strip(), v.strip(), True
    k, _, v = cond.partition("=")
    return k.strip(), v.strip(), False


def _eq(rv, v):
    if v == "":
        return _blank(rv)
    if _blank(rv):
        return False
    if str(_norm(rv)) == v:
        return True
    a, b = _num(rv), _num(v)
    return a is not None and b is not None and a == b


def _match(row, where):
    for cond in where or []:
        k, v, neg = _cond(cond)
        if k not in row:
            sys.exit(f"Colonna sconosciuta nel filtro: {k}. Colonne: {list(row)}")
        if _eq(row.get(k), v) == neg:
            return False
    return True


def _kv(ws):
    return {r[0].value: _norm(r[1].value) for r in ws.iter_rows(min_row=1, max_col=2) if not _blank(r[0].value)}


def percorso_riservato(cartella, wb_condominio=None):
    """Percorso di registro-riservato.xlsx: la chiave `Percorso registro riservato` del foglio Condominio
    (assoluta o relativa alla cartella del condominio) se valorizzata, altrimenti la cartella stessa."""
    if wb_condominio is None:
        wb_condominio = _wb(os.path.join(cartella, DEFAULT_FILE))
    altro = _kv(wb_condominio["Condominio"]).get("Percorso registro riservato") if "Condominio" in wb_condominio.sheetnames else None
    base = os.path.join(cartella, os.path.expanduser(str(altro).strip())) if not _blank(altro) else cartella
    return os.path.join(base, "registro-riservato.xlsx")


def _out(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=2, default=str))


# ---------------------------------------------------------------- comandi

def cmd_info(a):
    wb = _wb(a.path)
    _refresh(wb, a.oggi)
    out = {"file": a.path, "fogli": wb.sheetnames}
    if "Condominio" in wb.sheetnames:
        out["condominio"] = _kv(wb["Condominio"])
    for s in wb.sheetnames:
        if s not in KV_SHEETS and s not in GENERATED_SHEETS:
            righe = _rows(wb[s])
            out[f"righe_{s}"] = len([r for r in righe if str(r.get(_headers(wb[s])[0]) or "").strip().upper() != "TOTALE"])
    _out(out)


def cmd_read(a):
    wb = _wb(a.path)
    if a.sheet not in wb.sheetnames:
        sys.exit(f"Foglio inesistente: {a.sheet}. Disponibili: {wb.sheetnames}")
    if a.sheet in GENERATED_SHEETS:
        sys.exit(f"'{a.sheet}' è un foglio di sola consultazione, generato dagli altri: leggere quelli")
    _refresh(wb, a.oggi)
    ws = wb[a.sheet]
    if a.sheet in KV_SHEETS:
        _out(_kv(ws))
        return
    rows = [r for r in _rows(ws) if _match(r, a.where)]
    rows = [{k: _norm(v) for k, v in r.items()} for r in rows]
    if a.csv:
        import csv
        w = csv.DictWriter(sys.stdout, fieldnames=_headers(ws))
        w.writeheader()
        w.writerows(rows)
    else:
        _out(rows)


def aggiungi(ws, data):
    """Aggiunge una riga (dict colonna→valore) prima della riga TOTALE o in coda; assegna l'ID progressivo.
    Restituisce (numero di riga, ID). Solleva ValueError se ci sono colonne sconosciute."""
    hdr = _headers(ws)
    data = dict(data)
    unknown = [k for k in data if k not in hdr]
    if unknown:
        raise ValueError(f"Colonne sconosciute {unknown}. Colonne del foglio: {hdr}")
    if "ID" in hdr and _blank(data.get("ID")):
        ids = [_num(r.get("ID")) for r in _rows(ws)]
        ids = [int(i) for i in ids if i is not None]
        data["ID"] = max(ids) + 1 if ids else 1
    tot = _totale_row(ws)
    if tot:  # inserisce prima della riga TOTALE
        ws.insert_rows(tot)
        insert_at = tot
    else:
        insert_at = _next_free_row(ws, len(hdr))
    for col, h in enumerate(hdr, start=1):
        if h in data:
            _put(ws, insert_at, col, h, data[h])
    return insert_at, data.get("ID")


def cmd_append(a):
    wb = _wb(a.path)
    if a.sheet not in wb.sheetnames or a.sheet in KV_SHEETS or a.sheet in GENERATED_SHEETS:
        sys.exit(f"Foglio non valido per append: {a.sheet}")
    try:
        riga, nuovo_id = aggiungi(wb[a.sheet], json.loads(a.json))
    except ValueError as e:
        sys.exit(str(e))
    _save(wb, a.path, a.oggi)
    print(json.dumps({"ok": True, "foglio": a.sheet, "riga": riga, "ID": nuovo_id}, ensure_ascii=False))


def cmd_update(a):
    wb = _wb(a.path)
    if a.sheet not in wb.sheetnames or a.sheet in KV_SHEETS or a.sheet in GENERATED_SHEETS:
        sys.exit(f"Foglio non valido per update: {a.sheet}")
    ws = wb[a.sheet]
    hdr = _headers(ws)
    data = json.loads(a.json)
    for k in data:
        if k not in hdr:
            sys.exit(f"Colonna sconosciuta: {k}. Colonne del foglio: {hdr}")
    n = 0
    for r in range(2, ws.max_row + 1):
        row = {h: ws.cell(r, c).value for c, h in enumerate(hdr, start=1)}
        if all(_blank(v) for v in row.values()):
            continue
        if _match(row, a.where):
            for k, v in data.items():
                _put(ws, r, hdr.index(k) + 1, k, v)
            n += 1
    _save(wb, a.path, a.oggi)
    print(json.dumps({"ok": True, "righe_aggiornate": n}))


def cmd_set(a):
    wb = _wb(a.path)
    if a.sheet not in wb.sheetnames:
        sys.exit(f"Foglio inesistente: {a.sheet}")
    ws = wb[a.sheet]
    for row in ws.iter_rows(min_row=1, max_col=2):
        if row[0].value == a.key:
            row[1].value = _coerce(a.key, a.value)
            row[1].font = Font(name=FONT)
            _save(wb, a.path, a.oggi)
            print(json.dumps({"ok": True, a.key: _norm(row[1].value)}, ensure_ascii=False))
            return
    sys.exit(f"Chiave non trovata: {a.key}")


def cmd_diario(a):
    wb = _wb(a.path)
    if "Diario" not in wb.sheetnames:
        sys.exit("Foglio Diario assente")
    ws = wb["Diario"]
    hdr = _headers(ws)
    r = _next_free_row(ws, len(hdr))
    valori = [dt.datetime.now().isoformat(timespec="seconds"), a.operazione, a.dettaglio or "", a.eseguito, a.approvato or ""]
    for c, (h, v) in enumerate(zip(hdr, valori), start=1):
        _put(ws, r, c, h, v)
    _save(wb, a.path, a.oggi)
    print(json.dumps({"ok": True, "riga": r}))


def verifica(wb):
    """Controlli di coerenza del registro. Restituisce (problemi bloccanti, avvisi); non scrive nulla."""
    problemi, avvisi = [], []
    ids_a, conduttori = [], []
    if "Anagrafica" in wb.sheetnames:
        an = _rows(wb["Anagrafica"])
        ids_a = [str(r["ID unità"]).strip() for r in an if not _blank(r.get("ID unità"))]
        if not ids_a:
            avvisi.append("Anagrafica vuota: inserire le unità")
        dup = sorted({u for u in ids_a if ids_a.count(u) > 1})
        if dup:
            problemi.append(f"Anagrafica: ID unità duplicati {dup}")
        senza_email = [str(r["ID unità"]) for r in an if not _blank(r.get("ID unità")) and _blank(r.get("Email"))]
        if senza_email and ids_a:
            avvisi.append(f"Unità senza email (comunicazioni solo cartacee): {senza_email}")
        conduttori = [str(r["ID unità"]) for r in an if not _blank(r.get("ID unità")) and not _blank(r.get("Conduttore"))]
        if "Email conduttore" not in _headers(wb["Anagrafica"]):
            avvisi.append("Anagrafica: manca la colonna 'Email conduttore' (eseguire `registro.py schema`)")
        elif conduttori:
            senza = [str(r["ID unità"]) for r in an if not _blank(r.get("Conduttore")) and _blank(r.get("Email conduttore"))]
            if senza:
                avvisi.append(f"Conduttori senza email: {senza}")
    if "Millesimi" in wb.sheetnames:
        ws = wb["Millesimi"]
        hdr = _headers(ws)
        righe = [r for r in _rows(ws) if str(r.get("ID unità") or "").strip().upper() != "TOTALE"]
        ids_m = [str(r["ID unità"]).strip() for r in righe if not _blank(r.get("ID unità"))]
        if not _totale_row(ws):
            avvisi.append("Millesimi: manca la riga TOTALE (gli script la aggiornano da soli)")
        for h in hdr:
            if isinstance(h, str) and h.startswith("Tab "):
                s = sum((_num(r.get(h)) or Decimal(0)) for r in righe)
                if s == 0:
                    avvisi.append(f"{h}: tutta a zero (tabella non usata?)")
                elif abs(s - 1000) > CENT:
                    problemi.append(f"{h} somma a {_plain(s)}, non a 1000")
        dup = sorted({u for u in ids_m if ids_m.count(u) > 1})
        if dup:
            problemi.append(f"Millesimi: ID unità duplicati {dup}")
        for u in ids_a:
            if u not in ids_m:
                problemi.append(f"Unità {u} in Anagrafica ma non in Millesimi")
        for u in ids_m:
            if u not in ids_a:
                problemi.append(f"Unità {u} in Millesimi ma non in Anagrafica")
    if "Spese" in wb.sheetnames:
        sp = _rows(wb["Spese"])
        senza_tab = [r.get("ID") for r in sp if _blank(r.get("Tabella")) and not _blank(r.get("Importo"))]
        if senza_tab:
            avvisi.append(f"Spese senza Tabella (il riparto le rifiuta): ID {senza_tab}")
        if "Quota conduttore %" not in _headers(wb["Spese"]):
            avvisi.append("Spese: manca la colonna 'Quota conduttore %' (eseguire `registro.py schema`)")
        else:
            fuori = [r.get("ID") for r in sp if (_num(r.get("Quota conduttore %")) is not None and not 0 <= _num(r.get("Quota conduttore %")) <= 100)]
            if fuori:
                problemi.append(f"Spese con 'Quota conduttore %' fuori da 0-100 (il riparto le rifiuta): ID {fuori}")
            if conduttori:
                senza_pct = [r.get("ID") for r in sp if not _blank(r.get("Importo")) and _blank(r.get("Quota conduttore %"))]
                if senza_pct:
                    avvisi.append(f"Spese senza 'Quota conduttore %' (nel prospetto andranno tutte al proprietario): ID {senza_pct}")
    if "Condominio" in wb.sheetnames:
        kv = _kv(wb["Condominio"])
        for k in ("Nome", "Amministratore", "Email amministratore"):
            if _blank(kv.get(k)):
                avvisi.append(f"Foglio Condominio: '{k}' vuoto")
        if str(kv.get("Modalità collaudo") or "").upper() == "SI" and _blank(kv.get("Email collaudo")):
            problemi.append("Modalità collaudo SI ma 'Email collaudo' vuota: nessuna email potrà partire")
        n = _num(kv.get("Numero unità"))
        if n is not None and ids_a and int(n) != len(ids_a):
            avvisi.append(f"'Numero unità' = {int(n)} ma in Anagrafica ci sono {len(ids_a)} unità")
    return problemi, avvisi


def cmd_verifica(a):
    wb = _wb(a.path)
    _save(wb, a.path, a.oggi)  # riscrive i valori derivati: utile dopo modifiche fatte a mano nel foglio
    problemi, avvisi = verifica(wb)
    _out({"ok": not problemi, "problemi": problemi, "avvisi": avvisi})
    sys.exit(0 if not problemi else 1)


def situazione(wb, oggi=None, esercizio=None, unita=None):
    """Situazione per unità dal registro riservato: una voce per riga di `Situazione`, con rate e versamenti."""
    oggi = oggi or dt.date.today()
    if "Situazione" not in wb.sheetnames:
        raise ValueError("foglio Situazione assente: usare --file registro-riservato.xlsx")
    out = []
    for r in _rows(wb["Situazione"]):
        uid = str(r.get("ID unità") or "").strip()
        es = _num(r.get("Esercizio"))
        es = int(es) if es is not None else None
        if not uid or (unita and uid != unita) or (esercizio and es != esercizio):
            continue
        c = _conti(wb, uid, es, r.get("Dovuto"), oggi)
        out.append({
            "ID unità": uid, "Intestatario": r.get("Intestatario"), "Esercizio": es,
            "Dovuto": _plain(c["Dovuto"]), "Versato": _plain(c["Versato"]), "Saldo": _plain(c["Saldo"]),
            "Scaduto": _plain(c["Scaduto"]) if c["Scaduto"] is not None else None,
            "Livello sollecito": _norm(r.get("Livello sollecito")), "Ultimo sollecito": _norm(r.get("Ultimo sollecito")),
            "rate": c["rate"],
            "versamenti": [{k: _norm(v.get(k)) for k in ("Data", "Importo", "Rata", "Riferimento")} for v in c["versamenti"]],
        })
    return out


def cmd_situazione(a):
    wb = _wb(a.path)
    try:
        dati = situazione(wb, a.oggi, a.esercizio, a.unita)
    except ValueError as e:
        sys.exit(str(e))
    if a.unita and not dati:
        sys.exit(f"Nessuna riga in Situazione per l'unità {a.unita}")
    _out({"ok": True, "oggi": a.oggi.isoformat(), "unita": dati})


def cmd_schema(a):
    wb = _wb(a.path)
    schema = SCHEMA.get(os.path.basename(a.path))
    if not schema:
        sys.exit(f"Nessuno schema noto per {os.path.basename(a.path)} (attesi: {list(SCHEMA)})")
    aggiunte = []
    for sheet, cols in schema.items():
        if sheet in KV_SHEETS:
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            presenti = {r[0].value for r in ws.iter_rows(min_row=1, max_col=1)}
            for k in cols:
                if k not in presenti:
                    ws.append([k, None])
                    ws.cell(ws.max_row, 1).font = Font(name=FONT, bold=True)
                    ws.cell(ws.max_row, 1).fill = HDR_FILL
                    aggiunte.append(f"{sheet}: chiave '{k}'")
            continue
        if sheet not in wb.sheetnames:
            ws = wb.create_sheet(sheet)
            ws.append(cols)
            ws.freeze_panes = "A2"
            for c in ws[1]:
                c.font = Font(name=FONT, bold=True); c.fill = HDR_FILL
            aggiunte.append(f"foglio '{sheet}'")
            continue
        ws = wb[sheet]
        hdr = _headers(ws)
        for k in cols:
            if k not in hdr:
                col = len(hdr) + 1
                cell = ws.cell(1, col)
                cell.value = k; cell.font = Font(name=FONT, bold=True); cell.fill = HDR_FILL
                ws.column_dimensions[get_column_letter(col)].width = max(14, len(k) + 2)
                hdr.append(k)
                aggiunte.append(f"{sheet}: colonna '{k}' (in coda)")
    if os.path.basename(a.path) == DEFAULT_FILE and IN_BREVE not in wb.sheetnames:
        aggiunte.append(f"foglio '{IN_BREVE}' (in prima posizione, generato)")
    if aggiunte:
        _save(wb, a.path, a.oggi)
    _out({"ok": True, "aggiunte": aggiunte})


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--file", default=DEFAULT_FILE)
    p.add_argument("--dir", default=".")
    p.add_argument("--oggi", type=dt.date.fromisoformat, default=dt.date.today(),
                   help="data di riferimento AAAA-MM-GG per rate e scaduto (default: oggi)")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("info")
    r = sub.add_parser("read"); r.add_argument("sheet"); r.add_argument("--where", action="append"); r.add_argument("--csv", action="store_true")
    ap = sub.add_parser("append"); ap.add_argument("sheet"); ap.add_argument("json")
    up = sub.add_parser("update"); up.add_argument("sheet"); up.add_argument("json"); up.add_argument("--where", action="append", required=True)
    st = sub.add_parser("set"); st.add_argument("sheet"); st.add_argument("key"); st.add_argument("value")
    d = sub.add_parser("diario"); d.add_argument("operazione"); d.add_argument("--dettaglio"); d.add_argument("--eseguito", default="IA"); d.add_argument("--approvato")
    sub.add_parser("verifica")
    sub.add_parser("schema")
    si = sub.add_parser("situazione"); si.add_argument("--unita"); si.add_argument("--esercizio", type=int)
    a = p.parse_args()
    a.path = os.path.join(a.dir, a.file)
    {"info": cmd_info, "read": cmd_read, "append": cmd_append, "update": cmd_update,
     "set": cmd_set, "diario": cmd_diario, "verifica": cmd_verifica, "schema": cmd_schema,
     "situazione": cmd_situazione}[a.cmd](a)


if __name__ == "__main__":
    main()
