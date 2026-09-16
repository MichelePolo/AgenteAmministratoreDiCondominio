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

Opzioni comuni: --file registro-riservato.xlsx (default registro-condominio.xlsx), --dir <cartella>

I registri NON contengono formule: i valori derivati (riga TOTALE dei millesimi, Versato e Saldo
del foglio Situazione) vengono ricalcolati e scritti da questo strumento a ogni lettura e a ogni
scrittura. Così il file si legge correttamente anche dall'anteprima di Google Drive su cellulare,
che mostra solo i valori salvati nel file.
Richiede openpyxl (pip install openpyxl).
"""
import argparse
import datetime as dt
import json
import os
import sys
from decimal import Decimal, InvalidOperation

try:
    import openpyxl
    from openpyxl.styles import Font
except ImportError:  # pragma: no cover
    sys.exit("openpyxl non installato: pip install openpyxl")

DEFAULT_FILE = "registro-condominio.xlsx"
KV_SHEETS = {"Condominio"}
FONT = "Arial"
FMT_EURO = "#,##0.00 €"
FMT_DATE = "yyyy-mm-dd"
CENT = Decimal("0.01")
COL_IMPORTI = {"importo", "quota", "dovuto", "versato", "saldo"}
COL_DATE = {"data", "data pagamento", "ultimo sollecito", "data elaborazione"}
COL_INT = {"esercizio", "piano", "id spesa", "livello", "livello sollecito", "id", "rata",
           "esercizio corrente", "numero unità"}


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


def _refresh(wb):
    """Ricalcola i valori derivati e li scrive come numeri (mai formule)."""
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
            vers = _rows(wb["Versamenti"])
            for r in range(2, ws.max_row + 1):
                uid = ws.cell(r, ci["ID unità"]).value
                if _blank(uid):
                    continue
                es = _num(ws.cell(r, ci["Esercizio"]).value) if "Esercizio" in ci else None
                es = int(es) if es is not None else None
                versato = Decimal(0)
                for v in vers:
                    if str(v.get("ID unità") or "").strip() != str(uid).strip():
                        continue
                    ev = _esercizio_di(v)
                    if es is not None and ev is not None and ev != es:
                        continue
                    versato += _num(v.get("Importo")) or Decimal(0)
                dovuto = _num(ws.cell(r, ci["Dovuto"]).value) or Decimal(0)
                _put(ws, r, ci["Versato"], "Versato", versato)
                _put(ws, r, ci["Saldo"], "Saldo", dovuto - versato)


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


def _out(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=2, default=str))


# ---------------------------------------------------------------- comandi

def cmd_info(a):
    wb = _wb(a.path)
    _refresh(wb)
    out = {"file": a.path, "fogli": wb.sheetnames}
    if "Condominio" in wb.sheetnames:
        out["condominio"] = _kv(wb["Condominio"])
    for s in wb.sheetnames:
        if s not in KV_SHEETS:
            righe = _rows(wb[s])
            out[f"righe_{s}"] = len([r for r in righe if str(r.get(_headers(wb[s])[0]) or "").strip().upper() != "TOTALE"])
    _out(out)


def cmd_read(a):
    wb = _wb(a.path)
    if a.sheet not in wb.sheetnames:
        sys.exit(f"Foglio inesistente: {a.sheet}. Disponibili: {wb.sheetnames}")
    _refresh(wb)
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


def cmd_append(a):
    wb = _wb(a.path)
    if a.sheet not in wb.sheetnames or a.sheet in KV_SHEETS:
        sys.exit(f"Foglio non valido per append: {a.sheet}")
    ws = wb[a.sheet]
    hdr = _headers(ws)
    data = json.loads(a.json)
    unknown = [k for k in data if k not in hdr]
    if unknown:
        sys.exit(f"Colonne sconosciute {unknown}. Colonne del foglio: {hdr}")
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
    _refresh(wb)
    _touch(wb)
    wb.save(a.path)
    print(json.dumps({"ok": True, "foglio": a.sheet, "riga": insert_at, "ID": data.get("ID")}, ensure_ascii=False))


def cmd_update(a):
    wb = _wb(a.path)
    if a.sheet not in wb.sheetnames or a.sheet in KV_SHEETS:
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
    _refresh(wb)
    _touch(wb)
    wb.save(a.path)
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
            _touch(wb)
            wb.save(a.path)
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
    _touch(wb)
    wb.save(a.path)
    print(json.dumps({"ok": True, "riga": r}))


def cmd_verifica(a):
    wb = _wb(a.path)
    _refresh(wb)
    wb.save(a.path)  # riscrive i valori derivati: utile dopo modifiche fatte a mano nel foglio
    problemi, avvisi = [], []
    ids_a = []
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
        senza_tab = [r.get("ID") for r in _rows(wb["Spese"]) if _blank(r.get("Tabella")) and not _blank(r.get("Importo"))]
        if senza_tab:
            avvisi.append(f"Spese senza Tabella (il riparto le rifiuta): ID {senza_tab}")
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
    _out({"ok": not problemi, "problemi": problemi, "avvisi": avvisi})
    sys.exit(0 if not problemi else 1)


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--file", default=DEFAULT_FILE)
    p.add_argument("--dir", default=".")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("info")
    r = sub.add_parser("read"); r.add_argument("sheet"); r.add_argument("--where", action="append"); r.add_argument("--csv", action="store_true")
    ap = sub.add_parser("append"); ap.add_argument("sheet"); ap.add_argument("json")
    up = sub.add_parser("update"); up.add_argument("sheet"); up.add_argument("json"); up.add_argument("--where", action="append", required=True)
    st = sub.add_parser("set"); st.add_argument("sheet"); st.add_argument("key"); st.add_argument("value")
    d = sub.add_parser("diario"); d.add_argument("operazione"); d.add_argument("--dettaglio"); d.add_argument("--eseguito", default="IA"); d.add_argument("--approvato")
    sub.add_parser("verifica")
    a = p.parse_args()
    a.path = os.path.join(a.dir, a.file)
    {"info": cmd_info, "read": cmd_read, "append": cmd_append, "update": cmd_update,
     "set": cmd_set, "diario": cmd_diario, "verifica": cmd_verifica}[a.cmd](a)


if __name__ == "__main__":
    main()
