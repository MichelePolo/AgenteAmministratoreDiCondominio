#!/usr/bin/env python3
"""
riparto.py — calcola il riparto delle spese per unità secondo le tabelle millesimali.

  riparto.py --dir "<cartella>" [--esercizio 2026] [--ids 1,2,3] [--dal 2026-01-01] [--al 2026-06-30]
             [--titolo "Riparto 1° semestre 2026"] [--rate 4] [--out prospetti/riparto-2026-1sem.xlsx]

Regole:
- Tabella A/B/C/…: quota = importo × millesimi / 1000, arrotondata al centesimo.
- Il residuo di arrotondamento (max qualche centesimo) viene assegnato con il metodo del
  "resto maggiore" così che la somma delle quote sia ESATTAMENTE uguale all'importo.
- Importi negativi (note di credito, rimborsi) sono ripartiti con lo stesso criterio, con segno.
- UNITA:<ID>: tutto a quell'unità.
- MANUALE: quote lette dal foglio "Riparti manuali" (devono sommare all'importo).
- Si rifiuta di partire se una tabella usata non somma a 1000 (±0,01) o se una spesa non ha Tabella.
- Proprietario/conduttore: per le unità con `Conduttore` in Anagrafica, la quota di ogni spesa è divisa
  secondo `Quota conduttore %` della spesa (vuoto = 0 = tutto al proprietario). Verso il condominio il
  dovuto resta del proprietario; la parte del conduttore è informativa (foglio "Conduttori").

Produce un xlsx in prospetti/ (nome con suffisso _bozza; non sovrascrive mai un file esistente) con
fogli Riepilogo, Dettaglio, Millesimi usati — solo valori, nessuna formula — e stampa un riepilogo JSON.

Dopo l'approvazione dell'assemblea:

  riparto.py approva --dir "<cartella>" --prospetto prospetti/<file>_bozza.xlsx --scadenze 2026-10-31,2027-01-31
                     --approvato "<nome>" [--esercizio 2026] [--verbale "<riferimento>"] [--sostituisce prospetti/<file>.xlsx]

scrive le rate nel foglio Rate del registro riservato (Dovuto e Scaduto si ricalcolano da lì), crea le
righe mancanti di Situazione, aggiunge le scadenze delle rate al foglio Scadenze (senza importi per
unità), annota il Diario e, per ultimo, toglie _bozza dal nome del prospetto. Rifiuta di approvare due
volte lo stesso prospetto.
"""
import argparse
import datetime as dt
import json
import os
import re
import sys
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl non installato. Installarlo con: python3 -m pip install openpyxl (Windows: py -m pip install openpyxl; Linux con errore 'externally-managed-environment': sudo apt install python3-openpyxl)")

BASE_SCRIPTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "condominio-base", "scripts")

FONT = "Arial"
CENT = Decimal("0.01")
FMT_EURO = "#,##0.00 €"
HDR_FILL = PatternFill("solid", fgColor="DDE7F0")


def D(x):
    if x is None or x == "":
        return Decimal(0)
    return Decimal(str(x).replace(",", "."))


def blank(v):
    return v is None or (isinstance(v, str) and not v.strip())


def headers(ws):
    out = []
    for c in ws[1]:
        if blank(c.value):
            break
        out.append(c.value)
    return out


def rows(ws):
    hdr = headers(ws)
    out = []
    for r in ws.iter_rows(min_row=2, max_col=len(hdr), values_only=True):
        if all(blank(v) for v in r):
            continue
        out.append(dict(zip(hdr, r)))
    return out


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str) and v:
        return dt.date.fromisoformat(v[:10])
    return None


def ripartisci(importo, pesi):
    """pesi: dict id->Decimal (millesimi o quote). Ritorna dict id->Decimal, somma esatta = importo."""
    importo = importo.quantize(CENT)
    if importo < 0:
        return {u: -q for u, q in ripartisci(-importo, pesi).items()}
    tot = sum(pesi.values())
    if tot == 0:
        raise ValueError("somma dei pesi nulla")
    grezze = {u: importo * p / tot for u, p in pesi.items()}
    troncate = {u: q.quantize(CENT, rounding=ROUND_DOWN) for u, q in grezze.items()}
    residuo = (importo - sum(troncate.values())).quantize(CENT)
    n_cent = int((residuo / CENT).to_integral_value())
    # assegna i centesimi mancanti alle unità con resto decimale maggiore (metodo di Hamilton)
    ordine = sorted(pesi, key=lambda u: (grezze[u] - troncate[u], pesi[u]), reverse=True)
    for i in range(n_cent):
        troncate[ordine[i % len(ordine)]] += CENT
    assert sum(troncate.values()) == importo, "riparto non quadra"
    return troncate


def slug(s, n=40):
    s = re.sub(r"[^a-z0-9]+", "-", s.lower().replace("à", "a").replace("è", "e").replace("é", "e").replace("ì", "i").replace("ò", "o").replace("ù", "u"))
    return s.strip("-")[:n].strip("-")


def nome_libero(path):
    """Non sovrascrive mai: aggiunge _2, _3… se il file esiste."""
    base, ext = os.path.splitext(path)
    k, cand = 1, path
    while os.path.exists(cand):
        k += 1
        cand = f"{base}_{k}{ext}"
    return cand


def fmt_row(ws, r, bold=False, euro_from=None, skip=()):
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(r, c)
        cell.font = Font(name=FONT, bold=bold)
        if euro_from and c >= euro_from and c not in skip and isinstance(cell.value, (int, float)):
            cell.number_format = FMT_EURO


def calcola(a):
    reg = os.path.join(a.dir, "registro-condominio.xlsx")
    if not os.path.exists(reg):
        sys.exit("registro-condominio.xlsx non trovato: eseguire condominio-setup")
    wb = openpyxl.load_workbook(reg)
    cond = {r[0].value: r[1].value for r in wb["Condominio"].iter_rows(min_row=1, max_col=2) if r[0].value}
    esercizio = a.esercizio or int(cond.get("Esercizio corrente") or dt.date.today().year)

    anag = {str(r["ID unità"]).strip(): r for r in rows(wb["Anagrafica"]) if not blank(r.get("ID unità"))}
    mill_rows = [r for r in rows(wb["Millesimi"]) if not blank(r.get("ID unità")) and str(r["ID unità"]).strip().upper() != "TOTALE"]
    tabelle = [k for k in headers(wb["Millesimi"]) if isinstance(k, str) and k.startswith("Tab ")]
    mill = {t: {str(r["ID unità"]).strip(): D(r.get(t)) for r in mill_rows} for t in tabelle}
    unita = [str(r["ID unità"]).strip() for r in mill_rows]
    if not unita:
        sys.exit("Nessuna unità nel foglio Millesimi: completare il setup")
    conduttori = {u: str(anag[u]["Conduttore"]).strip() for u in unita if u in anag and not blank(anag[u].get("Conduttore"))}

    errori = []
    for t, pesi in mill.items():
        s = sum(pesi.values())
        if abs(s - 1000) > CENT and s != 0:
            errori.append(f"{t} somma a {s}, non a 1000")
    manuali = {}
    if "Riparti manuali" in wb.sheetnames:
        for r in rows(wb["Riparti manuali"]):
            if not blank(r.get("ID spesa")):
                manuali.setdefault(int(D(r["ID spesa"])), {})[str(r.get("ID unità")).strip()] = D(r.get("Quota"))

    # selezione spese
    spese = []
    ids = set(int(x) for x in a.ids.split(",")) if a.ids else None
    dal = dt.date.fromisoformat(a.dal) if a.dal else None
    al = dt.date.fromisoformat(a.al) if a.al else None
    for r in rows(wb["Spese"]):
        if blank(r.get("ID")) or blank(r.get("Importo")):
            continue
        if ids is not None and int(D(r["ID"])) not in ids:
            continue
        if ids is None and not blank(r.get("Esercizio")) and int(D(r["Esercizio"])) != esercizio:
            continue
        d = as_date(r.get("Data"))
        if dal and d and d < dal:
            continue
        if al and d and d > al:
            continue
        spese.append(r)
    if not spese:
        sys.exit("Nessuna spesa selezionata")

    # calcolo
    dettaglio = []  # (spesa, chiave tabella normalizzata, {unità: quota}, % conduttore, {unità: quota conduttore})
    for s in spese:
        importo = D(s["Importo"]).quantize(CENT)
        raw = str(s.get("Tabella") or "").strip()
        sid = int(D(s["ID"]))
        if not raw:
            errori.append(f"Spesa {sid} ({s.get('Descrizione')}): manca la Tabella (A, B, C…, UNITA:<ID> o MANUALE)"); continue
        up = raw.upper()
        if up.startswith("UNITA:") or up.startswith("UNITÀ:"):
            uid = raw.split(":", 1)[1].strip()
            if uid not in anag:
                errori.append(f"Spesa {sid}: unità {uid} inesistente"); continue
            chiave = f"UNITA:{uid}"
            quote = {u: (importo if u == uid else Decimal(0)) for u in unita}
        elif up == "MANUALE":
            q = manuali.get(sid)
            if not q:
                errori.append(f"Spesa {sid}: riparto MANUALE senza righe in 'Riparti manuali'"); continue
            if sum(q.values()).quantize(CENT) != importo:
                errori.append(f"Spesa {sid}: quote manuali sommano a {sum(q.values())}, importo {importo}"); continue
            chiave = "MANUALE"
            quote = {u: q.get(u, Decimal(0)).quantize(CENT) for u in unita}
        else:
            chiave = up
            key = f"Tab {up}"
            if key not in mill:
                errori.append(f"Spesa {sid}: tabella {raw} inesistente (disponibili: {', '.join(t[4:] for t in tabelle)})"); continue
            try:
                quote = ripartisci(importo, mill[key])
            except ValueError as e:
                errori.append(f"Spesa {sid}: {e}"); continue
        pct = D(s.get("Quota conduttore %")) if not blank(s.get("Quota conduttore %")) else Decimal(0)
        if not 0 <= pct <= 100:
            errori.append(f"Spesa {sid}: 'Quota conduttore %' = {pct} fuori da 0-100"); continue
        cond_q = {u: ((quote[u] * pct / 100).quantize(CENT, rounding=ROUND_HALF_UP) if u in conduttori else Decimal(0)) for u in unita}
        dettaglio.append((s, chiave, quote, pct, cond_q))
    if errori:
        print(json.dumps({"ok": False, "errori": errori}, ensure_ascii=False, indent=2))
        sys.exit(2)

    totali = {u: sum(q[u] for _, _, q, _, _ in dettaglio) for u in unita}
    tot_cond = {u: sum(cq[u] for _, _, _, _, cq in dettaglio) for u in unita}
    tot_prop = {u: totali[u] - tot_cond[u] for u in unita}
    per_tab = {}
    for _, chiave, q, _, _ in dettaglio:
        for u in unita:
            per_tab.setdefault(chiave, {}).setdefault(u, Decimal(0))
            per_tab[chiave][u] += q[u]
    per_tab = dict(sorted(per_tab.items()))
    totale = sum(D(s["Importo"]) for s, _, _, _, _ in dettaglio).quantize(CENT)
    assert sum(totali.values()) == totale
    rate = {u: ripartisci(totali[u], {i: Decimal(1) for i in range(a.rate)}) for u in unita} if a.rate else {}

    titolo = a.titolo or f"Riparto spese esercizio {esercizio}"
    oggi = dt.date.today().isoformat()
    nome = slug(titolo)
    nome = nome[len("riparto-"):] if nome.startswith("riparto-") else nome
    out = a.out or os.path.join("prospetti", f"{oggi}_riparto_{nome or f'esercizio-{esercizio}'}_bozza.xlsx")
    out_path = nome_libero(os.path.join(a.dir, out))
    out = os.path.relpath(out_path, a.dir)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    # --- xlsx (solo valori: leggibile anche dall'anteprima di Google Drive su cellulare)
    xw = Workbook()
    ws = xw.active; ws.title = "Riepilogo"
    ws["A1"] = f"{cond.get('Nome', '')} — {titolo}"; ws["A1"].font = Font(name=FONT, bold=True, size=13)
    ws["A2"] = f"Generato il {oggi} · BOZZA: valida solo dopo approvazione dell'assemblea"; ws["A2"].font = Font(name=FONT, italic=True)
    extra = ["A carico proprietà", "A carico conduttore"] if conduttori else []
    hdr = ["ID unità", "Interno", "Intestatario"] + [f"Quota {t}" for t in per_tab] + ["Totale dovuto"] + extra + [f"Rata {i + 1}" for i in range(a.rate)]
    ws.append([]); ws.append(hdr)
    for c in ws[4]:
        c.font = Font(name=FONT, bold=True); c.fill = HDR_FILL
    for u in unita:
        row = [u, anag.get(u, {}).get("Interno", ""), anag.get(u, {}).get("Intestatario", "")]
        row += [float(per_tab[t][u]) for t in per_tab] + [float(totali[u])]
        if conduttori:
            row += [float(tot_prop[u]), float(tot_cond[u])]
        row += [float(rate[u][i]) for i in range(a.rate)]
        ws.append(row)
        fmt_row(ws, ws.max_row, euro_from=4)
    tot_row = ["TOTALE", "", ""] + [float(sum(v.values())) for v in per_tab.values()] + [float(totale)]
    if conduttori:
        tot_row += [float(sum(tot_prop.values())), float(sum(tot_cond.values()))]
    tot_row += [float(sum(rate[u][i] for u in unita)) for i in range(a.rate)]
    ws.append(tot_row)
    fmt_row(ws, ws.max_row, bold=True, euro_from=4)
    for i, h in enumerate(hdr, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(h) + 2)
    ws.column_dimensions["C"].width = 30
    ws.freeze_panes = "D5"

    wd = xw.create_sheet("Dettaglio")
    wd.append(["ID spesa", "Data", "Fornitore", "Descrizione", "Importo", "Tabella"] + unita)
    for c in wd[1]:
        c.font = Font(name=FONT, bold=True); c.fill = HDR_FILL
    for s, chiave, q, _, _ in dettaglio:
        wd.append([int(D(s["ID"])), as_date(s.get("Data")), s.get("Fornitore"), s.get("Descrizione"), float(D(s["Importo"])), chiave] + [float(q[u]) for u in unita])
        fmt_row(wd, wd.max_row, euro_from=5)
        wd.cell(wd.max_row, 2).number_format = "yyyy-mm-dd"
    wd.append(["TOTALE", "", "", "", float(totale), ""] + [float(totali[u]) for u in unita])
    fmt_row(wd, wd.max_row, bold=True, euro_from=5)
    wd.column_dimensions["D"].width = 36; wd.column_dimensions["C"].width = 22
    wd.freeze_panes = "G2"

    if conduttori:  # riparto interno proprietario/conduttore, una sezione per unità affittata
        wc = xw.create_sheet("Conduttori")
        wc.append(["ID unità", "Interno", "Intestatario", "Conduttore", "ID spesa", "Data", "Descrizione", "Tabella",
                   "Quota unità", "% conduttore", "A carico conduttore", "A carico proprietà"])
        for c in wc[1]:
            c.font = Font(name=FONT, bold=True); c.fill = HDR_FILL
        for u in unita:
            if u not in conduttori:
                continue
            interno = anag[u].get("Interno", ""); intest = anag[u].get("Intestatario", "")
            for s, chiave, q, pct, cq in dettaglio:
                if q[u] == 0:
                    continue
                wc.append([u, interno, intest, conduttori[u], int(D(s["ID"])), as_date(s.get("Data")), s.get("Descrizione"), chiave,
                           float(q[u]), float(pct), float(cq[u]), float(q[u] - cq[u])])
                fmt_row(wc, wc.max_row, euro_from=9, skip=(10,))
                wc.cell(wc.max_row, 6).number_format = "yyyy-mm-dd"
            wc.append([u, interno, intest, conduttori[u], "", None, "TOTALE", "", float(totali[u]), "", float(tot_cond[u]), float(tot_prop[u])])
            fmt_row(wc, wc.max_row, bold=True, euro_from=9, skip=(10,))
        for col, w in {"C": 26, "D": 26, "G": 36, "K": 20, "L": 20}.items():
            wc.column_dimensions[col].width = w
        wc.freeze_panes = "E2"

    wm = xw.create_sheet("Millesimi usati")
    wm.append(["ID unità"] + tabelle)
    for c in wm[1]:
        c.font = Font(name=FONT, bold=True); c.fill = HDR_FILL
    for u in unita:
        wm.append([u] + [float(mill[t][u]) for t in tabelle]); fmt_row(wm, wm.max_row)
    wm.append(["TOTALE"] + [float(sum(mill[t].values())) for t in tabelle]); fmt_row(wm, wm.max_row, bold=True)
    xw.save(out_path)

    print(json.dumps({
        "ok": True, "titolo": titolo, "esercizio": esercizio, "prospetto": out,
        "spese": len(dettaglio), "totale": float(totale), "rate": a.rate,
        "per_unita": {u: {"interno": anag.get(u, {}).get("Interno"), "intestatario": anag.get(u, {}).get("Intestatario"),
                          "totale": float(totali[u]), "rate": [float(rate[u][i]) for i in range(a.rate)],
                          "conduttore": conduttori.get(u),
                          "a_carico_proprieta": float(tot_prop[u]), "a_carico_conduttore": float(tot_cond[u])} for u in unita},
        "unita_con_conduttore": len(conduttori),
        "per_tabella": {t: float(sum(v.values())) for t, v in per_tab.items()},
        "a_carico_singola_unita": {t: float(sum(v.values())) for t, v in per_tab.items() if t.startswith("UNITA:")},
    }, ensure_ascii=False, indent=2, default=str))


def _registro():
    """Importa registro.py dalla skill condominio-base (stessa installazione del plugin)."""
    sys.path.insert(0, os.path.abspath(BASE_SCRIPTS))
    try:
        import registro
    except ImportError:
        sys.exit(f"registro.py non trovato in {os.path.abspath(BASE_SCRIPTS)}: il plugin è installato in modo incompleto")
    return registro


def _fallisci(errori):
    print(json.dumps({"ok": False, "errori": errori}, ensure_ascii=False, indent=2))
    sys.exit(2)


def leggi_riepilogo(path):
    """Righe del foglio Riepilogo di un prospetto: (titolo, [{"ID unità", "Intestatario", "Totale dovuto", "rate": [...]}])."""
    ws = openpyxl.load_workbook(path)["Riepilogo"]
    titolo = str(ws["A1"].value or "")
    titolo = titolo.split(" — ", 1)[1] if " — " in titolo else titolo
    hdr = [c.value for c in ws[4]]
    if "ID unità" not in hdr or "Totale dovuto" not in hdr:
        raise ValueError("il foglio Riepilogo non ha le colonne 'ID unità' e 'Totale dovuto' in riga 4")
    col_rate = [i for i, h in enumerate(hdr) if isinstance(h, str) and re.fullmatch(r"Rata \d+", h)]
    righe = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        uid = r[hdr.index("ID unità")]
        if blank(uid) or str(uid).strip().upper() == "TOTALE":
            break
        righe.append({"ID unità": str(uid).strip(), "Intestatario": r[hdr.index("Intestatario")] if "Intestatario" in hdr else None,
                      "Totale dovuto": D(r[hdr.index("Totale dovuto")]), "rate": [D(r[i]) for i in col_rate]})
    return titolo, righe


def approva(a):
    reg = _registro()
    errori = []
    path_c = os.path.join(a.dir, "registro-condominio.xlsx")
    if not os.path.exists(path_c):
        sys.exit("registro-condominio.xlsx non trovato: eseguire condominio-setup")
    wb_c = openpyxl.load_workbook(path_c)
    path_r = reg.percorso_riservato(a.dir, wb_c)
    if not os.path.exists(path_r):
        sys.exit(f"Registro riservato non trovato: {path_r} (controllare 'Percorso registro riservato')")
    wb_r = openpyxl.load_workbook(path_r)
    if "Rate" not in wb_r.sheetnames:
        _fallisci(["Il registro riservato non ha il foglio Rate: eseguire `registro.py --file registro-riservato.xlsx schema`"])
    esercizio = a.esercizio or int(reg._kv(wb_c["Condominio"]).get("Esercizio corrente") or dt.date.today().year)

    # --- controlli: nessuna scrittura finché qualcosa non torna
    bozza = os.path.join(a.dir, a.prospetto)
    if not bozza.endswith("_bozza.xlsx"):
        _fallisci([f"{a.prospetto} non è una bozza (manca il suffisso _bozza): è già stato approvato?"])
    if not os.path.exists(bozza):
        gia = bozza[: -len("_bozza.xlsx")] + ".xlsx"
        _fallisci([f"Prospetto non trovato: {a.prospetto}" + (f" (esiste {os.path.relpath(gia, a.dir)}: già approvato?)" if os.path.exists(gia) else "")])
    # nome definitivo scelto prima dei controlli: se esiste già un prospetto approvato con lo stesso nome
    # (riparto rifatto lo stesso giorno), questo diventa _2, _3…
    definitivo = nome_libero(bozza[: -len("_bozza.xlsx")] + ".xlsx")
    nome_def = os.path.basename(definitivo)
    try:
        titolo, righe = leggi_riepilogo(bozza)
    except (KeyError, ValueError) as e:
        _fallisci([f"Prospetto illeggibile: {e}"])
    if not righe:
        _fallisci(["Il foglio Riepilogo del prospetto non contiene unità"])
    try:
        scadenze = [dt.date.fromisoformat(x.strip()) for x in a.scadenze.split(",") if x.strip()]
    except ValueError:
        _fallisci([f"Scadenze non valide: {a.scadenze} (formato AAAA-MM-GG separate da virgola)"])
    n_rate = len(righe[0]["rate"]) or 1
    if len(scadenze) != n_rate:
        errori.append(f"Il prospetto ha {n_rate} rat{'a' if n_rate == 1 else 'e'}, sono state indicate {len(scadenze)} scadenze")
    if scadenze != sorted(scadenze):
        errori.append("Le scadenze vanno indicate in ordine di data")
    rate_esistenti = reg._rows(wb_r["Rate"])
    if any(str(r.get("Prospetto") or "") == nome_def for r in rate_esistenti):
        errori.append(f"{nome_def} risulta già approvato: il foglio Rate ha già le sue rate")
    sostituito = os.path.basename(a.sostituisce) if a.sostituisce else None
    if sostituito and not any(str(r.get("Prospetto") or "") == sostituito for r in rate_esistenti):
        errori.append(f"--sostituisce {sostituito}: nessuna rata di quel prospetto nel foglio Rate")
    if errori:
        _fallisci(errori)

    # --- registro riservato: rate, righe di Situazione, dovuti 0.3 riportati
    ws_rate, ws_sit = wb_r["Rate"], wb_r["Situazione"]
    hdr_rate = reg._headers(ws_rate)
    sostituite = 0
    if sostituito:
        c_pro, c_val, c_note = (hdr_rate.index(h) + 1 for h in ("Prospetto", "Valida", "Note"))
        for r in range(2, ws_rate.max_row + 1):
            if str(ws_rate.cell(r, c_pro).value or "") == sostituito:
                reg._put(ws_rate, r, c_val, "Valida", "NO")
                reg._put(ws_rate, r, c_note, "Note", f"sostituita da {nome_def}")
                sostituite += 1
    sit = {(str(r.get("ID unità") or "").strip(), int(reg._num(r.get("Esercizio")) or 0)): r for r in reg._rows(ws_sit)}
    con_rate = {str(r.get("ID unità") or "").strip() for r in rate_esistenti if int(reg._num(r.get("Esercizio")) or 0) == esercizio}
    situazioni_create, riportati = [], {}
    for u in righe:
        uid = u["ID unità"]
        riga_sit = sit.get((uid, esercizio))
        if riga_sit is None:
            reg.aggiungi(ws_sit, {"ID unità": uid, "Intestatario": u["Intestatario"], "Esercizio": esercizio, "Livello sollecito": 0})
            situazioni_create.append(uid)
        elif uid not in con_rate and (reg._num(riga_sit.get("Dovuto")) or 0) != 0:
            # registro 0.3: il Dovuto scritto a mano diventa una rata senza scadenza, così non si perde
            dovuto = reg._num(riga_sit.get("Dovuto"))
            reg.aggiungi(ws_rate, {"ID unità": uid, "Esercizio": esercizio, "Prospetto": "Dovuto registrato prima della 0.4",
                                   "Rata": 0, "Importo": dovuto, "Valida": "SI",
                                   "Note": "riportato dal Dovuto di Situazione alla prima approvazione con le rate"})
            riportati[uid] = reg._plain(dovuto)
        importi = u["rate"] or [u["Totale dovuto"]]
        for n, (scad, imp) in enumerate(zip(scadenze, importi), start=1):
            reg.aggiungi(ws_rate, {"ID unità": uid, "Esercizio": esercizio, "Prospetto": nome_def, "Rata": n,
                                   "Scadenza": scad, "Importo": imp, "Valida": "SI"})

    # --- registro condiviso: scadenze delle rate (solo date) e Diario
    ws_scad = wb_c["Scadenze"]
    scadenze_create = []
    for n, scad in enumerate(scadenze, start=1):
        descr = f"Rata {n} di {n_rate} — {titolo}" if n_rate > 1 else f"Pagamento — {titolo}"
        _, sid = reg.aggiungi(ws_scad, {"Data": scad, "Tipo": "rata", "Descrizione": descr, "Ricorrenza": "nessuna",
                                        "Note": f"prospetto {nome_def}"})
        scadenze_create.append({"ID": sid, "Data": scad.isoformat(), "Descrizione": descr})
    totale = sum(u["Totale dovuto"] for u in righe)
    dettaglio = (f"{nome_def}: {len(righe)} unità, totale {totale:.2f} €, {n_rate} rat{'a' if n_rate == 1 else 'e'} "
                 f"({', '.join(d.isoformat() for d in scadenze)})")
    if a.verbale:
        dettaglio += f"; verbale: {a.verbale}"
    if sostituito:
        dettaglio += f"; sostituisce {sostituito}"
    reg.aggiungi(wb_c["Diario"], {"Data-ora": dt.datetime.now().isoformat(timespec="seconds"), "Operazione": "Approvato prospetto di riparto",
                                  "Dettaglio": dettaglio, "Eseguito da": "IA", "Approvato da": a.approvato})

    # --- salvataggi, poi la rinomina per ultima
    reg._save(wb_r, path_r)
    reg._save(wb_c, path_c)
    os.rename(bozza, definitivo)
    print(json.dumps({
        "ok": True, "prospetto": os.path.relpath(definitivo, a.dir), "esercizio": esercizio, "unita": len(righe),
        "totale": float(totale), "righe_rate": len(righe) * n_rate, "rate_sostituite": sostituite,
        "situazioni_create": situazioni_create, "dovuti_precedenti_riportati": riportati,
        "scadenze_create": scadenze_create,
        "prossimi_passi": ["scadenze: allineare il calendario per le nuove righe di Scadenze",
                           "comunicazioni: inviare il prospetto approvato ai condomini"],
    }, ensure_ascii=False, indent=2))


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "approva":
        p = argparse.ArgumentParser(prog="riparto.py approva", description="Registra l'approvazione di un prospetto di riparto.")
        p.add_argument("--dir", default=".")
        p.add_argument("--prospetto", required=True, help="percorso relativo del prospetto _bozza.xlsx")
        p.add_argument("--scadenze", required=True, help="date delle rate AAAA-MM-GG separate da virgola, una per rata")
        p.add_argument("--approvato", required=True, help="chi ha approvato (per il Diario)")
        p.add_argument("--esercizio", type=int)
        p.add_argument("--verbale")
        p.add_argument("--sostituisce", help="prospetto approvato che questo sostituisce (es. preventivo → consuntivo)")
        a = p.parse_args(sys.argv[2:])
        try:
            approva(a)
        except SystemExit:
            raise
        except Exception as e:  # mai un traceback all'utente
            _fallisci([f"{type(e).__name__}: {e}"])
        return
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--dir", default=".")
    p.add_argument("--esercizio", type=int)
    p.add_argument("--ids", help="ID spese separati da virgola")
    p.add_argument("--dal"); p.add_argument("--al")
    p.add_argument("--titolo")
    p.add_argument("--rate", type=int, default=0, help="numero di rate in cui suddividere il dovuto (0 = nessuna)")
    p.add_argument("--out")
    a = p.parse_args()
    try:
        calcola(a)
    except SystemExit:
        raise
    except Exception as e:  # mai un traceback all'utente: errore leggibile in JSON
        print(json.dumps({"ok": False, "errori": [f"{type(e).__name__}: {e}"]}, ensure_ascii=False, indent=2))
        sys.exit(2)


if __name__ == "__main__":
    main()
