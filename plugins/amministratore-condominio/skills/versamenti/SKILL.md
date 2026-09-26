---
name: versamenti
description: >
  Registra i pagamenti dei condomini, uno alla volta o da un estratto conto bancario in PDF,
  abbinandoli a unità e rate. Usare per "registra un versamento", "ha pagato", "bonifico",
  "estratto conto della banca", "chi ha pagato".
metadata:
  version: "0.4.0"
---

# Versamenti

Leggere prima `condominio-base`. Un versamento registrato sull'unità sbagliata è peggio di uno non
registrato: fa partire solleciti a chi ha pagato. Qui il modello **legge** (il PDF, la frase
dell'utente) e **presenta**; l'abbinamento lo propone uno script e lo conferma l'utente.

`abbina.py` è `${CLAUDE_SKILL_DIR}/scripts/abbina.py`; trova da solo il registro riservato
(`Percorso registro riservato`). Sempre con `--dir "<cartella del condominio>"`.

## Dove sta l'estratto conto

L'estratto conto contiene nomi, IBAN e causali di tutti: **non va nella cartella condivisa**. Il
posto è `<cartella riservata>/estratti conto/`, dove `<cartella riservata>` è quella di
`registro-riservato.xlsx`. Se l'utente lo ha messo in `da analizzare/` o altrove nella cartella
condivisa, proporre di spostarlo lì prima di tutto.

## Procedura

### 1. Movimenti

Scrivere un JSON con **solo i movimenti in entrata** (accrediti), uno per riga dell'estratto:

```json
[{"data": "2026-09-10", "importo": 150.25, "ordinante": "VERDI LUCA", "causale": "CONDOMINIO RATA 3 INT 3"}]
```

- **Da estratto conto PDF**: trascrivere data valuta o contabile (la stessa per tutto l'estratto),
  importo, ordinante e causale **così come sono scritti**, senza interpretare. Contare gli
  addebiti e i movimenti che non sono bonifici di condomini (interessi, giroconti) e dirlo in
  chiusura, senza trascriverli. Salvare il JSON accanto al PDF: `estratti conto/<nome pdf>.movimenti.json`.
- **Da una frase** ("registra 150 € dall'interno 3, bonifico del 10 settembre"): un solo
  movimento con `"unita": "U03"` (l'ID dell'unità indicata), senza ordinante. Se la data non è
  detta, chiederla: non usare oggi per default.

### 2. Proposta

```bash
python3 "${CLAUDE_SKILL_DIR}/scripts/abbina.py" --dir "<cartella>" --movimenti "<file.json>"
```

Non scrive nulla. Per ogni movimento l'esito è:
- `proposto` — unità, rata coperta (`copre`), motivi (nome, interno nella causale, importo
  uguale alla rata) e la riga `versamento` pronta;
- `da_abbinare` — `motivo` (nessun indizio, indizi deboli, più unità possibili) e fino a 3 `candidati`;
- `gia_registrato` — c'è già in `Versamenti` (stesso movimento: data, importo, ordinante);
- `ignorato` — addebito o importo nullo.

Mostrare una tabella unica, come in `archivio`:

```
1. 10/09 · 150,25 € · VERDI LUCA · "RATA 3 INT 3" → U03 Luca Verdi, rata 3   (nome, interno, importo)
2. 12/09 · 99,00 € · ROSSI · "quota"            → DA ABBINARE: U01 Mario Rossi? U05 Mario Rossi?
3. 15/09 · 150,25 € · VERDI LUCA                 → già registrato
```

Per ogni `da_abbinare` chiedere a chi attribuirlo (o se ignorarlo: può essere un rimborso, un
versamento di un fornitore). L'utente corregge per numero ("la 2 è U05"). Per le righe corrette
a mano, ricostruire `versamento` con i campi della proposta e `ID unità` cambiato; la `Rata` si
può lasciare vuota: la situazione imputa comunque i versamenti alle rate in ordine di scadenza.

### 3. Registrazione (solo dopo conferma)

Scrivere in un JSON l'elenco delle righe `versamento` confermate, poi:

```bash
python3 "${CLAUDE_SKILL_DIR}/scripts/abbina.py" registra --dir "<cartella>" --versamenti "<confermati.json>" --approvato "<nome>"
```

Lo script ricontrolla tutto prima di scrivere (unità esistenti, movimenti non già registrati),
crea le righe di `Situazione` mancanti e scrive nel Diario **solo il numero** dei versamenti. Se
esce con `"ok": false`, nulla è stato scritto: mostrare gli errori e correggere.

### 4. Chiusura

Riepilogo in tre righe: registrati, rimasti da abbinare, già presenti (e addebiti ignorati se
l'estratto era completo). Poi, da `registro.py --file registro-riservato.xlsx situazione`, le
unità con `Scaduto` > 0: "3 unità hanno rate scadute non pagate". Proporre la skill
`comunicazioni` per un sollecito o per l'invio della situazione personale; non sollecitare da qui.

## Regole

- Mai registrare un versamento senza conferma, nemmeno se l'esito è `proposto`.
- Mai scegliere tra due candidati: chiedere.
- Il conduttore che paga paga per l'unità: il versamento va all'unità (il debitore resta il
  proprietario, vedi `condominio-base`).
- Nel Diario, che è pubblico, mai nomi, unità o importi dei versamenti: ci pensa lo script. Se
  l'utente chiede di annotare qualcosa a mano, usare la colonna `Note` di `Versamenti`.
- Un versamento attribuito all'unità sbagliata non si cancella e non si compensa con un importo
  negativo (`registra` accetta solo importi positivi): si corregge la riga con
  `registro.py --file registro-riservato.xlsx update Versamenti --where "ID=<n>" '{"ID unità": "U05", "Note": "corretto il <data>: era U01"}'`.
- Con profilo `autogestione`, spiegare la prima volta perché un bonifico resta "da abbinare":
  meglio chiedere che attribuirlo alla persona sbagliata.
