# Modello dati

Due cartelle di lavoro Excel (`.xlsx`, modificabili anche da Google Sheets e LibreOffice).
Riga 1 = intestazioni, esattamente come sotto. Le colonne possono essere aggiunte in coda,
mai rinominate o spostate: gli script le cercano per nome e si fermano alla prima cella vuota
della riga 1.

**Nessuna formula.** I condomini consultano i file dall'app Google Drive, la cui anteprima mostra
solo i valori salvati (openpyxl non salva il risultato delle formule). I valori derivati (riga
`TOTALE` di `Millesimi`, `Versato` e `Saldo` di `Situazione`) sono numeri che `registro.py`
ricalcola e riscrive a ogni operazione (`append`, `update`, `verifica`).

Chiave di collegamento tra tutti i fogli: **`ID unità`** (testo breve, es. `U01`, `U02`…
oppure `A-1`, `B-3`). Stabilita al setup, non cambia mai.

---

## `registro-condominio.xlsx` — condiviso con tutti

### Foglio `Condominio` (chiave / valore)

Colonna A = chiave, B = valore (celle gialle), C = nota per chi compila a mano (non letta dagli script).

| Chiave | Valore atteso | Note |
|---|---|---|
| Nome | testo | es. "Condominio Le Betulle" |
| Indirizzo | testo | |
| Codice fiscale | testo | del condominio (obbligatorio se ha un amministratore, per fatture e conto corrente) |
| Profilo | `autogestione` \| `professionista` | orienta tono e spiegazioni |
| Amministratore | testo | nome e cognome |
| Email amministratore | email | mittente delle comunicazioni |
| Modalità collaudo | `SI` \| `NO` | SI = tutte le email vanno solo a `Email collaudo` |
| Email collaudo | email | di solito quella dell'amministratore |
| Esercizio corrente | anno (es. 2026) | esercizio di gestione in corso |
| Inizio esercizio | data ISO | es. 2026-01-01 |
| Fine esercizio | data ISO | es. 2026-12-31 |
| Numero unità | intero | coerente con `Anagrafica` |
| Tabelle millesimali | elenco, es. `A: proprietà; B: scale; C: ascensore` | descrizione delle colonne di `Millesimi` |
| Conto corrente | IBAN | conto intestato al condominio (art. 1129 c.c.) |
| Percorso registro riservato | percorso cartella | vuoto = stessa cartella; altrimenti cartella privata dove sta `registro-riservato.xlsx` |
| Ultima elaborazione | data-ora ISO | aggiornata dagli script |

### Foglio `Anagrafica`

Corrisponde all'anagrafe condominiale dell'art. 1130 n. 6 c.c.

| Colonna | Tipo | Note |
|---|---|---|
| ID unità | testo | chiave |
| Interno | testo | es. "3", "B" |
| Piano | intero | 0 = terra; usato per art. 1124 se serve |
| Intestatario | testo | proprietario (o comproprietari separati da `;`) |
| Email | email | destinatario delle comunicazioni; vuoto = solo cartaceo |
| Telefono | testo | |
| Conduttore | testo | inquilino, se presente |
| Dati catastali | testo | foglio/particella/subalterno |
| Note | testo | |

### Foglio `Millesimi`

| Colonna | Tipo | Note |
|---|---|---|
| ID unità | testo | chiave |
| Intestatario | testo | copia per leggibilità |
| Tab A | numero | proprietà generale (art. 1123 c.c.) |
| Tab B | numero | scale (art. 1124) — già calcolata dal tecnico |
| Tab C | numero | ascensore (art. 1124) |
| Tab D … | numero | altre tabelle (riscaldamento, cortile…) aggiunte in coda |

Vincolo: ogni colonna `Tab X` somma a 1000 (tolleranza ±0,01). L'ultima riga del foglio,
`TOTALE`, contiene le somme come numeri, ricalcolate da `registro.py` (`append`, `update`,
`verifica`); `riparto.py` ricalcola per conto suo e rifiuta di partire se una tabella non quadra.
Una tabella tutta a zero è "non usata" (es. niente ascensore) e non blocca. Chi non usa una
tabella per un'unità mette 0, non vuoto. Le unità vanno inserite sopra la riga `TOTALE`
(`append` lo fa da solo).

### Foglio `Spese`

| Colonna | Tipo | Note |
|---|---|---|
| ID | intero progressivo | assegnato da `registro.py append` |
| Data | data ISO | data del documento |
| Fornitore | testo | |
| Descrizione | testo | breve, in italiano |
| Importo | numero (2 dec.) | lordo, come da documento; negativo per note di credito e rimborsi |
| Tabella | `A`,`B`,`C`… \| `UNITA:U03` \| `MANUALE` | criterio di riparto, **obbligatorio** per il riparto; `UNITA:` = a carico di una sola unità; `MANUALE` = riparto indicato nel foglio `Riparti manuali` |
| Esercizio | anno | |
| File | percorso relativo | dentro `archivio/` |
| Pagata | `SI` \| `NO` | |
| Data pagamento | data ISO | |
| Note | testo | |

### Foglio `Riparti manuali` (opzionale)

Per spese con `Tabella = MANUALE`: una riga per unità.

| ID spesa | ID unità | Quota | Note |

### Foglio `Scadenze`

| Colonna | Tipo | Note |
|---|---|---|
| ID | intero progressivo | |
| Data | data ISO | |
| Ora | HH:MM | opzionale |
| Tipo | `assemblea` \| `rata` \| `manutenzione` \| `contratto` \| `adempimento` \| `altro` | |
| Descrizione | testo | |
| Ricorrenza | `nessuna` \| `mensile` \| `trimestrale` \| `annuale` \| `biennale` | |
| ID evento | testo | ID dell'evento Google Calendar, valorizzato da `scadenze` |
| Note | testo | |

### Foglio `Diario`

Solo in append. Mai cancellare righe.

| Data-ora | Operazione | Dettaglio | Eseguito da | Approvato da |
|---|---|---|---|---|
| ISO | breve | libero | `IA` \| nome | nome (vuoto se non serviva approvazione) |

### Foglio `Indice`

Traccia i file già elaborati da `archivio` (idempotenza).

| Hash | Nome originale | Nome archivio | Percorso | Data elaborazione | ID spesa |
|---|---|---|---|---|---|

`Hash` = SHA-256 del contenuto. Se un file in `da analizzare/` ha un hash già presente,
è un duplicato: spostarlo in `archivio/<anno>/altro/duplicati/` e annotare nel Diario.

---

## `registro-riservato.xlsx` — solo amministratore

### Foglio `Situazione`

| Colonna | Tipo | Note |
|---|---|---|
| ID unità | testo | chiave |
| Intestatario | testo | |
| Esercizio | anno | |
| Dovuto | numero | somma delle quote dell'esercizio (da ultimo prospetto approvato) |
| Versato | numero | **calcolato** da `registro.py`: somma dei `Versamenti` dell'unità per quell'esercizio (colonna `Esercizio` del versamento, altrimenti anno della `Data`) |
| Saldo | numero | **calcolato**: `Dovuto − Versato` (positivo = deve ancora) |
| Ultimo sollecito | data ISO | |
| Livello sollecito | 0,1,2,3 | 0 nessuno, 1 cortese, 2 formale, 3 diffida |
| Note | testo | |

### Foglio `Versamenti`

| ID | Data | ID unità | Importo | Esercizio | Riferimento | Rata | Note |

`Esercizio` collega il versamento alla riga giusta di `Situazione`; se vuoto vale l'anno della `Data`.

### Foglio `Solleciti`

| ID | Data | ID unità | Livello | Inviato a | Canale | Esito | Note |

---

## Regole di scrittura

- Importi: numeri, non testo; due decimali; separatore decimale gestito da Excel.
- Date: ISO `AAAA-MM-GG` in tutti i fogli.
- Non lasciare celle "quasi vuote" (spazi). Vuoto = vuoto.
- Mai formule: solo valori. `registro.py` applica formati (`#,##0.00 €`, date ISO) e font.
- `registro.py` aggiorna `Ultima elaborazione` a ogni scrittura.

## Prospetti (`prospetti/`)

Generati da `riparto.py`: `<data>_riparto_<titolo>_bozza.xlsx`, fogli `Riepilogo` (una riga per
unità: quote per tabella, totale, rate), `Dettaglio` (una riga per spesa) e `Millesimi usati`.
Solo valori. Mai sovrascritti (`_2`, `_3`…). Dopo l'approvazione dell'assemblea si toglie `_bozza`.
