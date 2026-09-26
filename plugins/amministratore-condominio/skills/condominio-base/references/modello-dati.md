# Modello dati

Due cartelle di lavoro Excel (`.xlsx`, modificabili anche da Google Sheets e LibreOffice).
Riga 1 = intestazioni, esattamente come sotto. Le colonne possono essere aggiunte in coda,
mai rinominate o spostate: gli script le cercano per nome e si fermano alla prima cella vuota
della riga 1.

**Nessuna formula.** I condomini consultano i file dall'app Google Drive, la cui anteprima mostra
solo i valori salvati (openpyxl non salva il risultato delle formule). I valori derivati (riga
`TOTALE` di `Millesimi`, `Dovuto`, `Versato`, `Saldo` e `Scaduto` di `Situazione`) sono numeri che
`registro.py` ricalcola e riscrive a ogni scrittura (`append`, `update`, `set`, `diario`,
`verifica`, `schema`).

Chiave di collegamento tra tutti i fogli: **`ID unità`** (testo breve, es. `U01`, `U02`…
oppure `A-1`, `B-3`). Stabilita al setup, non cambia mai.

---

## `registro-condominio.xlsx` — condiviso con tutti

### Foglio `In breve` (generato, primo foglio)

Riepilogo per i condomini, pensato per l'anteprima di Google Drive su telefono: nome del
condominio, data di aggiornamento, prossima assemblea e prossima rata, altre scadenze (massimo 4),
spese registrate nell'esercizio corrente con il totale per tabella (etichette da `Tabelle
millesimali`), ultimi 5 documenti archiviati con una descrizione ricavata dal nome del file, numero
di documenti in attesa in `da analizzare/`.

`registro.py` lo **riscrive da zero a ogni salvataggio** e lo tiene in prima posizione e attivo:
non si compila e non si legge con `read`/`append`/`update`. Non contiene dati per unità (le spese
`UNITA:` sono sommate insieme come "a carico di singole unità"). `registro.py schema` lo crea nei
registri di versioni precedenti.

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
| Conduttore | testo | inquilino, se presente: attiva il riparto proprietario/conduttore nel prospetto |
| Email conduttore | email | per la nota informativa al conduttore (mai solleciti) |
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
| Quota conduttore % | numero 0–100 | parte della quota che, dentro un'unità affittata, spetta all'inquilino (art. 9 L. 392/1978); vuoto = 0 |
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
| Dovuto | numero | **calcolato** se l'unità ha righe in `Rate` per l'esercizio: somma delle rate valide. Altrimenti (registri 0.3) è il valore scritto |
| Versato | numero | **calcolato** da `registro.py`: somma dei `Versamenti` dell'unità per quell'esercizio (colonna `Esercizio` del versamento, altrimenti anno della `Data`) |
| Saldo | numero | **calcolato**: `Dovuto − Versato` (positivo = deve ancora, negativo = credito) |
| Ultimo sollecito | data ISO | |
| Livello sollecito | 0,1,2,3 | 0 nessuno, 1 cortese, 2 formale, 3 diffida |
| Note | testo | |
| Scaduto | numero | **calcolato**: rate valide con `Scadenza` passata (dal giorno dopo) non coperte dai versamenti. Vuoto se l'unità non ha rate: senza scadenze il ritardo non si conosce |

### Foglio `Versamenti`

| ID | Data | ID unità | Importo | Esercizio | Riferimento | Rata | Note | ID movimento |

`Esercizio` collega il versamento alla riga giusta di `Situazione`; se vuoto vale l'anno della `Data`.
`ID movimento` è l'impronta del movimento bancario da cui nasce il versamento (skill `versamenti`):
se è già presente, il movimento è già stato registrato. Vuoto per i versamenti inseriti a mano.

### Foglio `Solleciti`

| ID | Data | ID unità | Livello | Inviato a | Canale | Esito | Note |

### Foglio `Rate`

Una riga per unità e rata di ogni prospetto approvato. Le scrive `riparto.py approva`; non si
cancellano mai.

| Colonna | Tipo | Note |
|---|---|---|
| ID unità | testo | |
| Esercizio | anno | |
| Prospetto | testo | nome del file in `prospetti/` (senza `_bozza`) |
| Rata | intero | 1…N nel prospetto |
| Scadenza | data ISO | |
| Importo | numero | |
| Valida | `SI` \| `NO` | vuoto = `SI`; `NO` = prospetto sostituito (es. consuntivo al posto del preventivo) |
| Note | testo | |

I versamenti dell'unità nell'esercizio si imputano alle rate valide in ordine di scadenza. Stato
di una rata: `pagata` (coperta per intero), `scaduta` (non coperta e scadenza passata), `parziale`
(coperta in parte, non ancora scaduta), `da pagare`. `registro.py --file registro-riservato.xlsx
situazione` restituisce per ogni unità dovuto, versato, saldo, scaduto, rate con stato e versamenti.

---

## Regole di scrittura

- Importi: numeri, non testo; due decimali; separatore decimale gestito da Excel.
- Date: ISO `AAAA-MM-GG` in tutti i fogli.
- Non lasciare celle "quasi vuote" (spazi). Vuoto = vuoto.
- Mai formule: solo valori. `registro.py` applica formati (`#,##0.00 €`, date ISO) e font.
- `registro.py` aggiorna `Ultima elaborazione` a ogni scrittura.

## Prospetti (`prospetti/`)

Generati da `riparto.py`: `<data>_riparto_<titolo>_bozza.xlsx`, fogli `Riepilogo` (una riga per
unità: quote per tabella, totale, eventuali colonne "A carico proprietà" e "A carico conduttore",
rate), `Dettaglio` (una riga per spesa), `Conduttori` (solo se ci sono unità affittate: per ogni
unità e spesa, quota, percentuale, parte del conduttore e della proprietà) e `Millesimi usati`.
Solo valori. Mai sovrascritti (`_2`, `_3`…). Dopo l'approvazione dell'assemblea si toglie `_bozza`.

## Riparto proprietario / conduttore

I millesimi non si dividono tra proprietario e inquilino. Per ogni spesa, la quota dell'unità è
`importo × millesimi / 1000`; se l'unità ha un `Conduttore`, la parte dell'inquilino è
`quota × Quota conduttore % / 100` (arrotondata al centesimo) e il resto è della proprietà.
Il `Dovuto` in `Situazione` resta l'intera quota, a carico del proprietario.

## Aggiornamento dello schema

`registro.py schema` (anche con `--file registro-riservato.xlsx`) aggiunge in coda le colonne,
i fogli e le chiavi mancanti rispetto a questo modello, senza toccare i dati. Va eseguito quando
un registro creato da una versione precedente del plugin non ha una colonna prevista qui.
