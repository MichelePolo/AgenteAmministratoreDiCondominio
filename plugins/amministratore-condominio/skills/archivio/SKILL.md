---
name: archivio
description: >
  Svuota "da analizzare": legge i documenti caricati dai condomini, propone nome e
  classificazione, li sposta in archivio/ e registra le spese. Usare per "archivia",
  "sistema i documenti", "cosa hanno caricato", "registra questa fattura".
metadata:
  version: "0.3.0"
---

# Archivio

Leggere prima `condominio-base` e le convenzioni in
`../condominio-base/references/convenzioni-file.md`. Questa è la skill che i condomini "vedono"
lavorare: un file sparisce da `da analizzare/` e ricompare in `archivio/` con un nome sensato.
Precisione e prevedibilità contano più della velocità.

## Procedura

### 1. Scansione

```bash
python3 "${CLAUDE_SKILL_DIR}/scripts/scansiona.py" --dir "<cartella>"
```

Lo script restituisce per ogni file hash e stato: `nuovo`, `duplicato` (già in archivio),
`duplicato_in_inbox` (stesso contenuto di un altro file caricato ora). Non sposta nulla e ignora
scorciatoie e file temporanei di Google Drive. Se `totale = 0`, dirlo e fermarsi.

In questa skill `registro.py` è `R="${CLAUDE_SKILL_DIR}/../condominio-base/scripts/registro.py"`
(sempre con `--dir "<cartella>"`).

### 2. Lettura e classificazione (per ogni file nuovo)

Aprire il documento (PDF, immagine, docx, xlsx…). Per le immagini di bollette o fatture leggere
il testo visibile; se illeggibile, classificare come `foto` con oggetto descrittivo e chiedere.

Determinare:

| Campo | Come |
|---|---|
| Tipo | fattura, bolletta, preventivo, verbale, comunicazione, contratto, ricevuta, polizza, foto, altro |
| Data documento | data in intestazione; NON la data di caricamento |
| Controparte | fornitore / mittente |
| Oggetto | 2–4 parole |
| Importo | totale lordo, se presente |
| Tabella di riparto | vedi sotto; solo per documenti che generano una spesa |
| Anno di archivio | anno della data documento |

**Tabella di riparto** — proporre, mai imporre:
- luce scale, pulizie scale, manutenzione portone → `B` (scale) se esiste, altrimenti `A`
- ascensore → `C` se esiste
- assicurazione, amministratore, spese generali, tetto, facciata → `A`
- riscaldamento centralizzato → tabella dedicata se esiste
- lavori che riguardano una sola unità → `UNITA:<ID>`
- se in dubbio → `A` con nota "da confermare"
Le tabelle disponibili sono nella chiave `Tabelle millesimali` del foglio `Condominio`.

**Quota conduttore %** — per ogni spesa, la percentuale che all'interno di un'unità affittata
spetta all'inquilino (art. 9 L. 392/1978 e tabella oneri accessori; dettagli in
`../condominio-base/references/normativa.md`). Proporla sempre, anche se oggi nessuna unità è
affittata:
- pulizie, luce scale, acqua, riscaldamento e condizionamento, spurghi, manutenzione **ordinaria**
  ascensore e impianti, piccole riparazioni parti comuni → `100`
- portierato → `90`
- manutenzione **straordinaria** (facciata, tetto, sostituzione caldaia o ascensore), compenso
  amministratore, assicurazione fabbricato, imposte → `0`
- se in dubbio → `0` con nota "da confermare" (il proprietario resta comunque il debitore)

Un preventivo, un verbale, un contratto non generano una spesa. Un contratto genera
tipicamente una **scadenza** (rinnovo, verifica periodica): segnalarla per la skill `scadenze`.

### 3. Proposta (obbligatoria, in blocco)

Mostrare una tabella con: nome originale → nome proposto, cartella di destinazione, spesa da
registrare (importo, tabella) o "nessuna", duplicati rilevati. Esempio:

```
1. bolletta enel.pdf → archivio/2026/fatture/2026-03-14_fattura_enel_luce-scale_412.50.pdf
   spesa: 412,50 € · tabella B · conduttore 100% · esercizio 2026
2. IMG_2231.jpg → archivio/2026/altro/2026-05-03_foto_portone_danno-cerniera.jpg
   spesa: nessuna
3. enel marzo (1).pdf → DUPLICATO di archivio/2026/fatture/2026-03-14_fattura_enel_… → altro/duplicati/
```

Chiedere conferma. L'utente può correggere singole righe ("la 1 è tabella A"). Aggiornare la
proposta e richiedere conferma solo per le righe cambiate.

### 4. Esecuzione (solo dopo conferma)

Per ogni file confermato, nell'ordine:

1. `mkdir -p` della cartella di destinazione, `mv` del file. Se esiste già un file con lo stesso
   nome, aggiungere suffisso `_2` e annotarlo.
2. Se genera una spesa: `registro.py append Spese '{...}'` con `File` = percorso relativo di
   destinazione e `Quota conduttore %` valorizzata. Conservare l'`ID` restituito.
3. `registro.py append Indice '{"Hash":..., "Nome originale":..., "Nome archivio":..., "Percorso":..., "Data elaborazione":"<oggi>", "ID spesa": <id o vuoto>}'`.
4. Duplicati: `mv` in `archivio/<anno>/altro/duplicati/` mantenendo il nome originale; riga in
   `Indice` con `Nome archivio` = `DUPLICATO di <percorso copia>`.

Alla fine una sola riga di Diario riassuntiva:
`registro.py diario "Archiviati N documenti, M spese registrate, K duplicati" --approvato "<nome>"`.

### 5. Chiusura

Riepilogare in due righe cosa è stato fatto e segnalare le scadenze emerse (contratti, verifiche)
proponendo di passarle alla skill `scadenze`. Se qualche file è rimasto in `da analizzare/`
perché non classificabile, dirlo esplicitamente e chiedere all'utente cosa sia.

## Regole

- Mai spostare, rinominare o cancellare senza la conferma del punto 3.
- Mai cancellare: i duplicati si spostano.
- Mai registrare una spesa senza `File` valorizzato (tranne spese inserite a mano dall'utente).
- Importi: usare il totale del documento. Se ci sono più importi (acconto/saldo), chiedere.
- Non aprire né archiviare documenti che contengono evidentemente dati sensibili di terzi non
  pertinenti (documenti d'identità, referti): lasciarli in `da analizzare/` e avvisare l'utente.
- Con profilo `autogestione`, spiegare brevemente perché una spesa va in una tabella piuttosto
  che in un'altra la prima volta che capita; con `professionista`, no.
