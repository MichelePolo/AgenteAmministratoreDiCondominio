---
name: condominio-setup
description: >
  Primo avvio: verifica i connettori Gmail e Calendar, crea cartella e registri, guida
  anagrafica e tabelle millesimali. Usare per "configura il condominio", "iniziamo",
  "nuovo condominio", "setup", o se manca registro-condominio.xlsx.
metadata:
  version: "0.3.0"
---

# Setup del condominio

Leggere prima `condominio-base` (struttura, principi, modello dati). Questo è il momento in cui
l'utente decide di fidarsi: essere ordinati, brevi, e non chiedere due volte la stessa cosa.
Il setup ha tre fasi: **connettori**, **struttura**, **tabelle millesimali**.

## Fase 0 — Connettori (Gmail e Google Calendar)

Il plugin usa i connettori integrati di Claude, non server MCP propri. Verificare quali strumenti
sono disponibili in questa sessione:

- **Gmail**: strumenti per inviare email (in Claude Code compaiono come `mcp__claude_ai_Gmail__*`;
  in Claude Desktop e Cowork come connettore "Gmail").
- **Google Calendar**: strumenti per creare/cercare eventi (`mcp__claude_ai_Google_Calendar__*`
  o connettore "Google Calendar").

Se uno dei due manca, dire cosa fare seguendo `references/connettori.md` (Claude Desktop:
Impostazioni → Connettori; Claude Code: collegarli su claude.ai, poi riavviare la sessione).
**Non bloccarsi**: archivio, registro e riparto funzionano senza connettori; solo comunicazioni e
scadenze in calendario ne hanno bisogno. Annotare l'esito ("Gmail: sì, Calendar: no") nel Diario a
fine setup.

Verificare anche Python: `python3 -c "import openpyxl"`; se manca, `pip install openpyxl`.

## Fase 1 — Struttura

### 1. Raccogliere il minimo indispensabile

Chiedere in un solo messaggio: nome del condominio, anno di esercizio (proporre l'anno corrente),
profilo (`autogestione` / `professionista`), email dell'amministratore. Se l'utente vuole solo
provare, proporre i **dati di esempio** (4 unità, 3 spese, 2 versamenti) con cui si può fare subito
un riparto.

La cartella di lavoro deve essere la cartella del condominio dentro Google Drive per desktop
(es. `~/Google Drive/Il mio Drive/Condominio Le Betulle` o `G:\Il mio Drive\...`). Se l'utente ha
aperto una cartella locale non sincronizzata, avvisare: funziona lo stesso, ma i condomini non
vedranno nulla finché non viene spostata in Drive. Non bloccarsi.

### 2. Creare la struttura

```bash
python3 "${CLAUDE_SKILL_DIR}/scripts/crea_registro.py" --dir "<cartella di lavoro>" --nome "<Nome>" --esercizio <anno> [--esempio]
```

Lo script non sovrascrive registri esistenti. Se i registri esistono ma sono di una versione
precedente del plugin, `registro.py schema` (e `--file registro-riservato.xlsx schema`) aggiunge
fogli e colonne mancanti senza toccare i dati. Poi impostare i valori raccolti:

```bash
R="${CLAUDE_SKILL_DIR}/../condominio-base/scripts/registro.py"
python3 "$R" --dir "<cartella>" set Condominio "Profilo" "autogestione"
python3 "$R" --dir "<cartella>" set Condominio "Amministratore" "Nome Cognome"
python3 "$R" --dir "<cartella>" set Condominio "Email amministratore" "x@y.it"
python3 "$R" --dir "<cartella>" set Condominio "Email collaudo" "x@y.it"
```

`Modalità collaudo` nasce a `SI` e resta `SI` finché l'utente non chiede esplicitamente di
disattivarla dopo aver visto almeno una email di prova arrivare a sé stesso.

## Fase 2 — Anagrafica e tabelle millesimali

Senza `--esempio`, `Anagrafica` e `Millesimi` nascono vuoti (solo la riga TOTALE a zero). Le
tabelle predefinite sono `Tab A` (proprietà), `Tab B` (scale), `Tab C` (ascensore); se il
condominio ne usa altre (riscaldamento, cortile…) si aggiungono colonne `Tab D`, `Tab E`… in coda
e si aggiorna la chiave `Tabelle millesimali` del foglio `Condominio`. Se una tabella non esiste
(es. niente ascensore) si lascia a zero: `verifica` la segnala come "non usata" senza bloccare.

Tre strade, lasciare scegliere:

- **Dettatura in chat**: l'utente elenca le unità ("interno 1, piano terra, Mario Rossi,
  mario@..., millesimi 300/100/0"); inserire con `registro.py append Anagrafica` e
  `append Millesimi`, un'unità per volta (stesso `ID unità`, es. `U01`), poi mostrare il riepilogo.
- **Da tabella millesimale in PDF o foto**: leggerla, proporre le righe da inserire e far
  confermare i numeri riga per riga: un millesimo sbagliato si propaga a ogni riparto.
- **A mano nel foglio**: dire quali colonne compilare (una riga per unità, sopra la riga TOTALE)
  e che ogni colonna `Tab` deve sommare a 1000.

Chiedere se ci sono **appartamenti affittati**: per quelli, `Conduttore` ed `Email conduttore` in
`Anagrafica`. I millesimi restano dell'unità; la divisione tra proprietario e inquilino avviene
spesa per spesa con `Quota conduttore %` (vedi `archivio` e `riparto`).

In tutti i casi chiudere con:

```bash
python3 "$R" --dir "<cartella>" verifica
```

che controlla che ogni tabella sommi a 1000, che Anagrafica e Millesimi abbiano le stesse unità,
che l'email di collaudo ci sia, e riscrive la riga TOTALE (utile se il foglio è stato compilato a
mano). Ripetere finché `"ok": true`. Non dichiarare il setup concluso prima.

Con profilo `autogestione` e più di otto unità, ricordare una volta sola che la nomina di un
amministratore è obbligatoria (art. 1129 c.c.).

## Fase 3 — Condivisione della cartella (istruzioni per l'utente, non automatizzabile)

Spiegare cosa fare in Google Drive (web):

1. Condividere la cartella del condominio con **tutti i condomini** (email in `Anagrafica`)
   come **Visualizzatore**.
2. Sulla sottocartella `da analizzare/`, alzare il permesso a **Collaboratore** (possono aggiungere).
3. **Rimuovere la condivisione da `registro-riservato.xlsx`** o, meglio, spostarlo in una cartella
   privata fuori da quella condivisa. Se lo si sposta, scrivere il nuovo percorso nella chiave
   `Percorso registro riservato` del foglio `Condominio`: tutte le skill lo leggono da lì e
   passano `--dir` a `registro.py`.

Dire all'utente che i condomini possono consultare tutto anche **da cellulare** con l'app Google
Drive: toccando `registro-condominio.xlsx` o un prospetto si apre l'anteprima con i fogli, senza
bisogno di Excel. È per questo che i file non contengono formule.

Ricordare che `LEGGIMI.txt` è già nella cartella ed è pensato per chi non usa l'IA.

## Chiudere

Registrare nel Diario: `registro.py diario "Setup completato" --dettaglio "Gmail: sì/no, Calendar: sì/no, N unità" --approvato "<nome>"`.
Proporre il passo successivo concreto: mettere un documento in `da analizzare/` e chiedere
"archivia", oppure "calcola il riparto delle spese" se ci sono dati di esempio.

## Cosa non fare

- Non inventare millesimi, codici fiscali o IBAN. Vuoto è meglio di finto.
- Non promettere che il calendario o le email funzionino se il connettore non è collegato.
- Non disattivare la modalità collaudo di propria iniziativa.
- Non scrivere formule nei registri.
