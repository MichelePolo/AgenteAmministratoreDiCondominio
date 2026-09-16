# Tutorial — dal primo avvio al primo riparto

Tempo: 30 minuti la prima volta. Serve: Claude Code oppure Claude Desktop con Cowork, Google Drive
per desktop, un account Google con Gmail e Calendar.

## 0. Prima di iniziare

1. In Google Drive (browser) crea la cartella `Condominio <nome>`. Per ora non condividerla.
2. Apri Google Drive per desktop e verifica che la cartella compaia sul tuo computer
   (Windows: `G:\Il mio Drive\Condominio <nome>`; Mac: `~/Google Drive/Il mio Drive/Condominio <nome>`).
3. Collega i connettori **Gmail** e **Google Calendar** di Claude al tuo account Google:
   in Claude Desktop da Impostazioni → *Connettori*; su claude.ai da Impostazioni → *Connettori*.
   Se usi Claude Code, i connettori collegati su claude.ai sono disponibili automaticamente.
4. Installa il plugin (vedi README).

## 1. Setup

Apri la cartella del condominio come cartella di lavoro (in Cowork: *Project or folder*; in
Claude Code: avvia `claude` dentro la cartella). Scrivi:

> configura il condominio

L'assistente verifica i connettori, poi ti chiede nome, anno di esercizio, profilo (autogestione
o professionista) e la tua email. Se vuoi solo provare, rispondi "usa i dati di esempio": avrai
4 unità, 3 spese e 2 versamenti finti.

Alla fine la cartella contiene `da analizzare/`, `archivio/`, `prospetti/`, i due registri e
`LEGGIMI.txt`.

## 2. Anagrafica e millesimi

Senza dati di esempio, `Anagrafica` e `Millesimi` sono vuoti. Puoi dettarli in chat:

> interno 1, piano terra, Mario Rossi, mario.rossi@…, millesimi A 300, B 100, C 0

oppure fotografare la tabella millesimale e chiedere di leggerla (confermerai i numeri riga per
riga), oppure compilare i fogli a mano con Excel, LibreOffice o Google Sheets (lasciando il
formato `.xlsx`): una riga per unità, sopra la riga TOTALE. Ogni colonna `Tab` deve sommare a
1000. Alla fine:

> controlla che i millesimi quadrino

L'assistente esegue la verifica, riscrive la riga TOTALE e ti dice cosa manca. Se il condominio
non ha l'ascensore, lascia `Tab C` a zero.

## 3. Condivisione (una volta sola, da Drive nel browser)

1. Condividi la cartella `Condominio <nome>` con tutti i condomini come **Visualizzatore**.
2. Entra in `da analizzare/` → Condividi → alza i condomini a **Collaboratore**.
3. Sposta `registro-riservato.xlsx` in una cartella privata (es. `Il mio Drive/Riservato condominio`)
   e scrivi il nuovo percorso nella chiave `Percorso registro riservato` del foglio `Condominio`.
   In alternativa, rimuovi la condivisione solo da quel file.

I condomini ora vedono l'archivio, il registro e il `LEGGIMI.txt`, e possono caricare documenti.

**Dal cellulare**: con l'app Google Drive un condomino apre la cartella condivisa e tocca
`registro-condominio.xlsx` o un prospetto: l'anteprima mostra i fogli (`Spese`, `Scadenze`,
`Diario`…) senza bisogno di Excel. I file non contengono formule, quindi i totali e i saldi si
vedono sempre. Per caricare una bolletta dal telefono: `da analizzare/` → `+` → *Carica*, o una
foto scattata al momento.

## 4. Primo archivio

Metti in `da analizzare/` una bolletta in PDF (o una foto). Scrivi:

> archivia i documenti

Vedrai una proposta: nuovo nome, cartella di destinazione, spesa da registrare, tabella di
riparto. Correggi se serve ("è tabella A"), poi "procedi". Guarda `archivio/2026/fatture/`
e il foglio `Spese`.

## 5. Primo riparto

> calcola il riparto delle spese di quest'anno, in 4 rate

Il prospetto finisce in `prospetti/` con il suffisso `_bozza`. Apri il foglio `Riepilogo`: una
riga per unità, le quote per tabella, il totale e le rate. Nel foglio `Dettaglio` c'è ogni spesa
ripartita. La somma delle quote coincide al centesimo con il totale delle spese.

Il prospetto è una bozza finché l'assemblea non lo approva; quando succede, dillo ("l'assemblea
ha approvato") e l'assistente toglie il suffisso e aggiorna il dovuto di ogni unità.

## 6. Convocazione dell'assemblea (in modalità collaudo)

> convoca l'assemblea ordinaria per il 20 novembre alle 18:30 in sala condominiale, seconda
> convocazione il 21 alle 18:30, ordine del giorno: approvazione rendiconto 2026 e riparto,
> preventivo caldaia

L'assistente prepara il testo, ti mostra i destinatari, e, poiché la modalità collaudo è
attiva, invia tutto **solo a te**, con oggetto `[COLLAUDO]`. Controlla la mail ricevuta.

Quando sei soddisfatto:

> disattiva la modalità collaudo

Da quel momento le email vanno ai condomini, sempre e solo dopo il tuo "sì".

## 7. Scadenze

> metti in calendario l'assemblea e le 4 rate

Il foglio `Scadenze` e Google Calendar restano allineati; rilanciare il comando non crea doppioni.

## Ogni settimana

1. "archivia i documenti": svuota l'inbox dei condomini.
2. "cosa scade nelle prossime settimane"
3. Quando arriva un bonifico: "registra un versamento di 150 € dall'interno 3, rata 2". Il saldo
   dell'unità si aggiorna da solo nel registro riservato.

## Se qualcosa non torna

- *"Non trovo registro-condominio.xlsx"*: la cartella di lavoro non è quella del condominio.
  Cambiala (Cowork: *Project or folder*; Claude Code: riavvia dentro la cartella giusta).
- *"openpyxl non installato"*: scrivi "installa openpyxl" e riprova.
- *"Tab B somma a 990"*: correggi i millesimi nel registro; il riparto non parte finché non quadra.
- *"Spesa 7: manca la Tabella"*: assegna la tabella a quella spesa ("la spesa 7 è tabella B").
- *Le email non partono / il calendario non si aggiorna*: il connettore non è collegato. Vedi il
  passo 0 e `skills/condominio-setup/references/connettori.md`.
- *Un condomino vede totali vuoti dal telefono*: il file è stato salvato con formule da un altro
  programma. Scrivi "verifica il registro": i valori vengono riscritti.

Tutto ciò che l'assistente fa è nel foglio `Diario`. Se non c'è nel Diario, non è successo.
