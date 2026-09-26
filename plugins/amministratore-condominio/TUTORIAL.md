# Tutorial — dal primo avvio al primo riparto

Tempo: 30 minuti la prima volta. Serve: Claude Code oppure Claude Desktop con Cowork, Google Drive
per desktop, un account Google con Gmail e Calendar, Python 3 con `openpyxl` (il setup lo controlla
e dice come installarlo).

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
> interno 2, primo piano, Anna Bianchi, anna@…, affittato a Paolo Gialli, paolo@…, millesimi A 250, B 250, C 300

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
`registro-condominio.xlsx`: si apre sul foglio **In breve**, che in una schermata dice la prossima
assemblea, la prossima rata, le spese dell'anno per tabella, gli ultimi documenti archiviati e
quanti documenti caricati aspettano di essere archiviati. Gli altri fogli (`Spese`, `Scadenze`,
`Diario`…) e i prospetti si aprono allo stesso modo, senza bisogno di Excel. `In breve` si
aggiorna da solo a ogni operazione e non contiene dati delle singole unità. I file non contengono formule, quindi i totali e i saldi si
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

Se ci sono appartamenti affittati, il `Riepilogo` ha anche le colonne "A carico proprietà" e
"A carico conduttore", e il foglio `Conduttori` mostra spesa per spesa la parte dell'inquilino,
secondo la percentuale che l'assistente propone per ogni spesa quando la archivia (100% per
servizi come luce scale e pulizie, 0% per lavori straordinari). Il proprietario resta il debitore
verso il condominio; con "manda la nota al conduttore dell'interno 2" l'inquilino riceve la sua
parte per informazione.

Il prospetto è una bozza finché l'assemblea non lo approva. Quando succede, dillo con le date delle
rate decise in assemblea:

> l'assemblea del 20 settembre ha approvato il riparto, rate il 31 ottobre, 31 gennaio, 30 aprile e 31 luglio

L'assistente registra le rate di ogni unità nel registro riservato (da lì si calcolano dovuto,
saldo e rate scadute), aggiunge le date al foglio `Scadenze`, toglie `_bozza` dal nome del prospetto
e annota tutto nel Diario. Se poi approvi il consuntivo che sostituisce il preventivo, dillo ("il
consuntivo sostituisce il preventivo"): le rate vecchie restano, marcate come non più valide.

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

## 8. Versamenti ed estratto conto

Scarica l'estratto conto del condominio in PDF dall'home banking e mettilo nella cartella del
registro riservato, sottocartella `estratti conto/`: **non** in quella condivisa, perché contiene
nomi e IBAN di tutti. Poi:

> registra i bonifici dell'estratto conto di settembre

L'assistente legge i bonifici in entrata e propone per ciascuno l'unità e la rata, con il motivo
(nome dell'ordinante, interno nella causale, importo uguale alla rata). Quelli dubbi (un cognome
solo, due omonimi, un familiare che paga per un altro) restano "da abbinare": di' tu a chi vanno
("la 3 è l'interno 3"). Dopo la tua conferma li registra; rilanciare lo stesso estratto non crea
doppioni. Per un pagamento singolo basta: "registra 150 € dall'interno 3, bonifico del 10 settembre".

## 9. La situazione di ciascuno

> manda a tutti la situazione

Ogni condomino riceve un'email solo con i propri conti: dovuto, versato, rate con lo stato
(pagata, da pagare, scaduta), versamenti registrati e IBAN. Vedi la prima email per intero e una
tabella di tutte, poi confermi una volta. In modalità collaudo parte una sola email, a te.

## Ogni settimana

1. "come siamo messi?": in cinque righe documenti da archiviare, scadenze, assemblea, rate
   scadute e le tre cose da fare adesso, in ordine di urgenza. Rispondi "fai la 1".
2. "archivia i documenti": svuota l'inbox dei condomini.
3. "registra i bonifici dell'estratto conto": il saldo e le rate scadute di ogni unità si
   aggiornano da soli nel registro riservato. Prima di sollecitare qualcuno, registra sempre i
   bonifici arrivati: si sollecita solo chi ha rate scadute non pagate.

## Se qualcosa non torna

- *"Non trovo registro-condominio.xlsx"*: la cartella di lavoro non è quella del condominio.
  Cambiala (Cowork: *Project or folder*; Claude Code: riavvia dentro la cartella giusta).
- *"openpyxl non installato"*: scrivi "installa openpyxl". Su Linux, se `pip` non c'è o risponde
  `externally-managed-environment`, l'assistente ti chiederà di eseguire tu
  `sudo apt install python3-openpyxl` (serve la password).
- *"Tab B somma a 990"*: correggi i millesimi nel registro; il riparto non parte finché non quadra.
- *"Spesa 7: manca la Tabella"*: assegna la tabella a quella spesa ("la spesa 7 è tabella B").
- *"Manca la colonna Email conduttore"*, oppure il registro non si apre su *In breve*: il registro è
  di una versione precedente. Scrivi "aggiorna lo schema del registro" (per entrambi i registri):
  fogli e colonne nuove (`In breve`, `Rate`, `Scaduto`…) vengono aggiunti senza toccare i dati. Il
  dovuto scritto con la versione 0.3 viene conservato alla prima approvazione con le rate.
- *Un bonifico resta "da abbinare"*: è voluto quando gli indizi sono deboli o due unità sono
  possibili. Di' a quale unità va; meglio una domanda in più che un versamento sulla persona sbagliata.
- *"Scaduto" è vuoto per un'unità*: l'unità non ha rate registrate (riparto approvato prima della
  0.4). Si riempie dalla prossima approvazione con le date delle rate.
- *Le email non partono / il calendario non si aggiorna*: il connettore non è collegato. Vedi il
  passo 0 e `skills/condominio-setup/references/connettori.md`.
- *Un condomino vede totali vuoti dal telefono*: il file è stato salvato con formule da un altro
  programma. Scrivi "verifica il registro": i valori vengono riscritti.

Tutto ciò che l'assistente fa è nel foglio `Diario`. Se non c'è nel Diario, non è successo.
