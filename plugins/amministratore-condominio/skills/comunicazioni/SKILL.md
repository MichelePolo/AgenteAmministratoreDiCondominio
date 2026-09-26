---
name: comunicazioni
description: >
  Prepara e, solo dopo conferma, invia via Gmail convocazioni, avvisi, prospetti, verbali e
  solleciti di pagamento; bozza di verbale. Usare per "convoca l'assemblea", "manda un
  avviso", "sollecita", "invia il riparto", "email ai condomini".
metadata:
  version: "0.3.0"
---

# Comunicazioni

Leggere prima `condominio-base`. Regola sopra tutte: **nessuna email parte senza una conferma
esplicita dell'utente dopo aver visto il testo finale e i destinatari**. I modelli di testo sono
in `references/modelli.md`.

## Procedura comune

1. **Contesto**: `registro.py info` per nome condominio, amministratore, `Modalità collaudo`,
   `Email collaudo`. `registro.py read Anagrafica` per i destinatari. (`registro.py` è
   `${CLAUDE_SKILL_DIR}/../condominio-base/scripts/registro.py`, sempre con `--dir`.)
2. **Destinatari**: tutte le unità con `Email` valorizzata; elencare a parte le unità senza
   email ("da avvisare a mano: interno 3, Rossi"). Per i solleciti, una sola unità.
3. **Bozza**: scrivere oggetto e corpo seguendo il modello. Firmare con nome dell'amministratore
   e nome del condominio. Niente formule vuote, niente emoji.
4. **Mostrare** oggetto, destinatari (To/Cc/Bcc) e corpo integrale. Chiedere: "Invio?".
5. **Modalità collaudo = SI**: sostituire TUTTI i destinatari con `Email collaudo`, prefissare
   l'oggetto con `[COLLAUDO]` e mettere in testa al corpo la riga
   `Destinatari reali: <elenco>`. Dirlo all'utente prima dell'invio.
6. **Invio** con il connettore Gmail solo dopo un "sì" inequivocabile. Per comunicazioni a tutti
   i condomini usare **Bcc** (i condomini non devono vedere gli indirizzi degli altri), To =
   amministratore. Allegati: i file di `prospetti/` o `archivio/` indicati dall'utente.
7. **Diario**: `registro.py diario "Inviata <tipo> a N destinatari" --dettaglio "<oggetto>" --approvato "<nome>"`.
   Salvare una copia del testo in `archivio/<anno>/comunicazioni/<data>_comunicazione_<oggetto>.md`.
   **Comunicazioni personali** (sollecito, nota al conduttore, situazione personale): il Diario è
   pubblico e l'archivio è condiviso, quindi nel Diario solo tipo e numero
   (`diario "Inviati 2 solleciti"`, senza `--dettaglio` né oggetto, che contiene l'interno) e la
   copia del testo in `<cartella riservata>/comunicazioni/`, mai in `archivio/`.

Se il connettore Gmail non è disponibile: preparare comunque il testo, salvarlo in
`archivio/<anno>/comunicazioni/` e dire all'utente di inviarlo dal proprio client, indicando come
collegare il connettore (`condominio-setup`, fase 0).

## Tipi

### Convocazione di assemblea

Chiedere: data/ora/luogo della prima e della seconda convocazione, ordine del giorno (punti
specifici, non generici: "approvazione rendiconto esercizio 2026 e relativo riparto", non
"varie ed eventuali" come unico punto). Verificare:

- almeno **5 giorni** tra invio e prima convocazione (art. 66 disp. att. c.c.); se meno, avvisare;
- seconda convocazione in giorno diverso dalla prima ed entro 10 giorni;
- se l'utente convoca via email ordinaria, ricordare **una volta** che i mezzi previsti sono
  raccomandata, PEC, fax, consegna a mano, salvo regolamento che ammetta l'email; proporre di
  inviare comunque l'email "a titolo di cortesia" e di curare la notifica formale.

Allegare il prospetto se all'ordine del giorno c'è un riparto. Registrare la data
dell'assemblea nel foglio `Scadenze` (tipo `assemblea`) e proporre l'evento in calendario.

### Avviso generico

Lavori programmati, interruzioni di servizi, richieste (es. "caricate le letture dei contatori
in `da analizzare/` entro il 30"). Breve, con data e cosa deve fare il condomino.

### Invio prospetto / verbale

Corpo di tre righe, allegato, riferimento all'assemblea. Per il verbale: solo dopo che l'utente
conferma che è la versione firmata/definitiva.

### Sollecito di pagamento (una unità per volta)

Leggere la situazione dell'unità dal registro riservato (`Percorso registro riservato` se valorizzato):
`registro.py --file registro-riservato.xlsx situazione --unita <id>`: dovuto, versato, saldo,
**scaduto**, rate con stato, versamenti, livello e data dell'ultimo sollecito, già calcolati.

Si sollecita **lo scaduto, non il saldo**: chi ha pagato tutte le rate già scadute è in regola
anche se deve ancora le prossime. Se `Scaduto` è 0, fermarsi: non c'è nulla da sollecitare. Nel
testo, l'importo è lo `Scaduto` e si elencano le rate con stato `scaduta` (numero, scadenza,
parte non pagata). Se `Scaduto` è vuoto (registro senza rate, precedente alla 0.4) usare il
`Saldo` e dire all'utente che senza le date delle rate non si può distinguere il ritardo.

Livelli, con testo da `references/modelli.md`:
1. **cortese** — prima volta, tono neutro, IBAN e importo;
2. **formale** — dopo 15–30 giorni dal primo, cita l'art. 63 disp. att. c.c. e chiede riscontro
   entro un termine;
3. **diffida** — testo da far verificare a un legale; il plugin prepara solo la bozza e lo dice.

Nel testo: solo i dati di quella unità. Mai il confronto con altri condomini. Dopo l'invio:
`update Situazione --where "ID unità=<id>" --where "Esercizio=<anno>" '{"Ultimo sollecito":"<oggi>","Livello sollecito":<n>}'`
e `append Solleciti`. Prima di sollecitare, chiedere se ci sono bonifici arrivati e non ancora
registrati: si registrano con la skill `versamenti`. Con profilo `autogestione`, ricordare che dopo 6 mesi di morosità
l'amministratore deve attivarsi per il recupero (art. 1129 co. 9 c.c.) e che conviene parlare
con la persona prima del secondo sollecito.

### Nota al conduttore (solo su richiesta esplicita)

Per un'unità con `Conduttore` e `Email conduttore` in `Anagrafica`, dopo un riparto: email
informativa all'inquilino con la sua parte (dal foglio `Conduttori` del prospetto o dal JSON di
`riparto.py`), Cc al proprietario, modello "Nota al conduttore" in `references/modelli.md`.
Non è una richiesta di pagamento al condominio: l'inquilino paga al proprietario, salvo diverso
accordo tra loro. Mai un sollecito al conduttore: il debitore verso il condominio è il proprietario.

### Bozza di verbale

Da appunti dell'utente o da registrazione trascritta: intestazione (condominio, data, ora,
convocazione, presenti con millesimi rappresentati, presidente e segretario), quorum
verificato in base a `Millesimi` (art. 1136 c.c.), un paragrafo per punto all'OdG con
esito della votazione (favorevoli/contrari/astenuti e millesimi), chiusura. Salvare in
`prospetti/` come `.md` o `.docx`; l'utente lo revisiona e lo carica firmato in
`da analizzare/` per l'archiviazione.

## Cosa non fare

- Non inviare a destinatari non presenti in `Anagrafica`.
- Non inviare solleciti al conduttore e non contare le sue quote come dovute al condominio.
- Non usare Cc per liste di condomini.
- Non allegare `registro-riservato.xlsx` o estratti per unità a comunicazioni collettive.
- Non minacciare azioni legali in un testo non richiesto esplicitamente come diffida.
