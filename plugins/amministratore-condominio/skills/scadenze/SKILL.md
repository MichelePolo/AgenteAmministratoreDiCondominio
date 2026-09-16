---
name: scadenze
description: >
  Scadenzario: assemblee, rate, manutenzioni, rinnovi contratti; allinea il foglio Scadenze a
  Google Calendar senza duplicati. Usare per "scadenze", "cosa scade", "metti in calendario",
  "ricordami", "quando è l'assemblea", "rinnovo contratto".
metadata:
  version: "0.3.0"
---

# Scadenze

Leggere prima `condominio-base`. Il foglio `Scadenze` del registro è la verità (visibile a tutti
i condomini); Google Calendar è lo specchio per l'amministratore. Ogni riga può avere al massimo
un evento (`ID evento`): questo rende l'operazione ripetibile senza duplicati.

## Procedura

### 1. Leggere

`registro.py read Scadenze` (`registro.py` è `${CLAUDE_SKILL_DIR}/../condominio-base/scripts/registro.py`,
sempre con `--dir`). Ordinare per data. Distinguere: passate, prossimi 30 giorni, oltre. Le scadenze con `Ricorrenza` diversa da `nessuna` e data passata vanno **riproposte**
alla data successiva (nuova riga, non modifica della vecchia), previa conferma.

### 2. Aggiungere (su richiesta o su segnalazione di altre skill)

Chiedere il minimo: data, tipo, descrizione, ricorrenza. Suggerire, senza insistere, le
scadenze tipiche mancanti se non ci sono nel foglio:

| Tipo | Quando | Base |
|---|---|---|
| assemblea | entro 180 giorni dalla fine esercizio | art. 1130 n. 10 c.c. |
| rata | date decise in assemblea (tipicamente 2–4 all'anno) | riparto approvato |
| manutenzione | verifica biennale ascensore; controllo impianto termico | D.P.R. 162/1999; D.P.R. 74/2013 |
| contratto | scadenza/rinnovo tacito (avvisare 60 giorni prima) | contratto in archivio |
| adempimento | CU/770 se il condominio è sostituto d'imposta; polizza fabbricato | verificare con commercialista |

Inserire: `registro.py append Scadenze '{"Data":"2026-11-15","Ora":"18:30","Tipo":"assemblea","Descrizione":"Assemblea ordinaria — approvazione rendiconto 2026","Ricorrenza":"annuale"}'`.

### 3. Allineare il calendario (solo dopo conferma dell'elenco)

Per ogni riga **senza** `ID evento` (`read Scadenze --where "ID evento="`) e con data futura:

1. Cercare nel calendario un evento con lo stesso titolo nello stesso giorno (idempotenza
   anche se il foglio è stato ricreato). Se esiste, riusarne l'ID.
2. Altrimenti creare l'evento: titolo `[Condominio <nome>] <Descrizione>`, data/ora (evento
   intero giorno se `Ora` è vuota), descrizione con tipo, riferimento al file in archivio se
   c'è, e la riga "Gestito da amministratore-condominio — riga Scadenze ID <n>". Promemoria:
   7 giorni e 1 giorno prima; per `contratto`, 60 e 7 giorni.
3. Scrivere l'ID: `registro.py update Scadenze --where "ID=<n>" '{"ID evento":"<id>"}'`.

Per righe **con** `ID evento` la cui data nel foglio è cambiata: aggiornare l'evento.
Non cancellare mai eventi dal calendario di propria iniziativa; se una scadenza viene
annullata, chiedere e poi svuotare `ID evento` e annotare nel Diario.

### 4. Riepilogo

Rispondere sempre con l'elenco delle prossime scadenze (data, cosa, quanto manca), poi ciò che
è stato aggiunto/aggiornato. Una riga di Diario per sessione:
`registro.py diario "Scadenzario: N eventi creati, M aggiornati"`.

Se il connettore Google Calendar non è disponibile, mantenere solo il foglio e dirlo, indicando
come collegarlo (`condominio-setup`, fase 0).

## Regole

- Le date delle rate nascono dal riparto approvato: non inventarle.
- Un contratto archiviato con rinnovo tacito genera SEMPRE una scadenza di preavviso, non solo
  la scadenza finale.
- Con profilo `autogestione`, spiegare una volta perché l'assemblea di approvazione del rendiconto
  ha un termine (180 giorni) e cosa succede se si sfora (responsabilità dell'amministratore).
