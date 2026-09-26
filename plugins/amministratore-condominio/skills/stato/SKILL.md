---
name: stato
description: >
  Riepilogo di come sta il condominio: documenti da archiviare, problemi del registro,
  scadenze, assemblea, bozze, rate scadute e cosa fare adesso. Usare per "come siamo messi",
  "situazione del condominio", "cosa c'è da fare", "riepilogo".
metadata:
  version: "0.4.0"
---

# Stato del condominio

Leggere prima `condominio-base`. È il punto di partenza della settimana: una risposta breve che
dice come stanno le cose e cosa conviene fare adesso. **Non modifica nulla**: né registri, né
Diario, né file. Le azioni si fanno dopo, con le altre skill, se l'utente le chiede.

## Procedura

1. Eseguire:
   ```bash
   python3 "${CLAUDE_SKILL_DIR}/scripts/stato.py" --dir "<cartella del condominio>"
   ```
   Il JSON contiene `inbox`, `registro` (problemi e avvisi), `scadenze` (prossime 30 giorni, da
   riproporre, senza evento in calendario), `assemblea` (convocazione trovata nel Diario, termine
   per il rendiconto), `bozze`, `pagamenti` (rate scadute per unità, dal registro riservato),
   `condominio.collaudo` e `azioni` (al massimo 3, già in ordine di urgenza).
2. Controllare anche i **connettori**, che lo script non vede: se in questa sessione mancano gli
   strumenti Gmail o Google Calendar, dirlo in una riga (vedi `condominio-setup`, fase 0).
3. Rispondere così, **al massimo cinque righe** più le azioni, senza tabelle lunghe:

   ```
   Condominio Le Betulle, 26 settembre 2026
   · 2 documenti da archiviare (il più vecchio da 16 giorni)
   · Prossimo: verifica ascensore il 15 ottobre; assemblea il 20 novembre
   · Rate scadute: 2 unità, 193,92 €
   · Registro in ordine · Modalità collaudo attiva

   Cosa farei adesso:
   1. Registrare i bonifici arrivati, poi valutare i solleciti → "registra i bonifici dell'estratto conto"
   2. Archiviare i 2 documenti → "archivia i documenti"
   ```

   Una riga per argomento, solo se c'è qualcosa da dire (niente "0 documenti"); le azioni con la
   frase da dire, così l'utente può rispondere "fai la 1". Con profilo `autogestione`, aggiungere
   a ogni azione una frase sul perché (per esempio: "la convocazione deve arrivare almeno 5 giorni
   prima dell'assemblea, art. 66 disp. att. c.c."); con `professionista`, no.
4. Se tutto è in ordine e non ci sono azioni, dirlo in una riga e indicare la prossima scadenza.

## Regole

- Le rate scadute sono dati riservati: nominare le unità in ritardo **solo se l'utente lo
  chiede** ("chi è in ritardo?"), usando `pagamenti.dettaglio`. Mai in documenti o email collettive.
- Non ricalcolare numeri e non rifare le verifiche a mano: il JSON è la fonte.
- Se lo script dice che il registro non c'è, la cartella di lavoro non è quella del condominio:
  dirlo e proporre `condominio-setup` o di aprire la cartella giusta.
- Se `pagamenti` è `null`, il registro riservato non è stato trovato: dire di controllare la
  chiave `Percorso registro riservato` del foglio `Condominio`.
