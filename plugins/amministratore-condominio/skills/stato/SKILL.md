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
   python3 "${CLAUDE_SKILL_DIR}/scripts/stato.py" --dir "<cartella del condominio>" --testo
   ```
   Stampa il riepilogo già pronto: una riga per argomento (documenti da archiviare, prossime
   scadenze, assemblea, rate scadute, prospetti in bozza, stato del registro e collaudo) e poi al
   massimo tre azioni in ordine di urgenza, ciascuna con la frase da dire. Esempio:

   ```
   Condominio Le Betulle — riepilogo al 26 settembre 2026

   · 2 documenti da archiviare (il più vecchio da 16 giorni)
   · Prossime scadenze: Verifica biennale ascensore il 15 ottobre
   · Assemblea il 20 novembre
   · Rate scadute: 2 unità, 193,92 €
   · Registro in ordine · Modalità collaudo attiva

   Cosa fare adesso:
   1. 2 unità hanno rate scadute per 193,92 €: prima registrare i bonifici arrivati, poi valutare i solleciti → "registra i bonifici dell'estratto conto"
   2. 2 documenti da archiviare, il più vecchio caricato 16 giorni fa → "archivia i documenti"
   ```
2. Controllare anche i **connettori**, che lo script non vede: se in questa sessione mancano gli
   strumenti Gmail o Google Calendar, aggiungere una riga (vedi `condominio-setup`, fase 0).
3. Rispondere con **quel testo, così com'è**: è lo stesso che arriva per email con il riepilogo
   programmato (skill `automazioni`), e deve restare identico. L'utente può rispondere "fai la 1".
   Con profilo `autogestione`, aggiungere dopo il testo una frase sul perché di ogni azione (per
   esempio: "la convocazione deve arrivare almeno 5 giorni prima dell'assemblea, art. 66 disp.
   att. c.c."); con `professionista`, no.
4. Per domande di dettaglio ("quali scadenze?", "quali prospetti in bozza?", "chi è in ritardo?")
   rilanciare lo script **senza** `--testo`: il JSON contiene `inbox`, `registro` (problemi e
   avvisi), `scadenze` (prossime 30 giorni, da riproporre, senza evento in calendario),
   `assemblea`, `bozze`, `pagamenti` (con il `dettaglio` per unità) e `azioni`.

## Regole

- Le rate scadute sono dati riservati: nominare le unità in ritardo **solo se l'utente lo
  chiede** ("chi è in ritardo?"), usando `pagamenti.dettaglio`. Mai in documenti o email collettive.
- Non ricalcolare numeri e non rifare le verifiche a mano: il testo e il JSON dello script sono la fonte.
- Se lo script dice che il registro non c'è, la cartella di lavoro non è quella del condominio:
  dirlo e proporre `condominio-setup` o di aprire la cartella giusta.
- Se `pagamenti` è `null`, il registro riservato non è stato trovato: dire di controllare la
  chiave `Percorso registro riservato` del foglio `Condominio`.
