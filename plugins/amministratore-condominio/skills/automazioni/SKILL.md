---
name: automazioni
description: >
  Attiva, sospende o elimina il riepilogo settimanale che arriva per email all'amministratore
  senza aprire Claude. Usare per "mandami ogni lunedì come siamo messi", "riepilogo
  automatico", "programma", "automazione", "sospendi il riepilogo".
metadata:
  version: "0.5.0"
---

# Automazioni

Leggere prima `condominio-base`. Il plugin ha **una sola automazione**: il riepilogo settimanale
"come siamo messi?" (lo stesso testo della skill `stato`), inviato per email **solo
all'amministratore**. L'automazione legge e avvisa: non modifica file, non scrive nel Diario, non
scrive ai condomini, non archivia, non registra, non sollecita. Non proporre e non creare
automazioni diverse da questa, anche se richieste: ogni altra operazione del plugin richiede una
conferma al momento, che un'esecuzione programmata non può avere.

È l'unica eccezione al principio "nessuna email senza conferma" (vedi `condominio-base`): va
solo all'amministratore, l'ha attivata lui, contiene solo conteggi.

## Dove gira

| Ambiente | Come | Limite |
|---|---|---|
| Claude Desktop (Mac, Windows) | attività locale pianificata (scheda Code → Routines → Local) | gira solo con l'app aperta e il computer acceso; al risveglio recupera una sola esecuzione persa |
| Linux, o chi usa solo il terminale | timer `systemd --user` che lancia `claude -p` (vedi sotto, "Terminale") | il computer deve essere acceso; recupera l'esecuzione persa all'avvio |

Le routine cloud non vanno bene: non vedono la cartella del condominio sul computer.

## Attivazione

1. **Controlli**, fermandosi al primo che non va:
   - la cartella di lavoro è quella del condominio (c'è `registro-condominio.xlsx`);
   - `registro.py info`: `Email amministratore` compilata (è il destinatario) e `Nome`;
   - in questa sessione c'è lo strumento di invio di Gmail; se manca, spiegare come collegare il
     connettore (`condominio-setup`, fase 0) e fermarsi;
   - `stato.py --testo` funziona adesso sulla cartella (mostrarne il testo: è quello che arriverà).
2. **Chiedere** giorno e ora (proposta: lunedì alle 8). Il modello è Haiku: il compito è
   meccanico (eseguire uno script, inviare un'email).
3. **Riassumere e chiedere conferma**: nome dell'automazione `riepilogo-<nome condominio in
   minuscolo con trattini>`, cartella, giorno e ora, destinatario, modello, i due soli permessi
   (script di stato, invio Gmail).
4. **Istruzioni**: prendere `references/istruzioni-riepilogo.md` e sostituire `{nome}` con il nome
   del condominio e `{email}` con `Email amministratore`. Nessun percorso di file: la skill `stato`
   si trova da sola anche dopo un aggiornamento del plugin.
5. **Creare** secondo l'ambiente (sezioni sotto).
6. **Diario**: `registro.py diario "Attivato il riepilogo settimanale per l'amministratore" --dettaglio "<giorno e ora>" --approvato "<nome>"`.

Se l'automazione esiste già per questo condominio, non crearne un'altra: proporre di modificarla.
Se `Email amministratore` cambia, l'automazione va riattivata (l'indirizzo è nelle istruzioni).

### Claude Desktop

Creare un'**attività locale** (non una routine remota): si può fare direttamente da questa
sessione descrivendola, per esempio "crea un'attività locale settimanale, ogni lunedì alle 8,
nella cartella <cartella>, modello Haiku, con queste istruzioni: …". Se la sessione non riesce a
crearla, guidare l'utente nella scheda **Code** → **Routines** → **New routine** → **Local**:
Name, Description ("Riepilogo settimanale del condominio"), Instructions (il testo del punto 4,
con modello Haiku e permessi "Ask"), cartella del condominio, Schedule **Weekly** con giorno e ora.

Poi, insieme all'utente:
1. **Run now** sulla pagina dell'attività;
2. alle richieste di permesso scegliere **Always allow** solo per l'esecuzione dello script di
   stato e per l'invio Gmail; rifiutare qualsiasi altra richiesta (non dovrebbero essercene);
3. controllare che l'email sia arrivata;
4. consigliare **Keep computer awake** (Impostazioni → Desktop app → General) se il computer va
   in sospensione a quell'ora.

### Terminale (Linux)

Vedi la sezione "Terminale" aggiunta dalla fase 3 (script `pianifica.py`).

## Sospendere, modificare, eliminare

- **Desktop**: chiedere in una sessione ("sospendi il riepilogo settimanale", "spostalo al
  venerdì") oppure dalla pagina dell'attività in Routines (Status, Edit, Delete).
- Annotare nel Diario: "Sospeso il riepilogo settimanale", "Eliminato il riepilogo settimanale".

## Regole

- Un solo riepilogo per condominio, sempre e solo verso `Email amministratore`.
- Mai aggiungere permessi oltre ai due previsti; mai la modalità che salta i permessi.
- Mai inserire nel testo o nell'email dati per unità: il testo di `stato.py --testo` già non li
  contiene, e va inviato così com'è.
- Con profilo `autogestione`, spiegare in una frase perché l'automazione non fa altro che
  avvisare: tutto ciò che cambia registri o scrive ai condomini ha bisogno di un "sì" detto al momento.
