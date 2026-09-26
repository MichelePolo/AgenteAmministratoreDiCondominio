---
name: riparto
description: >
  Riparto delle spese tra le unità secondo i millesimi: prospetto xlsx (per unità, per spesa,
  rate) e, dopo l'approvazione, registrazione delle rate. Usare per "riparto", "quanto deve
  ciascuno", "quote", "rendiconto", "rate", "bilancio", "approvato".
metadata:
  version: "0.3.0"
---

# Riparto

Leggere prima `condominio-base`. Il calcolo è deterministico e lo fa uno script, non il modello:
il compito qui è scegliere le spese giuste, spiegare il risultato e farlo approvare.

## Procedura

### 1. Definire il perimetro

Chiedere (o dedurre dalla richiesta) cosa ripartire:
- tutto l'esercizio corrente (default), oppure
- un intervallo di date (`--dal`/`--al`), oppure
- spese specifiche per ID (`--ids 3,7,12`, es. un lavoro straordinario).

E se serve la suddivisione in rate (`--rate N`). Non chiedere ciò che è ovvio: "riparto del
primo semestre" = `--dal <inizio esercizio> --al <fine giugno>`.

Verificare con `registro.py verifica` (o `read Spese --where "Tabella="`) che le spese abbiano
tutte una `Tabella`: lo script di riparto **rifiuta** le spese senza tabella. Le righe senza vanno
completate prima con `update Spese --where "ID=<n>" '{"Tabella":"B"}'`, proponendo la tabella con i
criteri di `archivio`. `registro.py` è `${CLAUDE_SKILL_DIR}/../condominio-base/scripts/registro.py`.

### 2. Calcolare

```bash
python3 "${CLAUDE_SKILL_DIR}/scripts/riparto.py" --dir "<cartella>" [--esercizio 2026] [--dal ...] [--al ...] [--ids ...] [--rate 4] --titolo "Riparto 1° semestre 2026"
```

Se lo script esce con `"ok": false`, mostrare gli errori così come sono (tabella che non somma
a 1000, spesa senza tabella, unità inesistente, riparto manuale che non quadra) e aiutare a
correggerli nel registro. Non aggirarli e non calcolare a mano.

Il prospetto viene salvato in `prospetti/<data>_riparto_<titolo>_bozza.xlsx`; se il nome esiste
già, lo script aggiunge `_2`, `_3`… e non sovrascrive mai. Contiene solo valori (nessuna formula),
quindi si legge anche dall'anteprima di Google Drive su cellulare. Gli importi negativi (note di
credito, rimborsi) sono ripartiti con lo stesso criterio, con segno. Il JSON di uscita riporta per
ogni unità totale e rate: usarlo per la presentazione, senza ricalcolare.

### 3. Presentare

Mostrare in chat il riepilogo per unità (intestatario, totale, rate) e i totali per tabella,
con un rigo di spiegazione del criterio. Se `unita_con_conduttore` > 0, mostrare per quelle unità
anche "a carico proprietà" e "a carico conduttore" e dire che il foglio `Conduttori` del
prospetto riporta la divisione spesa per spesa. Ricordare che verso il condominio il dovuto è
tutto del proprietario: la parte dell'inquilino è un'indicazione per il loro rapporto interno. Con profilo `autogestione`, mostrare anche un esempio
del calcolo per una riga ("ENEL 412,50 € × 250/1000 = 103,13 €") e ricordare che la somma delle
quote coincide al centesimo con la spesa perché i centesimi residui vanno alle unità con il
resto decimale maggiore. Indicare dove è stato salvato il prospetto.

Il prospetto è una **bozza** finché l'assemblea non lo approva: dirlo sempre.

### 4. Dopo l'approvazione (solo su richiesta esplicita: "approvato", "l'assemblea ha approvato")

Chiedere, se non sono già noti:
- **le date delle rate** decise dall'assemblea, una per ogni colonna `Rata N` del prospetto (una sola
  data se il prospetto non ha rate). Non inventarle e non proporre date "tipiche";
- il **riferimento al verbale** (data dell'assemblea);
- se il prospetto **sostituisce** un prospetto già approvato (tipicamente: il consuntivo sostituisce
  il preventivo dello stesso esercizio). In quel caso le rate vecchie restano, marcate non valide, e
  i versamenti già fatti restano imputati.

Poi un solo comando:

```bash
python3 "${CLAUDE_SKILL_DIR}/scripts/riparto.py" approva --dir "<cartella>" \
  --prospetto "prospetti/<file>_bozza.xlsx" --scadenze 2026-10-31,2027-01-31,2027-04-30,2027-07-31 \
  --approvato "<nome>" --verbale "assemblea del 20 settembre 2026" [--esercizio 2026] \
  [--sostituisce "prospetti/<prospetto approvato>.xlsx"]
```

Lo script, dopo aver controllato tutto e prima di scrivere qualsiasi cosa:
1. scrive le rate di ogni unità nel foglio `Rate` del registro riservato (trovato da solo con
   `Percorso registro riservato`): `Dovuto`, `Saldo` e `Scaduto` di `Situazione` si ricalcolano da lì.
   Non aggiornare il `Dovuto` a mano;
2. crea le righe di `Situazione` mancanti per l'esercizio;
3. aggiunge al foglio `Scadenze` una riga `rata` per data (senza importi per unità: il foglio è
   pubblico) e una riga di Diario;
4. per ultimo toglie `_bozza` dal nome del prospetto (è l'unico caso in cui si rinomina un file in
   `prospetti/`).

Se esce con `"ok": false`, mostrare gli errori così come sono: numero di scadenze diverso dalle
rate, date non in ordine, prospetto già approvato. Non aggirarli. Se `dovuti_precedenti_riportati`
non è vuoto, dire che il dovuto scritto prima della 0.4 per quelle unità è stato conservato come
"rata 0" senza scadenza.

Infine proporre alla skill `scadenze` di mettere in calendario le righe in `scadenze_create` e alla
skill `comunicazioni` l'invio del prospetto approvato ai condomini.

## Regole

- Mai modificare i millesimi per "far tornare i conti". Se una tabella non somma a 1000, il
  problema è nel registro e va risolto lì, con l'utente.
- Spese `UNITA:<ID>` (a carico di uno solo) vanno mostrate separatamente nel riepilogo in chat,
  perché sono quelle su cui nascono le discussioni.
- Proprietario e conduttore: la divisione usa `Quota conduttore %` di ogni spesa (vuoto = tutto al
  proprietario) e vale solo per le unità con `Conduttore` in `Anagrafica`. I millesimi non si
  dividono mai tra proprietario e inquilino. Prima del riparto, se ci sono unità affittate,
  `registro.py verifica` segnala le spese senza percentuale: proporre di completarle.
- Il riparto usa le tabelle così come sono: non applica di propria iniziativa l'art. 1124 (metà
  valore / metà altezza) né altre regole. Se l'utente chiede "è giusto usare la tabella B per
  questo?", rispondere con il riferimento normativo e invitare a verificare il regolamento.
- Rendiconto consuntivo (art. 1130-bis c.c.): il prospetto di riparto è una parte; il rendiconto
  completo richiede anche riepilogo finanziario (entrate, uscite, fondi) e nota esplicativa.
  Se l'utente chiede il rendiconto, generare il riparto e poi proporre di redigere la nota
  esplicativa in un documento separato in `prospetti/`, sulla base di Spese e Versamenti.
