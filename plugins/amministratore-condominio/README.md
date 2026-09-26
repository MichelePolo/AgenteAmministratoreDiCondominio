# Amministratore di condominio agentico

Plugin per Claude (Claude Code, Claude Desktop e Cowork) che aiuta chi amministra un condominio,
un condomino eletto in autogestione o un amministratore di professione, a tenere in ordine
documenti, conti, comunicazioni e scadenze. **L'assistente prepara; una persona approva.**

## L'idea in tre oggetti

| | Cos'è | Chi lo vede |
|---|---|---|
| **Cartella Google Drive condivisa** | i documenti del condominio, con una inbox `da analizzare/` dove chiunque carica fatture e foto | tutti i condomini |
| **`registro-condominio.xlsx`** | si apre sul foglio *In breve* (prossima assemblea e rata, spese dell'anno, ultimi documenti); poi anagrafica, millesimi, spese, scadenze, diario delle operazioni | tutti i condomini |
| **`registro-riservato.xlsx`** | situazione contabile per unità, rate, versamenti, solleciti; accanto, gli estratti conto | solo l'amministratore |

Tutti i condomini vedono la stessa cartella, dal computer o **dal cellulare con l'app Google
Drive**, con o senza intelligenza artificiale. Chi usa il plugin è semplicemente il condomino più
diligente. I file non contengono formule: l'anteprima di Drive mostra sempre i valori giusti.

## Cosa fa

| Skill | Quando la usi | Cosa produce |
|---|---|---|
| `stato` | "come siamo messi?" | in cinque righe: documenti da archiviare, problemi del registro, scadenze, assemblea e termini di legge, rate scadute; le tre cose da fare adesso |
| `condominio-setup` | la prima volta | verifica dei connettori e di Python, cartella strutturata, i due registri, anagrafica e tabelle millesimali, `LEGGIMI.txt` per i condomini |
| `archivio` | "archivia i documenti" | legge `da analizzare/`, propone nomi e classificazione, sposta in `archivio/`, registra le spese |
| `riparto` | "quanto deve ciascuno?", "l'assemblea ha approvato" | prospetto xlsx per unità e per spesa, con rate; somma esatta al centesimo; per gli appartamenti affittati, parte del proprietario e parte dell'inquilino; all'approvazione registra le rate con le loro scadenze |
| `versamenti` | "registra i bonifici dell'estratto conto" | legge l'estratto conto PDF, propone per ogni bonifico unità e rata, registra dopo conferma senza doppioni |
| `comunicazioni` | "convoca l'assemblea", "sollecita", "manda a tutti la situazione" | convocazioni, avvisi, solleciti (solo per rate scadute), verbali, situazione personale di ogni condomino; invio via Gmail solo dopo conferma |
| `scadenze` | "cosa scade?" | scadenzario nel registro allineato a Google Calendar, senza duplicati |
| `condominio-base` | (automatica) | regole, modello dati, riferimenti al codice civile |

## Requisiti

- **Claude Code** oppure **Claude Desktop con Cowork**: serve un ambiente in cui Claude può
  leggere e scrivere i file di una cartella.
- **Google Drive per desktop** installato, con la cartella del condominio sincronizzata.
- **Python 3** con `openpyxl` (`python3 -m pip install openpyxl` se manca; su Linux, se il sistema
  lo impedisce, `sudo apt install python3-openpyxl`).
- Connettori **Gmail** e **Google Calendar** di Claude collegati al proprio account Google
  (opzionali: senza, archivio e riparto funzionano lo stesso). Il setup li verifica e spiega come
  collegarli; dettagli in `skills/condominio-setup/references/connettori.md`.

## Installazione

**Claude Code** (consigliato):

```
/plugin marketplace add MichelePolo/AgenteAmministratoreDiCondominio
/plugin install amministratore-condominio@condominio-agentico
```

**Claude Desktop / Cowork**: nel pannello *Plugin* aggiungere il marketplace con l'URL
`https://github.com/MichelePolo/AgenteAmministratoreDiCondominio` e installare
`amministratore-condominio`. In alternativa, scaricare `amministratore-condominio.plugin` dalle
[release](https://github.com/MichelePolo/AgenteAmministratoreDiCondominio/releases) e caricarlo
dove l'app offre *Carica plugin*; in Claude Code lo stesso file si prova con
`claude --plugin-dir amministratore-condominio.plugin`.

Poi: aprire come cartella di lavoro la cartella del condominio in Google Drive e scrivere
"configura il condominio". Il [TUTORIAL](TUTORIAL.md) accompagna passo per passo.

## Principi di sicurezza

- Nessuna email parte senza conferma esplicita. Alla prima installazione la **modalità collaudo**
  è attiva: tutte le email arrivano solo all'amministratore.
- Nulla viene cancellato o sovrascritto: i duplicati si spostano, i prospetti hanno sempre un
  nome nuovo, gli errori si annotano.
- Ogni operazione è scritta nel foglio `Diario`, leggibile da tutti i condomini. Per versamenti,
  solleciti e situazioni personali il Diario riporta solo quante operazioni, mai nomi o importi.
- I dati per unità (saldi, rate, morosità) stanno solo nel registro riservato e non entrano mai in
  comunicazioni collettive; gli estratti conto bancari stanno nella cartella riservata.
- I bonifici non vengono mai attribuiti "a intuito": se gli indizi sono deboli o due unità sono
  possibili, l'assistente chiede.
- Il plugin cita gli articoli del codice civile pertinenti ma **non fornisce consulenza legale**.

## Limiti noti (versione 0.4)

- Il registro è un file `.xlsx`, non un Google Sheet nativo: i condomini lo consultano da browser
  o dall'app Drive, ma la modifica va fatta dall'amministratore. È una scelta: evita un connettore
  Google Sheets, che oggi richiede un progetto Google Cloud.
- Le automazioni programmate (es. archiviazione ogni mattina) non sono incluse; il plugin è
  pensato per essere avviato dall'utente.
- Le tabelle millesimali vanno fornite dall'utente (dettate, da PDF o a mano); il plugin non le
  calcola dai dati catastali.
- La divisione tra proprietario e inquilino usa le percentuali standard dell'art. 9 L. 392/1978,
  proposte spesa per spesa e modificabili; non legge i contratti di locazione.
- I movimenti bancari si leggono da PDF o si dettano; l'import di file CSV/Excel dall'home banking
  non c'è ancora. Gli addebiti dell'estratto conto non segnano le spese come pagate, e
  l'abbinamento non ricorda chi paga per chi (un familiare con cognome diverso va indicato ogni volta).

## Struttura

```
amministratore-condominio/
  .claude-plugin/plugin.json
  skills/
    condominio-base/        ← modello dati, principi, normativa, scripts/registro.py
    condominio-setup/       ← scripts/crea_registro.py, assets/LEGGIMI.txt, references/connettori.md
    stato/                  ← scripts/stato.py
    archivio/               ← scripts/scansiona.py
    riparto/                ← scripts/riparto.py (anche `approva`)
    versamenti/             ← scripts/abbina.py
    comunicazioni/          ← references/modelli.md, scripts/situazione_personale.py
    scadenze/
  README.md · TUTORIAL.md · LICENSE
```

## Licenza

MIT. Contributi benvenuti: aprire una issue o una pull request.
