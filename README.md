# condominio-agentico — marketplace di plugin Claude

Repository-marketplace per Claude Code, Claude Desktop e Cowork. Contiene il plugin
**amministratore-condominio**: un assistente che aiuta chi amministra un condominio a tenere in
ordine documenti, conti, comunicazioni e scadenze in una cartella Google Drive condivisa con tutti
i condomini. L'assistente prepara, una persona approva.

## Installazione

**Claude Code:**

```
/plugin marketplace add MichelePolo/AgenteAmministratoreDiCondominio
/plugin install amministratore-condominio@condominio-agentico
```

**Claude Desktop / Cowork:** nel pannello *Plugin* aggiungere il marketplace con l'URL di questo
repository (`https://github.com/MichelePolo/AgenteAmministratoreDiCondominio`) e installare
**amministratore-condominio**. In alternativa scaricare `amministratore-condominio.plugin` dalle
[release](https://github.com/MichelePolo/AgenteAmministratoreDiCondominio/releases).

Poi seguire il [TUTORIAL](TUTORIAL.md): in mezz'ora si arriva dal primo avvio al primo riparto.

## Plugin

| Plugin | Descrizione |
|---|---|
| [amministratore-condominio](plugins/amministratore-condominio/) | Amministratore di condominio agentico: archivio condiviso su Drive, riparti millesimali, comunicazioni, scadenze. Guida: [README](plugins/amministratore-condominio/README.md) · [Tutorial](TUTORIAL.md) |

## Requisiti in breve

Claude Code o Claude Desktop con Cowork; Google Drive per desktop; Python 3 con `openpyxl`;
connettori Gmail e Google Calendar di Claude (opzionali, per email e calendario). Nessun server
MCP da installare: il setup del plugin verifica i connettori e spiega come collegarli.

## Sviluppo

```
./build.sh                                   # sincronizza il tutorial, valida, crea il file .plugin
claude plugin validate plugins/amministratore-condominio/.claude-plugin/plugin.json
claude --plugin-dir plugins/amministratore-condominio      # prova locale in Claude Code
/plugin marketplace add ./                   # prova del marketplace da una copia locale
```

`TUTORIAL.md` alla radice è la copia di riferimento; `build.sh` la ricopia dentro il plugin.
Gli script Python si provano con `python3 -m py_compile plugins/amministratore-condominio/skills/*/scripts/*.py`
e, con `openpyxl` installato, creando un condominio di prova:
`python3 plugins/amministratore-condominio/skills/condominio-setup/scripts/crea_registro.py --dir /tmp/prova --esempio`.

Licenza MIT.
