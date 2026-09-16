# Connettori Gmail e Google Calendar

Il plugin non installa server MCP propri: usa i **connettori integrati di Claude** per Gmail e
Google Calendar, che l'utente collega una volta sola al proprio account Google. Senza connettori
il plugin funziona lo stesso per archivio, registro e riparto.

## Come collegarli

| Dove si usa Claude | Cosa fare |
|---|---|
| **Claude Desktop / Cowork** | Impostazioni → *Connettori* → **Gmail** → Connetti (accesso Google). Ripetere per **Google Calendar**. |
| **claude.ai (browser)** | Impostazioni → *Connettori* (claude.ai/settings/connectors) → Gmail e Google Calendar → Connetti. |
| **Claude Code** | I connettori collegati sull'account claude.ai sono disponibili automaticamente quando si è autenticati con lo stesso account. Collegarli da claude.ai come sopra, poi riavviare la sessione; `/mcp` mostra lo stato. |

Nella sessione gli strumenti compaiono con nomi come `mcp__claude_ai_Gmail__send_message` e
`mcp__claude_ai_Google_Calendar__create_event`. Se non compaiono, il connettore non è collegato o
la sessione va riavviata.

## Cosa fa il plugin con i connettori

- **Gmail**: invio delle comunicazioni (convocazioni, avvisi, prospetti, solleciti), sempre e solo
  dopo conferma esplicita e in modalità collaudo finché l'utente non la disattiva. Non legge la
  posta dei condomini.
- **Google Calendar**: creazione e aggiornamento degli eventi del foglio `Scadenze`; mai
  cancellazioni di propria iniziativa.

## Alternativa avanzata: server MCP ufficiali di Google

Google offre server MCP remoti per Gmail (`https://gmailmcp.googleapis.com/mcp/v1`) e Calendar
(`https://calendarmcp.googleapis.com/mcp/v1`), in *Developer Preview*. Richiedono un progetto
Google Cloud, l'abilitazione delle API e dei servizi MCP, una schermata di consenso OAuth con gli
ambiti Gmail/Calendar e un client OAuth "Applicazione web" con URI di reindirizzamento
`https://claude.ai/api/mcp/auth_callback`; ID e segreto del client vanno inseriti come
*connettore personalizzato* in Claude. Sono adatti a chi ha già un progetto Cloud; per tutti gli
altri i connettori integrati bastano. Guida: https://developers.google.com/workspace/guides/configure-mcp-servers
