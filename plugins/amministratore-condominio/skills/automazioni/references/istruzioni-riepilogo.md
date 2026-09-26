Riepilogo settimanale del {nome} per l'amministratore. È un compito automatico: nessuno risponde
alle domande, quindi non chiedere nulla e non fare nient'altro.

1. Usa la skill `amministratore-condominio:stato` ed esegui il suo script con l'opzione `--testo`,
   con `--dir` uguale alla cartella di lavoro (è la cartella del condominio). Nessun altro comando.
2. Invia un'email con lo strumento di invio di Gmail:
   - destinatario: {email}. Nessun altro destinatario, nessun Cc, nessun Bcc;
   - oggetto: `[{nome}] Riepilogo del <data>`, dove `<data>` è quella scritta nella prima riga
     del testo (per esempio "28 settembre 2026");
   - corpo: il testo stampato dallo script, **così com'è**, senza aggiunte, tagli o
     riformulazioni; in fondo, dopo una riga vuota, la riga: "Messaggio automatico. Per
     sospenderlo, chiedi a Claude: sospendi il riepilogo settimanale."
3. Se lo script fallisce o non trova il registro, invia comunque l'email a {email}, con oggetto
   `[{nome}] Riepilogo non prodotto` e nel corpo il messaggio di errore così com'è.
4. Non modificare file, non scrivere nel Diario, non usare altri strumenti, non scrivere a
   nessun altro.
