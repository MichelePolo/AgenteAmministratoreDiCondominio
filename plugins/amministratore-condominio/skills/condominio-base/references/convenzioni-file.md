# Convenzioni per i nomi dei file in `archivio/`

Un condomino che sfoglia l'archivio da browser deve capire cosa contiene ogni file dal solo nome.

## Schema

```
<data>_<tipo>_<controparte>_<oggetto>[_<importo>].<ext>
```

- `data`: ISO `AAAA-MM-GG`, la data **del documento** (non quella di caricamento).
- `tipo`: uno tra `fattura`, `preventivo`, `verbale`, `comunicazione`, `contratto`,
  `ricevuta`, `polizza`, `bolletta`, `foto`, `altro`.
- `controparte`: fornitore o mittente, in minuscolo, senza spazi (usare `-`), max 20 caratteri.
- `oggetto`: 2–4 parole in minuscolo con `-`.
- `importo`: solo per documenti con importo, punto decimale, due cifre: `412.50`.
- Estensione originale, in minuscolo.

Solo lettere non accentate, cifre, `-`, `_` e `.`. Niente spazi.

## Esempi

```
2026-03-14_fattura_enel_luce-scale_412.50.pdf
2026-02-02_preventivo_rossi-impianti_sostituzione-caldaia_8900.00.pdf
2026-04-20_verbale_assemblea_ordinaria-approvazione-rendiconto.pdf
2026-01-10_contratto_otis_manutenzione-ascensore.pdf
2026-05-03_foto_portone_danno-cerniera.jpg
```

## Cartelle

```
archivio/<anno>/fatture/
archivio/<anno>/preventivi/
archivio/<anno>/verbali/
archivio/<anno>/comunicazioni/
archivio/<anno>/contratti/
archivio/<anno>/altro/
archivio/<anno>/altro/duplicati/
```

`<anno>` = anno della data del documento. I contratti pluriennali vanno nell'anno di firma.

## Prospetti

```
prospetti/<data>_riparto_<titolo>_bozza.xlsx     ← generato da riparto.py, in attesa di approvazione
prospetti/<data>_riparto_<titolo>.xlsx           ← dopo l'approvazione (si toglie solo il suffisso)
prospetti/<data>_verbale_<oggetto>.md            ← bozza di verbale
```

## Cosa NON fare

- Non rinominare file già in `archivio/`. Se il nome è sbagliato, annotarlo nel Diario e,
  se serve, aggiungere una riga in `Indice` con la nota; il file resta dov'è.
- Non spostare file fuori da `da analizzare/` senza aver mostrato la proposta all'utente.
- Non eliminare mai nulla. I duplicati vanno in `altro/duplicati/`.
