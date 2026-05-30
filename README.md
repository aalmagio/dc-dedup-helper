# dc-dedup-helper

Script Python per analizzare tracciati Excel di anagrafiche e individuare contatti duplicati.

Il progetto e stato pensato soprattutto per tracciati provenienti da Mentor. Per questo motivo include, come comodita, la generazione dei link di deduplica Direct Channel gia pronti per l'uso.

Questo pero non lo rende legato a Mentor: chiunque puo ignorare le colonne con i link, usare solo il motore di deduplica, oppure personalizzare il template del link e le altre variabili tramite file di configurazione.

Funzionalita principali:

- mappatura interattiva delle colonne del tracciato
- riutilizzo della mappatura colonne salvata tra esecuzioni successive
- identificazione di contatti `sicuramente doppia` e `probabilmente doppia`
- generazione automatica dei link di deduplica
- configurazione esterna del template URL e di alcune variabili operative

## Requisiti

```bash
pip install -r requirements.txt
```

## Configurazione

Il file `dedup_config.json` contiene le principali variabili personalizzabili:

- `dedup_url_template`: URL usato per costruire i link di deduplica
- `dedup_link_label`: etichetta mostrata nel link Excel
- `saved_mapping_file`: nome del file locale usato per ricordare la mappatura colonne
- `name_similarity_threshold`: soglia di similarita nome/cognome per i casi `probabilmente doppia`

Il template del link usa i placeholder `{id_cedente}` e `{id_vincente}`.

## Avvio

```bash
python dedup.py
```

Lo script cerca file Excel nella cartella corrente, chiede il foglio da elaborare e guida la mappatura delle colonne necessarie.
