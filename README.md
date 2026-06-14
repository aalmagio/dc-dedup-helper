# dc-dedup-helper

Script Python per analizzare tracciati Excel di anagrafiche e individuare contatti duplicati.

Il progetto e stato pensato soprattutto per tracciati provenienti da Mentor. Per questo motivo include, come comodita, la generazione dei link di deduplica Direct Channel gia pronti per l'uso.

Questo pero non lo rende legato a Mentor: chiunque puo ignorare le colonne con i link, usare solo il motore di deduplica, oppure personalizzare il template del link e le altre variabili tramite file di configurazione.

Funzionalita principali:

- mappatura interattiva delle colonne del tracciato
- riutilizzo della mappatura colonne salvata tra esecuzioni successive
- identificazione di contatti `sicuramente doppia` e `probabilmente doppia`
- supporto a email e cellulare secondari (`Email2`, `Cellulare2`) con confronto incrociato
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

## Mappatura consigliata per i tracciati Mentor

I tracciati Mentor che usiamo contengono questi campi, da mappare cosi:

| Campo del tracciato | Campo nello script |
|---------------------|--------------------|
| ID Cliente          | ID univoco         |
| Codice Fiscale      | Codice Fiscale     |
| Email               | Email              |
| Email2              | Email secondaria   |
| Cellulare           | Cellulare          |
| Cellulare2          | Cellulare secondario |
| Telefono            | Telefono fisso     |
| Nome                | Nome               |
| Cognome             | Cognome            |

Email/Email2 e Cellulare/Cellulare2 vengono confrontati in modo incrociato: due
righe sono considerate doppie se condividono una qualunque email o un qualunque
cellulare, anche se uno e nel campo principale e l'altro nel secondario.

**Importante:** come "ID univoco" usa sempre **ID Cliente**, non il Codice
Fiscale. La deduplica automatica tramite link in Mentor funziona solo con
l'ID Cliente; i link costruiti sul Codice Fiscale non sono validi.
