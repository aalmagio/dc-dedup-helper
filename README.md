# dedup_v2

Script Python per analizzare tracciati Excel di anagrafiche e individuare contatti duplicati.

Funzionalita principali:

- mappatura interattiva delle colonne del tracciato
- identificazione di contatti `sicuramente doppia` e `probabilmente doppia`
- generazione automatica dei link di deduplica per Direct Channel
- riutilizzo della mappatura colonne salvata tra esecuzioni successive

## Avvio

```bash
python dedup.py
```

Lo script cerca file Excel nella cartella corrente, chiede il foglio da elaborare e guida la mappatura delle colonne necessarie.
