# Generatore "Piano Vendite Full Immersion"

Crea il file Excel operativo (calendario giorno per giorno, script personalizzati, WhatsApp/email, pipeline e dashboard)
partendo dall'export del portafoglio clienti Allianz (`REPORT.xlsx`).

```bash
python3 tools/piano_vendite_portafoglio.py REPORT.xlsx Piano_Vendite_Full_Immersion.xlsx
```

Parametri principali in testa allo script: `START` (primo lunedì del piano), `CAP_FERIALE` / `CAP_SABATO`
(nuovi contatti al giorno), `AGENTE`.

⚠️ L'export e l'Excel generato contengono dati personali dei clienti: **non vanno mai committati** nel repository.
