# VVF Weekend Scheduler

Applicazione in Python per la gestione dei turni dei Vigili del Fuoco volontari. Consente di:

- mantenere un'anagrafica completa di autisti e vigili (ruolo, grado, contatti, limiti settimanali);
- configurare vincoli duri/morbidi, coppie preferenziali, ferie e impostazioni speciali (es. regola Varchi/Pogliani);
- generare il piano turni annuale in formato Excel e ICS;
- operare sia con database SQLite sia, in alternativa, con i file di testo legacy.

Una GUI Tkinter (`vvf_gui.py`) permette di gestire tutte le impostazioni e lanciare la generazione senza toccare il codice.

## Requisiti

- Python 3.10+
- Dipendenze Python (installabili con `pip install -r requirements.txt`):
  - `pandas`
  - `openpyxl`
- Per la GUI: Tkinter (su Linux installare il pacchetto di sistema `python3-tk`).

## Struttura principale

- `database.py`: layer SQLite, schema e operazioni di import/export.
- `turnivvf.py`: motore di generazione turni ed esportazione.
- `vvf_gui.py`: interfaccia grafica per la gestione dati e lanciare lo scheduler.
- `requirements.txt`: dipendenze Python.

## Uso rapido

1. **Installazione dipendenze**
   ```bash
   pip install -r requirements.txt
   ```
   Su Linux, se necessario: `sudo apt install python3-tk`.

2. **Avvio GUI**
   ```bash
   python vvf_gui.py
   ```
   - Tab "Personale": gestisci anagrafica (ruoli, contatti, limiti)
   - Tab "Coppie & Vincoli": definisci vincoli duri/morbidi e preferenze
   - Tab "Ferie": inserisci periodi di indisponibilità
   - Tab "Impostazioni": regola giorni pianificati, senior minimi, regola Varchi/Pogliani
   - Tab "Genera turni": scegli anno, seed, cartella output e premi "Genera"

3. **Esecuzione CLI** (senza GUI)
   ```bash
   python turnivvf_fixed.py --year 2025 --db vvf_data.db --out output
   ```
   Opzioni utili:
   - `--import-from-text --autisti autisti.txt --vigili vigili.txt --vigili-senior vigili_senior.txt`
   - `--skip-db` per usare solo i file legacy.

## Note

- I file generati finiscono nella cartella scelta (`turni_<anno>.xlsx`, `turni_<anno>.ics`, `log_<anno>.txt`).
- I messaggi di log riportano eventuali deroghe o applicazioni della regola Varchi/Pogliani.
- La GUI consente di attivare/disattivare la regola speciale con una spunta; i conteggi restano attribuiti all’autista reale.
