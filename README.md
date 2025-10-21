# VVF Weekend Scheduler

Applicazione in Python per la pianificazione dei turni dei Vigili del Fuoco volontari. Il progetto combina una GUI Tkinter e un motore CLI per:

- gestire l’anagrafica di autisti e vigili con ruoli, contatti, livelli di esperienza e limiti settimanali;
- configurare vincoli duri/morbidi (coppie vietate, preferenze, regole speciali, ferie);
- generare i turni su periodi selezionati esportando Excel, ICS e log testuali;
- operare sia con database SQLite sia (opzionalmente) con i file di testo legacy.

## Requisiti

- **Python 3.10+**  
  Installare le dipendenze con:
  ```bash
  pip install -r requirements.txt
  ```
  (Pacchetti principali: `pandas`, `openpyxl`.)

- **Tkinter** per la GUI (su Debian/Ubuntu: `sudo apt install python3-tk`).

## Struttura del progetto

```
vvf_scheduler/         # Motore di pianificazione e export
  core.py              # Algoritmo principale (Scheduler)
  config.py            # Costruzione ProgramConfig da file legacy
  constants.py         # Costanti condivise (nomi giorni, mesi ecc.)
  exports.py           # Esportazione Excel/ICS
  rules.py             # Definizioni regole Hard/Soft/Off
  runner.py            # Funzione esegui(...) usata da CLI e GUI

database.py            # Layer SQLite e accesso dati
turnivvf.py            # Entry point CLI
vvf_gui.py             # Interfaccia grafica Tkinter
requirements.txt       # Dipendenze Python
vvf_data.db            # Database SQLite (se presente)
autisti.txt / vigili.txt / vigili_senior.txt (facoltativi legacy)
```

## Modalità d’uso

### GUI
Avviare l’interfaccia con:
```bash
python vvf_gui.py
```

Tab principali:
- **Personale**: anagrafica (ruoli, contatti, limiti settimanali).
- **Coppie & Vincoli**: gestione coppie vietate/preferite con severità Hard/Soft.
- **Ferie**: configurazione periodi di indisponibilità.
- **Impostazioni**: selezione giorni attivi e pannello “Parametri di generazione” con toggle Hard/Soft/Off per regole come:
  - `Limite turni settimanali`
  - `Esclusione estiva`
  - `Regola Varchi/Pogliani`
  - `Minimo SENIOR in squadra` (valore configurabile)
- **Genera turni**: scelta anno, mesi (con “Tutti i mesi”), seed RNG, cartella output, eventuale import legacy o uso totale dei file di testo.

Lo stato delle regole viene ripristinato ai valori di default a ogni avvio; è possibile salvare le preferenze premendo “Salva impostazioni”.

### CLI

Esempi di utilizzo:
```bash
# Genera l'intero anno usando il database SQLite
python turnivvf.py --year 2025 --db vvf_data.db --out output

# Genera solo Gennaio-Febbraio 2026 con logging dettagliato
python turnivvf.py --year 2026 --months 1 2 --db vvf_data.db --out output --verbose

# Importa i file legacy nel DB prima di generare
python turnivvf.py --year 2025 --db vvf_data.db --out output \
    --import-from-text --autisti autisti.txt --vigili vigili.txt --vigili-senior vigili_senior.txt

# Modalità completamente legacy (senza DB)
python turnivvf.py --year 2025 --skip-db \
    --autisti autisti.txt --vigili vigili.txt --vigili-senior vigili_senior.txt --out output
```

Opzioni principali:
- `--year` anno di riferimento;
- `--months` lista di mesi (1-12) da includere;
- `--out` cartella output (genererà `turni_<anno>.xlsx`, `turni_<anno>.ics`, `log_<anno>.txt`);
- `--seed` per rendere deterministica la generazione;
- `--verbose` abilita log dettagliati su stdout.

## Log e output
- Il log testuale riporta le deroghe applicate (ad esempio quando un vincolo soft viene rilassato).
- L’Excel contiene un foglio per ogni mese generato e un report statistico (autisti e vigili).
- Il file ICS esporta gli eventi per autisti e vigili (timezone Europe/Rome).

## Suggerimenti
- Verificare che `pandas`/`openpyxl` siano installati nell’ambiente usato dalla GUI; in caso contrario la generazione darà errore.
- Se si importano dati legacy, assicurarsi che i file siano codificati in UTF-8 per evitare errori di decodifica.
- La GUI consente di selezionare solo alcuni mesi: utile per rigenerare periodi già approvati senza toccare l’intero anno.

Buon lavoro con lo Scheduler! Per segnalazioni o miglioramenti aprire una issue nel repository. 
