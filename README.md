# Audit Certificazione ISO – Streamlit App

Applicazione Streamlit per la gestione guidata di audit di certificazione ISO, con compilazione checklist, osservazioni finali, azioni conseguenti ed esportazione dei report.

Questa versione è pensata per la distribuzione su **Streamlit Cloud** o ambienti simili, quindi non utilizza un archivio permanente lato server per le bozze: il salvataggio avviene tramite **download del file JSON** e il ripristino tramite **upload drag & drop** dello stesso file.

---

## Funzionalità principali

- Configurazione audit da sidebar
- Checklist audit filtrata in base a:
  - norma selezionata
  - tipo audit
- Compilazione guidata di:
  - esito
  - audit evidence
  - note
  - azioni richieste
- Sezione osservazioni finali
- Sezione azioni conseguenti
- Export disponibili:
  - PDF
  - report Word
  - marker map JSON
  - DOCX template compilato
  - JSON audit per salvataggio bozza

---

## Logica di salvataggio in ambiente cloud

In ambiente Streamlit distribuito **non viene mantenuto uno storico persistente delle bozze** sul filesystem dell'app.

Per questo motivo il flusso corretto è:

1. compilare l'audit
2. scaricare il file **JSON audit**
3. conservare il file sul proprio PC o su cloud personale
4. quando serve, ricaricare la bozza tramite **drag & drop** nella sidebar

Questa è la logica consigliata per evitare perdite di dati tra sessioni o riavvii dell'app.

---

## File principali del progetto

- `main_cloud_ready.py` → applicazione principale Streamlit
- `checklist_audit_iso.json` → checklist base caricata dall'app
- `Template_CheckList_MARKER.docx` → template Word opzionale per compilazione marker
- `requirements.txt` → dipendenze Python
- `README.md` → documentazione del progetto

---

## Requisiti

Consigliato Python 3.11 o 3.12.

Dipendenze principali:

- streamlit
- pandas
- python-docx
- reportlab

Esempio minimale di `requirements.txt`:

```txt
streamlit
pandas
python-docx
reportlab
```

---

## Avvio in locale

Clonare il repository e installare le dipendenze:

```bash
pip install -r requirements.txt
```

Poi avviare l'app con:

```bash
streamlit run main_cloud_ready.py
```

---

## Deploy su Streamlit Cloud

### 1. Caricare il repository su GitHub
Inserire nel repo almeno questi file:

- `main_cloud_ready.py`
- `checklist_audit_iso.json`
- `Template_CheckList_MARKER.docx` (se usato)
- `requirements.txt`
- `README.md`

### 2. Creare la app su Streamlit Cloud
- collegare il repository GitHub
- selezionare branch e file principale:
  - `main_cloud_ready.py`

### 3. Verificare il deploy
Controllare che:
- la checklist venga caricata correttamente
- il PDF venga generato
- il Word report venga generato
- il caricamento della bozza JSON funzioni
- il template DOCX venga esportato se il file template è presente nel repo

---

## Struttura operativa consigliata per l'utente finale

### Nuovo audit
- compilare i dati audit in sidebar
- compilare la checklist
- compilare osservazioni finali
- compilare eventuali azioni
- scaricare:
  - PDF
  - Word report
  - JSON audit

### Ripresa di un audit in bozza
- aprire l'app
- caricare il file JSON precedentemente scaricato
- continuare la compilazione
- riscaricare il JSON aggiornato

---

## Note importanti

- Il file JSON è la **vera bozza di lavoro**
- Se l'utente non scarica il JSON, la sessione può andare persa
- Il DOCX template compilato richiede la presenza del file `Template_CheckList_MARKER.docx`
- Se il template non è presente, l'app continua a funzionare ma quell'export non sarà disponibile

---

## Migliorie future possibili

- salvataggio su Google Drive / OneDrive / S3
- archivio audit multiutente
- autenticazione accessi
- dashboard storica centralizzata
- firma digitale o workflow approvativo
- esportazione con layout grafico avanzato

---

## Stato del progetto

Versione predisposta per distribuzione cloud con gestione bozza tramite upload/download JSON.
