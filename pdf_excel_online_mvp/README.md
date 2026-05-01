# PDF Scan to Excel - MVP Online

Versione web semplice per caricare PDF scansionati, leggere il contenuto con OCR e scaricare un file Excel.

## Funzioni incluse

- Upload PDF da browser
- OCR italiano + inglese con Tesseract
- Pre-elaborazione immagine base
- Estrazione iniziale di:
  - numero documento probabile
  - data documento probabile
  - totale probabile
  - righe contenenti importi
- Export Excel con fogli:
  - Documento
  - Righe_Rilevate
  - Testo_OCR
  - Da_Verificare

## Avvio rapido con Docker

Prerequisiti:

- Docker
- Docker Compose

Comando:

```bash
docker compose up --build
```

Poi apri:

```text
http://localhost:5173
```

Backend API:

```text
http://localhost:8000/docs
```

## Avvio manuale backend

Su Linux/Mac devi installare anche:

```bash
sudo apt-get install tesseract-ocr tesseract-ocr-ita poppler-utils
```

Poi:

```bash
cd backend
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

## Avvio manuale frontend

```bash
cd frontend
npm install
npm run dev
```

## Deploy online semplice

### Opzione consigliata

- Backend: Render, Railway o Azure App Service
- Frontend: Vercel o Netlify

Variabile ambiente frontend:

```bash
VITE_API_BASE=https://URL-DEL-TUO-BACKEND
```

## Limiti della versione MVP

Questa versione usa OCR open-source. È utile per partire subito, ma su scansioni complesse, tabelle difficili, timbri o documenti molto disordinati può sbagliare.

Per una versione professionale conviene collegare:

- Azure Document Intelligence
- Google Document AI
- AWS Textract

## Evoluzioni consigliate

- Login utenti
- Storico documenti
- Upload multiplo
- Revisione manuale dei campi
- Template per fatture, DDT, ordini
- Integrazione email
- OCR cloud avanzato
- Database PostgreSQL
- Pagamenti SaaS
