# DDT Vision Enterprise

Piattaforma web per convertire DDT PDF scansionati in Excel usando Gemini Vision.

## Funzioni enterprise

- Upload PDF da browser
- Analisi LLM Vision senza Tesseract
- Estrazione JSON strutturata per DDT
- Ignora firme, timbri, scritte a mano e segni di spunta
- Validazione automatica totale pezzi / somma righe
- Excel con fogli:
  - Riepilogo
  - Righe_DDT
  - Campi_Da_Verificare
  - Errori
  - Risposta_AI

## Variabili ambiente backend Render

Imposta queste variabili su Render:

```text
GEMINI_API_KEY=la_tua_chiave
GEMINI_MODEL=gemini-2.5-flash
```

Modello economico alternativo:

```text
GEMINI_MODEL=gemini-2.5-flash-lite
```

## Variabili ambiente frontend Vercel

```text
VITE_API_BASE=https://pdf-to-excel-ai.onrender.com
```

## Avvio locale

```bash
docker compose up --build
```

Frontend:

```text
http://localhost:5173
```

Backend:

```text
http://localhost:8000/health
http://localhost:8000/docs
```

## Deploy

Backend su Render:

```text
Root Directory: pdf_excel_online_mvp/backend
Start Command: uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

Frontend su Vercel:

```text
Root Directory: pdf_excel_online_mvp/frontend
```
