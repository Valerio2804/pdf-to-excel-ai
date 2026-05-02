from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pathlib import Path
from uuid import uuid4
from datetime import datetime
from typing import List
import os
import json

import google.generativeai as genai
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "storage" / "outputs"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="PDF to Excel AI Enterprise")

genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

PROMPT_DDT = """
Analizza questo DDT scansionato come un operatore amministrativo esperto.

Estrai SOLO i dati stampati.
Ignora scritte a mano, firme, timbri, segni di spunta, correzioni a penna e rumore.

Rispondi SOLO in JSON valido.

Schema:

{
  "summary": {
    "tipo_documento": "DDT",
    "numero_documento": "",
    "data_documento": "",
    "mittente": "",
    "destinatario": "",
    "indirizzo_destinatario": "",
    "totale_pezzi": 0,
    "confidenza": 0
  },
  "righe": [
    {
      "codice": "",
      "descrizione": "",
      "ean": "",
      "quantita": 0,
      "confidenza": 0
    }
  ],
  "campi_da_verificare": [],
  "errori": []
}

Regole:
- leggi la tabella articoli riga per riga
- non inventare dati
- quantità sempre numerica
- data in formato GG/MM/AAAA
- confidenza da 0 a 100
"""


def clean_json_text(text: str) -> str:
    return text.replace("```json", "").replace("```", "").strip()


def normalize_data(data: dict) -> dict:
    summary = data.get("summary") or {}
    righe = data.get("righe") or []

    rows = []
    for row in righe:
        if not isinstance(row, dict):
            continue

        try:
            quantita = int(row.get("quantita") or 0)
        except Exception:
            quantita = 0

        try:
            confidenza = int(row.get("confidenza") or 0)
        except Exception:
            confidenza = 0

        rows.append({
            "codice": str(row.get("codice") or "").strip(),
            "descrizione": str(row.get("descrizione") or "").strip(),
            "ean": str(row.get("ean") or "").strip(),
            "quantita": quantita,
            "confidenza": confidenza,
        })

    try:
        totale_pezzi = int(summary.get("totale_pezzi") or 0)
    except Exception:
        totale_pezzi = 0

    try:
        confidenza_summary = int(summary.get("confidenza") or 0)
    except Exception:
        confidenza_summary = 0

    result = {
        "summary": {
            "tipo_documento": summary.get("tipo_documento") or "DDT",
            "numero_documento": str(summary.get("numero_documento") or "").strip(),
            "data_documento": str(summary.get("data_documento") or "").strip(),
            "mittente": str(summary.get("mittente") or "").strip(),
            "destinatario": str(summary.get("destinatario") or "").strip(),
            "indirizzo_destinatario": str(summary.get("indirizzo_destinatario") or "").strip(),
            "totale_pezzi": totale_pezzi,
            "confidenza": confidenza_summary,
        },
        "righe": rows,
        "campi_da_verificare": data.get("campi_da_verificare") or [],
        "errori": data.get("errori") or [],
    }

    somma = sum(r["quantita"] for r in rows)

    if totale_pezzi and somma and totale_pezzi != somma:
        result["errori"].append(
            f"Somma quantità righe ({somma}) diversa da totale pezzi ({totale_pezzi})"
        )

    if not result["summary"]["numero_documento"]:
        result["campi_da_verificare"].append("Numero documento mancante")

    if not result["summary"]["data_documento"]:
        result["campi_da_verificare"].append("Data documento mancante")

    if not rows:
        result["campi_da_verificare"].append("Nessuna riga articolo rilevata")

    return result


def build_excel_enterprise(result: dict, output_path: Path):
    wb = Workbook()

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    def style_header(sheet):
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

    summary = result.get("summary", {})

    ws = wb.active
    ws.title = "Riepilogo"
    ws.append(["Campo", "Valore"])
    ws.append(["File", result.get("file", "")])
    ws.append(["Data elaborazione", result.get("data_elaborazione", "")])
    ws.append(["Tipo documento", summary.get("tipo_documento", "")])
    ws.append(["Numero documento", summary.get("numero_documento", "")])
    ws.append(["Data documento", summary.get("data_documento", "")])
    ws.append(["Mittente", summary.get("mittente", "")])
    ws.append(["Destinatario", summary.get("destinatario", "")])
    ws.append(["Indirizzo destinatario", summary.get("indirizzo_destinatario", "")])
    ws.append(["Totale pezzi", summary.get("totale_pezzi", "")])
    ws.append(["Confidenza", summary.get("confidenza", "")])
    style_header(ws)

    ws_rows = wb.create_sheet("Righe_DDT")
    ws_rows.append(["Codice", "Descrizione", "EAN", "Quantità", "Confidenza"])

    for row in result.get("righe", []):
        ws_rows.append([
            row.get("codice", ""),
            row.get("descrizione", ""),
            row.get("ean", ""),
            row.get("quantita", ""),
            row.get("confidenza", ""),
        ])
    style_header(ws_rows)

    ws_check = wb.create_sheet("Campi_Da_Verificare")
    ws_check.append(["Campo / Nota"])
    for item in result.get("campi_da_verificare", []):
        ws_check.append([str(item)])
    style_header(ws_check)

    ws_errors = wb.create_sheet("Errori")
    ws_errors.append(["Errore"])
    for item in result.get("errori", []):
        ws_errors.append([str(item)])
    style_header(ws_errors)

    ws_ai = wb.create_sheet("Risposta_AI")
    ws_ai.append(["JSON AI"])
    ws_ai.append([result.get("testo_ai", "")])
    style_header(ws_ai)

    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            sheet.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 80)
        sheet.freeze_panes = "A2"

    wb.save(output_path)


async def process_single_pdf(file: UploadFile) -> dict:
    if not file.filename.lower().endswith(".pdf"):
        return {
            "file": file.filename,
            "status": "error",
            "error": "File non PDF"
        }

    contents = await file.read()

    try:
        model = genai.GenerativeModel("gemini-2.5-flash")

        response = model.generate_content([
            {"mime_type": "application/pdf", "data": contents},
            PROMPT_DDT
        ])

        text = clean_json_text(response.text)

        try:
            raw_data = json.loads(text)
        except Exception:
            raw_data = {
                "summary": {},
                "righe": [],
                "campi_da_verificare": ["JSON non valido restituito da Gemini"],
                "errori": [text],
            }

        data = normalize_data(raw_data)

        job_id = str(uuid4())
        output_path = OUTPUT_DIR / f"{job_id}.xlsx"

        result = {
            "file": file.filename,
            "data_elaborazione": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "summary": data.get("summary", {}),
            "righe": data.get("righe", []),
            "campi_da_verificare": data.get("campi_da_verificare", []),
            "errori": data.get("errori", []),
            "testo_ai": json.dumps(data, ensure_ascii=False, indent=2),
        }

        build_excel_enterprise(result, output_path)

        return {
            "file": file.filename,
            "status": "completed",
            "job_id": job_id,
            "download_url": f"/api/download/{job_id}",
            "summary": result["summary"],
            "righe": result["righe"],
            "campi_da_verificare": result["campi_da_verificare"],
            "errori": result["errori"],
        }

    except Exception as e:
        return {
            "file": file.filename,
            "status": "error",
            "error": str(e)
        }


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/api/convert")
async def convert(files: List[UploadFile] = File(...)):
    results = []

    for file in files:
        result = await process_single_pdf(file)
        results.append(result)

    return {
        "count": len(results),
        "documents": results
    }


@app.get("/api/download/{job_id}")
def download(job_id: str):
    xlsx_path = OUTPUT_DIR / f"{job_id}.xlsx"

    if not xlsx_path.exists():
        raise HTTPException(status_code=404, detail="File Excel non trovato.")

    return FileResponse(
        path=xlsx_path,
        filename=f"pdf_convertito_{job_id}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
