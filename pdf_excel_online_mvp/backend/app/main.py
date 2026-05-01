from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pathlib import Path
from uuid import uuid4
from datetime import datetime
import re
import shutil

from pdf2image import convert_from_path
import pytesseract
import google.generativeai as genai
import os
from PIL import Image, ImageOps, ImageFilter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "storage" / "uploads"
OUTPUT_DIR = BASE_DIR / "storage" / "outputs"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="PDF Scan to Excel Online MVP")
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


def preprocess_image(image: Image.Image) -> Image.Image:
    image = image.convert("L")
    image = ImageOps.autocontrast(image)
    image = image.filter(ImageFilter.SHARPEN)
    return image


def parse_text(text: str, filename: str) -> dict:
    clean = re.sub(r"[ \t]+", " ", text)
    lines = [line.strip() for line in clean.splitlines() if line.strip()]

    date_match = re.search(r"\b(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})\b", clean)
    doc_match = re.search(r"(?:fattura|invoice|documento|doc\.?|n\.?|numero)\s*[:#\-]?\s*([A-Z0-9\-/]{3,})", clean, re.IGNORECASE)

    money_values = re.findall(r"\b\d{1,3}(?:[\.\s]\d{3})*(?:,\d{2})\b|\b\d+\.\d{2}\b", clean)

    possible_rows = []
    for line in lines:
        if re.search(r"\d", line) and len(line) > 8:
            amounts = re.findall(r"\b\d{1,3}(?:[\.\s]\d{3})*(?:,\d{2})\b|\b\d+\.\d{2}\b", line)
            if amounts:
                possible_rows.append({
                    "descrizione": line,
                    "importo_rilevato": amounts[-1]
                })

    return {
        "file": filename,
        "data_elaborazione": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "numero_documento": doc_match.group(1) if doc_match else "",
        "data_documento": date_match.group(1) if date_match else "",
        "totale_probabile": money_values[-1] if money_values else "",
        "righe": possible_rows[:200],
        "testo_ocr": "\n".join(lines[:500]),
        "da_verificare": []
    }


def build_excel(result: dict, output_path: Path):
    wb = Workbook()
    ws_doc = wb.active
    ws_doc.title = "Documento"
    ws_rows = wb.create_sheet("Righe_Rilevate")
    ws_text = wb.create_sheet("Testo_OCR")
    ws_check = wb.create_sheet("Da_Verificare")

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    def style_header(ws):
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

    ws_doc.append(["Campo", "Valore"])
    ws_doc.append(["File", result["file"]])
    ws_doc.append(["Data elaborazione", result["data_elaborazione"]])
    ws_doc.append(["Numero documento", result["numero_documento"]])
    ws_doc.append(["Data documento", result["data_documento"]])
    ws_doc.append(["Totale probabile", result["totale_probabile"]])
    style_header(ws_doc)

    ws_rows.append(["N", "Descrizione / riga OCR", "Importo rilevato"])
    for idx, row in enumerate(result["righe"], start=1):
        ws_rows.append([idx, row["descrizione"], row["importo_rilevato"]])
    style_header(ws_rows)

    ws_text.append(["Testo OCR completo/parziale"])
    for line in result["testo_ocr"].splitlines():
        ws_text.append([line])
    style_header(ws_text)

    ws_check.append(["Controllo", "Stato"])
    checks = [
        ("Numero documento trovato", "OK" if result["numero_documento"] else "DA VERIFICARE"),
        ("Data documento trovata", "OK" if result["data_documento"] else "DA VERIFICARE"),
        ("Totale probabile trovato", "OK" if result["totale_probabile"] else "DA VERIFICARE"),
        ("Righe con importi trovate", "OK" if result["righe"] else "DA VERIFICARE"),
    ]
    for c in checks:
        ws_check.append(list(c))
    style_header(ws_check)

    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 80)
        ws.freeze_panes = "A2"

    wb.save(output_path)


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/api/convert")
async def convert(file: UploadFile = File(...)):
    contents = await file.read()

    model = genai.GenerativeModel("gemini-2.5-flash")

    prompt = """
    Analizza questo DDT scansionato come un operatore amministrativo esperto.

Devi estrarre SOLO i dati stampati del documento.
Ignora completamente:
- scritte a mano
- firme
- timbri
- segni di spunta
- correzioni a penna
- note manuali
- rumore della scansione

Rispondi SOLO in JSON valido, senza testo prima o dopo.

Schema obbligatorio:

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

Regole fondamentali:
- leggi la tabella articoli riga per riga
- non inventare dati mancanti
- la quantità deve essere quella nella colonna Q.tà stampata
- il totale_pezzi deve corrispondere al totale stampato nel documento
- se la somma quantità non coincide con totale_pezzi, aggiungi un errore
- usa formato data GG/MM/AAAA
- confidenza da 0 a 100

    """

    response = model.generate_content([
        {"mime_type": "application/pdf", "data": contents},
        prompt
    ])

    import json

text = response.text.strip()

# rimuove eventuali blocchi ```json se Gemini li aggiunge
text = text.replace("```json", "").replace("```", "").strip()

try:
    data = json.loads(text)
except Exception:
    data = {
        "summary": {},
        "righe": [],
        "campi_da_verificare": ["JSON non valido da Gemini"],
        "errori": [text]
    }

job_id = str(uuid4())
output_path = OUTPUT_DIR / f"{job_id}.xlsx"

result = {
    "file": file.filename,
    "data_elaborazione": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "summary": data.get("summary", {}),
    "righe": data.get("righe", []),
    "campi_da_verificare": data.get("campi_da_verificare", []),
    "errori": data.get("errori", []),
    "testo_ai": text
}

build_excel_enterprise(result, output_path)

return {
    "job_id": job_id,
    "download_url": f"/api/download/{job_id}",
    "summary": result["summary"],
    "righe": result["righe"],
    "campi_da_verificare": result["campi_da_verificare"],
    "errori": result["errori"]
}
@app.get("/api/download/{job_id}")
def build_excel_enterprise(result: dict, output_path: Path):
    wb = Workbook()

    ws = wb.active
    ws.title = "Riepilogo"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    def style_header(sheet):
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

    summary = result.get("summary", {})

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

    for r in result.get("righe", []):
        ws_rows.append([
            r.get("codice", ""),
            r.get("descrizione", ""),
            r.get("ean", ""),
            r.get("quantita", ""),
            r.get("confidenza", "")
        ])
    style_header(ws_rows)

    ws_check = wb.create_sheet("Campi_Da_Verificare")
    ws_check.append(["Campo / Nota"])
    for item in result.get("campi_da_verificare", []):
        ws_check.append([str(item)])
    style_header(ws_check)

    ws_err = wb.create_sheet("Errori")
    ws_err.append(["Errore"])
    for item in result.get("errori", []):
        ws_err.append([str(item)])
    style_header(ws_err)

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
def download(job_id: str):
    xlsx_path = OUTPUT_DIR / f"{job_id}.xlsx"
    if not xlsx_path.exists():
        raise HTTPException(status_code=404, detail="File Excel non trovato.")
    return FileResponse(
        path=xlsx_path,
        filename=f"pdf_convertito_{job_id}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
