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
async def convert_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Carica un file PDF.")

    job_id = str(uuid4())
    pdf_path = UPLOAD_DIR / f"{job_id}.pdf"
    xlsx_path = OUTPUT_DIR / f"{job_id}.xlsx"

    with pdf_path.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        pages = convert_from_path(str(pdf_path), dpi=300)
        full_text = []
        for page in pages:
            processed = preprocess_image(page)
            text = pytesseract.image_to_string(processed, lang="ita+eng")
            full_text.append(text)
        result = parse_text("\n".join(full_text), file.filename)
        build_excel(result, xlsx_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Errore elaborazione OCR: {exc}")

    return {
        "job_id": job_id,
        "filename": file.filename,
        "download_url": f"/api/download/{job_id}",
        "summary": {
            "numero_documento": result["numero_documento"],
            "data_documento": result["data_documento"],
            "totale_probabile": result["totale_probabile"],
            "righe_rilevate": len(result["righe"]),
        }
    }


@app.get("/api/download/{job_id}")
def download(job_id: str):
    xlsx_path = OUTPUT_DIR / f"{job_id}.xlsx"
    if not xlsx_path.exists():
        raise HTTPException(status_code=404, detail="File Excel non trovato.")
    return FileResponse(
        path=xlsx_path,
        filename=f"pdf_convertito_{job_id}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
