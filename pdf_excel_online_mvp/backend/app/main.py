from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pathlib import Path
from uuid import uuid4
from datetime import datetime
from typing import List
from io import BytesIO
import os
import json
import re

import google.generativeai as genai
from pypdf import PdfReader, PdfWriter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ================= CONFIG =================

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "storage" / "outputs"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="DDT AI Enterprise")

genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ================= PROMPT =================

PROMPT = """
Analizza questa pagina di un DDT.

Estrai SOLO dati stampati.
Ignora firme, scritte a mano, timbri, spunte.

Rispondi SOLO JSON:

{
  "summary": {
    "numero_documento": "",
    "data_documento": "",
    "mittente": "",
    "destinatario": ""
  },
  "righe": [
    {
      "codice": "",
      "descrizione": "",
      "ean": "",
      "quantita": 0
    }
  ]
}
"""


# ================= UTILS =================

def clean_json(text):
    return text.replace("```json", "").replace("```", "").strip()


def sanitize_filename(name: str):
    name = name.rsplit(".", 1)[0]
    return re.sub(r'[^\w\-. ]', '_', name)


def split_pdf(pdf_bytes):
    reader = PdfReader(BytesIO(pdf_bytes))
    pages = []

    for page in reader.pages:
        writer = PdfWriter()
        writer.add_page(page)

        buffer = BytesIO()
        writer.write(buffer)
        pages.append(buffer.getvalue())

    return pages


# ================= EXCEL =================

def build_excel(result, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Dettaglio_Righe"

    ws.append([
        "File",
        "Pagina",
        "Numero",
        "Data",
        "Mittente",
        "Destinatario",
        "Codice",
        "Descrizione",
        "EAN",
        "Quantità"
    ])

    summary = result["summary"]

    for row in result["righe"]:
        ws.append([
            result["file"],
            row.get("pagina"),
            summary.get("numero_documento"),
            summary.get("data_documento"),
            summary.get("mittente"),
            summary.get("destinatario"),
            row.get("codice"),
            row.get("descrizione"),
            row.get("ean"),
            row.get("quantita")
        ])

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 60)

    wb.save(path)


# ================= CORE =================

async def process_pdf(file: UploadFile):
    content = await file.read()
    pages = split_pdf(content)

    model = genai.GenerativeModel("gemini-2.5-flash")

    all_rows = []
    summary = {}

    for i, page in enumerate(pages, start=1):
        try:
            res = model.generate_content([
                {"mime_type": "application/pdf", "data": page},
                PROMPT
            ])

            data = json.loads(clean_json(res.text))

            if not summary and data.get("summary"):
                summary = data["summary"]

            for r in data.get("righe", []):
                r["pagina"] = i
                all_rows.append(r)

        except Exception as e:
            print("Errore pagina:", i, e)

    job_id = str(uuid4())
    safe_name = sanitize_filename(file.filename)

    output_path = OUTPUT_DIR / f"{job_id}__{safe_name}.xlsx"

    result = {
        "file": file.filename,
        "summary": summary,
        "righe": all_rows
    }

    build_excel(result, output_path)

    return {
        "file": file.filename,
        "job_id": job_id,
        "download_url": f"/api/download/{job_id}"
    }


# ================= API =================

@app.get("/health")
def health():
    return {"ok": True}


@app.post("/api/convert")
async def convert(files: List[UploadFile] = File(...)):
    results = []

    for file in files:
        res = await process_pdf(file)
        results.append(res)

    return {"documents": results}


@app.get("/api/download/{job_id}")
def download(job_id: str):
    files = list(OUTPUT_DIR.glob(f"{job_id}__*.xlsx"))

    if not files:
        raise HTTPException(404, "File non trovato")

    file_path = files[0]
    original_name = file_path.name.split("__", 1)[1]

    return FileResponse(
        file_path,
        filename=original_name
    )
