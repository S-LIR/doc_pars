from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
from typing import List
import tempfile
import shutil
import os

from pdfminer.high_level import extract_text
from docx import Document
from openpyxl import load_workbook
import xlrd 
import extract_msg

app = FastAPI()

@app.get("/test")
def test():
    return {"message": "сервер работает"}

def parse_pdf(path: str) -> str:
    try:
        return extract_text(path)
    except Exception as e:
        return f"[Ошибка при чтении PDF]: {str(e)}"

def parse_docx(path: str) -> str:
    try:
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as e:
        return f"[Ошибка при чтении DOCX]: {str(e)}"

def parse_xlsx(path: str) -> str:
    try:
        wb = load_workbook(path)
        sheet = wb.active
        rows = []
        for row in sheet.iter_rows(values_only=True):
            line = "\t".join(str(cell) if cell is not None else "" for cell in row)
            rows.append(line)
        return "\n".join(rows)
    except Exception as e:
        return f"[Ошибка при чтении XLSX]: {str(e)}"

def parse_txt(path: str) -> str:
    try:
        with open(path, encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        return f"[Ошибка при чтении TXT]: {str(e)}"

def parse_xls(path: str) -> str:
    try:
        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(0)
        rows = []
        for row_idx in range(sheet.nrows):
            row = sheet.row_values(row_idx)
            line = "\t".join(str(cell) for cell in row)
            rows.append(line)
        return "\n".join(rows)
    except Exception as e:
        return f"[Ошибка при чтении XLS]: {str(e)}"
    
def parse_msg(path: str) -> str:
    try:
        msg = extract_msg.Message(path)
        msg_sender = msg.sender or ""
        msg_subject = msg.subject or ""
        msg_date = msg.date or ""
        msg_body = msg.body or ""
        return f"From: {msg_sender}\nSubject: {msg_subject}\nDate: {msg_date}\n\n{msg_body}"
    except Exception as e:
        return f"[Ошибка при чтении MSG]: {str(e)}"

def parse_file(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return parse_pdf(path)
    elif ext == ".docx":
        return parse_docx(path)
    elif ext == ".xlsx":
        return parse_xlsx(path)
    elif ext == ".xls":
        return parse_xls(path)
    elif ext == ".msg":
        return parse_msg(path)
    elif ext == ".txt":
        return parse_txt(path)
    else:
        return f"[Формат {ext} не поддерживается]"

@app.post("/parse")
async def parse_files(files: List[UploadFile] = File(...)):
    results = []

    for upload in files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=upload.filename) as tmp:
            shutil.copyfileobj(upload.file, tmp)
            tmp_path = tmp.name

        text = parse_file(tmp_path)
        results.append({
            "filename": upload.filename,
            "text": text[:6000]  # можно убрать ограничение, если нужно
        })

    return JSONResponse(content={"parsed_files": results})
