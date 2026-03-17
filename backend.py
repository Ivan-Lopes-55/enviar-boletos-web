from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse
from pathlib import Path
import smtplib
import pandas as pd
import tempfile
import shutil
from email.message import EmailMessage
import re

app = FastAPI()

@app.get("/")
def serve_frontend():
    return FileResponse("index.html")

# Normalize text for matching
def normalize(s: str) -> str:
    s = s.lower()
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def load_email_mapping(xlsx_path):
    df = pd.read_excel(xlsx_path, engine='openpyxl', dtype=str).fillna('')
    mapping = {}
    for _, row in df.iterrows():
        nome_raw = row.iloc[0]
        for sep in ["[", "<", "("]:
            nome_raw = nome_raw.split(sep)[0]
        nome_clean = " ".join(nome_raw.split()).strip()
        key = normalize(nome_clean)
        contato = row.iloc[3].strip()
        if key and contato:
            mapping[key] = contato
    return mapping

def send_email_smtp(to, subject, html, attachments, smtp_user, smtp_pass):
    msg = EmailMessage()
    msg['From'] = smtp_user
    msg['To'] = to
    msg['Subject'] = subject
    msg.add_alternative(html, subtype='html')

    for att in attachments:
        with open(att, 'rb') as f:
            msg.add_attachment(f.read(), filename=Path(att).name)

    with smtplib.SMTP('smtp.office365.com', 587) as smtp:
        smtp.starttls()
        smtp.login(smtp_user, smtp_pass)
        smtp.send_message(msg)

@app.post('/enviar-boletos')
async def enviar_boletos(planilha: UploadFile = File(...), pdfs: list[UploadFile] = File(...), smtp_user: str = Form(...), smtp_pass: str = Form(...)):
    tempdir = Path(tempfile.mkdtemp())
    plan_path = tempdir / planilha.filename
    with open(plan_path, 'wb') as f:
        f.write(await planilha.read())

    pdf_paths = []
    for pdf in pdfs:
        p = tempdir / pdf.filename
        with open(p, 'wb') as f:
            f.write(await pdf.read())
        pdf_paths.append(p)

    mapping = load_email_mapping(plan_path)
    enviados, nao = [], []

    for p in pdf_paths:
        nome_pdf = p.stem.replace("Boleto -", "").strip()
        nome_pdf = " ".join(nome_pdf.split()[:-1])
        key_pdf = normalize(nome_pdf)

        found = None
        if key_pdf in mapping:
            found = key_pdf
        else:
            for mk in mapping:
                if mk in key_pdf or key_pdf in mk:
                    found = mk
                    break
        if not found:
            nao.append(p.name)
            continue

        send_email_smtp(mapping[found], f"Boletos {nome_pdf}", "<p>Segue boleto em anexo.</p>", [p], smtp_user, smtp_pass)
        enviados.append(p.name)

    shutil.rmtree(tempdir)
    return JSONResponse({'enviados': enviados, 'nao_encontrados': nao})
