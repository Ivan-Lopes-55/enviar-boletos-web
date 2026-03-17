from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse
from pathlib import Path
import smtplib
import pandas as pd
import tempfile
import shutil
from email.message import EmailMessage

app = FastAPI()

# === ROTA PRINCIPAL: SERVE O INDEX.HTML ===
@app.get("/")
def serve_frontend():
    return FileResponse("index.html")

# === CARREGA PLANILHA ===
def load_email_mapping(xlsx_path):
    df = pd.read_excel(xlsx_path, engine='openpyxl', dtype=str)
    df = df.fillna('')
    mapping = {}
    for _, row in df.iterrows():
        nome = row.iloc[0].strip()
        contato = row.iloc[3].strip()
        if nome and contato:
            mapping[nome.lower()] = contato
    return mapping

# === ENVIO SMTP ===
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

# === ROTA DE ENVIO DE BOLETOS ===
@app.post('/enviar-boletos')
async def enviar_boletos(
    planilha: UploadFile = File(...),
    pdfs: list[UploadFile] = File(...),
    smtp_user: str = Form(...),
    smtp_pass: str = Form(...)
):

    tempdir = Path(tempfile.mkdtemp())

    planilha_path = tempdir / planilha.filename
    with open(planilha_path, 'wb') as f:
        f.write(await planilha.read())

    pdf_paths = []
    for pdf in pdfs:
        p = tempdir / pdf.filename
        with open(p, 'wb') as f:
            f.write(await pdf.read())
        pdf_paths.append(p)

    mapping = load_email_mapping(planilha_path)

    enviados = []
    sem_cliente = []

    for p in pdf_paths:
        cliente = p.stem.split('-')[0].strip().lower()
        if cliente not in mapping:
            sem_cliente.append(p.name)
            continue

        to = mapping[cliente]
        subject = f'Boletos {cliente}'
        html = f'<p>Segue boleto: {p.name}</p>'

        send_email_smtp(to, subject, html, [p], smtp_user, smtp_pass)
        enviados.append(p.name)

    shutil.rmtree(tempdir)

    return JSONResponse({'enviados': enviados, 'nao_encontrados': sem_cliente})
