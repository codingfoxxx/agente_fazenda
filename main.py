from fastapi import FastAPI, UploadFile, File
from pydantic import BaseModel
import shutil
import os

app = FastAPI(title="Agente Fazenda")

PLANILHA_PATH = "/data/Planilha_fazenda.xlsm"


class RunRequest(BaseModel):
    text: str


@app.get("/")
def home():
    return {"status": "gado-agent rodando"}


@app.post("/run")
def run(req: RunRequest):
    return {"reply": f"Recebi sua mensagem: {req.text}"}


@app.post("/upload")
async def upload_planilha(file: UploadFile = File(...)):
    os.makedirs("/data", exist_ok=True)

    if not file.filename:
        return {"status": "erro", "detail": "Nenhum arquivo enviado."}

    if not file.filename.lower().endswith(".xlsm"):
        return {"status": "erro", "detail": "Envie um arquivo .xlsm"}

    with open(PLANILHA_PATH, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return {
        "status": "planilha enviada com sucesso",
        "filename": file.filename,
        "path": PLANILHA_PATH
    }
