from fastapi import FastAPI, UploadFile, File
from pydantic import BaseModel
import shutil
import os

app = FastAPI()

PLANILHA_PATH = "/data/Planilha_fazenda.xlsm"

@app.get("/")
def home():
    return {"status": "gado-agent rodando"}

class RunRequest(BaseModel):
    text: str

@app.post("/run")
def run(req: RunRequest):
    return {"reply": f"Recebi sua mensagem: {req.text}"}

@app.post("/upload")
def upload_planilha(file: UploadFile = File(...)):
    os.makedirs("/data", exist_ok=True)
    with open(PLANILHA_PATH, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    return {"status": "planilha enviada com sucesso", "path": PLANILHA_PATH}
