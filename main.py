from fastapi import FastAPI
from pydantic import BaseModel

app = FastAPI()

@app.get("/")
def home():
    return {"status": "gado-agent rodando"}

class RunRequest(BaseModel):
    text: str

@app.post("/run")
def run(req: RunRequest):
    return {"reply": f"Recebi sua mensagem: {req.text}"}
