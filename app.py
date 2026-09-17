"""Application web du baromètre social KAMA CI.

Les collaborateurs remplissent le questionnaire ; chaque soumission est écrite
dans la feuille "Saisie des réponses" du classeur Excel, dont l'onglet
"Résultats" se recalcule automatiquement.
"""
import json
import os
import threading
import time
from contextlib import asynccontextmanager
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field, field_validator

import excel_store

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = os.environ.get("BAROMETRE_EXCEL", str(Path.home() / "Downloads" / "Barometre_social_KAMA_CI.xlsx"))
if not Path(EXCEL_PATH).exists():
    raise SystemExit(f"Classeur introuvable : {EXCEL_PATH} (definir BAROMETRE_EXCEL)")
DATA_DIR = BASE_DIR / "donnees"
DATA_DIR.mkdir(exist_ok=True)
ADMIN_CODE = os.environ.get("BAROMETRE_ADMIN_CODE", "")
JOURNAL = DATA_DIR / "journal_reponses.jsonl"      # copie de sauvegarde de chaque soumission
EN_ATTENTE = DATA_DIR / "reponses_en_attente.json"  # réponses non écrites (Excel ouvert)
SERVICES = [s.strip() for s in os.environ.get("BAROMETRE_SERVICES", "").split(";") if s.strip()]

QUESTIONS = excel_store.read_questionnaire(EXCEL_PATH)
_pending_lock = threading.Lock()


class Reponse(BaseModel):
    service: str = Field(min_length=1, max_length=120)
    anciennete: float = Field(ge=0, le=60)
    notes: list[int]
    commentaire: str = Field(default="", max_length=2000)

    @field_validator("notes")
    @classmethod
    def check_notes(cls, v):
        if len(v) != len(QUESTIONS):
            raise ValueError(f"{len(QUESTIONS)} notes attendues")
        if any(n < 1 or n > 5 for n in v):
            raise ValueError("Chaque note doit être comprise entre 1 et 5")
        return v


def _load_pending() -> list[dict]:
    if EN_ATTENTE.exists():
        return json.loads(EN_ATTENTE.read_text(encoding="utf-8"))
    return []


def _save_pending(items: list[dict]) -> None:
    if items:
        EN_ATTENTE.write_text(json.dumps(items, ensure_ascii=False, indent=1), encoding="utf-8")
    elif EN_ATTENTE.exists():
        EN_ATTENTE.unlink()


def flush_pending() -> int:
    """Tente d'écrire les réponses en attente dans Excel. Renvoie le nombre restant."""
    with _pending_lock:
        pending = _load_pending()
        if not pending:
            return 0
        try:
            excel_store.append_reponses(EXCEL_PATH, pending, len(QUESTIONS))
            pending = []
        except PermissionError:
            pass
        _save_pending(pending)
        return len(pending)


def _retry_loop():
    while True:
        time.sleep(60)
        flush_pending()


@asynccontextmanager
async def lifespan(_app):
    flush_pending()
    threading.Thread(target=_retry_loop, daemon=True).start()
    yield


app = FastAPI(title="Baromètre social KAMA CI", lifespan=lifespan)


@app.get("/api/questionnaire")
def questionnaire():
    return {"questions": QUESTIONS, "services": SERVICES}


@app.post("/api/reponses")
def soumettre(rep: Reponse):
    data = rep.model_dump()
    data["service"] = data["service"].strip()
    data["commentaire"] = data["commentaire"].strip()
    if data["anciennete"] == int(data["anciennete"]):
        data["anciennete"] = int(data["anciennete"])

    with JOURNAL.open("a", encoding="utf-8") as f:
        f.write(json.dumps({"date": datetime.now().isoformat(timespec="seconds"), **data}, ensure_ascii=False) + "\n")

    with _pending_lock:
        pending = _load_pending() + [data]
        try:
            excel_store.append_reponses(EXCEL_PATH, pending, len(QUESTIONS))
            _save_pending([])
            return {"statut": "enregistre"}
        except PermissionError:
            # Le classeur est ouvert dans Excel : on réessaiera automatiquement.
            _save_pending(pending)
            return {"statut": "en_attente"}
        except Exception as exc:  # noqa: BLE001
            _save_pending(pending)
            raise HTTPException(500, f"Erreur d'écriture Excel : {exc}")


def _check_admin(code: str | None) -> None:
    if not ADMIN_CODE or code != ADMIN_CODE:
        raise HTTPException(401, "Code d'accès incorrect")


@app.get("/api/resultats")
def resultats(x_admin_code: str | None = Header(default=None)):
    _check_admin(x_admin_code)
    data = excel_store.read_resultats(EXCEL_PATH, QUESTIONS)
    data["fichier"] = EXCEL_PATH
    data["en_attente"] = len(_load_pending())
    return data


@app.get("/api/excel")
def telecharger_excel(code: str = ""):
    _check_admin(code)
    return FileResponse(EXCEL_PATH, filename=Path(EXCEL_PATH).name)


@app.get("/")
def index():
    return FileResponse(BASE_DIR / "static" / "index.html")


@app.get("/resultats")
def page_resultats():
    return FileResponse(BASE_DIR / "static" / "resultats.html")


app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host=os.environ.get("HOST", "0.0.0.0"), port=int(os.environ.get("PORT", "8080")))
