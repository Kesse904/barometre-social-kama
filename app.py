"""Application web du baromètre social KAMA CI.

Les collaborateurs remplissent le questionnaire. En local, chaque soumission est écrite
dans la feuille "Saisie des réponses" du classeur Excel ; sur Vercel, elle est stockée
dans Upstash Redis et le classeur est généré à la demande (voir stockage.py).
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
from pydantic import BaseModel, Field, field_validator
from starlette.background import BackgroundTask

import excel_store
import stockage

BASE_DIR = Path(__file__).resolve().parent
MODELE = str(BASE_DIR / "modele" / "Barometre_social_KAMA_CI.xlsx")
EXCEL_PATH = os.environ.get("BAROMETRE_EXCEL", str(Path.home() / "Downloads" / "Barometre_social_KAMA_CI.xlsx"))
ADMIN_CODE = os.environ.get("BAROMETRE_ADMIN_CODE", "")
DEPARTEMENTS = [
    "Direction Générale / Transport",
    "Finance & Comptabilité",
    "Ressources Humaines",
    "Achats",
    "Logistique & Gestion des Stocks",
    "Commercial & Marketing",
    "Informatique / Systèmes d’Information",
    "Projets / BTP",
    "Organisation, Qualité, Conformité & RSE",
    "Audit Interne",
    "Administration / Secrétariat",
    "Stagiaires",
]
# BAROMETRE_SERVICES (séparés par ;) remplace la liste par défaut si elle est définie
SERVICES = [s.strip() for s in os.environ.get("BAROMETRE_SERVICES", "").split(";") if s.strip()] or DEPARTEMENTS

STOCKAGE = stockage.depuis_environnement(EXCEL_PATH, MODELE)
if STOCKAGE is not None and STOCKAGE.mode == "excel" and not Path(EXCEL_PATH).exists():
    raise SystemExit(f"Classeur introuvable : {EXCEL_PATH} (definir BAROMETRE_EXCEL)")

QUESTIONS = excel_store.read_questionnaire(EXCEL_PATH if STOCKAGE and STOCKAGE.mode == "excel" else MODELE)

# Mode Excel uniquement : journal de sauvegarde et réponses en attente (fichier ouvert dans Excel)
DATA_DIR = BASE_DIR / "donnees"
JOURNAL = DATA_DIR / "journal_reponses.jsonl"
EN_ATTENTE = DATA_DIR / "reponses_en_attente.json"
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

    @field_validator("service")
    @classmethod
    def check_service(cls, v):
        if v.strip() not in SERVICES:
            raise ValueError("Département inconnu : choisir dans la liste")
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
    if STOCKAGE is not None and STOCKAGE.mode == "excel":
        DATA_DIR.mkdir(exist_ok=True)
        flush_pending()
        threading.Thread(target=_retry_loop, daemon=True).start()
    yield


app = FastAPI(title="Baromètre social KAMA CI", lifespan=lifespan)


def _stockage():
    if STOCKAGE is None:
        trouvees = sorted(n for n in os.environ if "REDIS" in n or "KV_" in n or "UPSTASH" in n)
        raise HTTPException(
            503,
            "Stockage non configuré : connecter une base Upstash Redis au projet Vercel puis redéployer. "
            f"Variables détectées : {', '.join(trouvees) or 'aucune'}.",
        )
    return STOCKAGE


@app.get("/api/sante")
def sante():
    """Diagnostic public : mode de stockage et noms (jamais les valeurs) des variables Redis."""
    etat = {"stockage": STOCKAGE.mode if STOCKAGE else "non_configure"}
    etat["variables"] = sorted(n for n in os.environ if "REDIS" in n or "KV_" in n or "UPSTASH" in n)
    etat["code_admin_defini"] = bool(ADMIN_CODE)
    if STOCKAGE is not None and STOCKAGE.mode == "redis":
        try:
            STOCKAGE._commande("PING")
            etat["connexion"] = "ok"
        except Exception as exc:  # noqa: BLE001
            etat["connexion"] = f"erreur : {type(exc).__name__}"
    return etat


@app.get("/api/questionnaire")
def questionnaire():
    return {"questions": QUESTIONS, "services": SERVICES}


@app.post("/api/reponses")
def soumettre(rep: Reponse):
    store = _stockage()
    data = rep.model_dump()
    data["service"] = data["service"].strip()
    data["commentaire"] = data["commentaire"].strip()
    if data["anciennete"] == int(data["anciennete"]):
        data["anciennete"] = int(data["anciennete"])

    if store.mode == "redis":
        store.ajouter(data, len(QUESTIONS))
        return {"statut": "enregistre"}

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
    store = _stockage()
    data = excel_store.calcul_resultats(QUESTIONS, store.lister(len(QUESTIONS)))
    data["mode"] = store.mode
    if store.mode == "excel":
        data["fichier"] = EXCEL_PATH
        data["en_attente"] = len(_load_pending())
    else:
        data["fichier"] = "Base Upstash Redis (Excel généré au téléchargement)"
        data["en_attente"] = 0
    return data


class Reinitialisation(BaseModel):
    confirmation: str


@app.post("/api/reinitialiser")
def reinitialiser(demande: Reinitialisation, x_admin_code: str | None = Header(default=None)):
    _check_admin(x_admin_code)
    if demande.confirmation.strip().upper() != "REINITIALISER":
        raise HTTPException(400, "Confirmation invalide : taper REINITIALISER")
    store = _stockage()
    nb = len(store.lister(len(QUESTIONS)))
    try:
        if store.mode == "excel":
            with _pending_lock:
                sauvegarde = store.reinitialiser(len(QUESTIONS))
                _save_pending([])
        else:
            sauvegarde = store.reinitialiser(len(QUESTIONS))
    except PermissionError:
        raise HTTPException(409, "Le fichier Excel est ouvert : fermez-le puis réessayez.")
    return {"supprimees": nb, "sauvegarde": sauvegarde}


@app.get("/api/excel")
def telecharger_excel(x_admin_code: str | None = Header(default=None)):
    _check_admin(x_admin_code)
    store = _stockage()
    path = store.fichier_excel(len(QUESTIONS))
    nettoyage = BackgroundTask(os.remove, path) if store.mode == "redis" else None
    return FileResponse(path, filename="Barometre_social_KAMA_CI.xlsx", background=nettoyage)


@app.get("/")
def index():
    return FileResponse(BASE_DIR / "static" / "index.html")


@app.get("/resultats")
def page_resultats():
    return FileResponse(BASE_DIR / "static" / "resultats.html")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host=os.environ.get("HOST", "0.0.0.0"), port=int(os.environ.get("PORT", "8080")))
