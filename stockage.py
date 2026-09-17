"""Stockage des réponses.

- Mode "excel" (PC / serveur Windows) : les réponses sont écrites directement dans le classeur.
- Mode "redis" (Vercel) : le disque n'est pas persistant, les réponses sont stockées dans
  Upstash Redis et le classeur Excel est généré à la demande à partir du modèle.
"""
import json
import os
import shutil
import tempfile
import urllib.request
from datetime import datetime

import excel_store

REDIS_KEY = "barometre:reponses"


class StockageExcel:
    mode = "excel"

    def __init__(self, path: str):
        self.path = path

    def ajouter(self, reponse: dict, nb_questions: int) -> None:
        excel_store.append_reponses(self.path, [reponse], nb_questions)

    def lister(self, nb_questions: int) -> list[dict]:
        return excel_store.read_reponses(self.path, nb_questions)

    def fichier_excel(self, nb_questions: int) -> str:
        return self.path


class StockageRedis:
    mode = "redis"

    def __init__(self, url: str, token: str, modele: str):
        self.url = url.rstrip("/")
        self.token = token
        self.modele = modele

    def _commande(self, *args):
        req = urllib.request.Request(
            self.url,
            data=json.dumps(args).encode("utf-8"),
            headers={"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=10) as res:
            data = json.load(res)
        if "error" in data:
            raise RuntimeError(data["error"])
        return data["result"]

    def ajouter(self, reponse: dict, nb_questions: int) -> None:
        entree = {"date": datetime.now().isoformat(timespec="seconds"), **reponse}
        self._commande("RPUSH", REDIS_KEY, json.dumps(entree, ensure_ascii=False))

    def lister(self, nb_questions: int) -> list[dict]:
        lignes = []
        for i, brut in enumerate(self._commande("LRANGE", REDIS_KEY, 0, -1)):
            rep = json.loads(brut)
            lignes.append({"ligne": excel_store.FIRST_DATA_ROW + i, "id": i + 1, **rep})
        return lignes

    def fichier_excel(self, nb_questions: int) -> str:
        """Génère le classeur complet (modèle + toutes les réponses) dans un fichier temporaire."""
        fd, path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        shutil.copyfile(self.modele, path)
        reponses = self.lister(nb_questions)
        if reponses:
            excel_store.append_reponses(path, reponses, nb_questions)
        return path


class StockageRedisTCP(StockageRedis):
    """Même stockage, via une URL redis:// ou rediss:// (variable REDIS_URL)."""

    def __init__(self, url: str, modele: str):
        import redis

        self.client = redis.Redis.from_url(url, decode_responses=True, socket_timeout=10)
        self.modele = modele

    def _commande(self, *args):
        return self.client.execute_command(*args)


def _variable(*suffixes: str) -> str | None:
    """Cherche une variable d'environnement par suffixe (Vercel peut ajouter un préfixe)."""
    for suffixe in suffixes:
        if os.environ.get(suffixe):
            return os.environ[suffixe]
    for nom, valeur in sorted(os.environ.items()):
        if valeur and any(nom.endswith("_" + suffixe) for suffixe in suffixes):
            return valeur
    return None


def depuis_environnement(excel_path: str, modele: str):
    url = _variable("KV_REST_API_URL", "UPSTASH_REDIS_REST_URL")
    token = _variable("KV_REST_API_TOKEN", "UPSTASH_REDIS_REST_TOKEN")
    if url and token:
        return StockageRedis(url, token, modele)
    redis_url = _variable("REDIS_URL", "KV_URL")
    if redis_url:
        return StockageRedisTCP(redis_url, modele)
    if os.environ.get("VERCEL"):
        return None  # sur Vercel, écrire sur le disque perdrait les réponses
    return StockageExcel(excel_path)
