"""Wspólne funkcje: wczytanie ledgera, ujednolicanie tekstu i blokada generatorów."""
import hashlib
import json
import re
import sys
from pathlib import Path

# Wyniki trafiają do folderu, w którym uruchomiono skill (folder projektu).
KATALOG_GLOWNY = Path.cwd()


def katalog_spolki(ticker):
    return KATALOG_GLOWNY / ticker.upper()


def wczytaj_ledger(ticker):
    return json.loads((katalog_spolki(ticker) / "ledger.json").read_text(encoding="utf-8"))


def skrot_ledgera(ticker):
    return hashlib.sha256((katalog_spolki(ticker) / "ledger.json").read_bytes()).hexdigest()


def ujednolic(tekst):
    """Ujednolica spacje, cudzysłowy i myślniki, żeby porównywać sam tekst."""
    tekst = tekst.replace(" ", " ")
    tekst = re.sub(r"[‘’‚′]", "'", tekst)
    tekst = re.sub(r"[“”„″]", '"', tekst)
    tekst = re.sub(r"[‐-―−]", "-", tekst)
    return re.sub(r"\s+", " ", tekst).strip()


def wymagaj_weryfikacji(ticker):
    """Wywołuje każdy skrypt tworzący dokumenty. Kończy pracę, jeśli ledger nie jest zweryfikowany."""
    plik = katalog_spolki(ticker) / "weryfikacja" / "status.json"
    if not plik.exists():
        sys.exit(f"STOP: ledger {ticker} nie był jeszcze weryfikowany. "
                 f"Uruchom najpierw: python scripts/weryfikuj.py {ticker}")
    status = json.loads(plik.read_text(encoding="utf-8"))
    if not status.get("ok"):
        sys.exit(f"STOP: weryfikacja ledgera {ticker} nie przeszła. Popraw wpisy opisane w "
                 f"{ticker}/weryfikacja/raport.md i uruchom weryfikację ponownie.")
    if status.get("sha256_ledgera") != skrot_ledgera(ticker):
        sys.exit(f"STOP: ledger {ticker} zmienił się po weryfikacji. "
                 f"Uruchom ponownie: python scripts/weryfikuj.py {ticker}")
