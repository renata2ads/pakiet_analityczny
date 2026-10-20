"""Wyszukuje frazę w stronach raportu 10-K (bez względu na wielkość liter).

Użycie: python scripts/szukaj.py TICKER "fraza"
"""
import re
import sys
from pathlib import Path

# Wyniki trafiają do folderu, w którym uruchomiono skill (folder projektu).
KATALOG_GLOWNY = Path.cwd()
KONTEKST = 80


def main():
    if len(sys.argv) != 3:
        sys.exit('Użycie: python scripts/szukaj.py TICKER "fraza"')
    ticker, fraza = sys.argv[1].upper(), sys.argv[2]
    katalog = KATALOG_GLOWNY / ticker / "zrodla" / "10k" / "strony"
    if not katalog.exists():
        sys.exit(f"Brak stron raportu. Uruchom najpierw: python scripts/pobierz_10k.py {ticker}")

    wzor = re.compile(re.escape(fraza), re.IGNORECASE)
    trafienia = 0
    for plik in sorted(katalog.glob("*.txt")):
        tekst = plik.read_text(encoding="utf-8").replace("\n", " ")
        for m in wzor.finditer(tekst):
            trafienia += 1
            pocz, kon = max(0, m.start() - KONTEKST), min(len(tekst), m.end() + KONTEKST)
            print(f"[{plik.stem}] ...{tekst[pocz:kon]}...")
    print(f"\nTrafień: {trafienia}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
