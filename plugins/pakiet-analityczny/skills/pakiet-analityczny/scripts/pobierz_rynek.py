"""Pobiera kurs akcji (Yahoo Finance) i rentowność 10-letnich obligacji USA (FRED).

Użycie: python scripts/pobierz_rynek.py TICKER
"""
import json
import sys
from datetime import date, datetime, timezone
from pathlib import Path

import requests

USER_AGENT = "pakiet-analityczny renata2ads@gmail.com"
# Wyniki trafiają do folderu, w którym uruchomiono skill (folder projektu).
KATALOG_GLOWNY = Path.cwd()


def main():
    if len(sys.argv) != 2:
        sys.exit("Użycie: python scripts/pobierz_rynek.py TICKER")
    ticker = sys.argv[1].upper()
    katalog = KATALOG_GLOWNY / ticker / "zrodla" / "rynek"
    katalog.mkdir(parents=True, exist_ok=True)
    dzis = date.today().isoformat()
    naglowki = {"User-Agent": USER_AGENT}

    url_kurs = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?range=5d&interval=1d"
    odp = requests.get(url_kurs, headers=naglowki, timeout=60)
    odp.raise_for_status()
    dane = odp.json()
    (katalog / f"kurs_{dzis}.json").write_text(
        json.dumps(dane, ensure_ascii=False, indent=2), encoding="utf-8")
    wynik = dane["chart"]["result"][0]
    zamkniecia = wynik["indicators"]["quote"][0]["close"]
    znaczniki = wynik["timestamp"]
    ostatni = max(i for i, c in enumerate(zamkniecia) if c is not None)
    dzien_kursu = datetime.fromtimestamp(znaczniki[ostatni], timezone.utc).date().isoformat()
    kurs = zamkniecia[ostatni]

    url_fred = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10"
    odp = requests.get(url_fred, headers=naglowki, timeout=60)
    odp.raise_for_status()
    (katalog / f"dgs10_{dzis}.csv").write_text(odp.text, encoding="utf-8")
    wiersze = [w.split(",") for w in odp.text.strip().splitlines()[1:]]
    wiersze = [w for w in wiersze if len(w) == 2 and w[1] not in ("", ".")]
    dzien_dgs, rentownosc = wiersze[-1]

    meta = {
        "data_pobrania": dzis,
        "kurs": {"url": url_kurs, "plik": f"kurs_{dzis}.json", "waluta": wynik["meta"].get("currency"),
                 "zamkniecie": kurs, "dzien": dzien_kursu},
        "dgs10": {"url": url_fred, "plik": f"dgs10_{dzis}.csv",
                  "rentownosc_proc": float(rentownosc), "dzien": dzien_dgs},
    }
    (katalog / "info.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"{ticker}: kurs zamknięcia {kurs:.2f} {meta['kurs']['waluta']} ({dzien_kursu})")
    print(f"Rentowność 10-letnich obligacji USA: {rentownosc}% ({dzien_dgs})")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
