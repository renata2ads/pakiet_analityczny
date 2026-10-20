"""Sprawdza, czy spółka pasuje do pakietu analitycznego: składa 10-K w SEC i nie jest bankiem ani ubezpieczycielem.

Użycie: python scripts/sprawdz_spolke.py TICKER
Kod wyjścia 0 = spółka pasuje, 1 = nie pasuje (komunikat po polsku).
"""
import sys

import requests

from pobierz_10k import NAGLOWKI

# Kody SIC instytucji finansowych, do których prosty model DCF (FCF firmy) nie pasuje:
# 6000-6199 banki i instytucje kredytowe, 6200-6299 domy maklerskie, 6300-6411 ubezpieczyciele.
SIC_FINANSOWE = range(6000, 6412)


def main():
    if len(sys.argv) != 2:
        sys.exit("Użycie: python scripts/sprawdz_spolke.py TICKER")
    ticker = sys.argv[1].upper()
    tickery = requests.get("https://www.sec.gov/files/company_tickers.json", headers=NAGLOWKI, timeout=60).json()
    wpis = next((w for w in tickery.values() if w["ticker"].upper() == ticker), None)
    if not wpis:
        sys.exit(f"NIE PASUJE: tickera {ticker} nie ma w bazie SEC. Pakiet działa tylko dla spółek "
                 "notowanych w USA, które składają raporty w SEC.")
    cik = int(wpis["cik_str"])
    dane = requests.get(f"https://data.sec.gov/submissions/CIK{cik:010d}.json", headers=NAGLOWKI, timeout=60).json()
    nazwa = dane.get("name", wpis["title"])
    formy = set(dane["filings"]["recent"]["form"])
    if "10-K" not in formy:
        inne = ", ".join(sorted(f for f in formy if f in ("20-F", "40-F", "10-KT")))
        sys.exit(f"NIE PASUJE: {nazwa} ({ticker}) nie składa rocznego raportu 10-K w SEC"
                 + (f" (składa: {inne})" if inne else "") + ". Pakiet opiera się na raporcie 10-K.")
    try:
        sic = int(dane.get("sic") or 0)
    except ValueError:
        sic = 0
    if sic in SIC_FINANSOWE:
        sys.exit(f"NIE PASUJE: {nazwa} ({ticker}) to instytucja finansowa (kod SIC {sic}: "
                 f"{dane.get('sicDescription', '')}). Banków i ubezpieczycieli nie wyceniamy tym modelem DCF, "
                 "bo dług i gotówka są u nich częścią działalności operacyjnej, a nie finansowaniem.")
    print(f"PASUJE: {nazwa} ({ticker}), CIK {cik}, SIC {sic} ({dane.get('sicDescription', '')}), składa 10-K.")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
