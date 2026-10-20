"""Pobiera najnowszy raport 10-K spółki z SEC i dzieli go na strony.

Użycie: python scripts/pobierz_10k.py TICKER
"""
import json
import re
import sys
from pathlib import Path

import requests
from bs4 import BeautifulSoup

USER_AGENT = "pakiet-analityczny renata2ads@gmail.com"
NAGLOWKI = {"User-Agent": USER_AGENT, "Accept-Encoding": "gzip, deflate"}
# Wyniki trafiają do folderu, w którym uruchomiono skill (folder projektu).
KATALOG_GLOWNY = Path.cwd()

# Znacznik podziału strony: dowolny element ze stylem page-break-before/after.
WZOR_PODZIALU = re.compile(
    r"<(?:hr|div|p)[^>]*page-break-(?:before|after)\s*:\s*always[^>]*>",
    re.IGNORECASE,
)


def pobierz(url):
    odp = requests.get(url, headers=NAGLOWKI, timeout=60)
    odp.raise_for_status()
    return odp


def znajdz_cik(ticker):
    dane = pobierz("https://www.sec.gov/files/company_tickers.json").json()
    for wpis in dane.values():
        if wpis["ticker"].upper() == ticker.upper():
            return int(wpis["cik_str"]), wpis["title"]
    sys.exit(f"Nie znaleziono tickera {ticker} w bazie SEC.")


def najnowszy_10k(cik):
    dane = pobierz(f"https://data.sec.gov/submissions/CIK{cik:010d}.json").json()
    ostatnie = dane["filings"]["recent"]
    for i, forma in enumerate(ostatnie["form"]):
        if forma == "10-K":
            numer = ostatnie["accessionNumber"][i].replace("-", "")
            dokument = ostatnie["primaryDocument"][i]
            return {
                "nazwa_spolki": dane["name"],
                "cik": cik,
                "okres_raportu": ostatnie["reportDate"][i],
                "data_zlozenia": ostatnie["filingDate"][i],
                "numer_zgloszenia": ostatnie["accessionNumber"][i],
                "url": f"https://www.sec.gov/Archives/edgar/data/{cik}/{numer}/{dokument}",
            }
    sys.exit("Spółka nie ma raportu 10-K w ostatnich zgłoszeniach SEC.")


def tekst_strony(html):
    soup = BeautifulSoup(html, "html.parser")
    for el in soup(["script", "style"]):
        el.decompose()
    # Ukryte sekcje XBRL (ix:header) nie są częścią drukowanego dokumentu.
    for el in soup.find_all(style=re.compile(r"display\s*:\s*none", re.I)):
        el.decompose()
    tekst = soup.get_text("\n")
    linie = [re.sub(r"[ \t\xa0]+", " ", l).strip() for l in tekst.splitlines()]
    return "\n".join(l for l in linie if l)


def numer_ze_stopki(tekst):
    """Numer strony wydrukowany na dole: ostatnia krótka linia z samą liczbą."""
    for linia in reversed(tekst.splitlines()[-4:]):
        if re.fullmatch(r"\d{1,3}", linia):
            return int(linia)
    return None


def main():
    if len(sys.argv) != 2:
        sys.exit("Użycie: python scripts/pobierz_10k.py TICKER")
    ticker = sys.argv[1].upper()
    cik, _ = znajdz_cik(ticker)
    info = najnowszy_10k(cik)

    katalog = KATALOG_GLOWNY / ticker / "zrodla" / "10k"
    katalog_stron = katalog / "strony"
    katalog_stron.mkdir(parents=True, exist_ok=True)
    for stary in katalog_stron.glob("*.txt"):
        stary.unlink()

    html = pobierz(info["url"]).content.decode("utf-8", errors="replace")
    (katalog / "oryginal.htm").write_text(html, encoding="utf-8")

    fragmenty = WZOR_PODZIALU.split(html)
    spis = []
    bez_numeru = 0
    for fragment in fragmenty:
        tekst = tekst_strony(fragment)
        if not tekst:
            continue
        numer = numer_ze_stopki(tekst)
        if numer is None:
            bez_numeru += 1
            nazwa = f"bez-numeru-{bez_numeru:03d}.txt"
        else:
            nazwa = f"strona-{numer:03d}.txt"
            if (katalog_stron / nazwa).exists():
                # Powtórzony numer (np. w załącznikach) - nie nadpisujemy.
                bez_numeru += 1
                nazwa = f"bez-numeru-{bez_numeru:03d}.txt"
                numer = None
        (katalog_stron / nazwa).write_text(tekst, encoding="utf-8")
        spis.append({
            "strona": numer,
            "plik": nazwa,
            "poczatek": tekst.replace("\n", " ")[:100],
        })

    info["liczba_stron"] = len(spis)
    info["strony_z_numerem"] = sum(1 for s in spis if s["strona"] is not None)
    (katalog / "info.json").write_text(
        json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    (katalog / "spis.json").write_text(
        json.dumps(spis, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"{info['nazwa_spolki']} (CIK {cik}), 10-K za okres {info['okres_raportu']}, "
          f"złożony {info['data_zlozenia']}")
    print(f"Stron: {info['liczba_stron']}, w tym z numerem w stopce: {info['strony_z_numerem']}")
    print(f"Zapisano w {katalog}")


if __name__ == "__main__":
    main()
