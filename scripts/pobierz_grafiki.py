"""Szuka zdjęć w Wikimedia Commons i pobiera te z wolną licencją.

Użycie: python scripts/pobierz_grafiki.py TICKER "fraza" [limit]
Pobrane pliki trafiają do TICKER/zrodla/grafiki/, a ich metadane (autor, licencja,
adres strony pliku) do TICKER/zrodla/grafiki/kandydaci.json.
"""
import html
import json
import re
import sys
from pathlib import Path

import requests

USER_AGENT = "pakiet-analityczny renata2ads@gmail.com"
API = "https://commons.wikimedia.org/w/api.php"
SZEROKOSC = 1200
# Wyniki trafiają do folderu, w którym uruchomiono skill (folder projektu).
KATALOG_GLOWNY = Path.cwd()

DOZWOLONE_LICENCJE = re.compile(r"^(cc0|public domain|pd\b|cc by(-sa)? \d)", re.IGNORECASE)
WYKLUCZONE_NAZWY = re.compile(r"logo|screenshot|screen shot|zrzut|icon|wordmark", re.IGNORECASE)


def bez_html(tekst):
    return html.unescape(re.sub(r"<[^>]+>", "", tekst or "")).strip()


def szukaj(fraza, limit):
    params = {
        "action": "query", "format": "json", "generator": "search",
        "gsrsearch": f"{fraza} filetype:bitmap", "gsrnamespace": 6, "gsrlimit": limit,
        "prop": "imageinfo", "iiprop": "url|extmetadata|mime", "iiurlwidth": SZEROKOSC,
    }
    odp = requests.get(API, params=params, headers={"User-Agent": USER_AGENT}, timeout=60)
    odp.raise_for_status()
    return odp.json().get("query", {}).get("pages", {}).values()


def main():
    if len(sys.argv) not in (3, 4):
        sys.exit('Użycie: python scripts/pobierz_grafiki.py TICKER "fraza" [limit]')
    ticker, fraza = sys.argv[1].upper(), sys.argv[2]
    limit = int(sys.argv[3]) if len(sys.argv) == 4 else 8
    katalog = KATALOG_GLOWNY / ticker / "zrodla" / "grafiki"
    katalog.mkdir(parents=True, exist_ok=True)
    plik_kandydatow = katalog / "kandydaci.json"
    kandydaci = json.loads(plik_kandydatow.read_text(encoding="utf-8")) if plik_kandydatow.exists() else {}

    for strona in szukaj(fraza, limit):
        tytul = strona["title"]
        info = strona["imageinfo"][0]
        meta = info.get("extmetadata", {})
        licencja = bez_html(meta.get("LicenseShortName", {}).get("value"))
        if info.get("mime") == "image/svg+xml" or WYKLUCZONE_NAZWY.search(tytul):
            print(f"pomijam (logo/zrzut/SVG): {tytul}")
            continue
        if not DOZWOLONE_LICENCJE.search(licencja):
            print(f"pomijam (licencja {licencja!r}): {tytul}")
            continue
        autor = bez_html(meta.get("Artist", {}).get("value")) or "nieznany"
        nazwa = re.sub(r"[^A-Za-z0-9._-]+", "_", tytul.removeprefix("File:"))
        nazwa = Path(nazwa).stem[:80] + ".jpg"
        obraz = requests.get(info.get("thumburl") or info["url"], headers={"User-Agent": USER_AGENT}, timeout=120)
        obraz.raise_for_status()
        (katalog / nazwa).write_bytes(obraz.content)
        kandydaci[nazwa] = {
            "tytul": tytul, "autor": autor, "licencja": licencja,
            "url": info["descriptionurl"], "fraza": fraza,
            "opis_commons": bez_html(meta.get("ImageDescription", {}).get("value"))[:300],
        }
        print(f"pobrano: {nazwa} | {autor} | {licencja}")

    plik_kandydatow.write_text(json.dumps(kandydaci, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
