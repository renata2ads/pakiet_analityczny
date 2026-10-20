"""Bramka 2: inny model (Claude Sonnet, osobny proces, bez narzędzi) sprawdza, czy cytat potwierdza wpis.

Użycie: python scripts/weryfikuj_modelem.py TICKER
Wymaga wcześniejszego uruchomienia bramki 1 (TICKER/weryfikacja/bramka1.json).
"""
import json
import re
import shutil
import subprocess
import sys

from wspolne import KATALOG_GLOWNY, katalog_spolki, ujednolic, wczytaj_ledger

PARTIA = 6
LIMIT_CZASU = 600
MAKS_ZNAKOW_PLIKU = 20000

INSTRUKCJA = """Jesteś niezależnym kontrolerem danych finansowych. Dostajesz wpisy z rejestru faktów o spółce
oraz pełny tekst stron źródłowych. Dla każdego wpisu sprawdź, czy cytat i tekst źródła naprawdę potwierdzają wpis:
- liczbę (wartosc),
- okres (rok lub datę),
- jednostkę (np. mln czy mld USD, %),
- zakres (cała spółka czy segment),
- czy opis wpisu pasuje do cytatu.
Werdykt: "potwierdza", "nie potwierdza" albo "niejasne".
Odpowiedz WYŁĄCZNIE tablicą JSON, bez żadnego innego tekstu i bez bloków kodu, w formacie:
[{"id": "F-01", "werdykt": "potwierdza", "uzasadnienie": "jedno zdanie po polsku"}]
"""


def tekst_zrodla(wpis):
    tekst = (KATALOG_GLOWNY / wpis["zrodlo"]["plik"]).read_text(encoding="utf-8")
    if len(tekst) <= MAKS_ZNAKOW_PLIKU:
        return tekst
    # Duże pliki danych rynkowych (np. wieloletni CSV): fragment wokół cytatu.
    pozycja = max(tekst.find(wpis["cytat"]), 0)
    return "[fragment pliku]\n" + tekst[max(0, pozycja - 3000):pozycja + 3000]


def zapytaj_model(partia):
    claude = shutil.which("claude")
    if not claude:
        sys.exit("Nie znaleziono programu claude w PATH.")
    pliki = {}
    for w in partia:
        pliki.setdefault(w["zrodlo"]["plik"], tekst_zrodla(w))
    wpisy = [{k: w.get(k) for k in ("id", "klasa", "opis", "wartosc", "jednostka", "okres", "cytat")}
             | {"strona": w["zrodlo"].get("strona"), "plik": w["zrodlo"]["plik"]} for w in partia]
    tresc = INSTRUKCJA + "\n\nWPISY:\n" + json.dumps(wpisy, ensure_ascii=False, indent=2)
    for plik, tekst in pliki.items():
        tresc += f"\n\n===== ŹRÓDŁO: {plik} =====\n{tekst}"
    wynik = subprocess.run(
        [claude, "-p", "--model", "sonnet", "--tools", "", "--strict-mcp-config",
         "--no-session-persistence", "--output-format", "text"],
        input=tresc, capture_output=True, text=True, encoding="utf-8", timeout=LIMIT_CZASU,
    )
    odp = wynik.stdout.strip()
    m = re.search(r"\[.*\]", odp, re.DOTALL)
    try:
        werdykty = json.loads(m.group(0)) if m else []
    except json.JSONDecodeError:
        werdykty = []
    znalezione = {v.get("id"): v for v in werdykty}
    return [znalezione.get(w["id"]) or {
        "id": w["id"], "werdykt": "niejasne",
        "uzasadnienie": f"Model nie zwrócił werdyktu (kod {wynik.returncode}): {(odp or wynik.stderr)[:200]}"}
        for w in partia]


def bramka2(ticker):
    katalog = katalog_spolki(ticker) / "weryfikacja"
    b1 = {w["id"]: w["wynik"] for w in json.loads((katalog / "bramka1.json").read_text(encoding="utf-8"))}
    do_sprawdzenia = [w for w in wczytaj_ledger(ticker)
                      if w["klasa"] in ("fakt", "dana_rynkowa") and b1.get(w["id"]) == "ok"]
    wyniki = []
    for i in range(0, len(do_sprawdzenia), PARTIA):
        partia = do_sprawdzenia[i:i + PARTIA]
        print(f"Bramka 2: partia {i // PARTIA + 1} ({', '.join(w['id'] for w in partia)})", flush=True)
        wyniki += zapytaj_model(partia)
    (katalog / "bramka2.json").write_text(json.dumps(wyniki, ensure_ascii=False, indent=2), encoding="utf-8")
    return wyniki


def main():
    if len(sys.argv) != 2:
        sys.exit("Użycie: python scripts/weryfikuj_modelem.py TICKER")
    wyniki = bramka2(sys.argv[1])
    for w in wyniki:
        if w["werdykt"] != "potwierdza":
            print(f"{w['id']}: {w['werdykt']} - {w['uzasadnienie']}")
    ok = sum(1 for w in wyniki if w["werdykt"] == "potwierdza")
    print(f"Bramka 2: {ok}/{len(wyniki)} wpisów potwierdzonych")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
