"""Bramka 1: sprawdzenie ledgera bez AI (za każdym razem ten sam wynik).

Użycie: python scripts/sprawdz_cytaty.py TICKER
"""
import json
import re
import sys

from wspolne import KATALOG_GLOWNY, katalog_spolki, ujednolic, wczytaj_ledger


def warianty_liczby(wartosc):
    """Zapisy liczby, jakie mogą wystąpić w raporcie: 1234567, 1,234,567, (1,234,567), -1,234,567."""
    modul = abs(wartosc)
    if isinstance(wartosc, float) and not modul.is_integer():
        cyfry = repr(modul)
        z_sep = f"{int(modul):,}" + cyfry[cyfry.index("."):]
    else:
        cyfry = str(int(modul))
        z_sep = f"{int(modul):,}"
    wzory = []
    for zapis in {cyfry, z_sep}:
        z = re.escape(zapis)
        if wartosc < 0:
            wzory += [rf"\(\s*\$?\s*{z}\s*\)", rf"-\s*\$?\s*{z}"]
        else:
            wzory.append(z)
    return [rf"(?<![\d,.]){w}(?![\d]|,\d|\.\d)" for w in wzory]


def wartosc_w_cytacie(wartosc, cytat):
    if isinstance(wartosc, bool):
        return False
    if isinstance(wartosc, (int, float)):
        return any(re.search(w, cytat) for w in warianty_liczby(wartosc))
    return ujednolic(str(wartosc)).lower() in cytat.lower()


def sprawdz_wpis(w):
    klasa = w.get("klasa")
    if klasa in ("fakt", "dana_rynkowa"):
        if not w.get("cytat"):
            return "brak cytatu"
        plik = KATALOG_GLOWNY / (w.get("zrodlo") or {}).get("plik", "")
        if not plik.is_file():
            return f"brak pliku źródłowego {plik.relative_to(KATALOG_GLOWNY)}"
        cytat = ujednolic(w["cytat"])
        if cytat not in ujednolic(plik.read_text(encoding="utf-8")):
            return "cytatu nie ma w pliku źródłowym (na podanej stronie)"
        if klasa == "fakt":
            strona = w["zrodlo"].get("strona")
            if isinstance(strona, int) and plik.name != f"strona-{strona:03d}.txt":
                return f"numer strony {strona} nie zgadza się z plikiem {plik.name}"
            if not wartosc_w_cytacie(w.get("wartosc"), cytat):
                return f"wartości {w.get('wartosc')!r} nie ma w cytacie"
        return None
    if klasa == "grafika":
        if not (KATALOG_GLOWNY / (w.get("plik") or "")).is_file():
            return "brak pliku zdjęcia"
        braki = [p for p, v in (("autor", w.get("autor")), ("licencja", w.get("licencja")),
                                ("zrodlo.url", (w.get("zrodlo") or {}).get("url"))) if not v]
        return f"puste pola: {', '.join(braki)}" if braki else None
    if klasa == "zalozenie":
        return None if (w.get("uzasadnienie") or "").strip() else "brak uzasadnienia"
    return f"nieznana klasa {klasa!r}"


def bramka1(ticker):
    wyniki = []
    for w in wczytaj_ledger(ticker):
        powod = sprawdz_wpis(w)
        wyniki.append({"id": w.get("id"), "wynik": "ok" if powod is None else "błąd", "powod": powod or ""})
    katalog = katalog_spolki(ticker) / "weryfikacja"
    katalog.mkdir(parents=True, exist_ok=True)
    (katalog / "bramka1.json").write_text(json.dumps(wyniki, ensure_ascii=False, indent=2), encoding="utf-8")
    return wyniki


def main():
    if len(sys.argv) != 2:
        sys.exit("Użycie: python scripts/sprawdz_cytaty.py TICKER")
    wyniki = bramka1(sys.argv[1])
    bledy = [w for w in wyniki if w["wynik"] != "ok"]
    for b in bledy:
        print(f"{b['id']}: {b['powod']}")
    print(f"Bramka 1: {len(wyniki) - len(bledy)}/{len(wyniki)} wpisów ok")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
