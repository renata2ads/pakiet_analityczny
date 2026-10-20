"""Uruchamia obie bramki i zapisuje raport.md oraz status.json.

Użycie: python scripts/weryfikuj.py TICKER
"""
import json
import sys
from datetime import datetime

from sprawdz_cytaty import bramka1
from weryfikuj_modelem import bramka2
from wspolne import katalog_spolki, skrot_ledgera, wczytaj_ledger


def main():
    if len(sys.argv) != 2:
        sys.exit("Użycie: python scripts/weryfikuj.py TICKER")
    ticker = sys.argv[1].upper()
    skrot = skrot_ledgera(ticker)
    ledger = wczytaj_ledger(ticker)
    opisy = {w["id"]: w.get("opis", "") for w in ledger}

    w1 = bramka1(ticker)
    w2 = bramka2(ticker)

    problemy = [(w["id"], "bramka 1", w["powod"]) for w in w1 if w["wynik"] != "ok"]
    problemy += [(w["id"], f"bramka 2 ({w['werdykt']})", w["uzasadnienie"])
                 for w in w2 if w["werdykt"] != "potwierdza"]
    ok = not problemy and skrot == skrot_ledgera(ticker)

    katalog = katalog_spolki(ticker) / "weryfikacja"
    teraz = datetime.now().isoformat(timespec="seconds")
    (katalog / "status.json").write_text(json.dumps(
        {"ok": ok, "sha256_ledgera": skrot, "data": teraz}, ensure_ascii=False, indent=2), encoding="utf-8")

    przeszly_1 = sum(1 for w in w1 if w["wynik"] == "ok")
    przeszly_2 = sum(1 for w in w2 if w["werdykt"] == "potwierdza")
    linie = [
        f"# Weryfikacja ledgera {ticker}", "",
        f"Data: {teraz}", "",
        f"**Status: {'OK - wolno tworzyć dokumenty' if ok else 'BŁĘDY - popraw ledger'}**", "",
        f"- Wpisów w ledgerze: {len(ledger)}",
        f"- Bramka 1 (skrypt, bez AI): przeszło {przeszly_1} z {len(w1)}",
        f"- Bramka 2 (model Sonnet, fakty i dane rynkowe): potwierdzono {przeszly_2} z {len(w2)}", "",
        "## Problemy", "",
    ]
    if problemy:
        linie += ["| id | opis | bramka | powód |", "|---|---|---|---|"]
        linie += [f"| {i} | {opisy.get(i, '')} | {b} | {p} |" for i, b, p in problemy]
    else:
        linie.append("Brak problemów.")
    linie += ["", "## Werdykty bramki 2", "", "| id | werdykt | uzasadnienie |", "|---|---|---|"]
    linie += [f"| {w['id']} | {w['werdykt']} | {w['uzasadnienie']} |" for w in w2]
    (katalog / "raport.md").write_text("\n".join(linie) + "\n", encoding="utf-8")

    print("\n".join(linie[:10]))
    for i, b, p in problemy:
        print(f"- {i} [{b}]: {p}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
