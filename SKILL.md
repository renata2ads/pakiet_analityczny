---
name: pakiet-analityczny
description: Pakiet analityczny spółki notowanej w USA na podstawie raportu rocznego 10-K z SEC - ledger faktów z cytatami, weryfikacja dwiema bramkami, model DCF w Excelu, prezentacja PowerPoint i jednostronicowa notatka dla zarządu w Wordzie, w których każda liczba ma źródło. Używaj, gdy użytkownik wpisze /pakiet-analityczny TICKER albo poprosi o analizę, wycenę DCF, prezentację lub notatkę o spółce z giełdy amerykańskiej (np. "zrób pakiet dla GOOG", "wyceń Apple na podstawie 10-K", "przygotuj analizę Microsoftu dla zarządu").
argument-hint: TICKER
---

# Pakiet analityczny spółki

Dla spółki o tickerze z argumentu (`$ARGUMENTS`, np. `GOOG`) przygotuj w folderze `TICKER/` (w bieżącym folderze projektu):
`wyniki/TICKER_DCF.xlsx`, `wyniki/TICKER_prezentacja.pptx`, `wyniki/TICKER_notatka_zarzad.docx`.

Najważniejsza zasada: **każda informacja o spółce trafia najpierw do ledgera** (`TICKER/ledger.json`) ze źródłem i dosłownym cytatem, a dokumenty powstają wyłącznie ze skryptów, które czytają ledger. Dzięki temu każdą liczbę w dokumentach da się sprawdzić w raporcie. Nie przepisuj liczb ręcznie do dokumentów i nie dopisuj w skryptach liczb o spółce.

Teksty dla ludzi pisz po polsku z polskimi znakami, bez długich myślników (zamiast nich krótki dywiz -). Pliki w UTF-8, nazwy plików bez polskich znaków.

## Skrypty

Wszystkie skrypty są w `scripts/` obok tego pliku (dalej: `S` = pełna ścieżka do folderu `scripts` obok tego pliku SKILL.md, np. `~/.claude/skills/pakiet-analityczny/scripts` dla skilla globalnego - zawsze podawaj ścieżkę bezwzględną). Uruchamiaj je **z folderu, w którym użytkownik pracuje** (nie wchodź do folderu skilla) - wyniki trafiają do `TICKER/` w bieżącym folderze:

```
python "S/nazwa.py" TICKER
```

| Skrypt | Do czego |
|---|---|
| `sprawdz_spolke.py` | czy spółka składa 10-K i nie jest bankiem / ubezpieczycielem |
| `pobierz_10k.py` | pobiera najnowszy 10-K i dzieli go na strony (`TICKER/zrodla/10k/strony/strona-NNN.txt`, NNN = numer ze stopki) |
| `szukaj.py TICKER "fraza"` | wyszukuje frazę w stronach raportu |
| `pobierz_rynek.py` | kurs akcji (Yahoo) i rentowność 10-letnich obligacji USA (FRED) |
| `pobierz_grafiki.py TICKER "fraza" [limit]` | zdjęcia z Wikimedia Commons z wolną licencją |
| `weryfikuj.py` | bramka 1 (`sprawdz_cytaty.py`) + bramka 2 (`weryfikuj_modelem.py`), raport i status |
| `generuj_xlsx.py` | model DCF w Excelu, przeliczenie w LibreOffice, `wyniki/wycena.json`, podgląd PNG |
| `generuj_pptx.py` | prezentacja 16:9 + podgląd PNG |
| `generuj_docx.py` | notatka na 1 stronę A4 + PDF |
| `sprawdz_liczby.py` | każda liczba w pptx i docx ma pokrycie w ledgerze lub `wycena.json` |

Wymagania: Python 3.11+ z bibliotekami `requests beautifulsoup4 openpyxl python-pptx python-docx pymupdf pillow` (jeśli w projekcie jest `.venv`, używaj jego Pythona; jeśli bibliotek brakuje, utwórz `.venv` i zainstaluj je), LibreOffice (`soffice`), program `claude` w PATH (bramka 2). Adres e-mail w nagłówku User-Agent dla SEC jest wpisany w `pobierz_10k.py`.

## Krok 0. Czy spółka pasuje

Uruchom `sprawdz_spolke.py TICKER`. Jeśli wypisze `NIE PASUJE`, zakończ i przekaż użytkownikowi ten komunikat po polsku - nie twórz żadnych plików. Model DCF oparty na wolnych przepływach firmy nie pasuje do banków i ubezpieczycieli, a cały proces opiera się na raporcie 10-K.

## Krok 1. Źródła

1. `pobierz_10k.py TICKER`, potem `pobierz_rynek.py TICKER`.
2. Sprawdź na 2-3 stronach, że numer w nazwie pliku zgadza się z numerem na dole strony. Strony bez numeru w stopce (np. okładka) mają nazwy `bez-numeru-NNN.txt`.
3. Znajdź (`szukaj.py`) strony rachunku zysków i strat, bilansu i przepływów pieniężnych. Nazwy różnią się między spółkami (np. "Consolidated Statements of Income" / "of Operations" / "of Earnings"), więc szukaj kilku wariantów.

**Nigdy nie czytaj całego raportu naraz** - jest za długi. Szukaj frazami i czytaj tylko potrzebne pliki stron.

## Krok 2. Ledger

Plik `TICKER/ledger.json` to lista wpisów. Pola każdego wpisu:

- `id`: `F-01`... fakty, `R-01`... dane rynkowe, `Z-01`... założenia, `G-01`... grafiki,
- `klasa`: `fakt`, `dana_rynkowa`, `zalozenie` albo `grafika`,
- `klucz`: stała nazwa po angielsku (lista niżej - skrypty z niej korzystają),
- `opis` (po polsku), `wartosc` (liczba bez separatorów albo krótki tekst), `jednostka` (`mln USD`, `mld USD`, `%`, `szt.`...), `okres` (`2025`, `2026-01-23`...),
- `zrodlo`: `{dokument, plik, url, strona}`; `plik` to ścieżka od folderu projektu (np. `GOOG/zrodla/10k/strony/strona-052.txt`), `strona` to numer ze stopki (liczba) albo tekst dla okładki; dla danych z internetu `"n/d (www)"`,
- `cytat`: **dosłowny** fragment tekstu pliku źródłowego, bez poprawiania, tłumaczenia i skracania w środku; zawiera nazwę pozycji i liczbę. Tabele w plikach stron mają każdą komórkę w osobnej linii - w cytacie zastąp znaki nowej linii spacjami (bramka ujednolica spacje),
- `data_pobrania` (tylko dane rynkowe), `uzasadnienie` (tylko założenia: 1-3 zdania z odwołaniem do id faktów).

Dla faktu tekstowego `wartosc` musi być dosłownym fragmentem cytatu (np. `"artificial intelligence initiatives"`), a nie polskim streszczeniem - bramka 1 sprawdza, że wartość występuje w cytacie.

### Wymagane klucze (`RRRR` = rok)

| Klucz | Co | Uwagi |
|---|---|---|
| `revenue_RRRR` | przychody ogółem, 3 ostatnie lata | mln USD, jak w raporcie |
| `operating_income_RRRR` | zysk operacyjny, 3 lata | jeśli raport nie ma tej pozycji (np. "Income from operations" / "Operating income"), szukaj odpowiednika i napisz o tym użytkownikowi |
| `net_income_RRRR`, `pretax_income_RRRR`, `income_tax_RRRR` | ostatni rok | |
| `depreciation_amortization_RRRR`, `capex_RRRR` | ostatni rok, z rachunku przepływów | capex zapisz ze znakiem jak w raporcie (nawias = liczba ujemna) |
| `cash_and_equivalents_RRRR`, `marketable_securities_RRRR`, `long_term_debt_RRRR` | koniec ostatniego roku, z bilansu | osobne wpisy; jeśli spółka nie ma którejś pozycji, wpisz 0 z cytatem pokazującym jej brak albo napisz o tym użytkownikowi |
| `operating_margin_RRRR` | marża operacyjna, jeśli raport podaje ją wprost | opcjonalny (notatka) |
| `segment_<nazwa>_revenue_RRRR`, `segment_<nazwa>_operating_income_RRRR` | segmenty za ostatni rok | tyle segmentów, ile ma spółka; brak segmentów = brak wpisów |
| `shares_class_<klasa>` | liczba akcji w obrocie z okładki, każda klasa osobno, z datą w `okres` | np. Alphabet: `shares_class_a`, `shares_class_b`, `shares_class_c`; jedna klasa: `shares_class_common` |
| `project_<nazwa>` | 4-6 faktów o najważniejszych projektach | pierwszy liczbowy fakt trafia na kartę slajdu 2 |
| `project_capex_plan_RRRR_low` / `_high` | plan nakładów na kolejny rok, jeśli raport go podaje | dla jednej liczby wpisz tę samą wartość w obu |
| `risk_<nazwa>` | 2-3 ryzyka z Item 1A | `opis` zaczyna się od "Ryzyko: " i jest pełnym zdaniem po polsku |
| `effective_tax_rate_RRRR` | efektywna stopa podatkowa, jeśli podana | opcjonalny |
| `share_price`, `risk_free_rate` | dane rynkowe z plików w `TICKER/zrodla/rynek/` | cytat = fragment zapisanego pliku, np. `"regularMarketPrice": 172.3` albo wiersz CSV `2026-10-16,5.09` |

Założenia (`klasa: zalozenie`, `jednostka: "%"` z wartością w procentach, np. `17`): `revenue_growth_RRRR` (5 lat prognozy), `operating_margin`, `capex_pct_revenue_RRRR`, `da_pct_revenue_RRRR` (5 lat), `tax_rate`, `beta` (bez jednostki), `equity_risk_premium`, `terminal_growth`, `terminal_capex_pct_revenue`, `terminal_da_pct_revenue`.

Założenia mają być spójne: gdy nakłady są dużo wyższe niż amortyzacja, amortyzacja w kolejnych latach rośnie (nowe inwestycje zaczynają się amortyzować); w okresie rezydualnym spółka rośnie wolno, więc nakłady są tylko trochę wyższe niż amortyzacja. Beta i premia za ryzyko nie pochodzą z 10-K - zaznacz to w uzasadnieniu.

Nie wpisuj niczego, czego nie znalazłeś w źródle. Jeśli czegoś brakuje, napisz o tym użytkownikowi zamiast zgadywać.

## Krok 3. Zdjęcia projektów

1. Dla 3-4 projektów z ledgera uruchom `pobierz_grafiki.py TICKER "fraza"` (kilka fraz po angielsku, np. nazwa produktu). Skrypt sam pomija logotypy, zrzuty ekranu, SVG i licencje inne niż CC0, domena publiczna, CC BY, CC BY-SA.
2. **Obejrzyj pobrane zdjęcia** (Read na plikach .jpg) i wybierz po jednym, które naprawdę pokazuje dany projekt. Odrzucaj trafienia z przypadkowej zbieżności nazw i zdjęcia konkretnych osób, gdy produkt jest tylko w tle. Jeśli do projektu nie ma dobrego zdjęcia, powiedz o tym. Pozostałe pobrane pliki usuń, a w `kandydaci.json` zostaw tylko wybrane.
3. Dopisz wpisy `grafika`: `id` G-01..., `klucz`, `plik` (ścieżka od folderu projektu), `opis`, `projekt` (id faktu), `autor`, `licencja`, `zrodlo.url` (strona pliku na Commons) oraz opcjonalnie `naglowek` (nazwa projektu po polsku, tytuł slajdu) i `fakty` (lista 2-3 id faktów do punktów slajdu). Dane autora i licencji weź z `TICKER/zrodla/grafiki/kandydaci.json`.

## Krok 4. Weryfikacja - aż do zielonego statusu

Uruchom `weryfikuj.py TICKER`. Bramka 1 (skrypt) sprawdza, że cytat jest na podanej stronie i zawiera wartość. Bramka 2 (osobny model Sonnet przez `claude -p`, bez narzędzi i bez znajomości rozmowy) sprawdza, czy cytat naprawdę potwierdza liczbę, rok, jednostkę i zakres.

Gdy są błędy (`TICKER/weryfikacja/raport.md`), **popraw wpisy w ledgerze na podstawie źródła** i uruchom weryfikację ponownie. Nigdy nie zmieniaj skryptów bramek, żeby przepuszczały błędy - to cały sens kontroli. Werdykt "niejasne" też traktuj jako błąd: zwykle znaczy, że opis wpisu mówi więcej niż cytat. Każda zmiana ledgera unieważnia status (skrót SHA-256), więc po każdej poprawce weryfikacja od nowa. Generatory dokumentów odmówią pracy, dopóki status nie jest zielony.

## Krok 5. Excel z DCF

`generuj_xlsx.py TICKER` buduje model formułami Excela (arkusze: Podsumowanie, Dane historyczne, Założenia z WACC metodą CAPM, DCF 5 lat + wartość rezydualna Gordona, Wrażliwość 5x5, Ledger), przelicza go w LibreOffice, sprawdza, że nie ma błędów formuł, zapisuje `wyniki/wycena.json` i podgląd `wyniki/podglad/xlsx-NN.png`.

Obejrzyj podgląd (czy wykresy są czytelne i nic się nie ucina). Opisz użytkownikowi wynik w 3-4 zdaniach prostym językiem. Jeśli wartość różni się od kursu o więcej niż 50%, wskaż założenia, które za to najbardziej odpowiadają - ale sam ich nie zmieniaj.

## Krok 6. Prezentacja

`generuj_pptx.py TICKER`: tytuł, spółka w liczbach (wykres + 3 karty), po jednym slajdzie na każde zdjęcie projektu, wycena DCF z ramką "To są założenia, nie fakty", ryzyka i źródła. Przy każdej liczbie jest id wpisu.

Obejrzyj każdy slajd w `wyniki/podglad/pptx-NN.png`. Jeśli coś się ucina, nachodzi na siebie albo punkt slajdu nie ma pokrycia w cytacie, popraw ledger (np. `naglowek`, `fakty`, krótszy `opis`) albo generator i wygeneruj ponownie.

## Krok 7. Notatka i kontrola liczb

1. `generuj_docx.py TICKER` - notatka na dokładnie 1 stronę A4 (skrypt sam zmniejsza czcionkę do minimum 9 pt; jeśli to nie wystarczy, skróć opisy w ledgerze). Obejrzyj `wyniki/podglad/docx-01.png`.
2. `sprawdz_liczby.py TICKER` - każda liczba w pptx i docx musi mieć pokrycie w ledgerze lub `wycena.json`. Jeśli coś nie przejdzie, popraw generator albo ledger (i zweryfikuj ledger ponownie), nigdy samą kontrolę.

## Na koniec

Pokaż użytkownikowi krótkie podsumowanie: rok raportu, liczba wpisów ledgera (fakty / dane rynkowe / założenia / grafiki), wynik weryfikacji, wycena wobec kursu, lista plików w `TICKER/wyniki/` oraz wszystko, czego nie udało się znaleźć w źródłach. Dopisz, że to materiał edukacyjny, a nie rekomendacja inwestycyjna.
