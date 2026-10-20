# Skill pakiet-analityczny dla Claude Code

Skill przygotowuje pakiet analityczny spółki notowanej w USA na podstawie raportu rocznego 10-K z SEC:

- ledger faktów - każda liczba ma źródło (stronę raportu) i dosłowny cytat,
- weryfikację ledgera dwiema bramkami (skrypt i osobny model),
- model DCF w Excelu, prezentację PowerPoint i jednostronicową notatkę dla zarządu w Wordzie.

Użycie w Claude Code: `/pakiet-analityczny TICKER`, np. `/pakiet-analityczny GOOG`.

## Instalacja w projekcie

1. Skopiuj zawartość tego repozytorium do folderu `.claude/skills/pakiet-analityczny` w swoim projekcie (plik `SKILL.md` ma leżeć bezpośrednio w tym folderze).
2. Zainstaluj Python 3.11+ i biblioteki: `pip install requests beautifulsoup4 openpyxl python-pptx python-docx pymupdf pillow`.
3. Zainstaluj LibreOffice (polecenie `soffice` służy do przeliczenia arkusza i podglądów).
4. W plikach `scripts/pobierz_10k.py`, `scripts/pobierz_grafiki.py` i `scripts/pobierz_rynek.py` wpisz swój adres e-mail w `USER_AGENT` (SEC tego wymaga).
5. Uruchom Claude Code w folderze projektu i wpisz `/pakiet-analityczny TICKER`.

Materiał edukacyjny, a nie rekomendacja inwestycyjna.
