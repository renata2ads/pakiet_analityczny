# Skill pakiet-analityczny dla Claude Code

Skill przygotowuje pakiet analityczny spółki notowanej w USA na podstawie raportu rocznego 10-K z SEC:

- ledger faktów - każda liczba ma źródło (stronę raportu) i dosłowny cytat,
- weryfikację ledgera dwiema bramkami (skrypt i osobny model),
- model DCF w Excelu, prezentację PowerPoint i jednostronicową notatkę dla zarządu w Wordzie.

Użycie w Claude Code: `/pakiet-analityczny TICKER`, np. `/pakiet-analityczny GOOG`.

## Instalacja

Repozytorium jest marketplace'em pluginów Claude Code `moje-skille` z jednym pluginem `pakiet-analityczny`.

1. W Claude Code dodaj marketplace:
   ```
   /plugin marketplace add renata2ads/pakiet_analityczny
   ```
2. Zainstaluj plugin:
   ```
   /plugin install pakiet-analityczny@moje-skille
   ```
3. Zainstaluj Python 3.11+ i biblioteki: `pip install requests beautifulsoup4 openpyxl python-pptx python-docx pymupdf pillow`.
4. Zainstaluj LibreOffice (polecenie `soffice` służy do przeliczenia arkusza i podglądów).
5. W plikach `pobierz_10k.py`, `pobierz_grafiki.py` i `pobierz_rynek.py` (folder `plugins/pakiet-analityczny/skills/pakiet-analityczny/scripts/`) wpisz swój adres e-mail w `USER_AGENT` (SEC tego wymaga).
6. Uruchom Claude Code w folderze projektu i wpisz `/pakiet-analityczny TICKER`.

## Struktura

```
.claude-plugin/marketplace.json
plugins/pakiet-analityczny/
  .claude-plugin/plugin.json
  skills/pakiet-analityczny/
    SKILL.md
    scripts/
```

Materiał edukacyjny, a nie rekomendacja inwestycyjna.
