"""Model DCF w Excelu, liczony formułami, z danymi wyłącznie z ledgera.

Użycie: python scripts/generuj_xlsx.py TICKER
Tworzy TICKER/wyniki/TICKER_DCF.xlsx, przelicza go w LibreOffice, zapisuje
TICKER/wyniki/wycena.json i podgląd arkuszy w TICKER/wyniki/podglad/.
"""
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from wspolne import katalog_spolki, wczytaj_ledger, wymagaj_weryfikacji

CZCIONKA = "Arial"
NIEBIESKI = Font(name=CZCIONKA, color="0000FF")
ZIELONY = Font(name=CZCIONKA, color="008000")
ZWYKLY = Font(name=CZCIONKA)
POGRUBIONY = Font(name=CZCIONKA, bold=True)
TYTUL = Font(name=CZCIONKA, bold=True, size=14)
SZARY = Font(name=CZCIONKA, color="808080", size=9)
NAGLOWEK_TLO = PatternFill("solid", fgColor="DCE6F1")
WYNIK_TLO = PatternFill("solid", fgColor="FFF2CC")
RAMKA = Border(top=Side(style="thin"), bottom=Side(style="double"))

MLN = '#,##0;(#,##0);"-"'
USD = '#,##0.00;(#,##0.00);"-"'
PROC = '0.0%;(0.0%);"-"'
SZT = '#,##0'

ERRORY = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")


def soffice():
    for kandydat in (shutil.which("soffice"),
                     r"C:\Program Files\LibreOffice\program\soffice.exe",
                     r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                     "/usr/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if kandydat and Path(kandydat).exists():
            return kandydat
    sys.exit("Nie znaleziono LibreOffice (soffice). Zainstaluj LibreOffice.")


class Ledger:
    def __init__(self, wpisy):
        self.wpisy = wpisy
        self.po_kluczu = {w["klucz"]: w for w in wpisy}

    def __getitem__(self, klucz):
        if klucz not in self.po_kluczu:
            sys.exit(f"W ledgerze brakuje wpisu o kluczu {klucz!r}.")
        return self.po_kluczu[klucz]

    def lata(self, prefiks):
        lata = sorted(int(m.group(1)) for k in self.po_kluczu if (m := re.fullmatch(prefiks + r"_(\d{4})", k)))
        return lata

    def z_prefiksem(self, prefiks):
        return [w for w in self.wpisy if w["klucz"].startswith(prefiks)]


def komentarz(w):
    z = w.get("zrodlo") or {}
    if w["klasa"] == "zalozenie":
        tekst = f"{w['id']} (założenie): {w.get('uzasadnienie', '')}"
    else:
        tekst = f"{w['id']} | {z.get('dokument', '')} | strona: {z.get('strona', '')}\nCytat: {w.get('cytat', '')}"
    return Comment(tekst, "ledger", width=420, height=160)


def wejscie(ark, adres, w, procent=False, format_=None, kolumna_id=None):
    """Komórka wejściowa: niebieska wartość z ledgera, id obok, cytat w komentarzu."""
    wartosc = w["wartosc"]
    komorka = ark[adres]
    komorka.value = wartosc / 100 if procent else wartosc
    komorka.font = NIEBIESKI
    komorka.number_format = PROC if procent else (format_ or MLN)
    komorka.comment = komentarz(w)
    if kolumna_id:
        wiersz = komorka.row
        ark[f"{kolumna_id}{wiersz}"] = w["id"]
        ark[f"{kolumna_id}{wiersz}"].font = SZARY
    return komorka


def naglowek(ark, wiersz, wartosci, kol_start=1):
    for i, v in enumerate(wartosci):
        c = ark.cell(row=wiersz, column=kol_start + i, value=v)
        c.font = POGRUBIONY
        c.fill = NAGLOWEK_TLO
        c.alignment = Alignment(horizontal="center" if i else "left", vertical="center", wrap_text=True)


def szerokosci(ark, szer):
    for kol, s in szer.items():
        ark.column_dimensions[kol].width = s


def drukowanie(ark, poziomo=True):
    ark.page_setup.orientation = "landscape" if poziomo else "portrait"
    ark.page_setup.paperSize = ark.PAPERSIZE_A4
    ark.page_setup.fitToWidth = 1
    ark.page_setup.fitToHeight = 0
    ark.sheet_properties.pageSetUpPr.fitToPage = True
    ark.print_options.gridLines = False


def zbuduj(ticker, plik):
    L = Ledger(wczytaj_ledger(ticker))
    lata_hist = L.lata("revenue")[-3:]
    ostatni = lata_hist[-1]
    lata_prog = L.lata("revenue_growth")
    nazwa = ticker
    info_10k = katalog_spolki(ticker) / "zrodla" / "10k" / "info.json"
    if info_10k.exists():
        nazwa = json.loads(info_10k.read_text(encoding="utf-8"))["nazwa_spolki"]

    wb = Workbook()
    pod = wb.active
    pod.title = "Podsumowanie"
    dh = wb.create_sheet("Dane historyczne")
    zal = wb.create_sheet("Założenia")
    dcf = wb.create_sheet("DCF")
    wr = wb.create_sheet("Wrażliwość")
    led = wb.create_sheet("Ledger")
    DH, ZAL = "'Dane historyczne'!", "'Założenia'!"

    # ---------- Dane historyczne ----------
    dh["A1"] = f"{nazwa} - dane historyczne (mln USD)"
    dh["A1"].font = TYTUL
    naglowek(dh, 3, ["Pozycja"] + [str(r) for r in lata_hist] + ["id wpisów ledgera"])
    wiersze_hist = {}
    for i, (prefiks, etykieta) in enumerate([("revenue", "Przychody ogółem"),
                                             ("operating_income", "Zysk operacyjny")]):
        r = 4 + i
        dh.cell(row=r, column=1, value=etykieta).font = ZWYKLY
        idy = []
        for j, rok in enumerate(lata_hist):
            w = L[f"{prefiks}_{rok}"]
            wejscie(dh, f"{get_column_letter(2 + j)}{r}", w)
            idy.append(w["id"])
        dh.cell(row=r, column=2 + len(lata_hist), value=", ".join(idy)).font = SZARY
        wiersze_hist[prefiks] = r
    r = 6
    dh.cell(row=r, column=1, value="Marża operacyjna").font = ZWYKLY
    for j in range(len(lata_hist)):
        k = get_column_letter(2 + j)
        dh[f"{k}{r}"] = f"=IF({k}4=0,0,{k}5/{k}4)"
        dh[f"{k}{r}"].number_format = PROC
        dh[f"{k}{r}"].font = ZWYKLY
    kol_ost = get_column_letter(1 + len(lata_hist))  # kolumna ostatniego roku

    naglowek(dh, 8, ["Pozostałe fakty", "Wartość", "id", "Jednostka", "Okres"])
    pozostale = [f"net_income_{ostatni}", f"pretax_income_{ostatni}", f"income_tax_{ostatni}",
                 f"depreciation_amortization_{ostatni}", f"capex_{ostatni}",
                 f"cash_and_equivalents_{ostatni}", f"marketable_securities_{ostatni}",
                 f"long_term_debt_{ostatni}"]
    pozostale += [w["klucz"] for w in L.z_prefiksem("segment_")]
    pozostale += [w["klucz"] for w in L.z_prefiksem("effective_tax_rate_")]
    adres = {}
    r = 9
    for klucz in pozostale:
        w = L[klucz]
        dh.cell(row=r, column=1, value=w["opis"]).font = ZWYKLY
        wejscie(dh, f"B{r}", w, procent=w["jednostka"] == "%", kolumna_id="C")
        dh.cell(row=r, column=4, value=w["jednostka"]).font = ZWYKLY
        dh.cell(row=r, column=5, value=w["okres"]).font = ZWYKLY
        adres[klucz] = f"B{r}"
        r += 1
    r += 1
    naglowek(dh, r, ["Akcje i rynek", "Wartość", "id", "Jednostka", "Okres"])
    r += 1
    wiersze_akcji = []
    for w in L.z_prefiksem("shares_class_"):
        dh.cell(row=r, column=1, value=w["opis"]).font = ZWYKLY
        wejscie(dh, f"B{r}", w, format_=SZT, kolumna_id="C")
        dh.cell(row=r, column=4, value=w["jednostka"]).font = ZWYKLY
        dh.cell(row=r, column=5, value=w["okres"]).font = ZWYKLY
        wiersze_akcji.append(r)
        r += 1
    dh.cell(row=r, column=1, value="Liczba akcji razem (wszystkie klasy)").font = POGRUBIONY
    dh[f"B{r}"] = f"=SUM(B{wiersze_akcji[0]}:B{wiersze_akcji[-1]})"
    dh[f"B{r}"].number_format = SZT
    dh[f"B{r}"].font = POGRUBIONY
    adres["shares_total"] = f"B{r}"
    r += 1
    w = L["share_price"]
    dh.cell(row=r, column=1, value=w["opis"]).font = ZWYKLY
    wejscie(dh, f"B{r}", w, format_=USD, kolumna_id="C")
    dh.cell(row=r, column=4, value=w["jednostka"]).font = ZWYKLY
    dh.cell(row=r, column=5, value=w["okres"]).font = ZWYKLY
    adres["share_price"] = f"B{r}"
    szerokosci(dh, {"A": 62, "B": 16, "C": 14, "D": 14, "E": 14})
    dh.freeze_panes = "B4"
    drukowanie(dh)

    # ---------- Założenia ----------
    zal["A1"] = f"{nazwa} - założenia modelu DCF"
    zal["A1"].font = TYTUL
    zal["A2"] = "Niebieskie komórki to dane wejściowe z ledgera. Zmień założenie, a wycena przeliczy się sama."
    zal["A2"].font = SZARY
    naglowek(zal, 4, ["Założenie", "Wartość", "id", "Okres", "Uzasadnienie"])
    za = {}
    r = 5
    for w in [w for w in L.wpisy if w["klasa"] == "zalozenie"]:
        zal.cell(row=r, column=1, value=w["opis"]).font = ZWYKLY
        procent = w["jednostka"] == "%"
        wejscie(zal, f"B{r}", w, procent=procent, format_="0.00", kolumna_id="C")
        zal.cell(row=r, column=4, value=w["okres"]).font = ZWYKLY
        c = zal.cell(row=r, column=5, value=w["uzasadnienie"])
        c.font = ZWYKLY
        c.alignment = Alignment(wrap_text=True, vertical="top")
        zal.row_dimensions[r].height = 45
        za[w["klucz"]] = f"$B${r}"
        r += 1

    r += 1
    zal.cell(row=r, column=1, value="Koszt kapitału (WACC) - metoda CAPM").font = POGRUBIONY
    zal.cell(row=r, column=1).fill = NAGLOWEK_TLO
    r += 1
    w = L["risk_free_rate"]
    zal.cell(row=r, column=1, value="Stopa wolna od ryzyka (rentowność 10-letnich obligacji USA)").font = ZWYKLY
    wejscie(zal, f"B{r}", w, procent=True, kolumna_id="C")
    za["rf"] = f"$B${r}"
    r += 1

    def formula(etykieta, wzor, fmt=PROC, klucz=None, opis=None, font=ZWYKLY):
        nonlocal r
        zal.cell(row=r, column=1, value=etykieta).font = font
        zal[f"B{r}"] = wzor
        zal[f"B{r}"].number_format = fmt
        zal[f"B{r}"].font = font
        if opis:
            zal.cell(row=r, column=5, value=opis).font = SZARY
        if klucz:
            za[klucz] = f"$B${r}"
        r += 1

    formula("Koszt kapitału własnego = stopa wolna od ryzyka + beta × premia za ryzyko",
            f"={za['rf']}+{za['beta']}*{za['equity_risk_premium']}", klucz="ke")
    formula("Kapitalizacja rynkowa (mln USD) = kurs × liczba akcji",
            f"={DH}{adres['share_price']}*{DH}{adres['shares_total']}/1000000", fmt=MLN, klucz="mcap")
    formula("Dług długoterminowy (mln USD)", f"={DH}{adres[f'long_term_debt_{ostatni}']}", fmt=MLN, klucz="dlug")
    formula("Udział długu w finansowaniu = dług / (dług + kapitalizacja)",
            f"={za['dlug']}/({za['dlug']}+{za['mcap']})", klucz="udzial_dlugu",
            opis="Jeśli udział długu przekracza 10%, WACC uwzględnia dług po koszcie równym stopie wolnej od ryzyka po podatku.")
    formula("WACC", f"=IF({za['udzial_dlugu']}>0.1,(1-{za['udzial_dlugu']})*{za['ke']}"
                    f"+{za['udzial_dlugu']}*{za['rf']}*(1-{za['tax_rate']}),{za['ke']})",
            klucz="wacc", font=POGRUBIONY,
            opis="Dług jest nieistotny w porównaniu z wartością spółki (poniżej 10%), więc WACC = koszt kapitału własnego.")
    zal[za["wacc"].replace("$", "")].fill = WYNIK_TLO
    szerokosci(zal, {"A": 80, "B": 12, "C": 8, "D": 12, "E": 90})
    drukowanie(zal)

    # ---------- DCF ----------
    dcf["A1"] = f"{nazwa} - model DCF (mln USD)"
    dcf["A1"].font = TYTUL
    n = len(lata_prog)
    kol = [get_column_letter(2 + i) for i in range(n)]
    kt = get_column_letter(2 + n)  # kolumna okresu rezydualnego
    naglowek(dcf, 3, ["Pozycja"] + [str(r) for r in lata_prog] + [f"Rezydualny (po {lata_prog[-1]})"])
    przych_bazowy = f"{DH}{kol_ost}4"
    wiersze = {}

    def wiersz_dcf(nr, etykieta, wzory, wzor_t, fmt=MLN, font=ZWYKLY, kolor_linku=False):
        dcf.cell(row=nr, column=1, value=etykieta).font = font
        for i, k in enumerate(kol):
            dcf[f"{k}{nr}"] = wzory(i, k)
            dcf[f"{k}{nr}"].number_format = fmt
            dcf[f"{k}{nr}"].font = ZIELONY if kolor_linku else font
        if wzor_t is not None:
            dcf[f"{kt}{nr}"] = wzor_t
            dcf[f"{kt}{nr}"].number_format = fmt
            dcf[f"{kt}{nr}"].font = ZIELONY if kolor_linku else font
        wiersze[etykieta] = nr

    wiersz_dcf(4, "Wzrost przychodów", lambda i, k: f"={ZAL}{za[f'revenue_growth_{lata_prog[i]}']}",
               f"={ZAL}{za['terminal_growth']}", PROC, kolor_linku=True)
    wiersz_dcf(5, "Przychody", lambda i, k: f"={przych_bazowy if i == 0 else kol[i - 1] + '5'}*(1+{k}4)",
               f"={kol[-1]}5*(1+{kt}4)")
    wiersz_dcf(6, "Marża operacyjna", lambda i, k: f"={ZAL}{za['operating_margin']}",
               f"={ZAL}{za['operating_margin']}", PROC, kolor_linku=True)
    wiersz_dcf(7, "Zysk operacyjny", lambda i, k: f"={k}5*{k}6", f"={kt}5*{kt}6")
    wiersz_dcf(8, "Stopa podatkowa", lambda i, k: f"={ZAL}{za['tax_rate']}", f"={ZAL}{za['tax_rate']}",
               PROC, kolor_linku=True)
    wiersz_dcf(9, "Podatek", lambda i, k: f"={k}7*{k}8", f"={kt}7*{kt}8")
    wiersz_dcf(10, "NOPAT (zysk operacyjny po podatku)", lambda i, k: f"={k}7-{k}9", f"={kt}7-{kt}9")
    wiersz_dcf(11, "Amortyzacja jako % przychodów", lambda i, k: f"={ZAL}{za[f'da_pct_revenue_{lata_prog[i]}']}",
               f"={ZAL}{za['terminal_da_pct_revenue']}", PROC, kolor_linku=True)
    wiersz_dcf(12, "Amortyzacja", lambda i, k: f"={k}5*{k}11", f"={kt}5*{kt}11")
    wiersz_dcf(13, "Nakłady inwestycyjne jako % przychodów",
               lambda i, k: f"={ZAL}{za[f'capex_pct_revenue_{lata_prog[i]}']}",
               f"={ZAL}{za['terminal_capex_pct_revenue']}", PROC, kolor_linku=True)
    wiersz_dcf(14, "Nakłady inwestycyjne", lambda i, k: f"={k}5*{k}13", f"={kt}5*{kt}13")
    wiersz_dcf(15, "Wolne przepływy pieniężne (FCF)", lambda i, k: f"={k}10+{k}12-{k}14",
               f"={kt}10+{kt}12-{kt}14", font=POGRUBIONY)
    wiersz_dcf(16, "Okres dyskonta (lata)", lambda i, k: f"={i + 1}", None, "0")
    wiersz_dcf(17, "Czynnik dyskontowy", lambda i, k: f"=1/(1+{ZAL}{za['wacc']})^{k}16", None, "0.000")
    wiersz_dcf(18, "Wartość bieżąca FCF", lambda i, k: f"={k}15*{k}17", None)
    dcf[f"{kt}19"] = f"=IF({kt}5=0,0,{kt}15/{kt}5)"
    dcf[f"{kt}19"].number_format = PROC
    dcf.cell(row=19, column=1, value="FCF rezydualny jako % przychodów (do tabeli wrażliwości)").font = SZARY

    r = 21
    naglowek(dcf, r, ["Wycena", "Wartość"])
    wyc = {}

    def poz(etykieta, wzor, klucz, fmt=MLN, font=ZWYKLY, tlo=None):
        nonlocal r
        r += 1
        dcf.cell(row=r, column=1, value=etykieta).font = font
        dcf[f"B{r}"] = wzor
        dcf[f"B{r}"].number_format = fmt
        dcf[f"B{r}"].font = font
        if tlo:
            dcf[f"B{r}"].fill = tlo
        wyc[klucz] = f"B{r}"

    poz("WACC", f"={ZAL}{za['wacc']}", "wacc", PROC)
    poz("Wzrost rezydualny", f"={ZAL}{za['terminal_growth']}", "g", PROC)
    poz("Suma wartości bieżących FCF", f"=SUM({kol[0]}18:{kol[-1]}18)", "suma_pv")
    poz(f"Wartość rezydualna na koniec {lata_prog[-1]} (Gordon: FCF rezydualny / (WACC - g))",
        f"=IF({wyc['wacc']}<={wyc['g']},0,{kt}15/({wyc['wacc']}-{wyc['g']}))", "tv")
    poz("Wartość bieżąca wartości rezydualnej", f"={wyc['tv']}*{kol[-1]}17", "pv_tv")
    poz("Wartość firmy (EV)", f"={wyc['suma_pv']}+{wyc['pv_tv']}", "ev", font=POGRUBIONY)
    poz("+ gotówka i ekwiwalenty", f"={DH}{adres[f'cash_and_equivalents_{ostatni}']}", "gotowka")
    poz("+ papiery wartościowe", f"={DH}{adres[f'marketable_securities_{ostatni}']}", "papiery")
    poz("- dług długoterminowy", f"={DH}{adres[f'long_term_debt_{ostatni}']}", "dlug")
    poz("Wartość kapitału własnego", f"={wyc['ev']}+{wyc['gotowka']}+{wyc['papiery']}-{wyc['dlug']}",
        "equity", font=POGRUBIONY)
    poz("Liczba akcji (mln, wszystkie klasy)", f"={DH}{adres['shares_total']}/1000000", "akcje", "#,##0.0")
    poz("Wartość jednej akcji (USD)", f"=IF({wyc['akcje']}=0,0,{wyc['equity']}/{wyc['akcje']})", "na_akcje",
        USD, POGRUBIONY, WYNIK_TLO)
    poz("Aktualny kurs (USD)", f"={DH}{adres['share_price']}", "kurs", USD)
    poz("Różnica wartości wobec kursu", f"=IF({wyc['kurs']}=0,0,{wyc['na_akcje']}/{wyc['kurs']}-1)",
        "roznica", PROC, POGRUBIONY)
    for c in ("ev", "equity", "na_akcje"):
        dcf[wyc[c]].border = RAMKA
    szerokosci(dcf, {"A": 82, **{k: 13 for k in kol}, kt: 20})
    dcf.freeze_panes = "B4"
    drukowanie(dcf)
    dcf.page_setup.fitToHeight = 1

    # ---------- Wrażliwość ----------
    wr["A1"] = "Wartość jednej akcji (USD) dla różnych WACC i wzrostu rezydualnego"
    wr["A1"].font = TYTUL
    wr["A2"] = "Każda komórka liczy pełną wycenę formułą na podstawie arkusza DCF. Środek tabeli = wariant bazowy."
    wr["A2"].font = SZARY
    wr["B4"] = "Wzrost rezydualny →"
    wr["B4"].font = POGRUBIONY
    wr["A5"] = "WACC ↓"
    wr["A5"].font = POGRUBIONY
    kroki = [-0.01, -0.005, 0, 0.005, 0.01]
    for j, d in enumerate(kroki):
        c = wr.cell(row=5, column=2 + j, value=f"=DCF!{wyc['g']}+({d})")
        c.number_format = PROC
        c.font = POGRUBIONY
        c.fill = NAGLOWEK_TLO
    for i, d in enumerate(kroki):
        c = wr.cell(row=6 + i, column=1, value=f"=DCF!{wyc['wacc']}+({d})")
        c.number_format = PROC
        c.font = POGRUBIONY
        c.fill = NAGLOWEK_TLO
    fcf = f"DCF!${kol[0]}$15:${kol[-1]}$15"
    okresy = f"DCF!${kol[0]}$16:${kol[-1]}$16"
    for i in range(5):
        for j in range(5):
            w_ = f"$A{6 + i}"
            g_ = f"{get_column_letter(2 + j)}$5"
            wzor = (f"=IF({w_}<={g_},0,(SUMPRODUCT({fcf}/(1+{w_})^{okresy})"
                    f"+DCF!${kol[-1]}$5*(1+{g_})*DCF!${kt}$19/({w_}-{g_})/(1+{w_})^DCF!${kol[-1]}$16"
                    f"+DCF!${wyc['gotowka'][0]}${wyc['gotowka'][1:]}+DCF!${wyc['papiery'][0]}${wyc['papiery'][1:]}"
                    f"-DCF!${wyc['dlug'][0]}${wyc['dlug'][1:]})/DCF!${wyc['akcje'][0]}${wyc['akcje'][1:]})")
            c = wr.cell(row=6 + i, column=2 + j, value=wzor)
            c.number_format = USD
            c.font = POGRUBIONY if (i, j) == (2, 2) else ZWYKLY
    wr.conditional_formatting.add("B6:F10", ColorScaleRule(
        start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50,
        mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    wr["A12"] = "Aktualny kurs (USD):"
    wr["B12"] = f"=DCF!{wyc['kurs']}"
    wr["B12"].number_format = USD
    wr["A12"].font = ZWYKLY
    wr["B12"].font = ZIELONY
    szerokosci(wr, {"A": 22, **{get_column_letter(2 + j): 14 for j in range(5)}})
    drukowanie(wr, poziomo=False)

    # ---------- Ledger ----------
    naglowek(led, 1, ["id", "klasa", "klucz", "opis", "wartość", "jednostka", "okres", "źródło",
                      "strona", "cytat / uzasadnienie / autor i licencja"])
    for i, w in enumerate(L.wpisy, start=2):
        z = w.get("zrodlo") or {}
        if w["klasa"] == "zalozenie":
            tekst = w.get("uzasadnienie")
        elif w["klasa"] == "grafika":
            tekst = f"{w.get('autor')} | {w.get('licencja')} | {w.get('plik')}"
        else:
            tekst = w.get("cytat")
        wartosci = [w["id"], w["klasa"], w["klucz"], w.get("opis"), w.get("wartosc"), w.get("jednostka"),
                    w.get("okres"), z.get("url") or z.get("dokument"), str(z.get("strona", "")), tekst]
        for j, v in enumerate(wartosci, start=1):
            c = led.cell(row=i, column=j, value=v)
            c.font = ZWYKLY
            c.alignment = Alignment(wrap_text=j in (4, 8, 10), vertical="top")
    szerokosci(led, {"A": 7, "B": 13, "C": 30, "D": 40, "E": 14, "F": 10, "G": 11, "H": 40, "I": 12, "J": 70})
    led.freeze_panes = "A2"
    drukowanie(led)

    # ---------- Podsumowanie ----------
    pod["A1"] = f"{nazwa} ({ticker}) - wycena metodą DCF"
    pod["A1"].font = TYTUL
    pod["A2"] = "Wszystkie liczby pochodzą z ledgera (arkusz Ledger) i są liczone formułami. Założenia to nie fakty."
    pod["A2"].font = SZARY
    naglowek(pod, 4, ["Wynik", "Wartość"])
    wyniki_pod = [("Wartość jednej akcji z DCF (USD)", f"=DCF!{wyc['na_akcje']}", USD),
                  ("Aktualny kurs (USD)", f"=DCF!{wyc['kurs']}", USD),
                  ("Różnica wartości wobec kursu", f"=DCF!{wyc['roznica']}", PROC),
                  ("Wartość firmy, EV (mln USD)", f"=DCF!{wyc['ev']}", MLN),
                  ("Wartość kapitału własnego (mln USD)", f"=DCF!{wyc['equity']}", MLN)]
    for i, (et, wz, fmt) in enumerate(wyniki_pod, start=5):
        pod.cell(row=i, column=1, value=et).font = POGRUBIONY if i == 5 else ZWYKLY
        c = pod.cell(row=i, column=2, value=wz)
        c.number_format = fmt
        c.font = POGRUBIONY if i == 5 else ZIELONY
        if i == 5:
            c.fill = WYNIK_TLO
    naglowek(pod, 11, ["Kluczowe założenia", "Wartość"])
    kluczowe = [("WACC", f"={ZAL}{za['wacc']}"),
                ("Stopa wolna od ryzyka", f"={ZAL}{za['rf']}"),
                ("Beta", f"={ZAL}{za['beta']}"),
                ("Premia za ryzyko rynkowe", f"={ZAL}{za['equity_risk_premium']}"),
                ("Wzrost rezydualny", f"={ZAL}{za['terminal_growth']}"),
                ("Marża operacyjna", f"={ZAL}{za['operating_margin']}"),
                (f"Wzrost przychodów {lata_prog[0]}", f"={ZAL}{za[f'revenue_growth_{lata_prog[0]}']}"),
                ("Stopa podatkowa", f"={ZAL}{za['tax_rate']}")]
    for i, (et, wz) in enumerate(kluczowe, start=12):
        pod.cell(row=i, column=1, value=et).font = ZWYKLY
        c = pod.cell(row=i, column=2, value=wz)
        c.number_format = "0.00" if et == "Beta" else PROC
        c.font = ZIELONY

    # Tabele pomocnicze do wykresów (kolumny H-K)
    naglowek(pod, 4, ["Rok", "Przychody (mln USD)", "Marża operacyjna"], kol_start=8)
    r = 5
    for j, rok in enumerate(lata_hist):
        k = get_column_letter(2 + j)
        pod.cell(row=r, column=8, value=f"{rok}")
        pod.cell(row=r, column=9, value=f"={DH}{k}4").number_format = MLN
        pod.cell(row=r, column=10, value=f"={DH}{k}6").number_format = PROC
        r += 1
    for i, rok in enumerate(lata_prog):
        pod.cell(row=r, column=8, value=f"{rok}P")
        pod.cell(row=r, column=9, value=f"=DCF!{kol[i]}5").number_format = MLN
        pod.cell(row=r, column=10, value=f"=DCF!{kol[i]}6").number_format = PROC
        r += 1
    ost_przych = r - 1

    r += 1
    naglowek(pod, r, ["Rok", "FCF (mln USD)"], kol_start=8)
    fcf_start = r + 1
    for i, rok in enumerate(lata_prog):
        pod.cell(row=fcf_start + i, column=8, value=f"{rok}")
        pod.cell(row=fcf_start + i, column=9, value=f"=DCF!{kol[i]}15").number_format = MLN
    fcf_kon = fcf_start + n - 1

    r = fcf_kon + 2
    naglowek(pod, r, ["Krok (USD na akcję)", "Podstawa", "Wartość"], kol_start=8)
    m_start = r + 1
    akcje = f"DCF!{wyc['akcje']}"
    most = [("Wartość firmy", "=0", f"=DCF!{wyc['ev']}/{akcje}"),
            ("+ gotówka i papiery", f"=I{m_start + 1 - 1}+J{m_start}-J{m_start}+J{m_start}",
             f"=(DCF!{wyc['gotowka']}+DCF!{wyc['papiery']})/{akcje}"),
            ("- dług", f"=J{m_start}+J{m_start + 1}-J{m_start + 2}", f"=DCF!{wyc['dlug']}/{akcje}"),
            ("Wartość jednej akcji", "=0", f"=DCF!{wyc['na_akcje']}")]
    most[1] = ("+ gotówka i papiery", f"=J{m_start}", most[1][2])
    for i, (et, baza, wart) in enumerate(most):
        pod.cell(row=m_start + i, column=8, value=et)
        pod.cell(row=m_start + i, column=9, value=baza).number_format = USD
        pod.cell(row=m_start + i, column=10, value=wart).number_format = USD
    m_kon = m_start + 3
    for wiersz in pod.iter_rows(min_row=4, max_row=m_kon, min_col=8, max_col=10):
        for c in wiersz:
            if c.font != POGRUBIONY:
                c.font = ZWYKLY

    # Wykres 1: przychody (słupki) + marża (linia, oś pomocnicza)
    w1 = BarChart()
    w1.title = "Przychody (mln USD) i marża operacyjna"
    w1.add_data(Reference(pod, min_col=9, min_row=4, max_row=ost_przych), titles_from_data=True)
    w1.set_categories(Reference(pod, min_col=8, min_row=5, max_row=ost_przych))
    w1.y_axis.title = "mln USD"
    w1.y_axis.numFmt = "#,##0"
    w1.y_axis.majorGridlines = None
    l1 = LineChart()
    l1.add_data(Reference(pod, min_col=10, min_row=4, max_row=ost_przych), titles_from_data=True)
    l1.y_axis.axId = 200
    l1.y_axis.title = "marża"
    l1.y_axis.numFmt = "0%"
    l1.y_axis.crosses = "max"
    l1.series[0].smooth = False
    w1 += l1
    w1.legend.position = "b"
    w1.height, w1.width = 8, 17
    pod.add_chart(w1, "A22")

    # Wykres 2: FCF
    w2 = BarChart()
    w2.title = "Wolne przepływy pieniężne (FCF) w prognozie, mln USD"
    w2.add_data(Reference(pod, min_col=9, min_row=fcf_start - 1, max_row=fcf_kon), titles_from_data=True)
    w2.set_categories(Reference(pod, min_col=8, min_row=fcf_start, max_row=fcf_kon))
    w2.y_axis.numFmt = "#,##0"
    w2.legend = None
    w2.x_axis.tickLblPos = "low"
    w2.height, w2.width = 8, 17
    pod.add_chart(w2, "A39")

    # Wykres 3: most od wartości firmy do wartości jednej akcji (słupki skumulowane z niewidoczną podstawą)
    w3 = BarChart()
    w3.type = "col"
    w3.grouping = "stacked"
    w3.overlap = 100
    w3.title = "Od wartości firmy do wartości jednej akcji (USD na akcję)"
    w3.add_data(Reference(pod, min_col=9, max_col=10, min_row=m_start - 1, max_row=m_kon), titles_from_data=True)
    w3.set_categories(Reference(pod, min_col=8, min_row=m_start, max_row=m_kon))
    w3.series[0].graphicalProperties.noFill = True
    w3.series[0].graphicalProperties.line.noFill = True
    w3.series[1].graphicalProperties.solidFill = "4472C4"
    w3.legend = None
    w3.y_axis.numFmt = "#,##0"
    w3.height, w3.width = 8, 17
    pod.add_chart(w3, "A56")

    szerokosci(pod, {"A": 40, "B": 16, "H": 22, "I": 18, "J": 16})
    drukowanie(pod, poziomo=False)
    pod.print_area = "A1:J73"

    wb.save(plik)
    return wyc


def przelicz(plik):
    """Przelicza skoroszyt w LibreOffice (bez okna) i nadpisuje plik wersją z policzonymi wartościami."""
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        profil = Path(tmp) / "profil"
        wyj = Path(tmp) / "wyj"
        subprocess.run([soffice(), f"-env:UserInstallation={profil.as_uri()}", "--headless",
                        "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(wyj), str(plik)],
                       check=True, capture_output=True, timeout=180)
        przeliczony = wyj / plik.name
        if not przeliczony.exists():
            sys.exit("LibreOffice nie przeliczył pliku.")
        return przeliczony.read_bytes()


def podglad(plik, katalog):
    import pymupdf
    katalog.mkdir(parents=True, exist_ok=True)
    for stary in katalog.glob("xlsx-*.png"):
        stary.unlink()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        profil = Path(tmp) / "profil"
        subprocess.run([soffice(), f"-env:UserInstallation={profil.as_uri()}", "--headless",
                        "--convert-to", "pdf", "--outdir", tmp, str(plik)],
                       check=True, capture_output=True, timeout=180)
        with pymupdf.open(Path(tmp) / (plik.stem + ".pdf")) as pdf:
            for i, strona in enumerate(pdf, start=1):
                strona.get_pixmap(dpi=90).save(katalog / f"xlsx-{i:02d}.png")
            return len(pdf)


def main():
    if len(sys.argv) != 2:
        sys.exit("Użycie: python scripts/generuj_xlsx.py TICKER")
    ticker = sys.argv[1].upper()
    wymagaj_weryfikacji(ticker)
    wyniki = katalog_spolki(ticker) / "wyniki"
    wyniki.mkdir(parents=True, exist_ok=True)
    plik = wyniki / f"{ticker}_DCF.xlsx"

    wyc = zbuduj(ticker, plik)
    przeliczony = przelicz(plik)

    # Szukamy błędów w przeliczonych wartościach.
    tmp_plik = wyniki / f".{ticker}_DCF_przeliczony.xlsx"
    tmp_plik.write_bytes(przeliczony)
    wartosci = load_workbook(tmp_plik, data_only=True)
    bledy = [f"{ark.title}!{c.coordinate}: {c.value}" for ark in wartosci for wiersz in ark.iter_rows()
             for c in wiersz if isinstance(c.value, str) and c.value in ERRORY]
    if bledy:
        tmp_plik.unlink()
        sys.exit("Błędy w formułach:\n" + "\n".join(bledy))

    d = wartosci["DCF"]
    wycena = {
        "ticker": ticker,
        "wartosc_jednej_akcji_usd": round(d[wyc["na_akcje"]].value, 2),
        "kurs_usd": d[wyc["kurs"]].value,
        "roznica_proc": round(d[wyc["roznica"]].value * 100, 1),
        "wacc_proc": round(d[wyc["wacc"]].value * 100, 2),
        "wzrost_rezydualny_proc": round(d[wyc["g"]].value * 100, 2),
        "wartosc_firmy_mln_usd": round(d[wyc["ev"]].value),
        "wartosc_kapitalu_mln_usd": round(d[wyc["equity"]].value),
        "zrodlo": f"{ticker}/wyniki/{plik.name} (arkusz DCF, przeliczony w LibreOffice)",
    }
    (wyniki / "wycena.json").write_text(json.dumps(wycena, ensure_ascii=False, indent=2), encoding="utf-8")

    # Plik z formułami zostaje głównym wynikiem (Excel przeliczy go przy otwarciu);
    # podgląd robimy z wersji przeliczonej.
    stron = podglad(tmp_plik, wyniki / "podglad")
    tmp_plik.unlink()

    print(f"Zapisano {plik}")
    print(f"Podgląd: {stron} stron w {wyniki / 'podglad'}")
    print(json.dumps(wycena, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
