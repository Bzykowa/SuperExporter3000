import pandas as pd
import pathlib
import re
import xml.etree.ElementTree as parser
from math import ceil
from xml_parsing.xml_parser import XMLParser


class Invoices(XMLParser):
    """The class responsible for parsing invoices in Excel files to xml
    file in Optima format."""

    def __init__(
        self, company_code: str, data_path: str,
        exchange_rates: dict, holidays: list, month: int
    ):
        """
        Invoices parser constructor

        :param company_code: Short code for a company in Optima
        :type company_code: str
        :param data_path: A path to a directory containing files to export
        :type data_path: str
        :param exchange_rates: A dict with data from NBP API EUR to PLN
        :type exchange_rates: dict
        :param holidays: A list of dates for holidays (no exchange rate)
        :type holidays: list
        :param month: Number of the month of the invoices batch
        :type month: int
        """
        super().__init__(company_code, data_path)
        self.client_data = pd.DataFrame(
            columns=[
                "IdFolder",
                "Kod", "Nazwa", "Nazwa2", "Nazwa3", "Telefon", "Telefon2",
                "TelefonSms", "Fax", "Ulica", "NrDomu", "NrLokalu",
                "KodPocztowy", "Poczta", "Miasto", "Kraj", "Wojewodztwo",
                "Powiat", "Gmina", "URL", "Grupa", "OsobaFizyczna", "NIP",
                "NIPKraj", "Zezwolenie", "Regon", "Pesel", "Email",
                "BankRachunekNr", "BankNazwa", "Osoba", "Opis", "Rodzaj",
                "PlatnikVAT", "PodatnikVatCzynny", "Eksport", "LimitKredytu",
                "Termin", "FormaPlatnosci", "Ceny", "CenyNazwa", "Upust",
                "NieNaliczajOdsetek", "MetodaKasowa", "WindykacjaEMail",
                "WindykacjaTelefonSms", "AlgorytmNettoBrutto", "Waluta"
            ],
            dtype=str
        )
        self.client_data.astype({"IdFolder": int})
        self.invoice_data = pd.DataFrame(
            columns=[
                "IdFolder", "Numer", "DataWystawienia", "Kwota",
                "KwotaEUR", "DataKursu"
            ]
        )
        self.exchange_data = exchange_rates
        self.holiday_data = holidays
        self.month = int(month)
        self.errors = []

        self.read_data()

    def read_data(self):
        """
        Extract data from invoices (Excel files in data_path directory)
        """
        # search for modern excel files
        path = pathlib.Path(self.data_path)

        files = [
            str(p.resolve()) for p in path.glob("[0-9]*.xlsx")
        ]

        # process invoices file by file
        for file in files:
            invoice = pd.ExcelFile(file)
            client_record = {
                "IdFolder": 0,
                "Kod": "", "Nazwa": "", "Nazwa2": "", "Nazwa3": "",
                "Telefon": "", "Telefon2": "", "TelefonSms": "",
                "Fax": "", "Ulica": "", "NrDomu": "", "NrLokalu": "",
                "KodPocztowy": "", "Poczta": "", "Miasto": "", "Kraj": "",
                "Wojewodztwo": "", "Powiat": "", "Gmina": "", "URL": "",
                "Grupa": "", "OsobaFizyczna": "", "NIP": "", "NIPKraj": "",
                "Zezwolenie": "", "Regon": "", "Pesel": "", "Email": "",
                "BankRachunekNr": "", "BankNazwa": "", "Osoba": "",
                "Opis": "", "Rodzaj": "", "PlatnikVAT": "",
                "PodatnikVatCzynny": "", "Eksport": "", "LimitKredytu": "",
                "Termin": "", "FormaPlatnosci": "", "Ceny": "",
                "CenyNazwa": "", "Upust": "", "NieNaliczajOdsetek": "",
                "MetodaKasowa": "", "WindykacjaEMail": "",
                "WindykacjaTelefonSms": "", "AlgorytmNettoBrutto": "",
                "Waluta": ""
            }
            invoice_record = {}
            # read data and process it
            # get the number and client code from file name 00 xxxx.xlsx
            data_file_name = (pathlib.Path(file).name).split(" ", maxsplit=1)
            invoice_record["IdFolder"] = int(data_file_name[0])
            client_record["IdFolder"] = int(data_file_name[0])
            client_record["Kod"] = data_file_name[1][:-5]

            # excel data
            # join split combo for double space removal
            client_record["Nazwa"] = " ".join(invoice.book["Tabelle1"].cell(
                row=12, column=7).value.strip().split())
            # regex is for spliting on , . and whitespace,
            # there will be errors anyway
            address1 = re.split(r'[,\.\s]+', invoice.book["Tabelle1"].cell(
                row=13, column=7).value.strip())
            # normal case, house number at the end (Kakestrasse 10)
            if any(char.isdigit() for char in address1[-1]):
                client_record["Ulica"] = (" ".join(address1[:-1])).strip()
                client_record["NrDomu"] = address1[-1].strip()
            # whitespace/,/. between house num and letter (Kakestrasse 10 A)
            elif any(char.isdigit() for char in address1[-2]):
                client_record["Ulica"] = (" ".join(address1[:-2])).strip()
                client_record["NrDomu"] = (" ".join(address1[-2:])).strip()
            # too verbose or too short address, impossible to predict the split
            # (Kakestrasse 10 A Stock 69. (Haus am See))
            else:
                client_record["Ulica"] = (" ".join(address1)).strip()
                client_record["NrDomu"] = ""
            address2 = invoice.book["Tabelle1"].cell(
                row=14, column=7).value.strip().split(" ")
            if any(char.isdigit() for char in address2[0]):
                client_record["Miasto"] = (" ".join(address2[1:])).strip()
                client_record["KodPocztowy"] = address2[0].strip()

            client_record["Kraj"] = "Niemcy"
            client_record["OsobaFizyczna"] = "0"
            client_record["PlatnikVAT"] = "1"
            client_record["PodatnikVatCzynny"] = "1"
            client_record["Eksport"] = "0"
            client_record["LimitKredytu"] = "0"
            client_record["Termin"] = "7"
            client_record["FormaPlatnosci"] = "przelew"
            client_record["Ceny"] = "0"
            client_record["CenyNazwa"] = "domyślna"
            client_record["Upust"] = "0"
            client_record["NieNaliczajOdsetek"] = "0"
            client_record["MetodaKasowa"] = "0"
            client_record["AlgorytmNettoBrutto"] = "0"

            invoice_record["DataWystawienia"] = self.read_date(
                invoice.book["Tabelle1"].cell(row=4, column=11).value
            )
            # date of return of first employee
            powrot1 = self.read_date(
                invoice.book["Tabelle1"].cell(row=20, column=6).value
            )
            # date of return of second employee
            powrot2 = self.read_date(
                invoice.book["Tabelle1"].cell(row=27, column=6).value
            )

            # exchange date is minimum of
            # max(powrot1, powrot2) and datawystawienia - 1 if months match
            # else datawystawienia - 1
            # if no rate for the day (sunday, saturday, holidays)
            # go back 1 day and check again
            if pd.isnull(powrot2) and not pd.isnull(powrot1):
                invoice_record["DataKursu"] = self.set_exchange_date(
                    powrot1
                ) if (powrot1 < invoice_record["DataWystawienia"]
                      and powrot1.month == self.month) else \
                    self.set_exchange_date(
                        invoice_record["DataWystawienia"]
                )
            elif not pd.isnull(powrot2) and not pd.isnull(powrot1):
                powrot = powrot1 if powrot1 > powrot2 else powrot2
                invoice_record["DataKursu"] = self.set_exchange_date(
                    powrot
                ) if (powrot < invoice_record["DataWystawienia"]
                      and powrot.month == self.month) else \
                    self.set_exchange_date(
                        invoice_record["DataWystawienia"]
                )
            else:
                invoice_record["DataKursu"] = pd.NaT \
                    if pd.isnull(invoice_record["DataWystawienia"]) else \
                    self.set_exchange_date(
                        invoice_record["DataWystawienia"]
                )

            invoice_record["Numer"] = invoice.book["Tabelle1"].cell(
                row=3, column=11).value

            if invoice.book["Tabelle1"].cell(row=39, column=11).value == 0:
                invoice_record["KwotaEUR"] = round(
                    invoice.book["Tabelle1"].cell(
                        row=40, column=11).value, 2)
            else:
                invoice_record["KwotaEUR"] = round(
                    invoice.book["Tabelle1"].cell(
                        row=37, column=11).value, 2)

            # checking data integrity
            dw_null = pd.isnull(invoice_record["DataWystawienia"])
            dk_null = pd.isnull(invoice_record["DataKursu"])
            if dw_null:
                self.errors.append(
                    "{} : Brak daty wystawienia / nieparsowalna".format(
                        pathlib.Path(file).name))
            if ((not dw_null) and
                    (invoice_record["DataWystawienia"].month != self.month)):
                self.errors.append(
                    "{} : Data wystawienia nie z miesiąca {}".format(
                        pathlib.Path(file).name, self.month)
                )
            if dk_null:
                self.errors.append(
                    "{} : Nie można ustalić daty kursu".format(
                        pathlib.Path(file).name))
            if not dk_null:
                try:
                    exchange_idx = invoice_record["DataKursu"].strftime(
                        "%Y-%m-%d")
                    invoice_record["Kwota"] = round(
                        invoice_record["KwotaEUR"] *
                        self.exchange_data[exchange_idx], 2)
                except KeyError:
                    self.errors.append(
                        "{} : Brak kursu na dzień {}".format(
                            pathlib.Path(file).name, exchange_idx))
            if invoice_record["KwotaEUR"] == 0:
                self.errors.append(
                    "{} : kwota na fakturze równa 0".format(
                        pathlib.Path(file).name)
                )

            self.client_data.loc[len(self.client_data)] = client_record
            self.invoice_data.loc[len(self.invoice_data)] = invoice_record

        # sort numerically
        self.client_data.sort_values(by=["IdFolder"], inplace=True)
        self.invoice_data.sort_values(by=["IdFolder"], inplace=True)
        self.client_data.reset_index(drop=True, inplace=True)
        self.invoice_data.reset_index(drop=True, inplace=True)
        self.check_gaps()

    def gen_xml_layout(self):
        """
        Generate the layout for invoice records.
        """
        self.records = parser.SubElement(self.root, "REJESTRY_SPRZEDAZY_VAT")
        self.records.set("xmlns", "")

        version = parser.SubElement(self.records, "WERSJA")
        version.text = self.cdata_wrap("2.00")
        zdr_id = parser.SubElement(self.records, "BAZA_ZRD_ID")
        zdr_id.text = self.cdata_wrap(self.company_code)
        doc_id = parser.SubElement(self.records, "BAZA_DOC_ID")
        doc_id.text = self.cdata_wrap(self.company_code)

        self.invoices = [
            self.gen_invoice_record(idx)
            for idx in range(len(self.invoice_data))
        ]

        self.records.extend(self.invoices)

    def gen_invoice_record(self, idx: int):
        """
        Generate a record for an invoice

        :param idx: Id of the document in exported data
        :type idx: int
        """
        # main tag
        invoice = parser.Element("REJESTR_SPRZEDAZY_VAT")

        # document type
        id_zrodla = parser.SubElement(invoice, "ID_ZRODLA")
        id_zrodla.text = self.cdata_wrap("")
        modul = parser.SubElement(invoice, "MODUL")
        modul.text = self.cdata_wrap("Rejestr Vat")
        typ = parser.SubElement(invoice, "TYP")
        typ.text = self.cdata_wrap("Rejestr sprzedazy")
        rejestr = parser.SubElement(invoice, "REJESTR")
        rejestr.text = self.cdata_wrap("SPRZEDAŻ")

        # document's data
        data_wystawienia = parser.SubElement(invoice, "DATA_WYSTAWIENIA")
        data_wystawienia.text = self.cdata_wrap(
            self.invoice_data.at[idx, "DataWystawienia"].strftime("%d.%m.%Y"))
        data_sprzedazy = parser.SubElement(invoice, "DATA_SPRZEDAZY")
        data_sprzedazy.text = self.cdata_wrap(
            self.invoice_data.at[idx, "DataWystawienia"].strftime("%d.%m.%Y"))
        termin = parser.SubElement(invoice, "TERMIN")
        termin.text = self.cdata_wrap(
            (self.invoice_data.at[idx, "DataWystawienia"] +
             pd.Timedelta(days=7)).strftime("%d.%m.%Y")
        )
        data_obowiazku_podatkowego = parser.SubElement(
            invoice, "DATA_DATAOBOWIAZKUPODATKOWEGO")
        data_obowiazku_podatkowego.text = self.cdata_wrap(
            self.invoice_data.at[idx, "DataWystawienia"].strftime("%d.%m.%Y"))
        data_prawa_odliczenia = parser.SubElement(
            invoice, "DATA_DATAPRAWAODLICZENIA")
        data_prawa_odliczenia.text = self.cdata_wrap(
            self.invoice_data.at[idx, "DataWystawienia"].strftime("%d.%m.%Y"))
        numer = parser.SubElement(invoice, "NUMER")
        numer.text = self.cdata_wrap(self.invoice_data.at[idx, "Numer"])
        korekta = parser.SubElement(invoice, "KOREKTA")
        korekta.text = self.cdata_wrap("Nie")
        korekta_numer = parser.SubElement(invoice, "KOREKTA_NUMER")
        korekta_numer.text = self.cdata_wrap("")
        wewnetrzna = parser.SubElement(invoice, "WEWNETRZNA")
        wewnetrzna.text = self.cdata_wrap("Nie")
        metoda_kasowa = parser.SubElement(invoice, "METODA_KASOWA")
        metoda_kasowa.text = self.cdata_wrap("Nie")
        fiskalna = parser.SubElement(invoice, "FISKALNA")
        fiskalna.text = self.cdata_wrap("Nie")
        detaliczna = parser.SubElement(invoice, "DETALICZNA")
        detaliczna.text = self.cdata_wrap("Nie")
        eksport = parser.SubElement(invoice, "EKSPORT")
        eksport.text = self.cdata_wrap("Nie")
        finalny = parser.SubElement(invoice, "FINALNY")
        finalny.text = self.cdata_wrap("Nie")

        # client's data
        podatnik_czynny = parser.SubElement(invoice, "PODATNIK_CZYNNY")
        podatnik_czynny.text = self.cdata_wrap("Tak")
        identyfikator_księgowy = parser.SubElement(
            invoice, "IDENTYFIKATOR_KSIEGOWY")
        identyfikator_księgowy.text = self.cdata_wrap("")
        typ_podmiotu = parser.SubElement(invoice, "TYP_PODMIOTU")
        typ_podmiotu.text = self.cdata_wrap("kontrahent")
        podmiot = parser.SubElement(invoice, "PODMIOT")
        podmiot.text = self.cdata_wrap(self.client_data.at[idx, "Kod"])
        podmiot_id = parser.SubElement(invoice, "PODMIOT_ID")
        podmiot_id.text = self.cdata_wrap("")
        podmiot_nip = parser.SubElement(invoice, "PODMIOT_NIP")
        podmiot_nip.text = self.cdata_wrap("")
        nazwa1 = parser.SubElement(invoice, "NAZWA1")
        nazwa1.text = self.cdata_wrap(
            " ".join(self.client_data.at[idx, "Nazwa"].split()))
        nazwa2 = parser.SubElement(invoice, "NAZWA2")
        nazwa2.text = self.cdata_wrap("")
        nazwa3 = parser.SubElement(invoice, "NAZWA3")
        nazwa3.text = self.cdata_wrap("")
        nip_kraj = parser.SubElement(invoice, "NIP_KRAJ")
        nip_kraj.text = self.cdata_wrap("")
        nip = parser.SubElement(invoice, "NIP")
        nip.text = self.cdata_wrap("")

        # client's address
        kraj = parser.SubElement(invoice, "KRAJ")
        kraj.text = self.cdata_wrap(self.client_data.at[idx, "Kraj"])
        wojewodztwo = parser.SubElement(invoice, "WOJEWODZTWO")
        wojewodztwo.text = self.cdata_wrap("")
        powiat = parser.SubElement(invoice, "POWIAT")
        powiat.text = self.cdata_wrap("")
        gmina = parser.SubElement(invoice, "GMINA")
        gmina.text = self.cdata_wrap("")
        ulica = parser.SubElement(invoice, "ULICA")
        ulica.text = self.cdata_wrap(self.client_data.at[idx, "Ulica"])
        nr_domu = parser.SubElement(invoice, "NR_DOMU")
        nr_domu.text = self.cdata_wrap(self.client_data.at[idx, "NrDomu"])
        nr_lokalu = parser.SubElement(invoice, "NR_LOKALU")
        nr_lokalu.text = self.cdata_wrap("")
        miasto = parser.SubElement(invoice, "MIASTO")
        miasto.text = self.cdata_wrap(self.client_data.at[idx, "Miasto"])
        kod_pocztowy = parser.SubElement(invoice, "KOD_POCZTOWY")
        kod_pocztowy.text = self.cdata_wrap(
            self.client_data.at[idx, "KodPocztowy"])
        poczta = parser.SubElement(invoice, "POCZTA")
        poczta.text = self.cdata_wrap(self.client_data.at[idx, "Miasto"])
        dodatkowe = parser.SubElement(invoice, "DODATKOWE")
        dodatkowe.text = self.cdata_wrap("")
        pesel = parser.SubElement(invoice, "PESEL")
        pesel.text = self.cdata_wrap("")
        rolnik = parser.SubElement(invoice, "ROLNIK")
        rolnik.text = self.cdata_wrap("Nie")

        # payment's data
        typ_platnika = parser.SubElement(invoice, "TYP_PLATNIKA")
        typ_platnika.text = self.cdata_wrap("kontrahent")
        platnik = parser.SubElement(invoice, "PLATNIK")
        platnik.text = self.cdata_wrap(self.client_data.at[idx, "Kod"])
        platnik_id = parser.SubElement(invoice, "PLATNIK_ID")
        platnik_id.text = self.cdata_wrap("")
        platnik_nip = parser.SubElement(invoice, "PLATNIK_NIP")
        platnik_nip.text = self.cdata_wrap("")
        platnik_rachunek_nr = parser.SubElement(invoice, "PLATNIK_RACHUNEK_NR")
        platnik_rachunek_nr.text = self.cdata_wrap("DE92480501610015540842")
        kategoria = parser.SubElement(invoice, "KATEGORIA")
        kategoria.text = self.cdata_wrap("USŁUGI")
        kategoria_id = parser.SubElement(invoice, "KATEGORIA_ID")
        kategoria_id.text = self.cdata_wrap("")
        opis = parser.SubElement(invoice, "OPIS")
        opis.text = self.cdata_wrap("Usługi opiekuńcze")
        forma_platnosci = parser.SubElement(invoice, "FORMA_PLATNOSCI")
        forma_platnosci.text = self.cdata_wrap("przelew")
        forma_platnosci_id = parser.SubElement(
            invoice, "FORMA_PLATNOSCI_ID")
        forma_platnosci_id.text = self.cdata_wrap("")
        deklaracja_vat7 = parser.SubElement(invoice, "DEKLARACJA_VAT7")
        deklaracja_vat7.text = self.cdata_wrap(
            self.invoice_data.at[idx, "DataWystawienia"].strftime("%Y-%m")
        )
        deklaracja_vatue = parser.SubElement(invoice, "DEKLARACJA_VATUE")
        deklaracja_vatue.text = self.cdata_wrap("Nie")
        waluta = parser.SubElement(invoice, "WALUTA")
        waluta.text = self.cdata_wrap("EUR")
        kurs_waluty = parser.SubElement(invoice, "KURS_WALUTY")
        kurs_waluty.text = self.cdata_wrap("NBP")
        notowanie_waluty_ile = parser.SubElement(
            invoice, "NOTOWANIE_WALUTY_ILE")
        notowanie_waluty_ile.text = self.cdata_wrap(
            self.exchange_data[self.invoice_data.at[
                idx, "DataKursu"].strftime("%Y-%m-%d")]
        )
        notowanie_waluty_za_ile = parser.SubElement(
            invoice, "NOTOWANIE_WALUTY_ZA_ILE")
        notowanie_waluty_za_ile.text = self.cdata_wrap("1")
        data_kursu = parser.SubElement(invoice, "DATA_KURSU")
        data_kursu.text = self.cdata_wrap(
            self.invoice_data.at[idx, "DataKursu"].strftime("%Y-%m-%d"))
        kurs_do_ksiegowania = parser.SubElement(invoice, "KURS_DO_KSIEGOWANIA")
        kurs_do_ksiegowania.text = self.cdata_wrap("Nie")
        kurs_waluty2 = parser.SubElement(invoice, "KURS_WALUTY_2")
        kurs_waluty2.text = self.cdata_wrap("NBP")
        notowanie_waluty_ile2 = parser.SubElement(
            invoice, "NOTOWANIE_WALUTY_ILE_2")
        notowanie_waluty_ile2.text = self.cdata_wrap(
            self.exchange_data[self.invoice_data.at[
                idx, "DataKursu"].strftime("%Y-%m-%d")]
        )
        notowanie_waluty_za_ile2 = parser.SubElement(
            invoice, "NOTOWANIE_WALUTY_ZA_ILE_2")
        notowanie_waluty_za_ile2.text = self.cdata_wrap("1")
        data_kursu2 = parser.SubElement(invoice, "DATA_KURSU_2")
        data_kursu2.text = self.cdata_wrap(
            self.invoice_data.at[idx, "DataKursu"].strftime("%Y-%m-%d"))
        platnosc_vat_w_pln = parser.SubElement(invoice, "PLATNOSC_VAT_W_PLN")
        platnosc_vat_w_pln.text = self.cdata_wrap("Nie")
        akcyza_za_wegiel = parser.SubElement(invoice, "AKCYZA_NA_WEGIEL")
        akcyza_za_wegiel.text = self.cdata_wrap(0)
        akcyza_za_wegiel_kolumna_kpr = parser.SubElement(
            invoice, "AKCYZA_NA_WEGIEL_KOLUMNA_KPR")
        akcyza_za_wegiel_kolumna_kpr.text = self.cdata_wrap("nie księgować")
        jpk_fa = parser.SubElement(invoice, "JPK_FA")
        jpk_fa.text = self.cdata_wrap("Tak")
        deklaracja_vat27 = parser.SubElement(invoice, "DEKLARACJA_VAT27")
        deklaracja_vat27.text = self.cdata_wrap("Nie")

        # invoice positions
        pozycje = parser.SubElement(invoice, "POZYCJE")
        pozycja = parser.SubElement(pozycje, "POZYCJA")
        kategoria_pos = parser.SubElement(pozycja, "KATEGORIA_POS")
        kategoria_pos.text = self.cdata_wrap("USŁUGI")
        kategoria_id_pos = parser.SubElement(pozycja, "KATEGORIA_ID_POS")
        kategoria_id_pos.text = self.cdata_wrap("")
        stawka_vat = parser.SubElement(pozycja, "STAWKA_VAT")
        stawka_vat.text = self.cdata_wrap(0)
        status_vat = parser.SubElement(pozycja, "STATUS_VAT")
        status_vat.text = self.cdata_wrap("zwolniona")
        netto = parser.SubElement(pozycja, "NETTO")
        netto.text = self.cdata_wrap(self.invoice_data.at[idx, "KwotaEUR"])
        vat = parser.SubElement(pozycja, "VAT")
        vat.text = self.cdata_wrap(0)
        netto_sys = parser.SubElement(pozycja, "NETTO_SYS")
        netto_sys.text = self.cdata_wrap(self.invoice_data.at[idx, "Kwota"])
        vat_sys = parser.SubElement(pozycja, "VAT_SYS")
        vat_sys.text = self.cdata_wrap(0)
        netto_sys2 = parser.SubElement(pozycja, "NETTO_SYS2")
        netto_sys2.text = self.cdata_wrap(self.invoice_data.at[idx, "Kwota"])
        vat_sys2 = parser.SubElement(pozycja, "VAT_SYS2")
        vat_sys2.text = self.cdata_wrap(0)
        rodzaj_sprzedazy = parser.SubElement(pozycja, "RODZAJ_SPRZEDAZY")
        rodzaj_sprzedazy.text = self.cdata_wrap("usługi")
        uwz_w_proporcji = parser.SubElement(pozycja, "UWZ_W_PROPORCJI")
        uwz_w_proporcji.text = self.cdata_wrap("warunkowo")
        kolumna_kpr = parser.SubElement(pozycja, "KOLUMNA_KPR")
        kolumna_kpr.text = self.cdata_wrap("Sprzedaż")
        kolumna_ryczalt = parser.SubElement(pozycja, "KOLUMNA_RYCZALT")
        kolumna_ryczalt.text = self.cdata_wrap("Nie księgować")
        opis_poz = parser.SubElement(pozycja, "OPIS_POZ")
        opis_poz.text = self.cdata_wrap("Usługi opiekuńcze")
        opis_poz2 = parser.SubElement(pozycja, "OPIS_POZ_2")
        opis_poz2.text = self.cdata_wrap("")

        parser.SubElement(invoice, "KWOTY_DODATKOWE")

        # payments' data section
        # payments' main tag
        platnosci = parser.SubElement(invoice, "PLATNOSCI")
        # payment's main tag
        platnosc = parser.SubElement(platnosci, "PLATNOSC")

        # other data
        id_zrodla_platnosci = parser.SubElement(platnosc, "ID_ZRODLA_PLAT")
        id_zrodla_platnosci.text = self.cdata_wrap("")
        termin_plat = parser.SubElement(platnosc, "TERMIN_PLAT")
        termin_plat.text = self.cdata_wrap(
            (self.invoice_data.at[idx, "DataWystawienia"] +
             pd.Timedelta(days=7)).strftime("%d.%m.%Y")
        )
        forma_platnosci_plat = parser.SubElement(
            platnosc, "FORMA_PLATNOSCI_PLAT")
        forma_platnosci_plat.text = self.cdata_wrap("przelew")
        forma_platnosci_id_plat = parser.SubElement(
            platnosc, "FORMA_PLATNOSCI_ID_PLAT")
        forma_platnosci_id_plat.text = self.cdata_wrap("")
        kwota_plat = parser.SubElement(platnosc, "KWOTA_PLAT")
        kwota_plat.text = self.cdata_wrap(
            self.invoice_data.at[idx, "KwotaEUR"])
        waluta_plat = parser.SubElement(platnosc, "WALUTA_PLAT")
        waluta_plat.text = self.cdata_wrap("EUR")
        kurs_waluty_plat = parser.SubElement(platnosc, "KURS_WALUTY_PLAT")
        kurs_waluty_plat.text = self.cdata_wrap("NBP")
        notowanie_waluty_ile_plat = parser.SubElement(
            platnosc, "NOTOWANIE_WALUTY_ILE_PLAT")
        notowanie_waluty_ile_plat.text = self.cdata_wrap(
            self.exchange_data[self.invoice_data.at[
                idx, "DataKursu"].strftime("%Y-%m-%d")]
        )
        notowanie_waluty_za_ile_plat = parser.SubElement(
            platnosc, "NOTOWANIE_WALUTY_ZA_ILE_PLAT")
        notowanie_waluty_za_ile_plat.text = self.cdata_wrap("1")
        kwota_pln_plat = parser.SubElement(platnosc, "KWOTA_PLN_PLAT")
        kwota_pln_plat.text = self.cdata_wrap(
            self.invoice_data.at[idx, "Kwota"])
        kierunek = parser.SubElement(platnosc, "KIERUNEK")
        kierunek.text = self.cdata_wrap("przychód")
        podlega_rozliczeniu = parser.SubElement(
            platnosc, "PODLEGA_ROZLICZENIU")
        podlega_rozliczeniu.text = self.cdata_wrap("tak")
        konto = parser.SubElement(platnosc, "KONTO")
        konto.text = self.cdata_wrap("")
        nie_naliczaj_odsetek = parser.SubElement(
            platnosc, "NIE_NALICZAJ_ODSETEK")
        nie_naliczaj_odsetek.text = self.cdata_wrap("Nie")
        przelew_sepa = parser.SubElement(platnosc, "PRZELEW_SEPA")
        przelew_sepa.text = self.cdata_wrap("Nie")
        data_kursu_plat = parser.SubElement(platnosc, "DATA_KURSU_PLAT")
        data_kursu_plat.text = self.cdata_wrap(
            self.invoice_data.at[idx, "DataKursu"].strftime("%Y-%m-%d"))
        waluta_dok = parser.SubElement(platnosc, "WALUTA_DOK")
        waluta_dok.text = self.cdata_wrap("EUR")
        platnosc_typ_podmiotu = parser.SubElement(
            platnosc, "PLATNOSC_TYP_PODMIOTU")
        platnosc_typ_podmiotu.text = self.cdata_wrap("kontrahent")
        platnosc_podmiot = parser.SubElement(platnosc, "PLATNOSC_PODMIOT")
        platnosc_podmiot.text = self.cdata_wrap(
            self.client_data.at[idx, "Kod"])
        platnosc_podmiot_id = parser.SubElement(
            platnosc, "PLATNOSC_PODMIOT_ID")
        platnosc_podmiot_id.text = self.cdata_wrap("")
        platnosc_podmiot_nip = parser.SubElement(
            platnosc, "PLATNOSC_PODMIOT_NIP")
        platnosc_podmiot_nip.text = self.cdata_wrap("")
        platnosc_podmiot_rachunek_nr = parser.SubElement(
            platnosc, "PLATNOSC_PODMIOT_RACHUNEK_NR")
        platnosc_podmiot_rachunek_nr.text = self.cdata_wrap(
            "DE92480501610015540842")
        plat_kategoria = parser.SubElement(platnosc, "PLAT_KATEGORIA")
        plat_kategoria.text = self.cdata_wrap("USŁUGI")
        plat_kategoria_id = parser.SubElement(platnosc, "PLAT_KATEGORIA_ID")
        plat_kategoria_id.text = self.cdata_wrap("")
        plat_elixir_01 = parser.SubElement(platnosc, "PLAT_ELIXIR_O1")
        plat_elixir_01.text = self.cdata_wrap(
            "Zapłata za {}".format(self.invoice_data.at[idx, "Numer"])
        )
        plat_elixir_02 = parser.SubElement(platnosc, "PLAT_ELIXIR_O2")
        plat_elixir_02.text = self.cdata_wrap("")
        plat_elixir_03 = parser.SubElement(platnosc, "PLAT_ELIXIR_O3")
        plat_elixir_03.text = self.cdata_wrap("")
        plat_elixir_04 = parser.SubElement(platnosc, "PLAT_ELIXIR_O4")
        plat_elixir_04.text = self.cdata_wrap("")

        return invoice

    def get_clients_data(self):
        """
        Returns a DataFrame with exported client data
        from Excel invoices.
        """
        return self.client_data

    def verify_data(self):
        """
        Return a list of errors found during data load.
        """
        print(self.errors)
        return self.errors

    def read_date(self, date):
        """
        Helper function to read dates from Excel in a proper format.
        Errors will return NaT.

        :param date: Value to cast to pandas.Timestamp
        """
        if isinstance(date, pd.Timestamp):
            return date
        else:
            return pd.to_datetime(date, errors="coerce")

    def set_exchange_date(self, date: pd.Timestamp):
        """
        Docstring for set_exchange_date

        :param date: Start date from which the exchange date is calculated
        :type date: pd.Timestamp
        """
        exchange = date - pd.Timedelta(days=1)

        while (exchange.weekday() == 6 or exchange.weekday() == 5 or
               exchange.strftime("%Y-%m-%d") in self.holiday_data):
            exchange = exchange - pd.Timedelta(days=1)

        return exchange

    def check_gaps(self):
        """
        Check if there are any missing invoices.

        """
        inv = self.invoice_data["IdFolder"].tolist()
        if not inv:
            return
        start = inv[0]
        j = 0

        for i in range(start, len(inv)+start):
            if inv[j] != i:
                self.errors.append(
                    "{} : nie ma faktury".format(i)
                )
            j += 1

    def split_xml(self, max_records):
        """
        Split one big xml file into smaller chunks with maximum number of
        records in one file equal to max_records.

        :param max_records: a limit for children in a single file
        :type max_records: int
        """

        if max_records >= len(self.invoices):
            return

        file_num = ceil(len(self.invoices) / max_records)

        for i in range(0, file_num):
            # set up the layout of the split document
            root = parser.Element("ROOT")
            root.set("xmlns", "http://www.comarch.pl/cdn/optima/offline")
            records = parser.SubElement(root, "REJESTRY_SPRZEDAZY_VAT")
            records.set("xmlns", "")
            version = parser.SubElement(records, "WERSJA")
            version.text = self.cdata_wrap("2.00")
            zdr_id = parser.SubElement(records, "BAZA_ZRD_ID")
            zdr_id.text = self.cdata_wrap(self.company_code)
            doc_id = parser.SubElement(records, "BAZA_DOC_ID")
            doc_id.text = self.cdata_wrap(self.company_code)

            inv_slice = self.invoices[
                i * max_records:max((i+1)*max_records, len(self.invoices))
            ]
            records.extend(inv_slice)

            self.split.append(root)
