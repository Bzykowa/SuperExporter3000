import pandas as pd
import xml.etree.ElementTree as parser
from xml_parsing.xml_parser import XMLParser


class Delegations(XMLParser):
    """The class responsible for parsing Excel data about delegations to xml
    file in Optima format."""

    def __init__(self, company_code: str, data_path: str):
        super().__init__(company_code, data_path)
        # read the data from submitted file
        self.read_data()

    def read_data(self):
        self.data = pd.read_excel(
            self.data_path, sheet_name="do 30", decimal=",", skiprows=1,
            header=None, keep_default_na=False, dtype={
                0: str, 1: str}
        )

    def gen_xml_layout(self):
        """
        Generate the layout for delegation records.
        """
        self.records = parser.SubElement(self.root, "DOKUMENTY_INNE_ROZCHOD")
        self.records.set("xmlns", "")

        version = parser.SubElement(self.records, "WERSJA")
        version.text = self.cdata_wrap("2.00")
        zdr_id = parser.SubElement(self.records, "BAZA_ZRD_ID")
        zdr_id.text = self.cdata_wrap(self.company_code)
        doc_id = parser.SubElement(self.records, "BAZA_DOC_ID")
        doc_id.text = self.cdata_wrap(self.company_code)

        delegations = [
            self.gen_delegation_xml(row)
            for row in self.data.itertuples(index=False, name="Delegation")
        ]

        self.records.extend(delegations)

    def gen_delegation_xml(self, row: tuple):
        """
        Generate xml element for a delegation record
        """
        # todo: figure out if the order of tags matters
        # main tag
        delegation = parser.Element("DOKUMENT_INNY_ROZCHOD")

        # document type
        id_zrodla = parser.SubElement(delegation, "ID_ZRODLA")
        id_zrodla.text = self.cdata_wrap("")
        typ = parser.SubElement(delegation, "TYP")
        typ.text = self.cdata_wrap("Ewidencja dodatkowa kosztow")
        symbol_dokumentu = parser.SubElement(delegation, "SYMBOL_DOKUMENTU")
        symbol_dokumentu.text = self.cdata_wrap("EDK")
        symbol_dokumentu_id = parser.SubElement(
            delegation, "SYMBOL_DOKUMENTU_ID")
        symbol_dokumentu_id.text = self.cdata_wrap("")

        # document's data
        data_wystawienia = parser.SubElement(delegation, "DATA_WYSTAWIENIA")
        data_wystawienia.text = self.cdata_wrap(row._36.strftime("%d.%m.%Y"))
        data_operacji = parser.SubElement(delegation, "DATA_OPERACJI")
        data_operacji.text = self.cdata_wrap(row._18.strftime("%d.%m.%Y"))
        data_wplywu = parser.SubElement(delegation, "DATA_WPLYWU")
        data_wplywu.text = self.cdata_wrap(row._18.strftime("%d.%m.%Y"))
        numer = parser.SubElement(delegation, "NUMER")
        numer.text = self.cdata_wrap(row._2)
        numer_obcy = parser.SubElement(delegation, "NUMER_OBCY")
        numer_obcy.text = self.cdata_wrap(row._4)
        rejestr = parser.SubElement(delegation, "REJESTR")
        rejestr.text = self.cdata_wrap("KOSZTY")
        typ_podmiotu = parser.SubElement(delegation, "TYP_PODMIOTU")
        typ_podmiotu.text = self.cdata_wrap("pracownik")

        # employee's data
        podmiot = parser.SubElement(delegation, "PODMIOT")
        podmiot.text = self.cdata_wrap(row._0)
        podmiot_id = parser.SubElement(delegation, "PODMIOT_ID")
        podmiot_id.text = self.cdata_wrap("")
        podmiot_nip = parser.SubElement(delegation, "PODMIOT_NIP")
        podmiot_nip.text = self.cdata_wrap("")
        nazwa1 = parser.SubElement(delegation, "NAZWA1")
        nazwa1.text = self.cdata_wrap(" ".join(row._5.split()))
        nazwa2 = parser.SubElement(delegation, "NAZWA2")
        nazwa2.text = self.cdata_wrap("")
        nazwa3 = parser.SubElement(delegation, "NAZWA3")
        nazwa3.text = self.cdata_wrap("")
        nip_kraj = parser.SubElement(delegation, "NIP_KRAJ")
        nip_kraj.text = self.cdata_wrap("")
        nip = parser.SubElement(delegation, "NIP")
        nip.text = self.cdata_wrap("")

        # employee's address, currently unused
        kraj = parser.SubElement(delegation, "KRAJ")
        kraj.text = self.cdata_wrap("Polska")
        wojewodztwo = parser.SubElement(delegation, "WOJEWODZTWO")
        wojewodztwo.text = self.cdata_wrap("")
        powiat = parser.SubElement(delegation, "POWIAT")
        powiat.text = self.cdata_wrap("")
        gmina = parser.SubElement(delegation, "GMINA")
        gmina.text = self.cdata_wrap("")
        ulica = parser.SubElement(delegation, "ULICA")
        ulica.text = self.cdata_wrap("")
        nr_domu = parser.SubElement(delegation, "NR_DOMU")
        nr_domu.text = self.cdata_wrap("")
        nr_lokalu = parser.SubElement(delegation, "NR_LOKALU")
        nr_lokalu.text = self.cdata_wrap("")
        miasto = parser.SubElement(delegation, "MIASTO")
        miasto.text = self.cdata_wrap("")
        kod_pocztowy = parser.SubElement(delegation, "KOD_POCZTOWY")
        kod_pocztowy.text = self.cdata_wrap("")
        poczta = parser.SubElement(delegation, "POCZTA")
        poczta.text = self.cdata_wrap("")
        dodatkowe = parser.SubElement(delegation, "DODATKOWE")
        dodatkowe.text = self.cdata_wrap("")

        # payment's data
        typ_platnika = parser.SubElement(delegation, "TYP_PLATNIKA")
        typ_platnika.text = self.cdata_wrap("pracownik")
        platnik = parser.SubElement(delegation, "PLATNIK")
        platnik.text = self.cdata_wrap(row._0)
        platnik_id = parser.SubElement(delegation, "PLATNIK_ID")
        platnik_id.text = self.cdata_wrap("")
        platnik_nip = parser.SubElement(delegation, "PLATNIK_NIP")
        platnik_nip.text = self.cdata_wrap("")
        kategoria = parser.SubElement(delegation, "KATEGORIA")
        kategoria.text = self.cdata_wrap("DELEGACJE ZAGR. OP")
        kategoria_id = parser.SubElement(delegation, "KATEGORIA_ID")
        kategoria_id.text = self.cdata_wrap("")
        opis = parser.SubElement(delegation, "OPIS")
        opis.text = self.cdata_wrap("Podróże służbowe zagraniczne")
        kwota_razem = parser.SubElement(delegation, "KWOTA_RAZEM")
        kwota_razem.text = self.cdata_wrap(row._23)
        waluta = parser.SubElement(delegation, "WALUTA")
        waluta.text = self.cdata_wrap("EUR")
        kurs_waluty = parser.SubElement(delegation, "KURS_WALUTY")
        kurs_waluty.text = self.cdata_wrap("NBP")
        notowanie_waluty_ile = parser.SubElement(
            delegation, "NOTOWANIE_WALUTY_ILE")
        notowanie_waluty_ile.text = self.cdata_wrap(row._30)
        notowanie_waluty_za_ile = parser.SubElement(
            delegation, "NOTOWANIE_WALUTY_ZA_ILE")
        notowanie_waluty_za_ile.text = self.cdata_wrap("1")
        data_kursu = parser.SubElement(delegation, "DATA_KURSU")
        data_kursu.text = self.cdata_wrap(self.set_exchange_date(row._36))
        kwota_razem_pln = parser.SubElement(delegation, "KWOTA_RAZEM_PLN")
        kwota_razem_pln.text = self.cdata_wrap(row._35)
        parser.SubElement(delegation, "POZYCJE")
        forma_platnosci = parser.SubElement(delegation, "FORMA_PLATNOSCI")
        forma_platnosci.text = self.cdata_wrap("przelew")
        forma_platnosci_id = parser.SubElement(
            delegation, "FORMA_PLATNOSCI_ID")
        forma_platnosci_id.text = self.cdata_wrap("")
        termin = parser.SubElement(delegation, "TERMIN")
        termin.text = self.cdata_wrap(
            (row._36 + pd.Timedelta(days=7)).strftime("%d.%m.%Y")
        )
        generacja_platnosci = parser.SubElement(
            delegation, "GENERACJA_PLATNOSCI")
        generacja_platnosci.text = self.cdata_wrap("Tak")

        # payments' data section
        # payments' main tag
        platnosci = parser.SubElement(delegation, "PLATNOSCI")
        # payment's main tag
        platnosc = parser.SubElement(platnosci, "PLATNOSC")

        # other data
        id_zrodla_platnosci = parser.SubElement(platnosc, "ID_ZRODLA_PLAT")
        id_zrodla_platnosci.text = self.cdata_wrap("")
        termin_plat = parser.SubElement(platnosc, "TERMIN_PLAT")
        termin_plat.text = self.cdata_wrap(
            (row._36 + pd.Timedelta(days=7)).strftime("%d.%m.%Y")
        )
        forma_platnosci_plat = parser.SubElement(
            delegation, "FORMA_PLATNOSCI_PLAT")
        forma_platnosci_plat.text = self.cdata_wrap("przelew")
        forma_platnosci_id_plat = parser.SubElement(
            delegation, "FORMA_PLATNOSCI_ID_PLAT")
        forma_platnosci_id_plat.text = self.cdata_wrap("")
        kwota_plat = parser.SubElement(platnosc, "KWOTA_PLAT")
        kwota_plat.text = self.cdata_wrap(row._23)
        waluta_plat = parser.SubElement(platnosc, "WALUTA_PLAT")
        waluta_plat.text = self.cdata_wrap("EUR")
        kurs_waluty_plat = parser.SubElement(platnosc, "KURS_WALUTY_PLAT")
        kurs_waluty_plat.text = self.cdata_wrap("NBP")
        notowanie_waluty_ile_plat = parser.SubElement(
            platnosc, "NOTOWANIE_WALUTY_ILE_PLAT")
        notowanie_waluty_ile_plat.text = self.cdata_wrap(row._30)
        notowanie_waluty_za_ile_plat = parser.SubElement(
            delegation, "NOTOWANIE_WALUTY_ZA_ILE_PLAT")
        notowanie_waluty_za_ile_plat.text = self.cdata_wrap("1")
        kwota_pln_plat = parser.SubElement(platnosc, "KWOTA_PLN_PLAT")
        kwota_pln_plat.text = self.cdata_wrap(row._35)
        kierunek = parser.SubElement(platnosc, "KIERUNEK")
        kierunek.text = self.cdata_wrap("rozchód")
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
        data_kursu_plat.text = self.cdata_wrap(self.set_exchange_date(row._36))
        waluta_dok_plat = parser.SubElement(platnosc, "WALUTA_DOK_PLAT")
        waluta_dok_plat.text = self.cdata_wrap("EUR")
        platnosc_typ_podmiotu = parser.SubElement(
            platnosc, "PLATNOSC_TYP_PODMIOTU")
        platnosc_typ_podmiotu.text = self.cdata_wrap("pracownik")
        platnosc_podmiot = parser.SubElement(platnosc, "PLATNOSC_PODMIOT")
        platnosc_podmiot.text = self.cdata_wrap(row._0)
        platnosc_podmiot_id = parser.SubElement(
            delegation, "PLATNOSC_PODMIOT_ID")
        platnosc_podmiot_id.text = self.cdata_wrap("")
        platnosc_podmiot_nip = parser.SubElement(
            delegation, "PLATNOSC_PODMIOT_NIP")
        platnosc_podmiot_nip.text = self.cdata_wrap("")
        plat_kategoria = parser.SubElement(platnosc, "PLAT_KATEGORIA")
        plat_kategoria.text = self.cdata_wrap("DELEGACJE ZAGR. OP")
        plat_kategoria_id = parser.SubElement(platnosc, "PLAT_KATEGORIA_ID")
        plat_kategoria_id.text = self.cdata_wrap("")
        plat_elixir_01 = parser.SubElement(platnosc, "PLAT_ELIXIR_O1")
        plat_elixir_01.text = self.cdata_wrap(
            "Zaplata za {}{}".format(row._2, row._3))
        plat_elixir_02 = parser.SubElement(platnosc, "PLAT_ELIXIR_O2")
        plat_elixir_02.text = self.cdata_wrap("")
        plat_elixir_03 = parser.SubElement(platnosc, "PLAT_ELIXIR_O3")
        plat_elixir_03.text = self.cdata_wrap("")
        plat_elixir_04 = parser.SubElement(platnosc, "PLAT_ELIXIR_O4")
        plat_elixir_04.text = self.cdata_wrap("")

        return delegation

    def set_exchange_date(self, date: pd.Timestamp):
        """
        Set the date of eur to pln exchange for a delegation
        """
        exchange_date = date - \
            pd.Timedelta(days=3) if date.weekday == 0 else date - \
            pd.Timedelta(days=1)
        return exchange_date.strftime("%d.%m.%Y")
