import pandas as pd
import xml.etree.ElementTree as parser
from xml_parser import XMLParser


class Delegations(XMLParser):
    """The class responsible for parsing Excel data about delegations to xml
    file in Optima format."""

    def __init__(self, company_code: str, data_path: str):
        super().__init__(company_code, data_path)
        # read the data from submitted file
        self.read_data()

    def read_data(self):
        self.data = pd.read_excel(
            self.data_path, sheet_name="Do 30", header=0, decimal=",")

    def gen_xml_layout(self):
        """
        Generate the layout for delegation records.
        """
        self.records = parser.SubElement(self.root, "DOKUMENTY_INNE_ROZCHOD")
        self.records.set("xmlns", "")

        version = parser.SubElement(self.records, "WERSJA")
        version.text = self.cdata_wrap("2.00")
