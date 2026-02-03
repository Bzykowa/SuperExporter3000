import xml.etree.ElementTree as parser
from xml.dom import minidom
from abc import ABC, abstractmethod


class XMLParser(ABC):
    """Abstract class with all methods needed for generating optima xmls"""

    def __init__(self, company_code: str, data_path: str):
        self.company_code = company_code
        self.data_path = data_path

        # set up the root of the document
        self.root = parser.Element("ROOT")
        self.root.set("xmlns", "http://www.comarch.pl/cdn/optima/offline")

    @abstractmethod
    def read_data(self):
        pass

    @abstractmethod
    def gen_xml_layout(self):
        pass

    def formatted_print(self):
        """
        Return a pretty-printed XML string for the Element.
        """
        rough_string = parser.tostring(self.root, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        final = reparsed.toprettyxml(indent="  ")
        final = final.replace("&lt;", "<")
        final = final.replace("&gt;", ">")
        return final

    def cdata_wrap(self, data):
        """Return data wrapped in cdata tag."""
        return "<![CDATA[{}]]>".format(data)

    def excel_text_number_to_float(self, x: str):
        """Change num from x xxx,xx to a float"""
        result = x.replace(" ", "")
        result = result.replace(",", ".")
        return float(result)
