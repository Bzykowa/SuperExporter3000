import xml.etree.ElementTree as parser
from xml.dom import minidom
from abc import ABC, abstractmethod


class XMLParser(ABC):
    """Abstract class with all methods needed for generating optima xmls"""

    def __init__(
        self, company_code: str, data_path: str, max_records: int = 0
    ):
        self.company_code = company_code
        self.data_path = data_path
        self.max_records = max_records

        # set up the root of the document
        self.root = parser.Element("ROOT")
        self.root.set("xmlns", "http://www.comarch.pl/cdn/optima/offline")
        # if needed can split root info smaller xmls
        self.split = []

    @abstractmethod
    def read_data(self):
        pass

    @abstractmethod
    def gen_xml_layout(self):
        pass

    def split_xml(self, max_records: int):
        """
        Split one big xml file into smaller chunks with maximum number of
        records in one file equal to max_records.

        :param max_records: a limit for children in a single file
        :type max_records: int
        """
        pass

    def formatted_print(self):
        """
        Return a pretty-printed XML string or a list of them.
        """
        if self.split:
            xmls = []
            for elem in self.split:
                rough_string = parser.tostring(elem, 'utf-8')
                reparsed = minidom.parseString(rough_string)
                final = reparsed.toprettyxml(indent="  ")
                final = final.replace("&lt;", "<")
                final = final.replace("&gt;", ">")
                xmls.append(final)
            if xmls:
                return xmls

        rough_string = parser.tostring(self.root, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        final = reparsed.toprettyxml(indent="  ")
        final = final.replace("&lt;", "<")
        final = final.replace("&gt;", ">")
        return final

    def cdata_wrap(self, data):
        """Return data wrapped in cdata tag."""
        return "<![CDATA[{}]]>".format(data)
