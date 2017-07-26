import zipfile
import os
import shutil
from lxml import etree


class DocxDocument:
    def __init__(self, filename):
        self.filename = filename;
        self.document_xml_path = "word/document.xml";
        self.app_xml_path = "docProps/app.xml";
        self.extract_xmls();
        self.clean_app_xml_root();
        self.document_xml = etree.parse(self.document_xml_path);
        self.document_xml_namespaces = self.document_xml.getroot().nsmap;
        self.app_xml = etree.parse(self.app_xml_path);
        self.app_xml_namespaces = self.app_xml.getroot().nsmap;
        self.get_app_xml_info();
        self.get_document_xml_info();

    def extract_xmls(self):
        archive = zipfile.ZipFile(self.filename);
        archive.extract(self.document_xml_path);
        archive.extract(self.app_xml_path)
        archive.close();

    def clean_app_xml_root(self):
        app_xml_file = open(self.app_xml_path, "r");
        app_xml_file_contents = app_xml_file.read().replace('xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ', "");
        app_xml_file.close();
        app_xml_file = open(self.app_xml_path, "w");
        app_xml_file.write(app_xml_file_contents);
        app_xml_file.close();

    def get_app_xml_info(self):
        self.n_pages = int(self.app_xml.xpath(".//Pages/text()", namespaces=self.app_xml_namespaces)[0]);
        self.n_paragraphs = int(self.app_xml.xpath(".//Paragraphs/text()", namespaces=self.app_xml_namespaces)[0]);
        self.n_words = int(self.app_xml.xpath(".//Words/text()", namespaces=self.app_xml_namespaces)[0]);
        self.n_characters = int(self.app_xml.xpath(".//Characters/text()", namespaces=self.app_xml_namespaces)[0]);

    def get_document_xml_info(self):
        self.n_drawings = len(self.document_xml.xpath(".//w:drawing", namespaces=self.document_xml_namespaces));
        self.n_tables = len(self.document_xml.xpath(".//w:tbl", namespaces=self.document_xml_namespaces));

    def __del__(self):
        path = os.path.join(os.path.abspath(os.path.dirname(__file__)), self.document_xml_path.split("/")[0]);
        shutil.rmtree(path);
        path = os.path.join(os.path.abspath(os.path.dirname(__file__)), self.app_xml_path.split("/")[0]);
        shutil.rmtree(path);