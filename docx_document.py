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
        self.document_xml = etree.parse(self.document_xml_path);
        self.app_xml = etree.parse(self.app_xml_path);

    def extract_xmls(self):
        archive = zipfile.ZipFile(self.filename);
        archive.extract(self.document_xml_path);
        archive.extract(self.app_xml_path)
        archive.close();

    def __del__(self):
        path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "word");
        shutil.rmtree(path);
        path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "docProps");
        shutil.rmtree(path);