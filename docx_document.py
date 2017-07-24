import zipfile
import os
import shutil
from lxml import etree


class DocxDocument:
    def __init__(self, filename):
        self.filename = filename;
        self.extract_xmls();
        self.document_xml = etree.ElementTree(file="word/document.xml");
        self.app_xml = etree.ElementTree(file="docProps/app.xml");

    def extract_xmls(self):
        archive = zipfile.ZipFile(self.filename);
        archive.extract("word/document.xml");
        archive.extract("docProps/app.xml")
        archive.close();

    def __del__(self):
        path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "word");
        shutil.rmtree(path);
        path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "docProps");
        shutil.rmtree(path);