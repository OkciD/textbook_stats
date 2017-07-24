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
        self.app_xml = etree.parse(self.app_xml_path);

    def extract_xmls(self):
        archive = zipfile.ZipFile(self.filename);
        archive.extract(self.document_xml_path);
        archive.extract(self.app_xml_path)
        archive.close();

    def clean_app_xml_root(self):
        app_xml_file = open(self.app_xml_path, "r");
        app_xml_file_contents = app_xml_file.read().replace('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">', "<Properties>");
        app_xml_file.close();
        app_xml_file = open(self.app_xml_path, "w");
        app_xml_file.write(app_xml_file_contents);
        app_xml_file.close();

    def __del__(self):
        path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "word");
        shutil.rmtree(path);
        path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "docProps");
        shutil.rmtree(path);