import zipfile
import os
import shutil


def extract_document_xml(docx_filename):
    archive = zipfile.ZipFile(docx_filename);
    archive.extract("word/document.xml");
    archive.close();


def delete_document_xml():
    path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "word");
    shutil.rmtree(path)


def main():
    filename = input("Введите имя файла: ");
    extract_document_xml(filename);
    # delete_document_xml();


main();
