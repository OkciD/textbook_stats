import zipfile


def extract_document_xml(docx_filename):
    archive = zipfile.ZipFile(docx_filename);

    archive.close();


def main():
    filename = input("Введите имя файла: ");
    extract_document_xml(filename);
