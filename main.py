import docx_document


def main():
    # from lxml import etree
    #
    # tree = etree.parse("lol.xml")  # Парсинг строки
    #
    # print(tree.xpath(".//item/text()"))


    docx = docx_document.DocxDocument("sample.docx");
    print(docx.get_n_pages());

main();
