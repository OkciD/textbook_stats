import docx_document


def main():
    filename = input("Input filename: ");
    docx = docx_document.DocxDocument(filename);
    print("Pages: ", docx.n_pages);
    print("Paragraphs: ", docx.n_paragraphs);
    print("Words: ", docx.n_words);
    print("Characters: ", docx.n_characters);
    print("Drawings: ", docx.n_drawings);
    print("Tables: ", docx.n_tables);

main();
