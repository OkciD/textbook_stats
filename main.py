import docx_document
import pandas as pd


def print_dict(dict):
    for key in dict:
        print(key + ": " + str(dict[key]));


def print_table(dict):
    keys = dict.keys();
    table = pd.DataFrame(columns=keys);
    for key_i in dict:
        values_row = [];
        for key_j in dict:
            if (dict[key_j] <= 0):
                values_row.append(None);
            elif (dict[key_i] <= 0):
                values_row.append(0);
            else:
                values_row.append(dict[key_j] / dict[key_i]);
        row = pd.DataFrame([values_row], columns=keys, index=[key_i]);
        table = table.append(row);
    print("\nFirst row items are numerators, first column items are denominators.");
    print(table);


def main():
    filename = input("Input filename: ");
    docx = docx_document.DocxDocument(filename);
    values_dict = {"Pages": docx.n_pages, "Paragraphs": docx.n_paragraphs, "Words": docx.n_words,
                   "Characters": docx.n_characters, "Drawings": docx.n_drawings, "Tables": docx.n_tables};

    print_dict(values_dict);
    print_table(values_dict);
    input("\nPress any key to exit");

main();