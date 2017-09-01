import docx_document
import pandas as pd
import numpy as np


def print_dict(dict):
    for key in dict:
        print(key + ": " + str(dict[key]));


def print_table(dict):
    keys = dict.keys();
    table = pd.DataFrame(columns=keys);
    for key_i in dict:
        values_row = [dict[key_j] / dict[key_i] for key_j in dict];
        row = pd.DataFrame([values_row], columns=keys, index=[key_i]);
        table = table.append(row);
    print(table);


def main():
    # filename = input("Input filename: ");
    filename = "sample.docx";
    docx = docx_document.DocxDocument(filename);
    values_dict = {"Pages": docx.n_pages, "Paragraphs": docx.n_paragraphs, "Words": docx.n_words,
                   "Characters": docx.n_characters, "Drawings": docx.n_drawings, "Tables": docx.n_tables};

    print_dict(values_dict);
    print_table(values_dict);

main();