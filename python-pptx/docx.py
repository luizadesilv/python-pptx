from docx2python import docx2python
import pandas as pd


class File:

    @staticmethod
    def read_docx(path):
        doc = docx2python(path)
        output = doc.text

        return output

    @staticmethod
    def split_text(output):
        text = output.split("->")

        return text

    @staticmethod
    def data_to_table():
        file = pd.read_csv(r'data/table-data.csv')
        table = pd.DataFrame(file)
        print(table)
        return table
