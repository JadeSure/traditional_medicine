from docx import Document
from pathlib import Path
import pandas as pd

BASE_PATH = Path(__file__).resolve().parent


def replace_string_in_docx(docx_filename, old_string, new_string, output_path):
    doc = Document(docx_filename)
    file_name = str(docx_filename).split("/")[-1]
    for p in doc.paragraphs:
        for run in p.runs:
            if old_string in run.text:
                run.text = run.text.replace(old_string, new_string)
    doc.save(f'{output_path}_{file_name}')


def read_names(file_path):
    df = pd.read_excel(str(file_path))
    output = {}
    for row in df.itertuples():
        output[row.name] = row.email
    return output


if __name__ == "__main__":
    source_cert = BASE_PATH / 'source' / '2024CPD证书.docx'
    source_people = BASE_PATH / 'source' / 'names.xlsx'

    print(BASE_PATH)
    people = read_names(source_people)
    for name in people.keys():
        print(type(name))
        replace_string_in_docx(source_cert, 'LEI    yao', name, f'{BASE_PATH}/res/{name}')
