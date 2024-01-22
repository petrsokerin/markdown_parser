import json
import argparse
from docx import Document
from typing import Dict, List

def read_notebook(path: str) -> Dict:
    with open(path, encoding='utf8') as f:
        file = json.load(f)
    return file

def parse_notebook(file: Dict) -> Document:
    document = Document()
    for cell in file['cells']:
        if cell['cell_type'] == 'markdown':
            paragraph = parse_cell(cell['source'])
            document.add_paragraph(paragraph)
    return document

def parse_cell(lines: List) -> str:
    paragraph = ''
    for line in lines:
        if "![" not in line: # drop pictures
            paragraph += line
    return paragraph

def main():
    parser = argparse.ArgumentParser(description='Process command line arguments.')
    parser.add_argument('--path', type=str, help='notebook path for parsing')
    args = parser.parse_args()
    path = args.path

    file = read_notebook(path)
    document = parse_notebook(file)
    document.save('docx_file.docx')

if __name__ == "__main__":
    main()

