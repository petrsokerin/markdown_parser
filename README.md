# Markdown parser

This Python utility converts Jupyter Notebooks (.ipynb) into Microsoft Word Documents (.docx). The main goal of the repo is spell and grama checking of markdown cells text in MS Word. 
## Quick Start

1. Clone the repository:

```
git clone https://github.com/petrsokerin/markdown_parser.git
cd markdown_parser
```

2. Install the required dependencies:

```
pip install -r requirements.txt
```

3. Convert a Jupyter Notebook to DOCX by running:

```
python markdown_parser.py --path path_to_your_notebook.ipynb
```

A new Word document docx_file.docx will be created in the current directory containing the markdown content from your Jupyter Notebook.
