from flask import Flask
import pandas as pd
import nbformat
from nbconvert.preprocessors import ExecutePreprocessor
from nbconvert import HTMLExporter
import os
app = Flask(__name__)

@app.route('/')
def hello_world():
    return run_nb('test.ipynb', '?')

def run_nb(ju_nb, para):
    os.environ["JUPYTER_PARAMETER"] = para
    # the jupyter notebook can retrieve os.environ["JUPYTER_PARAMETER"]
    
    html_exporter = HTMLExporter()
    nb = nbformat.read(open(ju_nb), as_version=4)
    ep = ExecutePreprocessor(timeout=600, kernel_name='python3')
    ep.preprocess(nb, {'metadata': {'path': os.path.dirname(ju_nb)}})
    (body, resources) = html_exporter.from_notebook_node(nb)

    return body

if __name__ == '__main__':
   app.run()