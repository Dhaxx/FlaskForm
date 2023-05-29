from flask import Flask, render_template, request
from docx import Document
from docx.shared import Inches
import subprocess

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/generate_document', methods=['POST'])
def generate_document():
    # Pega os dados do formulario
    name = request.form.get('name')
    area = request.form.get('area')
    subject = request.form.get('subject')

    # Cria o documento e escreve o mesmo
    document = Document()
    header = document.sections[0].header
    htable=header.add_table(1, 1, width=Inches(12))
    htab_cells=htable.rows[0].cells
    ht0=htab_cells[0].add_paragraph()
    kh=ht0.add_run()
    kh.add_picture('./static/logo.jpg', width=Inches(6), height=Inches(2.0))

    document.add_paragraph(f'Nome: {name}')
    document.add_paragraph(f'Setor: {area}')
    document.add_paragraph(f'Assunto: {subject}')

    # Save the document to a temporary file
    temp_file = 'temp.docx'
    document.save(temp_file)

    subprocess.Popen(['start', temp_file], shell=True)

    # Send the document as a download
    # return send_file(temp_file, attachment_filename='detalhes_funcionario.docx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)