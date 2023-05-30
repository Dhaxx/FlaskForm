from flask import Flask, render_template, request, make_response, send_file
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from docx.shared import Inches
import subprocess

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/generate_document', methods=['POST'])
def generate_document():
    # Pega os dados do formulário
    subject = request.form.get('subject')
    colaborador = request.form.get('colaborador')
    descricao = request.form.get("descricao")

    # Cria o documento e escreve o mesmo
    document = Document()
    header = document.sections[0].header
    htable = header.add_table(1, 1, width=Inches(12))
    htab_cells = htable.rows[0].cells
    ht0 = htab_cells[0].add_paragraph()
    kh = ht0.add_run()
    kh.add_picture('./static/logo.jpg', width=Inches(6), height=Inches(2.0))

    # Adiciona o título de seção centralizado com tamanho de fonte 20
    title_paragraph = document.add_paragraph()
    title_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    title_run = title_paragraph.add_run("ATESTADO DE VISITA TÉCNICA")
    title_run.bold = True
    title_run.font.size = Pt(20)

    # Adiciona a seção de informações
    document.add_paragraph()
    info_paragraph = document.add_paragraph()
    info_paragraph.add_run("Atestamos, para os devidos fins, que a ")
    info_paragraph.add_run("Empresa Amendola & Amendola Software Ltda").bold = True
    info_paragraph.add_run(", inscrita no CNPJ nº 04.326.049/0001-90, realizou visita técnica nesta entidade, "
                        "conforme informações abaixo:")

    # Insere os campos dinâmicos no documento
    fields = []
    for key, value in request.form.items():
        if key.startswith('name'):
            index = key.replace('name', '')
            field = {'name': value, 'area': request.form.get('area' + index), 'cpf': request.form.get('cpf' + index)}
            fields.append(field)
    
    p_colaborador = document.add_paragraph()
    p_colaborador.add_run(f'Colaborador: ').bold = True
    p_colaborador.add_run(colaborador)

    document.add_paragraph(f'Assunto: {subject}')
    document.add_paragraph() 

    document.add_heading('Observações do Treinamento', level=1)
    document.add_paragraph(descricao)

    # Insere os campos dinâmicos no documento
    for field in fields:
        p_name = document.add_paragraph()
        p_name.add_run('Nome: ').bold = True
        p_name.add_run(field['name'])

        p_area = document.add_paragraph()
        p_area.add_run('cpf: ').bold = True
        p_area.add_run(field['cpf'])

        p_area = document.add_paragraph()
        p_area.add_run('Setor: ').bold = True
        p_area.add_run(field['area'])
        document.add_paragraph()  # Adicione uma nova linha em branco entre cada conjunto de campos

    # Salva o documento em um arquivo temporário
    temp_file = 'temp.docx'
    document.save(temp_file)

    response = make_response(send_file(temp_file, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'))
    response.headers.set('Content-Disposition', 'attachment', filename='document.docx')
    return response

    # # Abre o documento
    # subprocess.Popen(['start', temp_file], shell=True)

    # return 'Documento gerado com sucesso!'
if __name__ == '__main__':
    app.run(host="0.0.0.0")