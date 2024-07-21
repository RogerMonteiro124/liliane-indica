from flask import Flask, request, redirect, url_for, send_from_directory, render_template
import openpyxl
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
FILENAME = 'form_data.xlsx'

# Cria um novo arquivo Excel ou abre o existente
def get_workbook():
    if os.path.exists(FILENAME):
        workbook = openpyxl.load_workbook(FILENAME)
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Indicação", "Nome Completo", "Telefone", "Escola de Votação"])
    return workbook

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Recebe dados do formulário
    indicacao = request.form.get('indicacao')
    nome_completo = request.form.get('nome-completo')
    telefone = request.form.get('telefone')
    escola_votacao = request.form.get('escola-votacao')

    # Abre o arquivo Excel e adiciona uma nova linha
    workbook = get_workbook()
    sheet = workbook.active
    sheet.append([indicacao, nome_completo, telefone, escola_votacao])
    workbook.save(FILENAME)

    return redirect(url_for('obrigado', nome_completo=nome_completo))

@app.route('/obrigado/<nome_completo>')
def obrigado(nome_completo):
    return render_template('obrigado.html', nome_completo=nome_completo)

@app.route('/download')
def download():
    #return send_from_directory(directory='.', filename=FILENAME, as_attachment=True)
    return send_from_directory(directory='.', path=FILENAME, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=False)
