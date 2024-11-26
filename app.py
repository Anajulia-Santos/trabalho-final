from flask import Flask, render_template, request, redirect, flash, send_file
from flask_mysqldb import MySQL
import openpyxl
from io import BytesIO

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Necessário para usar flash

# Configuração do banco de dados
app.config['MYSQL_HOST'] = 'localhost'
app.config['MYSQL_USER'] = 'root'
app.config['MYSQL_PASSWORD'] = ''
app.config['MYSQL_DB'] = 'MapeamentoCultural'

mysql = MySQL(app)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit-data', methods=['POST'])
def submit_data():
    nome = request.form['nome']
    email = request.form['email']
    municipio = request.form['municipio']
    manifestacao = request.form['manifestacao']

    # Inserir dados no banco de dados
    cursor = mysql.connection.cursor()
    cursor.execute("""
        INSERT INTO Usuarios (nome, email, municipio) VALUES (%s, %s, %s)
    """, (nome, email, municipio))
    usuario_id = cursor.lastrowid  # Pega o ID do usuário recém-criado

    # Inserir resposta no banco
    cursor.execute("""
        INSERT INTO Respostas (id_usuario, manifestacao) VALUES (%s, %s)
    """, (usuario_id, manifestacao))
    resposta_id = cursor.lastrowid  # Pega o ID da resposta
    
    # Dados específicos conforme a manifestação
    if manifestacao == "Musica":
        nome_artista = request.form.get("nome_artista")
        genero = request.form.get("genero")
        local_apresentacao = request.form.get("local_apresentacao")
        sobre = request.form.get("sobre")
        cursor.execute("""
            INSERT INTO Musica (id_resposta, nome_artista, genero, local_apresentacao, sobre)
            VALUES (%s, %s, %s, %s, %s)
        """, (resposta_id, nome_artista, genero, local_apresentacao, sobre))

    elif manifestacao == "Danca":
        estilo = request.form.get("estilo")
        grupo_danca = request.form.get("grupo_danca")
        local_apresentacao_danca = request.form.get("local_apresentacao_danca")
        sobre_danca = request.form.get("sobre_danca")
        cursor.execute("""
            INSERT INTO Danca (id_resposta, estilo, grupo_danca, local_apresentacao, sobre)
            VALUES (%s, %s, %s, %s, %s)
        """, (resposta_id, estilo, grupo_danca, local_apresentacao_danca, sobre_danca))

    elif manifestacao == "Artesanato":
        tipo_material = request.form.get("tipo_material")
        tecnica = request.form.get("tecnica")
        origem_historica = request.form.get("origem_historica")
        cursor.execute("""
            INSERT INTO Artesanato (id_resposta, tipo_material, tecnica, origem_historica)
            VALUES (%s, %s, %s, %s)
        """, (resposta_id, tipo_material, tecnica, origem_historica))

    elif manifestacao == "Gastronomia":
        prato_tipico = request.form.get("prato_tipico")
        ingredientes = request.form.get("ingredientes")
        historia_prato = request.form.get("historia_prato")
        cursor.execute("""
            INSERT INTO Gastronomia (id_resposta, prato_tipico, ingredientes, historia_prato)
            VALUES (%s, %s, %s, %s)
        """, (resposta_id, prato_tipico, ingredientes, historia_prato))

    elif manifestacao == "Literatura":
        autor = request.form.get("autor")
        obra = request.form.get("obra")
        tipo_oralidade = request.form.get("tipo_oralidade")
        sinopse = request.form.get("sinopse")
        cursor.execute("""
            INSERT INTO LiteraturaOralidade (id_resposta, autor, obra, tipo_oralidade, sinopse)
            VALUES (%s, %s, %s, %s, %s)
        """, (resposta_id, autor, obra, tipo_oralidade, sinopse))

    elif manifestacao == "Religiosidade":
        rito = request.form.get("rito")
        data_celebracao = request.form.get("data_celebracao")
        local_celebracao = request.form.get("local_celebracao")
        sobre_religiosidade = request.form.get("sobre_religiosidade")
        cursor.execute("""
            INSERT INTO Religiosidade (id_resposta, rito, data_celebracao, local_celebracao, sobre)
            VALUES (%s, %s, %s, %s, %s)
        """, (resposta_id, rito, data_celebracao, local_celebracao, sobre_religiosidade))

    elif manifestacao == "Patrimonio":
        nome_patrimonio = request.form.get("nome_patrimonio")
        localizacao_patrimonio = request.form.get("localizacao_patrimonio")
        data_criacao = request.form.get("data_criacao")
        importancia_cultural = request.form.get("importancia_cultural")
        cursor.execute("""
            INSERT INTO Patrimonio (id_resposta, nome, tipo, localizacao, data_criacao, importancia_cultural)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (resposta_id, nome_patrimonio, 'Patrimônio', localizacao_patrimonio, data_criacao, importancia_cultural))

    elif manifestacao == "Eventos":
        nome_evento = request.form.get("nome_evento")
        data_inicio_evento = request.form.get("data_inicio_evento")
        data_fim_evento = request.form.get("data_fim_evento")
        local_evento = request.form.get("local_evento")
        cursor.execute("""
            INSERT INTO Eventos (id_resposta, nome_evento, tipo_evento, data_inicio, data_fim, local_evento)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (resposta_id, nome_evento, 'Evento Cultural', data_inicio_evento, data_fim_evento, local_evento))

    # Confirmar transação e fechar o cursor
    mysql.connection.commit()
    cursor.close()

    flash('Informações enviadas com sucesso!', 'success')

    return redirect('/')

@app.route('/relatorios')
def relatorios():
    # Consultar dados para relatórios
    cursor = mysql.connection.cursor()
    cursor.execute("""
        SELECT Usuarios.nome, Usuarios.email, Usuarios.municipio, Respostas.manifestacao,
               Musica.nome_artista, Musica.genero, Musica.local_apresentacao, Musica.sobre,
               Danca.estilo, Danca.grupo_danca, Danca.local_apresentacao, Danca.sobre,
               Artesanato.tipo_material, Artesanato.tecnica, Artesanato.origem_historica,
               Gastronomia.prato_tipico, Gastronomia.ingredientes, Gastronomia.historia_prato,
               LiteraturaOralidade.autor, LiteraturaOralidade.obra, LiteraturaOralidade.tipo_oralidade, LiteraturaOralidade.sinopse,
               Religiosidade.rito, Religiosidade.data_celebracao, Religiosidade.local_celebracao, Religiosidade.sobre,
               Patrimonio.nome, Patrimonio.localizacao, Patrimonio.data_criacao, Patrimonio.importancia_cultural,
               Eventos.nome_evento, Eventos.data_inicio, Eventos.data_fim, Eventos.local_evento
        FROM Respostas
        JOIN Usuarios ON Respostas.id_usuario = Usuarios.id_usuario
        LEFT JOIN Musica ON Respostas.id_resposta = Musica.id_resposta
        LEFT JOIN Danca ON Respostas.id_resposta = Danca.id_resposta
        LEFT JOIN Artesanato ON Respostas.id_resposta = Artesanato.id_resposta
        LEFT JOIN Gastronomia ON Respostas.id_resposta = Gastronomia.id_resposta
        LEFT JOIN LiteraturaOralidade ON Respostas.id_resposta = LiteraturaOralidade.id_resposta
        LEFT JOIN Religiosidade ON Respostas.id_resposta = Religiosidade.id_resposta
        LEFT JOIN Patrimonio ON Respostas.id_resposta = Patrimonio.id_resposta
        LEFT JOIN Eventos ON Respostas.id_resposta = Eventos.id_resposta
    """)
    dados = cursor.fetchall()
    cursor.close()

    return render_template('relatorios.html', dados=dados)

@app.route('/relatorios/download')
def download_relatorio():
    # Consultar dados para o relatório completo
    cursor = mysql.connection.cursor()
    cursor.execute("""
        SELECT Usuarios.nome, Usuarios.email, Usuarios.municipio, Respostas.manifestacao,
               Musica.nome_artista, Musica.genero, Musica.local_apresentacao, Musica.sobre,
               Danca.estilo, Danca.grupo_danca, Danca.local_apresentacao, Danca.sobre,
               Artesanato.tipo_material, Artesanato.tecnica, Artesanato.origem_historica,
               Gastronomia.prato_tipico, Gastronomia.ingredientes, Gastronomia.historia_prato,
               LiteraturaOralidade.autor, LiteraturaOralidade.obra, LiteraturaOralidade.tipo_oralidade, LiteraturaOralidade.sinopse,
               Religiosidade.rito, Religiosidade.data_celebracao, Religiosidade.local_celebracao, Religiosidade.sobre,
               Patrimonio.nome, Patrimonio.localizacao, Patrimonio.data_criacao, Patrimonio.importancia_cultural,
               Eventos.nome_evento, Eventos.data_inicio, Eventos.data_fim, Eventos.local_evento
        FROM Respostas
        JOIN Usuarios ON Respostas.id_usuario = Usuarios.id_usuario
        LEFT JOIN Musica ON Respostas.id_resposta = Musica.id_resposta
        LEFT JOIN Danca ON Respostas.id_resposta = Danca.id_resposta
        LEFT JOIN Artesanato ON Respostas.id_resposta = Artesanato.id_resposta
        LEFT JOIN Gastronomia ON Respostas.id_resposta = Gastronomia.id_resposta
        LEFT JOIN LiteraturaOralidade ON Respostas.id_resposta = LiteraturaOralidade.id_resposta
        LEFT JOIN Religiosidade ON Respostas.id_resposta = Religiosidade.id_resposta
        LEFT JOIN Patrimonio ON Respostas.id_resposta = Patrimonio.id_resposta
        LEFT JOIN Eventos ON Respostas.id_resposta = Eventos.id_resposta
    """)
    dados = cursor.fetchall()
    cursor.close()

    # Gerar o Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatório Cultural"

    # Cabeçalhos
    headers = ['Nome', 'Email', 'Município', 'Manifestação', 'Artista', 'Gênero', 'Local Apresentação', 'Sobre',
               'Estilo', 'Grupo de Dança', 'Local Apresentação Dança', 'Sobre Dança', 'Material', 'Técnica', 'Origem Histórica',
               'Prato Típico', 'Ingredientes', 'História do Prato', 'Autor', 'Obra', 'Tipo de Oralidade', 'Sinopse',
               'Rito', 'Data Celebração', 'Local Celebração', 'Sobre Religiosidade', 'Nome Patrimônio', 'Localização',
               'Data de Criação', 'Importância Cultural', 'Nome Evento', 'Data Início', 'Data Fim', 'Local Evento']
    ws.append(headers)

    for row in dados:
        ws.append(row)

    # Salvar Excel em memória
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # Enviar arquivo Excel para o usuário
    return send_file(output, as_attachment=True, download_name="relatorio_cultural.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    app.run(debug=True)
