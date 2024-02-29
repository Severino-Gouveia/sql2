import pyodbc
import pandas as pd
from datetime import date, timedelta
from flask import Flask, render_template, send_from_directory, request, session

app = Flask(__name__)
app.secret_key = 'chave' 

# Configurações de conexão com o banco de dados
SERVIDOR = '172.20.20.38'
BANCO_DE_DADOS = 'EAMPROD'
USUARIO = 'excel'
SENHA = 'excel'
NOME_DO_DRIVER = 'ODBC Driver 17 for SQL Server'

# String de conexão formatada
string_de_conexao = f'DRIVER={{{NOME_DO_DRIVER}}};SERVER={SERVIDOR};DATABASE={BANCO_DE_DADOS};UID={USUARIO};PWD={SENHA}'

# Funções de consulta ao banco de dados

def consultar_banco(consulta_sql, *params):
    try:
        with pyodbc.connect(string_de_conexao) as conexao:
            with conexao.cursor() as cursor:
                cursor.execute(consulta_sql, *params)
                registros = cursor.fetchall()
                df = pd.DataFrame.from_records(registros, columns=[desc[0] for desc in cursor.description])

                if 'DATA_TEXTO' in df.columns:
                    # Se a coluna 'DATA_TEXTO' estiver presente, processe as colunas de data
                    df['Hora'] = df['DATA_TEXTO'].dt.strftime('%H:%M:%S')
                    df['Semana'] = df['DATA_TEXTO'].dt.strftime('%U')
                    df['Data'] = df['DATA_TEXTO'].dt.strftime('%d/%m/%Y')

                return df

    except pyodbc.Error as erro:
        print(f"Erro: {erro}")
        return None


def construir_consulta_sql(data_inicial, data_final, departamento):
    return '''
        SELECT 
            R5ADDETAILS.ADD_CREATED as DATA_TEXTO, 
            R5ADDETAILS.ADD_CODE as AES,
            R5EVENTS.EVT_DESC as Descrição, 
            CASE R5EVENTS.EVT_STATUS
                WHEN 'FE' THEN 'Fechado'
                WHEN 'RP' THEN 'Reprogramar'
                WHEN 'CM' THEN 'Aguardando Chegada Material'
                WHEN 'P' THEN 'Programada'
                WHEN 'REJ' THEN 'Rejeitada'
                WHEN 'AP' THEN 'Aprovada'
                WHEN 'PM' THEN 'Aguardar Parada de Maquina'
                WHEN 'R' THEN 'Emitido'
                WHEN 'CL' THEN 'Concluida'
                WHEN 'EE' THEN 'Em Execucao'
                ELSE 'OUTROS'
            END as Status,
            CASE R5ADDETAILS.ADD_TYPE
                WHEN '*' THEN 'Sim'
                ELSE 'Nao'
            END as Observacao,
            R5EVENTS.EVT_STATUS as EVT_STAT, 
            R5ADDETAILS.ADD_LINE as Linha,
            R5ADDETAILS.ADD_TEXT as Observacões, 
            R5ADDETAILS.ADD_USER as matricula,
            R5PERSONNEL.PER_DESC as Colaborador, 
            R5CREWS.CRW_DESC AS turno,
            R5EVENTS.EVT_MRC AS Departamento
        FROM 
            EAMPROD.dbo.R5ADDETAILS R5ADDETAILS
            INNER JOIN R5EVENTS ON R5EVENTS.EVT_CODE = R5ADDETAILS.ADD_CODE
            INNER JOIN R5PERSONNEL ON R5PERSONNEL.PER_CODE = R5ADDETAILS.ADD_USER
            INNER JOIN R5CREWEMPLOYEES ON R5CREWEMPLOYEES.CRE_PERSON = R5ADDETAILS.ADD_USER
            INNER JOIN R5CREWS ON R5CREWEMPLOYEES.CRE_CREW = R5CREWS.CRW_CODE
        WHERE 
            R5ADDETAILS.ADD_ENTITY = 'EVNT' 
            AND (R5ADDETAILS.ADD_CREATED >= ?) 
            AND (R5ADDETAILS.ADD_CREATED < ?) 
            AND (R5EVENTS.EVT_MRC = ?);
    '''

def consultar_banco2(departamento, data_inicio, data_fim, responsavel):
    # Conecta ao banco de dados
    conexao = pyodbc.connect(string_de_conexao)
    # Define a consulta SQL parametrizada
    consulta_sql = """
    SELECT
        EVT_CODE AS 'OS',
        EVT_DESC AS 'DESCRIÇÃO',
        CASE EVT_STATUS 
            WHEN 'FE' THEN 'Fechado'
            WHEN 'RP' THEN 'Reprogramar'
            WHEN 'CM' THEN 'Aguardando Chegada Material'
            WHEN 'P'  THEN 'Programada'
            WHEN 'REJ' THEN 'Rejeitada'
            WHEN 'AP' THEN 'Aprovada'
            WHEN 'PM' THEN 'Aguardar Parada de Máquina'
            WHEN 'R'  THEN 'Emitido'
            WHEN 'CL' THEN 'Concluída'
            WHEN 'EE' THEN 'Em Execução'
            WHEN 'AREL' THEN 'Aguardando relatório'
            WHEN 'Q' THEN 'Solicitação de serviço'
            ELSE 'Outros'
        END AS 'STATUS',
        CASE EVT_PRIORITY 
            WHEN '' THEN 'CRÍTICA - DO'
            WHEN '1' THEN 'EMERGÊNCIA D+1'
            WHEN '2' THEN 'URGENCIA D+7'
            WHEN '3'  THEN 'URGENCIA PROG. D+15'
            WHEN '4' THEN 'PROGRAMADA D+30'
            WHEN '5' THEN 'PLANEJADA'
            ELSE 'OUTROS'
        END AS 'PRIORIDADE',
        EVT_CLASS  AS 'CLASSE',
        EVT_MRC AS 'DEPARTAMENTO',
        CONVERT(varchar, EVT_CREATED, 103) AS 'DATA DE CRIAÇÃO',
        CONVERT(varchar, EVT_UDFDATE04, 103) AS 'DATA PROGRAMADA',
        CASE ACS_RESPONSIBLE
            WHEN '3501040' THEN 'LEANDRO MAYRON'
            WHEN '5600023' THEN 'EDILBERTO GOIS'
            WHEN '3501061' THEN 'JANSEN JAQUES'
            WHEN '3004928'  THEN 'ALEX DE SOUSA'
            WHEN '3004933' THEN 'JACKSON FERNANDES'
            WHEN '5600045' THEN 'ANTÔNIO MARTINS'
            WHEN '3501125' THEN 'SIDINEY RODRIGUES VIEIRA'
            WHEN '3501060'  THEN 'SEVERINO GOUVEIA'
            WHEN '3006637' THEN 'LEONARDO FERNANDES'
            WHEN '3501017' THEN 'PABLO XAVIER'
            WHEN '3501123' THEN 'FERNANDA'
            WHEN '3007314' THEN 'RÔMULO HENRIQUE'
            WHEN '3501063' THEN 'EDUARDO ALMEIDA'
            WHEN '3004941' THEN 'ROSINILDO GOMES'
            WHEN '3007312' THEN 'ALCIKLEBER PEREIRA'
            ELSE 'Outros'
        END AS 'EXECUTANTE',
        CASE ACS_RESPONSIBLE
            WHEN '3501040' THEN 'COMERCIAL'
            WHEN '5600023' THEN 'COMERCIAL'
            WHEN '3501061' THEN 'COMERCIAL'
            WHEN '3004928'  THEN 'COMERCIAL'
            WHEN '3004933' THEN 'COMERCIAL'
            WHEN '5600045' THEN 'COMERCIAL'
            WHEN '3501125' THEN 'COMERCIAL'
            WHEN '3501060'  THEN 'COMERCIAL'
            WHEN '3006637' THEN 'COMERCIAL'
            WHEN '3501017' THEN 'COMERCIAL'
            WHEN '3501123' THEN 'COMERCIAL'
            WHEN '3007314' THEN 'TARDE'
            WHEN '3501063' THEN 'MANHÃ'
            WHEN '3004941' THEN 'MANHÃ'
            WHEN '3007312' THEN 'TARDE'
            ELSE 'Outros'
        END AS 'TURNO'
    FROM
        R5EVENTS
    INNER JOIN
        R5ACTSCHEDULES ON R5EVENTS.EVT_CODE = R5ACTSCHEDULES.ACS_EVENT
    WHERE 
        EVT_MRC = ?
        AND ACS_SCHED >= ?
        AND ACS_SCHED <= ?
        AND ACS_RESPONSIBLE = ?
    """
    # Executa a consulta SQL parametrizada
    df = pd.read_sql_query(consulta_sql, conexao, params=(departamento, data_inicio, data_fim, responsavel))
    # Fecha a conexão com o banco de dados
    conexao.close()
    return df

# Rota para servir arquivos estáticos, incluindo imagens
@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory('static', filename)

# Rota para renderizar a página HTML com o filtro por departamento
@app.route('/filtrar/<departamento>')
def filtrar_departamento(departamento):
    # Armazena o departamento na sessão
    session['departamento'] = departamento

    data_inicial = date.today() - timedelta(days=1)
    data_final = date.today()

    consulta_sql = construir_consulta_sql(data_inicial, data_final, departamento)

    df = consultar_banco(consulta_sql, data_inicial, data_final, departamento)

    if df is not None:
        # Select only the desired columns for the RDM
        relatorio_estilizado = df[['AES', 'Descrição', 'Colaborador', 'Observacões', 'Hora', 'Status', 'Semana', 'Data', 'Departamento']]
        imagem_path = 'img/LOGO_EPASA.png'
        return render_template('index.html', data=relatorio_estilizado.to_dict('records'), imagem_path=imagem_path, departamento=departamento)
    else:
        return "Erro na consulta do banco de dados."

# Rota para renderizar a página HTML com filtro por data
@app.route('/filtrar/data', methods=['POST'])
def filtrar_data():
    data_inicial = request.form['data_inicial']
    data_final = request.form['data_final']
    departamento = request.form.get('departamento', default=None)

    # Se o departamento não foi especificado, defina um valor padrão
    if departamento is None:
        # Verifica se o departamento está armazenado na sessão
        departamento = session.get('departamento', 'DP02')

    # Converta as datas para o formato desejado
    data_inicial = pd.to_datetime(data_inicial).date()
    data_final = pd.to_datetime(data_final).date() + timedelta(days=1)

    consulta_sql = construir_consulta_sql(data_inicial, data_final, departamento)

    df = consultar_banco(consulta_sql, data_inicial, data_final, departamento)

    if df is not None:
        # Restante do código permanece inalterado
        relatorio_estilizado = df[['AES', 'Descrição', 'Colaborador', 'Observacões', 'Hora', 'Status', 'Semana', 'Data', 'Departamento']]
        imagem_path = 'img/LOGO_EPASA.png'
        return render_template('index.html', data=relatorio_estilizado.to_dict('records'), imagem_path=imagem_path, departamento=departamento)
    else:
        return "Erro na consulta do banco de dados."

@app.route('/prog', methods=['GET', 'POST'])
def prog():
    if request.method == 'POST':
        departamento = request.form['departamento']
        data_inicio = request.form['data_inicio']
        data_fim = request.form['data_fim']
        responsavel = request.form['responsavel']
        df = consultar_banco2(departamento, data_inicio, data_fim, responsavel)
        return render_template('prog.html', tables=[df.to_html(classes='data')], titles=df.columns.values)
    return render_template('prog.html')

@app.route('/')
def index():
    return filtrar_departamento('DP01')  # Defina o departamento padrão aqui

if __name__ == '__main__':
    app.run(debug=True)


    
