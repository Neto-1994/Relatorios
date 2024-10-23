import mysql.connector
from decouple import config

# Função para criar uma conexão
def obter_conexao():
    try:
        host = config('HOST')
        port = config('PORT_DB')
        database = config('DATABASE')
        user = config('USER')
        password = config('PASSWORD')
        conexao = mysql.connector.connect(host=host, port=port, database=database, user=user, password=password)

        db_info = conexao.get_server_info()
        print("Conectado ao servidor MySQL versão: ", db_info)
        cursor = conexao.cursor()
        cursor.execute("select database();")
        banco = cursor.fetchone()
        print("Conectado ao banco de dados: ", banco)

        return conexao

    except Exception as e:
        print("Erro ao criar conexão com o banco: ", e)
        return None
