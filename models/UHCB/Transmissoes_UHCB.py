from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from Conexao import obter_conexao
from Validacao import validar

class Trans_UHCB():
    def main(self, Planilha, data1, data2, mes, ano, dias):
        def obter_valores():
            # Lista com os codigos
            codigos = [1207, 1200, 1201, 1204, 1209]
            # Listas para adicionar os dados
            valores = []
            resultado = []
            # Abre conexao com o banco de dados
            cursor = obter_conexao().cursor()
            # Execucao da query para todos os codigos registrados
            consulta_sql = "SELECT c.codigo_sec, DAY(m.hora_transmissao), COUNT(m.hora_transmissao) \
                            FROM ( \
                                SELECT DISTINCT codigo_sec \
                                FROM mensagens \
                                WHERE Codigo_Sec IN (1207, 1200, 1201, 1204, 1209) \
                            ) c \
                            LEFT JOIN mensagens m \
                                ON c.codigo_sec = m.codigo_sec \
                                AND hora_transmissao >= %s AND hora_transmissao <= %s \
                                AND status_mensagem = 'G' \
                            GROUP BY c.codigo_sec, DAY(m.hora_transmissao) \
                            ORDER BY CASE c.codigo_sec \
                                WHEN 1207 THEN 1 \
                                WHEN 1200 THEN 2 \
                                WHEN 1201 THEN 3 \
                                WHEN 1204 THEN 4 \
                                WHEN 1209 THEN 5 \
                                END, DAY(m.hora_transmissao);"
            cursor.execute(consulta_sql, (data1, data2))
            # Extrai o valor da contagem dos dados de retorno
            c = 0  # Variavel para percorrer as estacoes
            dia = 1  # Variavel para contagem e comparação de dias
            for dados in cursor:
                if dados[0] == codigos[c]:  # Compara os codigos para alinhar os dados em listas
                    if dados[1] == dia:  # Compara o dia do dado com a variavel de comparacao
                        valores.append(dados[2])
                        dia += 1
                    else:
                        while (dia < dados[1]):  # Percorre o periodo ate o dia do dado
                            # Vai adicionando zero enquanto percorre o periodo
                            valores.append(0)
                            dia += 1
                        # Adiciona o valor do dado quando a variavel dia igualar com o dia do dado
                        valores.append(dados[2])
                        dia += 1
                else:
                    # Verifica se a quantidade de valores é menor que a quantidade de dias no mês
                    if len(valores) < dias:
                        # Adiciona 0 na lista até completar a quantidade de dias no mês
                        while (len(valores) < dias):
                            valores.append(0)
                    # Adiciona uma lista de dados diarios dentro de outra lista acumulativa
                    resultado.append(valores)
                    # Reiniciar valores para o novo código
                    dia = 1
                    # Verificação em caso de não houver nenhum dado de uma estação
                    if dados[1] is None:
                        # Inicia uma nova lista com a quantidade de dados encontrado no período
                        valores = [dados[2]]
                        # Adiciona 0 na lista até completar a quantidade de dias no mês
                        while (len(valores) < dias):
                            valores.append(0)
                        # Atualiza as variáveis para a próxima linha de dados
                        c += 1
                        dia = 1
                    else:
                        # Se o primeiro dia do proximo codigo for igual a variavel de comparacao
                        if dados[1] == dia:
                            # Inicia uma nova lista com o novo dado
                            valores = [dados[2]]
                            c += 1  # Incrementa a variavel para buscar o proximo codigo para comparacao dos dados
                            dia += 1  # Incrementa a variavel de comparacao de dias
                        else:
                            valores = [0]  # Inicia nova lista com zero
                            dia += 1  # Incrementa a variavel de comparacao de dias
                            # Percorre o periodo ate o dia do dado, caso o primeiro dia do codigo seja diferente da variavel de comparacao de dias
                            while (dia < dados[1]):
                                # Vai adicionando zero enquanto percorre o periodo
                                valores.append(0)
                                dia += 1
                            # Adiciona o valor do dado quando a variavel dia igualar com o dia do dado
                            valores.append(dados[2])
                            c += 1
                            dia += 1
            # Verifica se a quantidade de valores é menor que a quantidade de dias no mês
            if len(valores) < dias:
                # Adiciona 0 na lista até completar a quantidade de dias no mês
                while (len(valores) < dias):
                    valores.append(0)
            # Adiciona uma lista de dados diarios dentro de outra lista acumulativa
            resultado.append(valores)
            return resultado

        def registrar_valores(resultado, meteorologica):
            try:
                creds = validar()
                service = build("sheets", "v4", credentials=creds)
                if mes == 1:
                    # Atualizar nomes das abas
                    requests = []
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1509426409,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}12",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 561935798,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}11",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1372082395,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}10",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 761991810,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}09",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 2075093274,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}08",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1251683214,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}07",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1058572307,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}06",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1025584686,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}05",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 199622018,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}04",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 959470095,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}03",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1436683883,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}02",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1566820967,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}01",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 653130901,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Transmissoes {ano}",
                            },
                            "fields": "title",
                        }
                    })
                    body = {"requests": requests}
                    # Executa atualização dos nomes das abas
                    response = (
                        service.spreadsheets()
                        .batchUpdate(spreadsheetId=Planilha, body=body)
                        .execute())
                    Posicao_dados = f"Transmissoes {ano}01!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}01!B11"
                elif mes == 2:
                    Posicao_dados = f"Transmissoes {ano}02!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}02!B11"
                elif mes == 3:
                    Posicao_dados = f"Transmissoes {ano}03!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}03!B11"
                elif mes == 4:
                    Posicao_dados = f"Transmissoes {ano}04!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}04!B11"
                elif mes == 5:
                    Posicao_dados = f"Transmissoes {ano}05!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}05!B11"
                elif mes == 6:
                    Posicao_dados = f"Transmissoes {ano}06!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}06!B11"
                elif mes == 7:
                    Posicao_dados = f"Transmissoes {ano}07!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}07!B11"
                elif mes == 8:
                    Posicao_dados = f"Transmissoes {ano}08!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}08!B11"
                elif mes == 9:
                    Posicao_dados = f"Transmissoes {ano}09!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}09!B11"
                elif mes == 10:
                    Posicao_dados = f"Transmissoes {ano}10!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}10!B11"
                elif mes == 11:
                    Posicao_dados = f"Transmissoes {ano}11!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}11!B11"
                elif mes == 12:
                    Posicao_dados = f"Transmissoes {ano}12!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}12!B11"
                # Chamada da API de Planilhas
                sheet = service.spreadsheets()
                # Escrever dados em uma planilha
                result = (
                    sheet.values()
                    .update(spreadsheetId=Planilha, range=Posicao_dados, valueInputOption="USER_ENTERED", body={"values": resultado})
                    .execute())
                result = (
                    sheet.values()
                    .update(spreadsheetId=Planilha, range=Posicao_dados_meteorologica, valueInputOption="USER_ENTERED", body={"values": meteorologica})
                    .execute())
            except HttpError as err:
                print(err)
        # Função de busca no banco
        resultado = obter_valores()
        meteorologica = [resultado.pop()]
        # Função de registro na planilha
        registrar_valores(resultado, meteorologica)
