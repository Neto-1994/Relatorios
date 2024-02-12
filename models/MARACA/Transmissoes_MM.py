from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from Conexao import obter_conexao
from Validacao import validar

class Trans_MM():
    def main(self, Planilha, data1, data2, mes, ano):
        def obter_valores():
            # Lista com os codigos
            codigos = [1243, 1245, 1244, 1247, 1246, 1222,
                       1221, 1226, 1227, 1228, 1229, 1230, 1231, 1232]
            # Listas para adicionar os dados
            valores = []
            resultado = []
            # Abre conexao com o banco de dados
            cursor = obter_conexao().cursor()
            # Execucao da query para todos os codigos registrados
            consulta_sql = "SELECT codigo_sec, COUNT(hora_transmissao) FROM mensagens WHERE Codigo_Sec IN (1243, 1245, 1244, 1247, 1246, 1222, 1221, 1226, 1227, 1228, 1229, 1230, 1231, 1232) \
                AND hora_transmissao >= %s AND hora_transmissao <= %s AND status_mensagem = 'G' \
                GROUP BY codigo_sec, DATE(hora_transmissao) \
                ORDER BY CASE codigo_sec \
                    WHEN 1243 THEN 1 \
                    WHEN 1245 THEN 2 \
                    WHEN 1244 THEN 3 \
                    WHEN 1247 THEN 4 \
                    WHEN 1246 THEN 5 \
                    WHEN 1222 THEN 6 \
                    WHEN 1221 THEN 7 \
                    WHEN 1226 THEN 8 \
                    WHEN 1227 THEN 9 \
                    WHEN 1228 THEN 10 \
                    WHEN 1229 THEN 11 \
                    WHEN 1230 THEN 12 \
                    WHEN 1231 THEN 13 \
                    WHEN 1232 THEN 14 \
                    END, DATE(hora_transmissao);"
            cursor.execute(consulta_sql, (data1, data2))
            # Extrai o valor da contagem dos dados de retorno
            c = 0
            dia = 1
            for dados in cursor:
                if dados[0] == codigos[c]:
                    # if dados[1] == dia:
                    valores.append(dados[1])
                    # dia += 1
                    '''else:
                        while dia != dados[1]:
                            valores.append(0)
                            dia += 1
                        valores.append(dados[2])
                        dia += 1'''
                else:
                    '''# Preencher com zeros para os dias sem dados[2]
                    while dia != dados[1]:
                        valores.append(0)
                        dia += 1'''
                    resultado.append(valores)
                    # Reiniciar valores para o novo código
                    valores = [dados[1]]
                    c += 1

            resultado.append(valores)
            print(f'Resultado: {resultado}')
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
                    Posicao_dados_meteorologica = f"Transmissoes {ano}01!B19"
                elif mes == 2:
                    Posicao_dados = f"Transmissoes {ano}02!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}02!B19"
                elif mes == 3:
                    Posicao_dados = f"Transmissoes {ano}03!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}03!B19"
                elif mes == 4:
                    Posicao_dados = f"Transmissoes {ano}04!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}04!B19"
                elif mes == 5:
                    Posicao_dados = f"Transmissoes {ano}05!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}05!B19"
                elif mes == 6:
                    Posicao_dados = f"Transmissoes {ano}06!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}06!B19"
                elif mes == 7:
                    Posicao_dados = f"Transmissoes {ano}07!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}07!B19"
                elif mes == 8:
                    Posicao_dados = f"Transmissoes {ano}08!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}08!B19"
                elif mes == 9:
                    Posicao_dados = f"Transmissoes {ano}09!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}09!B19"
                elif mes == 10:
                    Posicao_dados = f"Transmissoes {ano}10!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}10!B19"
                elif mes == 11:
                    Posicao_dados = f"Transmissoes {ano}11!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}11!B19"
                elif mes == 12:
                    Posicao_dados = f"Transmissoes {ano}12!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}12!B19"
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
        meteorologica = resultado[-2:]
        resultado = resultado[:-2]
        # Função de registro na planilha
        registrar_valores(resultado, meteorologica)
