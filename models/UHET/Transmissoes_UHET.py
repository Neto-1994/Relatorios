from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from Conexao import obter_conexao
from Validacao import validar

class Trans_UHET():
    def main(self, Planilha, data1, data2, mes, ano):
        def obter_valores():
            # Lista com os codigos
            codigos = [813, 817, 812, 796, 808, 797, 814, 800, 799, 803,815, 811,
                       801, 806, 816, 818, 805, 809, 798, 807, 804, 802, 810]
            # Listas para adicionar os dados
            valores = []
            resultado = []
            # Abre conexao com o banco de dados
            cursor = obter_conexao().cursor()
            # Execucao da query para todos os codigos registrados
            consulta_sql = "SELECT codigo_sec, COUNT(hora_transmissao) FROM mensagens WHERE Codigo_Sec IN (813, 817, 812, 796, 808, 797, 814, 800, 799, 803, 815, 811, 801, 806, 816, 818, 805, 809, 798, 807, 804, 802, 810) \
                AND hora_transmissao >= %s AND hora_transmissao <= %s AND status_mensagem = 'G' \
                GROUP BY codigo_sec, DATE(hora_transmissao) \
                ORDER BY CASE codigo_sec \
                    WHEN 813 THEN 1 \
                    WHEN 817 THEN 2 \
                    WHEN 812 THEN 3 \
                    WHEN 796 THEN 4 \
                    WHEN 808 THEN 5 \
                    WHEN 797 THEN 6 \
                    WHEN 814 THEN 7 \
                    WHEN 800 THEN 8 \
                    WHEN 799 THEN 9 \
                    WHEN 803 THEN 10 \
                    WHEN 815 THEN 11 \
                    WHEN 811 THEN 12 \
                    WHEN 801 THEN 13 \
                    WHEN 806 THEN 14 \
                    WHEN 816 THEN 15 \
                    WHEN 818 THEN 16 \
                    WHEN 805 THEN 17 \
                    WHEN 809 THEN 18 \
                    WHEN 798 THEN 19 \
                    WHEN 807 THEN 20 \
                    WHEN 804 THEN 21 \
                    WHEN 802 THEN 22 \
                    WHEN 810 THEN 23 \
                    END, DATE(hora_transmissao);"
            cursor.execute(consulta_sql, (data1, data2))
            # Extrai o valor da contagem dos dados de retorno
            c = 0
            for dados in cursor:
                if dados[0] == codigos[c]:
                    valores.append(dados[1])
                else:
                    resultado.append(valores)
                    # Reiniciar valores para o novo código
                    valores = [dados[1]]
                    c += 1
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
                    Posicao_dados_meteorologica = f"Transmissoes {ano}01!B29"
                elif mes == 2:
                    Posicao_dados = f"Transmissoes {ano}02!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}02!B29"
                elif mes == 3:
                    Posicao_dados = f"Transmissoes {ano}03!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}03!B29"
                elif mes == 4:
                    Posicao_dados = f"Transmissoes {ano}04!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}04!B29"
                elif mes == 5:
                    Posicao_dados = f"Transmissoes {ano}05!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}05!B29"
                elif mes == 6:
                    Posicao_dados = f"Transmissoes {ano}06!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}06!B29"
                elif mes == 7:
                    Posicao_dados = f"Transmissoes {ano}07!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}07!B29"
                elif mes == 8:
                    Posicao_dados = f"Transmissoes {ano}08!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}08!B29"
                elif mes == 9:
                    Posicao_dados = f"Transmissoes {ano}09!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}09!B29"
                elif mes == 10:
                    Posicao_dados = f"Transmissoes {ano}10!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}10!B29"
                elif mes == 11:
                    Posicao_dados = f"Transmissoes {ano}11!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}11!B29"
                elif mes == 12:
                    Posicao_dados = f"Transmissoes {ano}12!B5"
                    Posicao_dados_meteorologica = f"Transmissoes {ano}12!B29"
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
        # Função de registro na tabela
        registrar_valores(resultado, meteorologica)
