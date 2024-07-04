from googleapiclient.discovery import build
from urllib.error import HTTPError
from Conexao import obter_conexao
from Validacao import validar

class Envios_UHJA():
    def main(self, Planilha, data1, data2, mes, ano):
        def obter_valores():
            # Listas para adicionar os dados
            valores = []
            # Abre conexao com o banco de dados
            cursor = obter_conexao().cursor()
            # Execucao da query para todos os codigos registrados
            consulta_sql = "SELECT COUNT(*) \
                            FROM ( \
                                SELECT DISTINCT codigo_sec \
                                FROM log_ANA \
                                WHERE codigo_sec IN (879) \
                            ) c \
                            LEFT JOIN log_ANA l \
                                ON c.codigo_sec = l.codigo_sec \
                                AND l.dt_medicao >= %s AND l.dt_medicao <= %s \
                                AND (TIMESTAMPDIFF(MINUTE, DATE_SUB(l.dt_medicao, INTERVAL 3 HOUR), l.dt_transmissao)) <= 180 \
                            GROUP BY c.codigo_sec;"
            cursor.execute(consulta_sql, (data1, data2))
            # Extrai o valor da contagem dos dados de retorno
            for dados in cursor:
                d = [dado for dado in dados]
                valores.append(d)
            return valores

        def registrar_valores(valores):
            try:
                creds = validar()
                service = build("sheets", "v4", credentials=creds)
                if mes == 1:
                    # Atualizar nomes das abas
                    requests = []
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 440522839,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}12",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 394485444,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}11",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 847405118,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}10",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 904372573,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}09",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1508716150,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}08",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 335583717,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}07",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1907808253,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}06",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1644918058,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}05",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1814106985,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}04",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 661690805,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}03",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 805434920,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}02",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1201086059,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}01",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 955315068,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Envio ANA {ano}",
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
                    Posicao_Escrever = f"Envio ANA {ano}01!B6"
                elif mes == 2:
                    Posicao_Escrever = f"Envio ANA {ano}02!B6"
                elif mes == 3:
                    Posicao_Escrever = f"Envio ANA {ano}03!B6"
                elif mes == 4:
                    Posicao_Escrever = f"Envio ANA {ano}04!B6"
                elif mes == 5:
                    Posicao_Escrever = f"Envio ANA {ano}05!B6"
                elif mes == 6:
                    Posicao_Escrever = f"Envio ANA {ano}06!B6"
                elif mes == 7:
                    Posicao_Escrever = f"Envio ANA {ano}07!B6"
                elif mes == 8:
                    Posicao_Escrever = f"Envio ANA {ano}08!B6"
                elif mes == 9:
                    Posicao_Escrever = f"Envio ANA {ano}09!B6"
                elif mes == 10:
                    Posicao_Escrever = f"Envio ANA {ano}10!B6"
                elif mes == 11:
                    Posicao_Escrever = f"Envio ANA {ano}11!B6"
                elif mes == 12:
                    Posicao_Escrever = f"Envio ANA {ano}12!B6"
                # Chamada da API de Planilhas
                sheet = service.spreadsheets()
                # Executa atualização dos dados na planilha
                result = (
                    sheet.values()
                    .update(spreadsheetId=Planilha, range=Posicao_Escrever, valueInputOption="USER_ENTERED", body={"values": valores})
                    .execute())
            except HTTPError as err:
                print(err)
        # Funções
        valores = obter_valores()
        registrar_valores(valores)
