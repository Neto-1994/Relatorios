from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from Conexao import obter_conexao
from Validacao import validar

class Disp_UHET():
    def main(self, Planilha, data1, data2, mes, ano):
        def obter_valores():
            # Listas para adicionar os dados
            valores = []
            # Abre conexao com o banco de dados
            cursor = obter_conexao().cursor()
            # Execucao da query para todos os codigos registrados
            consulta_sql = "SELECT COUNT(Dt_Medicao) FROM medicoes WHERE Codigo_Sec IN (813, 817, 812, 796, 808, 797, 814, 800, 799, 803, 815, 811, 801, 806, 816, 818, 805, 809, 798, 807, 804, 802, 810) AND Dt_Medicao >= %s AND Dt_Medicao <= %s GROUP BY Codigo_Sec \
                ORDER BY CASE Codigo_Sec \
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
                    END;"
            cursor.execute(consulta_sql, (data1, data2))
            # Extrai o valor da contagem dos dados de retorno
            for dados in cursor:
                d = [dado for dado in dados]
                valores.append(d)
            return valores

        def registrar_valores(valores, meteorologica):
            try:
                creds = validar()
                service = build("sheets", "v4", credentials=creds)
                if mes == 1:
                    # Atualizar nomes das abas
                    requests = []
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1399899933,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}12",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 639326302,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}11",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 293280777,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}10",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 969146266,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}09",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 406943217,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}08",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 632803758,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}07",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 541754527,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}06",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 78174347,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}05",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 334261136,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}04",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 354835193,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}03",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1761630436,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}02",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 737873737,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}01",
                            },
                            "fields": "title",
                        }
                    })
                    requests.append({
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": 1607083518,  # Substitua pelo ID da aba
                                # Novo nome da aba
                                "title": f"Disponibilidade {ano}",
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
                    Posicao_Escrever = f"Disponibilidade {ano}01!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}01!G5"
                elif mes == 2:
                    Posicao_Escrever = f"Disponibilidade {ano}02!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}02!G5"
                elif mes == 3:
                    Posicao_Escrever = f"Disponibilidade {ano}03!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}03!G5"
                elif mes == 4:
                    Posicao_Escrever = f"Disponibilidade {ano}04!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}04!G5"
                elif mes == 5:
                    Posicao_Escrever = f"Disponibilidade {ano}05!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}05!G5"
                elif mes == 6:
                    Posicao_Escrever = f"Disponibilidade {ano}06!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}06!G5"
                elif mes == 7:
                    Posicao_Escrever = f"Disponibilidade {ano}07!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}07!G5"
                elif mes == 8:
                    Posicao_Escrever = f"Disponibilidade {ano}08!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}08!G5"
                elif mes == 9:
                    Posicao_Escrever = f"Disponibilidade {ano}09!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}09!G5"
                elif mes == 10:
                    Posicao_Escrever = f"Disponibilidade {ano}10!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}10!G5"
                elif mes == 11:
                    Posicao_Escrever = f"Disponibilidade {ano}11!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}11!G5"
                elif mes == 12:
                    Posicao_Escrever = f"Disponibilidade {ano}12!B5"
                    Posicao_dados_meteorologica = f"Disponibilidade {ano}12!G5"
                # Chamada da API de Planilhas
                sheet = service.spreadsheets()
                # Executa atualização dos dados na planilha
                result = (
                    sheet.values()
                    .update(spreadsheetId=Planilha, range=Posicao_Escrever, valueInputOption="USER_ENTERED", body={"values": valores})
                    .execute())
                result = (
                    sheet.values()
                    .update(spreadsheetId=Planilha, range=Posicao_dados_meteorologica, valueInputOption="USER_ENTERED", body={"values": meteorologica})
                    .execute())
            except HttpError as err:
                print(err)
        # Funções
        valores = obter_valores()
        meteorologica = [valores.pop()]
        registrar_valores(valores, meteorologica)
