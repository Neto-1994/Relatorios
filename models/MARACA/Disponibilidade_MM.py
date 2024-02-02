import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from Conexao import obter_conexao


class Disp_MM():
    def main(self, Planilha, data1, data2, mes, ano):
        def obter_valores():
            # Listas para adicionar os dados
            valores = []
            # Abre conexao com o banco de dados
            cursor = obter_conexao().cursor()
            # Execucao da query para todos os codigos registrados
            consulta_sql = "SELECT COUNT(Dt_Medicao) FROM medicoes WHERE Codigo_Sec IN (1243, 1245, 1244, 1247, 1246, 1222, 1221, 1226, 1227, 1228, 1229, 1230, 1231, 1232) AND Dt_Medicao >= %s AND Dt_Medicao <= %s GROUP BY Codigo_Sec \
                ORDER BY CASE Codigo_Sec \
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
                    END;"
            cursor.execute(consulta_sql, (data1, data2))
            # Extrai o valor da contagem dos dados de retorno
            for dados in cursor:
                d = [dado for dado in dados]
                valores.append(d)
            return valores

        def registrar_valores(valores, meteorologica):
            # Parte de log da API Google Sheets
            # If modifying these scopes, delete the file token.json.
            SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
            creds = None
            # The file token.json stores the user's access and refresh tokens, and is
            # created automatically when the authorization flow completes for the first
            # time.
            if os.path.exists("E:/Relatorios/token.json"):
                creds = Credentials.from_authorized_user_file(
                    "E:/Relatorios/token.json", SCOPES)
            # If there are no (valid) credentials available, let the user log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        "E:/Relatorios/CredenciaisJair.json", SCOPES
                    )
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open("E:/Relatorios/token.json", "w") as token:
                    token.write(creds.to_json())

            try:
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
        meteorologica = valores[-2:]
        valores = valores[:-2]
        registrar_valores(valores, meteorologica)

