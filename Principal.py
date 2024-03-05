from datetime import datetime, timedelta
import calendar
from Conexao import obter_conexao
from models.UHET.Transmissoes_UHET import Trans_UHET
from models.UHET.Disponibilidade_UHET import Disp_UHET
from models.UHET.EnviosANA_UHET import Envios_UHET
from models.SAE.Transmissoes_SAE import Trans_SAE
from models.SAE.Disponibilidade_SAE import Disp_SAE
from models.UHJA.Transmissoes_UHJA import Trans_UHJA
from models.UHJA.Disponibilidade_UHJA import Disp_UHJA
from models.UHJA.EnviosANA_UHJA import Envios_UHJA
from models.UHPP.Transmissoes_UHPP import Trans_UHPP
from models.UHPP.Disponibilidade_UHPP import Disp_UHPP
from models.UHPP.EnviosANA_UHPP import Envios_UHPP
from models.PHRO.Transmissoes_PHRO import Trans_PHRO
from models.PHRO.Disponibilidade_PHRO import Disp_PHRO
from models.PHRO.EnviosANA_PHRO import Envios_PHRO
from models.PHJG.Transmissoes_PHJG import Trans_PHJG
from models.PHJG.Disponibilidade_PHJG import Disp_PHJG
from models.PHJG.EnviosANA_PHJG import Envios_PHJG
from models.UHMI.Transmissoes_UHMI import Trans_UHMI
from models.UHMI.Disponibilidade_UHMI import Disp_UHMI
from models.UHMI.EnviosANA_UHMI import Envios_UHMI
from models.UHSA.Transmissoes_UHSA import Trans_UHSA
from models.UHSA.Disponibilidade_UHSA import Disp_UHSA
from models.UHSA.EnviosANA_UHSA import Envios_UHSA
from models.UHCB.Transmissoes_UHCB import Trans_UHCB
from models.UHCB.Disponibilidade_UHCB import Disp_UHCB
from models.UHCB.EnviosANA_UHCB import Envios_UHCB
from models.MARACA.Transmissoes_MM import Trans_MM
from models.MARACA.Disponibilidade_MM import Disp_MM


try:
    class Principal():
        hoje = datetime.today()
        mes = hoje.month
        ano = hoje.year
        if mes == 1:
            mes = 12
            ano -= 1
        else:
            mes -= 1
        # Formatando as datas
        dias = calendar.monthrange(ano, mes)[1]
        data1 = datetime(ano, mes, 1)
        data2 = datetime(ano, mes, dias) + timedelta(days=1, seconds=-1)
        # ID da planilha
        Planilha_UHET = "1AAyf62BhBSr1sVu86XRNwmgPNrs7AdEvu-0LYqqHD5I"
        Planilha_SAE = "17pHFrV60BWQrLcXoB5ljWxAJDI8pnsLBTUZDKwTWXhg"
        Planilha_UHJA = "13AXalT_IUrho_MXftddxPiMKSNVbq3iTGUAN_2jBi14"
        Planilha_UHPP = "1WwaYrc8la-Q83zyZirIyEJg3HPB7sNBrYekWHh3YJ2M"
        Planilha_PHRO = "1xWiaGYEnCN65yE9fW209UnohUCE4k_Nc82vVCmJSNOM"
        Planilha_PHJG = "10BIkqGcbYBLq-wVVFRKsZLjlShnm6KqzLWCfS5xzpmc"
        Planilha_UHMI = "1EPMz7Dh_o83bBu5m9eh9GaoTFCBXF1JZCpsIsD2FFLY"
        Planilha_UHSA = "1ng-JqjnlF0qQJ_slf7ufU-oY4q5Aln9ZX9rb1kbCgzs"
        Planilha_UHCB = "1VOKLFjn0XtZ-u3_TMjvgxSN3dDrYiCL3eKbbwgRrdXA"
        # "1ecdZkGcG5XfbkaW94fYk6shbJTk190IanLfQwTZSUNw"
        Planilha_MM = "1pOGcsXG_4aAV2Vt1gGgckkbT426T9Qk7YZLUwjqHaW0"

        # Funções
        def Disponibilidade(Planilha_UHET, Planilha_SAE, Planilha_UHJA, Planilha_UHPP, Planilha_PHRO, Planilha_PHJG, Planilha_UHMI, Planilha_UHSA, Planilha_UHCB, Planilha_MM, data1, data2, mes, ano):
            '''objeto = Disp_UHET()
            instancia = objeto.main(
                Planilha_UHET, data1, data2, mes, ano)
            objeto = Disp_SAE()
            instancia = objeto.main(
                Planilha_SAE, data1, data2, mes, ano)
            objeto = Disp_UHJA()
            instancia = objeto.main(
                Planilha_UHJA, data1, data2, mes, ano)
            objeto = Disp_UHPP()
            instancia = objeto.main(
                Planilha_UHPP, data1, data2, mes, ano)
            objeto = Disp_PHRO()
            instancia = objeto.main(
                Planilha_PHRO, data1, data2, mes, ano)
            objeto = Disp_PHJG()
            instancia = objeto.main(
                Planilha_PHJG, data1, data2, mes, ano)
            objeto = Disp_UHMI()
            instancia = objeto.main(
                Planilha_UHMI, data1, data2, mes, ano)
            objeto = Disp_UHSA()
            instancia = objeto.main(
                Planilha_UHSA, data1, data2, mes, ano)
            objeto = Disp_UHCB()
            instancia = objeto.main(
                Planilha_UHCB, data1, data2, mes, ano)'''
            objeto = Disp_MM()
            instancia = objeto.main(
                Planilha_MM, data1, data2, mes, ano)

        def Transmissoes(Planilha_UHET, Planilha_SAE, Planilha_UHJA, Planilha_UHPP, Planilha_PHRO, Planilha_PHJG, Planilha_UHMI, Planilha_UHSA, Planilha_UHCB, Planilha_MM, data1, data2, mes, ano, dias):
            '''objeto = Trans_UHET()
            instancia = objeto.main(Planilha_UHET, data1, data2, mes, ano, dias)
            objeto = Trans_SAE()
            instancia = objeto.main(Planilha_SAE, data1, data2, mes, ano, dias)
            objeto = Trans_UHJA()
            instancia = objeto.main(Planilha_UHJA, data1, data2, mes, ano, dias)
            objeto = Trans_UHPP()
            instancia = objeto.main(Planilha_UHPP, data1, data2, mes, ano, dias)
            objeto = Trans_PHRO()
            instancia = objeto.main(Planilha_PHRO, data1, data2, mes, ano, dias)
            objeto = Trans_PHJG()
            instancia = objeto.main(Planilha_PHJG, data1, data2, mes, ano, dias)
            objeto = Trans_UHMI()
            instancia = objeto.main(Planilha_UHMI, data1, data2, mes, ano, dias)
            objeto = Trans_UHSA()
            instancia = objeto.main(Planilha_UHSA, data1, data2, mes, ano, dias)
            objeto = Trans_UHCB()
            instancia = objeto.main(Planilha_UHCB, data1, data2, mes, ano, dias)'''
            objeto = Trans_MM()
            instancia = objeto.main(Planilha_MM, data1, data2, mes, ano, dias)

        def EnviosANA(Planilha_UHET, Planilha_UHJA, Planilha_UHPP, Planilha_PHRO, Planilha_PHJG, Planilha_UHMI, Planilha_UHSA, Planilha_UHCB, data1, data2, mes, ano):
            objeto = Envios_UHET()
            instancia = objeto.main(
                Planilha_UHET, data1, data2, mes, ano)
            objeto = Envios_UHJA()
            instancia = objeto.main(
                Planilha_UHJA, data1, data2, mes, ano)
            objeto = Envios_UHPP()
            instancia = objeto.main(
                Planilha_UHPP, data1, data2, mes, ano)
            objeto = Envios_PHRO()
            instancia = objeto.main(
                Planilha_PHRO, data1, data2, mes, ano)
            objeto = Envios_PHJG()
            instancia = objeto.main(
                Planilha_PHJG, data1, data2, mes, ano)
            objeto = Envios_UHMI()
            instancia = objeto.main(
                Planilha_UHMI, data1, data2, mes, ano)
            objeto = Envios_UHSA()
            instancia = objeto.main(
                Planilha_UHSA, data1, data2, mes, ano)
            objeto = Envios_UHCB()
            instancia = objeto.main(
                Planilha_UHCB, data1, data2, mes, ano)

        # Execuções
        Disponibilidade(Planilha_UHET, Planilha_SAE, Planilha_UHJA, Planilha_UHPP, Planilha_PHRO, Planilha_PHJG,
                        Planilha_UHMI, Planilha_UHSA, Planilha_UHCB, Planilha_MM, data1, data2, mes, ano)
        Transmissoes(Planilha_UHET, Planilha_SAE, Planilha_UHJA, Planilha_UHPP, Planilha_PHRO,
                     Planilha_PHJG, Planilha_UHMI, Planilha_UHSA, Planilha_UHCB, Planilha_MM, data1, data2, mes, ano, dias)
        '''EnviosANA(Planilha_UHET, Planilha_UHJA, Planilha_UHPP, Planilha_PHRO, Planilha_PHJG,
                  Planilha_UHMI, Planilha_UHSA, Planilha_UHCB, data1, data2, mes, ano)'''

        # Encerrar conexao com o banco de dados
        conexao = obter_conexao()
        conexao.close()

except OSError as e:
    print(f'Erro: ${e}')
