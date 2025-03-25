import os
import socket
import getpass
import platform
import psutil
from openpyxl import Workbook, load_workbook

def coletar_informacoes():
    informacoes = {}
    informacoes['Nome do Computador'] = os.getenv('COMPUTERNAME')
    informacoes['Endereço IP'] = socket.gethostbyname(socket.gethostname())
    informacoes['Nome do Usuário'] = getpass.getuser()
    informacoes['Sistema Operacional'] = platform.system()
    informacoes['Versão do SO'] = platform.version()
    informacoes['Processador'] = platform.processor()
    informacoes['Memória Total (GB)'] = round(psutil.virtual_memory().total / (1024 ** 3), 2)
    informacoes['Memória Disponível (GB)'] = round(psutil.virtual_memory().available / (1024 ** 3), 2)
    informacoes['Uso de CPU (%)'] = psutil.cpu_percent(interval=1)
    return informacoes

def atualizar_planilha(informacoes, arquivo='informacoes_computadores.xlsx'):
    if os.path.exists(arquivo):
        workbook = load_workbook(arquivo)
        sheet = workbook.active
    else:
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(['Nome do Computador', 'Endereço IP', 'Nome do Usuário', 'Sistema Operacional', 'Versão do SO', 'Processador', 'Memória Total (GB)', 'Memória Disponível (GB)', 'Uso de CPU (%)'])

    sheet.append([informacoes['Nome do Computador'], informacoes['Endereço IP'], informacoes['Nome do Usuário'], informacoes['Sistema Operacional'], informacoes['Versão do SO'], informacoes['Processador'], informacoes['Memória Total (GB)'], informacoes['Memória Disponível (GB)'], informacoes['Uso de CPU (%)']])
    workbook.save(arquivo)

def main():
    informacoes = coletar_informacoes()
    atualizar_planilha(informacoes)

if __name__ == "__main__":
    main()