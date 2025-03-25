import os
import socket
import getpass
import platform
import psutil
from openpyxl import Workbook, load_workbook
import cpuinfo
import winreg
import tkinter as tk
from tkinter import messagebox

def coletar_informacoes():
    informacoes = {}
    informacoes['Nome do Computador'] = os.getenv('COMPUTERNAME')
    informacoes['Endereço IP'] = socket.gethostbyname(socket.gethostname())
    informacoes['Nome do Usuário'] = getpass.getuser()
    
    # Nome do Domínio ou Workgroup
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Services\Tcpip\Parameters")
        domain, _ = winreg.QueryValueEx(key, "Domain")
        if domain:
            informacoes['Domínio'] = domain
        else:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters")
            workgroup, _ = winreg.QueryValueEx(key, "Workgroup")
            informacoes['Domínio'] = workgroup
    except Exception as e:
        informacoes['Domínio'] = "Desconhecido"
    
    # Sistema Operacional e Versão
    so_info = platform.uname()
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        edition, _ = winreg.QueryValueEx(key, "EditionID")
        informacoes['Sistema Operacional'] = f"{so_info.system} {so_info.release} {edition}"
    except Exception as e:
        informacoes['Sistema Operacional'] = f"{so_info.system} {so_info.release} Desconhecida"
    
    # Processador
    cpu_info = cpuinfo.get_cpu_info()
    # Ajustar o formato do processador
    processador = cpu_info['brand_raw']
    processador = processador.split(' ')[-3] + ' ' + ' '.join(processador.split(' ')[-2:])
    informacoes['Processador'] = processador
    
    # Memória Total
    informacoes['Memória Total (GB)'] = round(psutil.virtual_memory().total / (1024 ** 3))  # Arredondar para cima
    
    return informacoes

def atualizar_planilha(informacoes):
    # Obter o diretório onde o executável está localizado
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    arquivo = os.path.join(diretorio_atual, 'informacoes_computadores.xlsx')

    try:
        if os.path.exists(arquivo):
            workbook = load_workbook(arquivo)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(['Nome do Computador', 'Endereço IP', 'Nome do Usuário', 'Domínio', 'Processador', 'Memória Total (GB)', 'Sistema Operacional'])

        sheet.append([informacoes['Nome do Computador'], informacoes['Endereço IP'], informacoes['Nome do Usuário'], informacoes['Domínio'], informacoes['Processador'], informacoes['Memória Total (GB)'], informacoes['Sistema Operacional']])
        workbook.save(arquivo)
        mostrar_alerta(f"Planilha atualizada com sucesso: {arquivo}")
    except PermissionError:
        mostrar_alerta(f"Erro de permissão ao tentar salvar o arquivo: {arquivo}. Verifique se o arquivo está aberto ou se você tem permissões de escrita.")
    except Exception as e:
        mostrar_alerta(f"Ocorreu um erro ao tentar salvar o arquivo: {e}")

def mostrar_alerta(mensagem):
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela principal
    messagebox.showinfo("Aviso", mensagem)
    root.after(2000, root.destroy)  # Fechar a janela após 2 segundos

def main():
    informacoes = coletar_informacoes()
    atualizar_planilha(informacoes)

if __name__ == "__main__":
    main()