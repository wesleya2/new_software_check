#Bibliotecas
from glob import iglob
from os.path import getmtime
import pandas as pd
import win32com.client as win32
import subprocess
from subprocess import DEVNULL

#mapear unidade
subprocess.call(r'net use w: \\applicationsrepository$\ZIP', stdout=DEVNULL, stderr=DEVNULL)

#lista programa verificar
programas_x64 = ['Google Chrome*','Mozilla Firefox*','Microsoft Power BI Desktop*','Microsoft Project Standard 2016 C2R*','Microsoft Project Professional 2016 C2R*','Microsoft Visio Standard 2016 C2R*','Microsoft Visio Professional 2016 C2R*','Tableau Reader*','Steelray Project Viewer*','Microsoft Edge*','Microsoft 365 Apps for Enterprise 16.0.14*']
programas_x32 = ['Adobe Acrobat Reader DC Patch*','Zoom Skype for Business Plugin*','Zoom Outlook Plugin*','Zoom 5.1*']
#listas
new_version = []
new_version2 = []
old_version = []

#Carregar a ultima verificação
software_antigo = pd.read_csv('/PycharmProjects/Projeto_SCCM/Arquivos_Temp/programas.csv')
for i, sof in enumerate(software_antigo['Software']):
    old_version.append(software_antigo['Software'][i])

#Verifica novas versões x64
for i, programa in enumerate(programas_x64):
    files = iglob(f'W:/{programas_x64[i]}.zip')
    arquivo_mais_recente = max(files, key=getmtime)
    if 'x86' in arquivo_mais_recente:
        arquivo_mais_recente = arquivo_mais_recente.replace('x86','')
        new_version.append(arquivo_mais_recente[3:])

    else:
        new_version.append(arquivo_mais_recente[3:])

#Verifica novas versões x32
for i, programa in enumerate(programas_x32):
    files = iglob(f'W:/{programas_x32[i]}.zip')
    arquivo_mais_recente = max(files, key=getmtime)
    new_version.append(arquivo_mais_recente[3:])

for i, novo in enumerate(new_version):
    if new_version[i] in old_version:
        continue
    else:
        new_version2.append(new_version[i])

#remover mapeamento
subprocess.call(r'net use w: /delete /yes', stdout=DEVNULL, stderr=DEVNULL)

#Salvar arquivo com a ultima verificação
old_version = new_version
arquivo_final = pd.DataFrame(old_version, columns=['Software'])
arquivo_final.to_csv('folderpath/Projeto_SCCM/Arquivos_Temp/programas.csv', index=False)
#criando anexo
anexo_final = pd.DataFrame(new_version2, columns=['Nome'])
anexo_final.to_csv(r'/PycharmProjects/Projeto_SCCM/Arquivos_Temp/novos_programas.csv', index=False)

#Enviar Email
base_contato = pd.read_csv(r'/PycharmProjects/Projeto_SCCM/base_contatos.csv', delimiter=';')
if len(new_version2) > 0:
    outlook = win32.Dispatch('outlook.application')
    for i, contato in enumerate(base_contato['Email']):
        email = outlook.CreateItem(0)
        email.To = base_contato['Email'][i]
        email.Subject = 'Atualizar SCCM'
        email.Body = f"""{len(new_version2)} Softwares foram atualizados!"""
        anexo = '/PycharmProjects/Projeto_SCCM/Arquivos_Temp/novos_programas.csv'
        email.Attachments.Add(anexo)
        email.Send()
else:
    outlook = win32.Dispatch('outlook.application')
    for i, contato in enumerate(base_contato['Email']):
        email = outlook.CreateItem(0)
        email.To = base_contato['Email'][i]
        email.Subject = 'Atualizar SCCM'
        email.Body = f"""{len(new_version2)} Softwares foram atualizados!"""
        email.Send()
        
        
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = 'teste@teste.com'
email.Subject = 'Atualizar SCCM'
email.Body = """teste"""
email.Send()

