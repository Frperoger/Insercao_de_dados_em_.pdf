import os
import pandas as pd
import fitz #mexer diretamente no pdf
from datetime import datetime

#Caminho das pastas
PASTA_ENTRADA = r""
PASTA_SAIDA = r""
PLANILHA = r"" #Planilha onde haverá confirmação de dados
PASTA_LOG = r""

#LOGS
def escrever_log(mensagem):
    data = datetime.now().strftime("%d_%m_%Y") # 1 arquivo por dia
    arquivo_log = os.path.join(PASTA_LOG,f"log_{data}.txt")
    agora = datetime.now().strftime("%H:%M:%S")
    with open(arquivo_log,"a",encoding="utf-8") as f:
        f.write(f"[{agora}]{mensagem}\n")

#cria pastas
os.makedirs(os.path.join(PASTA_ENTRADA), exist_ok=True)
os.makedirs(PASTA_SAIDA, exist_ok=True)
os.makedirs(PASTA_LOG,exist_ok=True)

#Abre arquivo excel
df = pd.read_excel(PLANILHA)
for arquivo in os.listdir(PASTA_ENTRADA):
    if not arquivo.lower().endswith(".pdf"):
        continue
    caminho_pdf = os.path.join(PASTA_ENTRADA,arquivo)
    pdf = fitz.open(caminho_pdf)
    pagina = pdf[0]
    texto = pagina.get_text()
    cpf_encontrado = None
    
    #Variaveis a serem comparadas da planilha
    for _, linha in df.iterrows():
        coluna1 = str(linha["Coluna1"]).strip()
        coluna2 = str(linha["Coluna2"]).strip()
        coluna3 = str(linha["Coluna3"]).strip()
        
        if coluna1 in texto and coluna2 in texto:
            coluna3_encontrada = coluna3
            break
    if not coluna3_encontrada:
        escrever_log(f"Funcionário não encontrado: {arquivo}")
        pdf.close()
        continue
        
    #Posição do texto inserido no PDF
    largura = pagina.rect.width
    altura = pagina.rect.height
        
    x = 10
    y = altura - 130
    y2 = y - 415
    
    pagina.insert_text(
        (x,y),
        f"CPF: {coluna3_encontrada}",
        fontsize = 8,
        fontname="couriernew",
        fontfile=r"C:\Windows\Fonts\Cour.ttf"
    )
    
    pagina.insert_text(
        (x,y2),
        f"CPF: {coluna3_encontrada}",
        fontsize = 8,
        fontname="couriernew",
        fontfile=r"C:\Windows\Fonts\Cour.ttf"
    )
    
    coluna3_limpo = coluna3.replace(".-",'')
    novo_coluna1 = f"{coluna1}_{coluna3_limpo}.pdf"
    
    escrever_log(f"CPF inserido: {arquivo} --> {novo_coluna1}")
    caminho_saida = os.path.join(PASTA_SAIDA,novo_coluna1)
    pdf.save(caminho_saida)
    pdf.close()