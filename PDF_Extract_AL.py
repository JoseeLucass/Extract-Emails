import pdfplumber
import re
import pandas as pd
import os

mapa_inep_escolas = {
    "27011844": "ESCOLA ESTADUAL MARQUES DA SILVA",
    "27041603": "ESCOLA ESTADUAL JOSEFA CAVALCANTE SURUAGY",
    "27001750": "ESCOLA ESTADUAL DE ENSINO MEDIO NEZINHO PEREIRA",
    "27216616": "ESCOLA ESTADUAL MANOEL DE ARAUJO DORIA",
    "27035573": "ESCOLA ESTADUAL DR EDSON DOS SANTOS BERNARDES",
    "27038483": "ESCOLA ESTADUAL OVIDIO EDGAR DE ALBUQUERQUE",
    "27035646": "ESCOLA ESTADUAL ALFREDO GASPAR DE MENDONCA",
    "27253007": "ESCOLA ESTADUAL PROFESSOR LIBERALINO BONFIM DE OLIVEIRA",
    "27035590": "ESCOLA ESTADUAL JORNALISTA LAFAIETTE BELO",
    "27038556": "ESCOLA ESTADUAL DR MIGUEL GUEDES NOGUEIRA",
    "27038564": "ESCOLA ESTADUAL PROFª AURELINA PALMEIRA DE MELO",
    "27038599": "ESCOLA ESTADUAL PROF SEBASTIAO DA HORA",
    "27036499": "ESCOLA ESTADUAL TARCISIO DE JESUS",
    "27034887": "ESCOLA ESTADUAL PROFESSOR AFRANIO LAGES",
    "27037142": "ESCOLA ESTADUAL PROFª JOSEFA CONCEICAO DA COSTA",
    "27006590": "ESCOLA ESTADUAL PROFª ANA MARIA TEODOSIO",
    "27031365": "ESCOLA ESTADUAL PROFESSOR GUEDES DE MIRANDA",
    "27040240": "ESCOLA ESTADUAL FERNANDINA MALTA",
    "27008452": "ESCOLA ESTADUAL LUCILO JOSE RIBEIRO",
    "27051617": "ESCOLA ESTADUAL PROFA EDLEUZA OLIVEIRA DA SILVA"
}

def extrair_emails_por_inep_em_pdf(pdf_path, mapa_inep):
    resultados = {}
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
            for inep in mapa_inep.keys():
                if inep in texto and inep not in resultados:
                    bloco = texto.split(inep, 1)[1][:500]
                    email = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', bloco)
                    if email:
                        resultados[inep] = {
                            "INEP": inep,
                            "Escola": mapa_inep[inep],
                            "Email": email.group(0),
                            "Arquivo": os.path.basename(pdf_path)
                        }
    return list(resultados.values())

def processar_pasta_pdfs(pasta, mapa_inep):
    resultados_totais = []
    for nome_arquivo in os.listdir(pasta):
        if nome_arquivo.lower().endswith(".pdf"):
            caminho_pdf = os.path.join(pasta, nome_arquivo)
            print(f"Processando: {nome_arquivo}")
            resultados = extrair_emails_por_inep_em_pdf(caminho_pdf, mapa_inep)
            resultados_totais.extend(resultados)
    return resultados_totais

def salvar_em_excel(dados, nome_arquivo="Lista_de_Emails_AL.xlsx"):
    df = pd.DataFrame(dados)
    df.to_excel(nome_arquivo, index=False)
    print(f"\n Planilha salva como: {nome_arquivo}")


pasta_dos_pdfs = "PDF's das Escolas de AL"
dados = processar_pasta_pdfs(pasta_dos_pdfs, mapa_inep_escolas)
salvar_em_excel(dados)
