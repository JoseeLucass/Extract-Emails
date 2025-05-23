import os
import pandas as pd
import re
from rapidfuzz import process, fuzz
import csv
import unicodedata

lista_escolas_referencia = [
    "JOAO DE ALMEIDA ESCOLA MUNICIPAL", "JOSE BONIFACIO DO COUTO", "PROFESSORA ORNELLA RITA FERRARI SACILOTTO",
    "NIOMAR APPARECIDA MATTOS GOBBO AMARAL GURGEL PROFA", "MARIO PATARRA FRATTINI PROF", "JOAO SOLIDARIO PEDROSO PROF",
    "VICTOR CIVITA"
]

pasta_dos_arquivos = "Tabela das Escolas de SP"

def normalizar_texto(texto):
    texto = str(texto).upper().strip()
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join(c for c in texto if not unicodedata.combining(c))  
    texto = re.sub(r'\s+', ' ', texto)  
    return texto

lista_escolas_referencia_norm = [normalizar_texto(e) for e in lista_escolas_referencia]

def extrair_email(celula):
    if pd.isna(celula):
        return None
    encontrados = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', str(celula))
    return encontrados[0].lower() if encontrados else None

def validar_email(email):
    if email and email.endswith("@educacao.sp.gov.br"):
        return True
    return False

def encontrar_escola(nome, lista_escolas, limiar=90):
    nome_norm = normalizar_texto(nome)
    resultado = process.extractOne(nome_norm, lista_escolas, scorer=fuzz.partial_ratio)
    if resultado and resultado[1] >= limiar:
        return resultado[0]
    return None

emails_unicos = {}

for arquivo in os.listdir(pasta_dos_arquivos):
    caminho = os.path.join(pasta_dos_arquivos, arquivo)
    
    try:
        print(f"Processando arquivo: {arquivo}")

        if arquivo.endswith('.csv'):
            df = pd.read_csv(
                caminho, 
                encoding='utf-8', 
                engine='python', 
                quoting=csv.QUOTE_NONE, 
                on_bad_lines='skip',
                delimiter=';'
            )
        elif arquivo.endswith('.xlsx'):
            df = pd.read_excel(caminho)
        else:
            print(f"Arquivo ignorado (extensão não suportada): {arquivo}")
            continue
        
        print(f"Colunas encontradas ({len(df.columns)}): {list(df.columns)}")
        
        if df.shape[1] <= 19:
            print(f"Arquivo {arquivo} ignorado porque não tem pelo menos 20 colunas.")
            continue
        
        for idx, linha in df.iterrows():
            escola_raw = linha.iloc[6]
            escola = normalizar_texto(escola_raw)
            email_bruto = linha.iloc[19]
            email = extrair_email(email_bruto)

            escola_correspondida_norm = encontrar_escola(escola, lista_escolas_referencia_norm, limiar=90)

            if escola_correspondida_norm and email and validar_email(email):
                
                idx_ref = lista_escolas_referencia_norm.index(escola_correspondida_norm)
                escola_correspondida = lista_escolas_referencia[idx_ref]

                if escola_correspondida not in emails_unicos:
                    emails_unicos[escola_correspondida] = email
                    print(f"Email adicionado: {escola_correspondida} -> {email} (Origem: {escola_raw})")
            else:
                if idx < 5:  
                    print(f"Linha {idx} descartada: Escola='{escola_raw}', Normalizada='{escola}', Email='{email}', Match='{escola_correspondida_norm}'")

    except Exception as e:
        print(f"Erro ao processar {arquivo}: {e}")

if emails_unicos:
    df_resultado = pd.DataFrame([
        {"Escola": escola, "Email": email}
        for escola, email in emails_unicos.items()
    ])

    df_resultado.to_excel("Lista_de_Emails_SP.xlsx", index=False)
    print("Arquivo 'emails_extraidos.xlsx' salvo com sucesso.")
else:
    print("Nenhum e-mail extraído.")
