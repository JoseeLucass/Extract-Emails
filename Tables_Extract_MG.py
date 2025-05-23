import os
import re
import pandas as pd

mapa_inep = {
    31041556: "EE ANTONIO ALTICIANO",
    31145947: "EE DE AGUA QUENTE",
    31096512: "EE SAO JOSE",
    31184527: "EE CONDE AFONSO CELSO",
    31019046: "EE GOVERNADOR BIAS FORTES",

}

def carregar_planilha_xlsx(caminho: str) -> pd.DataFrame:
   
    bruto = pd.read_excel(caminho, header=None, dtype=str)
    header_row = bruto[bruto.apply(
        lambda r: r.astype(str).str.contains("Código da Escola", case=False, na=False).any(), axis=1
    )].index[0]
    df = pd.read_excel(caminho, header=header_row, dtype=str)
    df = df.dropna(how="all")
    return df
 
def extrair_emails(pasta: str, mapa: dict) -> list[dict]:
    resultados = {}

    mapa_6d = {}
    for inep, nome in mapa.items():
        chave = int(str(inep)[-6:])       
        mapa_6d.setdefault(chave, []).append((inep, nome))

    for arquivo in os.listdir(pasta):
        if not arquivo.lower().endswith(".xlsx"):
            continue                      

        caminho = os.path.join(pasta, arquivo)
        try:
            df = carregar_planilha_xlsx(caminho)
        except Exception as e:
            print(f"Erro ao abrir {arquivo}: {e}")
            continue

        col_codigo = next(c for c in df.columns if "Código da Escola" in c)
        col_email  = next(c for c in df.columns if "mail"            in c.lower())

        for _, row in df.iterrows():
            codigo_raw = str(row.get(col_codigo, "")).strip()
            email      = str(row.get(col_email,   "")).strip()

            if not codigo_raw or not email:
                continue 
            try:
                codigo_int = int(float(codigo_raw))
            except ValueError:
                continue

            if codigo_int not in mapa_6d:
                continue  

            for inep, nome in mapa_6d[codigo_int]:
                
                if inep in resultados:
                    continue
                resultados[inep] = {
                    "INEP":   inep,
                    "Escola": nome,
                    "Email":  email,
                    "Fonte":  arquivo,
                }

    return list(resultados.values())

def salvar_em_excel(registros: list[dict], saida="Lista_de_Emails_MG.xlsx"):
    if not registros:
        print("Nenhum e-mail correspondente encontrado.")
        return
    pd.DataFrame(registros).to_excel(saida, index=False)
    print(f"Arquivo salvo: {saida}  —  {len(registros)} escolas")

if __name__ == "__main__":
    pasta_dados = "Tabela das Escolas de MG"      
    lista = extrair_emails(pasta_dados, mapa_inep)
    salvar_em_excel(lista)
