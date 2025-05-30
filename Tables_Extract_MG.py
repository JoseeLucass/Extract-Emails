import os
import re
import pandas as pd

mapa_inep = {
    31041556: "EE ANTONIO ALTICIANO",
    31145947: "EE DE AGUA QUENTE",
    31096512: "EE SAO JOSE",
    31184527: "EE CONDE AFONSO CELSO",
    31019046: "EE GOVERNADOR BIAS FORTES",
    31128252: "EE ANTONIO CARLOS",
    31239372: "EE LIMA DUARTE",
    31166634: "EE DONA ELEONORA PIERUCCETTI",
    31015334: "EE CONEGO LUIZ GIAROLA CARLOS",
    31364932: "CENTRO INTERESCOLAR DE CULTURA ARTE LINGUAGENS E TECNOLOGIAS",
    31002330: "EE SIRIA MARQUES DA SILVA",
    31002151: "EE PADRE JOAO BOSCO PENIDO BURNIER",
    31001767: "EE DOUTOR JOSE DO PATROCINIO DA SILVA PONTES",
    31246425: "EE PROFESSORA ADIR ANDRADE ALBANO",
    31000175: "EE MARIA DE LOURDES DE OLIVEIRA",
    31001571: "EE PADRE EUSTAQUIO",
    31000493: "EE DEPUTADO ILACIR PEREIRA LIMA",
    31342440: "EE ZILDA ARNS NEUMANN",
    31000051: "EE GUIA LOPES",
    31001457: "EE MADRE CARMELITA",
    31001317: "EE CORACAO EUCARISTICO",
    31003000: "ESCOLA MUNICIPAL JONAS BARCELLOS CORREA",
    31362247: "ESCOLA MUNICIPAL JARDIM LEBLON",
    31002259: "EE AFRANIO DE MELO FRANCO",
    31001848: "EE DOUTOR ANTONIO AUGUSTO SOARES CANEDO",
    31317357: "EE PROFESSOR AGNELO CORREIA VIANA",
    31322563: "EE JOVEM PROTAGONISTA",
    31002801: "ESCOLA MUNICIPAL HELENA ANTIPOFF",
    31002241: "EE PASCHOAL COMANDUCCI",
    31307700: "COLEGIO TIRADENTES PMMG MINAS CAIXA",
    31001473: "EE CORONEL VICENTE TORRES JUNIOR",
    31001376: "EE LUCIO DOS SANTOS",
    31000558: "EE MARIO CASASSANTA",
    31007935: "EE TITO LIVIO DE SOUZA",
    31321061: "COLEGIO TIRADENTES PMMG",
    31008028: "EE DO BAIRRO SAO CAETANO",
    31294667: "EE CONEGO JOAO SEVERO",
    31079499: "EE JOAO OSORIO DE QUEIROZ",
    31311898: "EE CORONEL EGIDIO BENICIO DE ABREU",
    31019135: "EE DONA NHANHA",
    31008117: "EE MELO VIANA",
    31369810: "EE DE ENSINO FUNDAMENTAL ANOS FINAIS E ENSINO MEDIO",
    31123951: "EE JOAO DE SOUZA GONCALVES",
    31180785: "EE JOSE ALVES DE MAGALHAES",
    31140384: "EE NOSSA SENHORA DO CARMO",
    31246204: "EE ANALIA CARNEIRO DOS SANTOS",
    31108430: "EE JOSE GOMES PIMENTEL",
    31008486: "EE JOSE PEREIRA CANCADO",
    31054682: "EE ANTONIO FELIPE DE SALLES",
    31171361: "EE MARIA UMBELINA DE ANDRADE GOMES",
    31079642: "EE CIRILO PEREIRA DA FONSECA",
    31043885: "EE LEVINDO DIAS",
    31171786: "EE PROFESSOR WANDERLEY FERREIRA DE REZENDE",
    31032701: "EE PRESIDENTE TANCREDO NEVES",
    31193356: "EE GUSTAVO AUGUSTO DA SILVA",
    31196436: "EE BELCHIOR DE FARIA",
    31330671: "EE COMENDADOR GOMES",
    31310883: "EE DR LINDOLFO BERNARDES",
    31042081: "EE DE CONSELHEIRO PENA",
    31042200: "EE MARIA GARCIA PINTO",
    31042145: "EE MARIA GUILHERMINA PENA",
    31008681: "EE PRESIDENTE TANCREDO NEVES",
    31239232: "EE SAO SEBASTIAO",
    31140554: "EE MESTRE CANDINHO",
    31190900: "EE PROFESSORA CELINA MACHADO",
    31351067: "EE DE ENSINO FUNDAMENTAL E MEDIO",
    31140848: "EE IRMA RAIMUNDA MARQUES",
    31254355: "EE JULIANA CATARINA DA SILVEIRA",
    31023825: "EE PROFESSORA ISABEL MOTTA",
    31097551: "EE DOUTOR PEDRO PAULO NETO",
    31033014: "EE ARMANDO NOGUEIRA SOARES",
    31108375: "EE DOM BOSCO",
    31097764: "EE FAZENDA PARAISO",
    31055107: "EE EDUARDO AMARAL",
    31068292: "EE ANTONIO MACEDO",
    31023981: "EE FELICIO DOS SANTOS",
    31115207: "EE PROFESSOR JOAQUIM RODARTE",
    31043141: "EE MANOEL BYRRO",
    31205371: "EE DO BAIRRO JARDIM DO IPE",
    31043559: "EE DE SAO VITOR",
    31043184: "EE PROFESSOR PAULO FREIRE",
    31043681: "EE TENENTE JOSE COELHO DA ROCHA",
    31172766: "EE PREFEITO CELSO VIEIRA VILELA",
    31080501: "EE BOM JESUS DA VEREDA",
    31009181: "EE JUSCELINO KUBITSCHEK DE OLIVEIRA",
    31009229: "EE RACHEL IANCU STEURMAN",
    31009253: "EE JOAQUIM JOSE PEREIRA",
    31020699: "EE EUCLIDES PINTO DE OLIVEIRA",
    31020613: "EE JOSE FRANCISCO DE PAIVA CAMPOS",
    31190993: "EE ALMIRANTE TOYODA",
    31191124: "EE MAURILIO ALBANESE NOVAES",
    31191159: "EE HAYDEE MARIA IMACULADA SCHITTINI",
    31191001: "EE LAURA XAVIER SANTANA",
    31124460: "EE CRISTIANO MACHADO",
    31103187: "EE TRAJANO PROCOPIO DE ALVARENGA SILVA MONTEIRO",
    31172936: "EE PROFESSOR SOUZA NILO",
    31240800: "EE ARY PIMENTA BUGELLI",
    31033871: "EE DE ITAUNA",
    31146986: "EE MANOEL DA SILVA GUSMAO",
    31196592: "EE PROFESSORA MARIA DE BARROS",
    31159182: "EE NOSSA SENHORA DE LOURDES",
    31184799: "EE DO HAVAI",
    31231843: "EE DO NUCLEO HABITACIONAL I",
    31369870: "EE DE ENSINO FUNDAMENTAL",
    31080624: "EE PROFESSORA NHAGUI AZEVEDO",
    31103462: "EE JOAO XXIII",
    31103527: "EE DOUTOR GERALDO PARREIRAS",
    31111660: "ESCOLA MUNICIPAL ISRAEL PINHEIRO",
    31051951: "EM SERAFIM LOPES GODINHO",
    31034410: "EE MARIA RITA DUARTE",
    31068420: "EE ANTONIO CARLOS",
    31352110: "EM CORONEL EMILIO ESTEVES DOS REIS",
    31071692: "EM MARIA ALADIA SANT ANA",
    31071455: "EM SANTANA ITATIAIA",
    31071803: "EM CAMILO GUEDES",
    31068683: "EE DUARTE DE ABREU",
    31071315: "EM PEDRO NAGIB NASSER",
    31071498: "EM FERNAO DIAS PAES",
    31071242: "EM ALMERINDA DE OLIVEIRA TAVARES",
    31034185: "EE COMENDADOR ZICO TOBIAS",
    31173061: "EE IRACEMA RODRIGUES",
    31075094: "EE SAO VICENTE DE PAULO",
    31062961: "EE DE MONTALVANIA",
    31147583: "EE DE LAMBARI",
    31106488: "EE DE OURO PRETO",
    31082147: "EE SANTOS DUMONT",
    31034908: "EE GOVERNADOR VALADARES",
    31108839: "EE AFFONSO ROQUETTE",
    31115428: "EE DEUS UNIVERSO E VIRTUDE",
    31118958: "EE MARCOLINO DE BARROS",
    31118796: "EE DOUTOR PAULO BORGES",
    31118991: "EE PROFESSOR MODESTO",
    31199109: "EE MARIANA TAVARES",
    31199184: "EE ODILON BEHRENS",
    31128881: "EE CORONEL ANTONINHO",
    31129071: "EE PROFESSOR RAYMUNDO MARTINIANO FERREIRA",
    31362280: "COLEGIO TIRADENTES PMMG",
    31319163: "EE CORONEL PEDRO NERY",
    31015989: "EE GALDINO ANANIAS DE SANTANA",
    31212679: "EE MARIA DA GLORIA ASSUNCAO",
    31010111: "EE VEREADOR JOSE ROBERTO PEREIRA",
    31010081: "EE JOAO DE DEUS GOMES",
    31231711: "EE MARIA DA PIEDADE SOUZA ROCHA",
    31185396: "EE LIDIO ALMEIDA",
    31222470: "EE JUQUINHA DE ALMEIDA",
    31330680: "EE PROFESSORA MARGARET BARROSO PINTO",
    31159484: "EE BARAO DA RIFAINA",
    31134911: "EE AMELIA PASSOS",
    31082571: "EE TENENTE FELISMINO HENRIQUES DE SOUZA",
    31010740: "EE FRANCISCO TIBURCIO DE OLIVEIRA",
    31010821: "EE GERVASIO LARA",
    31075787: "EE DE RIBEIRAO DE SAO DOMINGOS",
    31173983: "EE PADRE JOSE RIBEIRO",
    31194395: "EE DR JOAO NOGUEIRA DE ALMEIDA",
    31103900: "EE CORONEL JOSE GOMES DE ARAUJO",
    31103993: "EE VICENTE DE PAULA FRAGA",
    31103969: "EE CRISTIANO MACHADO",
    31174084: "EE MINISTRO LUCIO DE MENDONCA",
    31338761: "EE ALINE DIAS NEVES",
    31134546: "EE BRIGHENTI CESARE",
    31045331: "EE MAJOR LERMINO PIMENTA",
    31070149: "EE OSWALDO CRUZ",
    31009202: "EE PROFESSOR ERNESTO CARNEIRO SANTIAGO",
    31025003: "EE MESTRA VIRGINIA REIS",
    31024996: "EE MESTRA ROSA MADUREIRA FAGUNDES",
    31141909: "EE SANTOS AZEREDO",
    31141593: "EE DOUTOR AFONSO VIANA",
    31141852: "EE PROFESSOR JOAO FERNANDINO JUNIOR",
    31342572: "EE VENCESLAU BRAS",
    31147354: "EE DE SETUBAL",
    31191531: "EE PROFESSORA ANA LETRO STAACKS",
    31057151: "ESCOLA MUNICIPAL AMBROSINA MARIA DE JESUS",
    31269981: "EM SAO JOSE",
    31174688: "EE DEPUTADO TEODOSIO BANDEIRA",
    31239364: "ESCOLA MUNICIPAL SAO JOAO BATISTA",
    31182036: "EE GOVERNADOR VALADARES",
    31165701: "EM FREDERICO PEIRO",
    31159867: "EE BRASIL",
    31165441: "EM MONTEIRO LOBATO",
    31159981: "EE DOM EDUARDO",
    31338893: "EE DE ENSINO FUNDAMENTAL E MEDIO",
    31245909: "E M PROF DOMINGOS PIMENTEL DE ULHOA",
    31167614: "EE SEGISMUNDO PEREIRA",
    31167339: "EE PROFESSOR JOSE IGNACIO DE SOUSA",
    31325473: "E M PROFESSORA ORLANDA NEVES STRACK",
    31322601: "EE DARCI RIBEIRO",
    31175048: "EE IRMAO MARIO ESDRAS"
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
