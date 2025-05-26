# Extrator de E-mails com Fuzzy Matching (PDF, CSV e XLSX)

Este script em Python realiza a **extração automática de e-mails** a partir de documentos nos formatos **PDF**, **CSV** e **XLSX**, com suporte a **fuzzy matching** para correspondência aproximada de nomes de instituições (escolas, empresas, etc). Ideal para padronizar e validar bases de dados heterogêneas.

## Funcionalidades principais

- **Leitura de arquivos PDF** usando `pdfplumber`
- **Leitura de planilhas Excel** (.xlsx) e arquivos CSV
- Extração de e-mails utilizando expressões regulares (`re`)
- Correspondência aproximada (fuzzy matching) de nomes com `rapidfuzz`
- Normalização de texto (remoção de acentos, espaços e símbolos) com `unicodedata`
- Exportação dos resultados em arquivos `.xlsx` ou `.csv`

## Tecnologias utilizadas

- Python 3.8+
- `pdfplumber` para ler texto de PDFs
- `pandas` para manipulação de dados tabulares
- `rapidfuzz` para fuzzy matching de strings
- `re` para extração de e-mails via regex
- `unicodedata` para normalização de texto
- `os` e `csv` para manipulação de arquivos
