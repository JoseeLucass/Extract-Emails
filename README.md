# Extração de E-mails (PDFs e Tabelas)

Projeto para extração automática de e-mails a partir de **PDFs** e **planilhas CSV/XLSX**, com opção de correspondência aproximada (*fuzzy matching*) com nomes de entidades.

---

## Funcionalidades

- Extração de e-mails em PDFs via **pdfplumber**  
- Leitura de e-mails em colunas específicas de **CSV/XLSX**  
- Correspondência aproximada com nomes via **RapidFuzz**  
- Exportação dos resultados em **.csv**

---

## Ferramentas Utilizadas

- `os`  
- `re`  
- `csv`  
- `unicodedata`  
- `pandas`  
- `pdfplumber`  
- `rapidfuzz` (`process`, `fuzz`)

---
