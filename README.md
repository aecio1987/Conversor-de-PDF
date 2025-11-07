# 🧾 Extração de Registros de Ponto em PDF

Este projeto converte relatórios de **registro de ponto (em PDF)** para uma planilha Excel (.xlsx).  
Ele identifica automaticamente as tabelas do PDF e as transforma em dados tabulares.

O script tenta duas abordagens:
1. **Docling** — extração direta de tabelas (método principal);
2. **PyMuPDF (fitz)** — extração via texto e regex (método alternativo).

---

## 🚀 Pré-requisitos

- Python 3.9 ou superior
- Instale as dependências executando:

```bash
pip install pandas numpy pymupdf
