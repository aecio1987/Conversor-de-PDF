"""
Script de extração de registros de ponto a partir de PDFs e exportação para Excel.

O script permite extrair tabelas de PDFs (relatórios de ponto) usando:
1️⃣ Docling — método principal (mais preciso para PDFs estruturados)
2️⃣ PyMuPDF (fitz) — método alternativo via texto + regex

"""

import os
import re
import pandas as pd
import numpy as np

# ----------------------------------------------------------------------
# TENTATIVA DE IMPORTAÇÃO DAS BIBLIOTECAS
# ----------------------------------------------------------------------
try:
    from docling.document import Document
    DOCLING_AVAILABLE = True
except Exception:
    Document = None
    DOCLING_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except Exception:
    fitz = None
    PYMUPDF_AVAILABLE = False


# ----------------------------------------------------------------------
# CONFIGURAÇÕES DO USUÁRIO
# ----------------------------------------------------------------------

# 🔧 Informe o nome do seu arquivo PDF (de entrada)
NOME_ARQUIVO_PDF = "entrada.pdf"

# 🔧 Informe o nome do arquivo Excel (de saída)
NOME_ARQUIVO_EXCEL = "saida.xlsx"

# 🔧 Defina aqui os nomes das colunas conforme aparecem no SEU PDF
#    Por exemplo: ["Nome", "Data", "Entrada 1", "Saída 1", "Entrada 2", "Saída 2", "Total"]
#    A quantidade deve corresponder à tabela que aparece no PDF.
COLUNAS_NOMES = [
    "Nome",         # Nome do servidor ou funcionário
    "Data",         # Data do registro (dd/mm/yyyy)
    "Entr.Manha",   # Horário de entrada no período da manhã
    "Saíd.Manha",   # Horário de saída no período da manhã
    "Entr.Tarde",   # Horário de entrada no período da tarde
    "Said.Tarde",   # Horário de saída no período da tarde
    "Total"         # Total de horas no dia
]

NUM_COLUNAS_ESPERADO = len(COLUNAS_NOMES)


# ----------------------------------------------------------------------
# FUNÇÃO: EXTRAÇÃO COM DOCLING
# ----------------------------------------------------------------------
def extrair_com_docling(pdf_path):
    """Tenta extrair tabelas usando a biblioteca Docling."""
    if not DOCLING_AVAILABLE:
        print("Docling não disponível. Pulando essa etapa.")
        return None

    if not os.path.exists(pdf_path):
        print(f"Arquivo não encontrado: {pdf_path}")
        return None

    try:
        print("🔹 Extraindo tabelas com Docling...")
        doc = Document.from_file(pdf_path)

        if not getattr(doc, "tables", None):
            print("Docling não detectou tabelas.")
            return None

        df_final = pd.DataFrame()

        def row_contains_keywords(row):
            """Remove cabeçalhos e rodapés com palavras-chave comuns."""
            keywords = [
                "nome", "data", "entr", "said", "tarde", "manha", "instituição",
                "página", "emissão", "estado de mato grosso", "relação de registro"
            ]
            try:
                return any(any(k in str(item).lower() for k in keywords) for item in row)
            except Exception:
                return True

        for i, tabela in enumerate(doc.tables):
            try:
                df_tabela = tabela.to_pandas(fill_na=True)
            except Exception as e:
                print(f"Falha ao converter tabela {i+1}: {e}")
                continue

            if df_tabela.empty:
                continue

            # Garante o número correto de colunas
            if len(df_tabela.columns) > NUM_COLUNAS_ESPERADO:
                df_tabela = df_tabela.iloc[:, :NUM_COLUNAS_ESPERADO]
            elif len(df_tabela.columns) < NUM_COLUNAS_ESPERADO:
                for _ in range(NUM_COLUNAS_ESPERADO - len(df_tabela.columns)):
                    df_tabela[len(df_tabela.columns)] = np.nan

            df_tabela.columns = range(NUM_COLUNAS_ESPERADO)
            df_tabela = df_tabela[~df_tabela.apply(row_contains_keywords, axis=1)]

            df_final = pd.concat([df_final, df_tabela], ignore_index=True)

        if df_final.empty:
            print("Docling produziu DataFrame vazio.")
            return None

        df_final.columns = COLUNAS_NOMES
        df_final = df_final.replace(r'^\s*$', '', regex=True).replace(['nan', 'None'], '', regex=True)

        def limpar_nome(texto):
            """Limpa o campo de nome, removendo ruídos e caracteres extras."""
            texto = re.sub(r'[\n;:]', ' ', str(texto))
            texto = re.sub(r'[^\w\sÁÀÃÂÉÈÊÍÌÓÒÕÔÚÙÇ\-]', ' ', texto)
            return re.sub(r'\s{2,}', ' ', texto).strip()

        df_final[COLUNAS_NOMES[0]] = df_final[COLUNAS_NOMES[0]].apply(limpar_nome)

        print("✅ Extração com Docling concluída.")
        return df_final

    except Exception as e:
        print(f"Erro crítico no Docling: {e}")
        return None


# ----------------------------------------------------------------------
# FUNÇÃO: EXTRAÇÃO COM PyMuPDF (fallback)
# ----------------------------------------------------------------------
def extrair_com_pymupdf(pdf_path):
    """Tenta extrair texto puro e reconstruir tabela com expressões regulares."""
    if not PYMUPDF_AVAILABLE:
        print("PyMuPDF não disponível. Instale com 'pip install pymupdf'.")
        return None

    if not os.path.exists(pdf_path):
        print(f"Arquivo não encontrado: {pdf_path}")
        return None

    try:
        print("🔹 Extraindo texto com PyMuPDF (modo fallback)...")
        doc = fitz.open(pdf_path)
        textos = [page.get_text("text") for page in doc]
        doc.close()

        text = "\n".join(textos)
        text = text.replace("\t", " ")
        text = re.sub(r'\n{2,}', '\n', text)

        registros = []
        for m in re.finditer(r'\d{2}/\d{2}/\d{4}', text):
            date_str = m.group(0)
            snippet = text[max(0, m.start() - 400):m.end()]
            times = re.findall(r'\d{2}:\d{2}:\d{2}', snippet)[-5:]
            while len(times) < 5:
                times.append('')

            first_time_match = re.search(r'\d{2}:\d{2}:\d{2}', snippet)
            nome_cand = snippet[:first_time_match.start()].strip() if first_time_match else snippet.strip()
            nome_cand = re.sub(r'[^\w\sÁÀÃÂÉÈÊÍÌÓÒÕÔÚÙÇ\-]', ' ', nome_cand)
            nome_cand = re.sub(r'\s{2,}', ' ', nome_cand).strip()

            registros.append({
                COLUNAS_NOMES[0]: nome_cand,
                COLUNAS_NOMES[1]: date_str,
                COLUNAS_NOMES[2]: times[0],
                COLUNAS_NOMES[3]: times[1],
                COLUNAS_NOMES[4]: times[2],
                COLUNAS_NOMES[5]: times[3],
                COLUNAS_NOMES[6]: times[4]
            })

        if not registros:
            print("Nenhum registro encontrado via PyMuPDF.")
            return None

        df = pd.DataFrame(registros)
        print(f"✅ PyMuPDF produziu {len(df)} registros.")
        return df

    except Exception as e:
        print(f"Erro no fallback PyMuPDF: {e}")
        return None


# ----------------------------------------------------------------------
# FUNÇÃO PRINCIPAL
# ----------------------------------------------------------------------
def extrair_tabelas(pdf_path):
    """Executa a extração, tentando Docling e depois PyMuPDF."""
    df = extrair_com_docling(pdf_path) if DOCLING_AVAILABLE else None
    if df is None or df.empty:
        print("Tentando fallback com PyMuPDF...")
        df = extrair_com_pymupdf(pdf_path)
    return df


# ----------------------------------------------------------------------
# EXECUÇÃO DO SCRIPT
# ----------------------------------------------------------------------
if __name__ == "__main__":
    df_result = extrair_tabelas(NOME_ARQUIVO_PDF)

    if df_result is None or df_result.empty:
        print("-" * 40)
        print("❌ Falha na extração. Nenhum dado encontrado.")
    else:
        print("-" * 40)
        print(f"✅ Total de registros extraídos: {len(df_result)}")

        try:
            df_result.to_excel(NOME_ARQUIVO_EXCEL, index=False)
            print(f"📁 Planilha salva com sucesso: {NOME_ARQUIVO_EXCEL}")
        except Exception as e:
            print(f"❌ Erro ao salvar Excel: {e}")
