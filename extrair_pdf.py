"""
Script de extração de tabelas de registros de ponto a partir de arquivos PDF.

Este script tenta extrair tabelas automaticamente de PDFs em duas etapas:
1. Primeira tentativa com a biblioteca Docling.
2. Caso Docling falhe ou não esteja instalada, usa PyMuPDF (fitz) como fallback.

O resultado é salvo em um arquivo Excel (.xlsx) com as colunas padronizadas.
"""

import os
import re
import pandas as pd
import numpy as np

# ----------------------------------------------------------------------
# TENTATIVA DE IMPORTAÇÃO DAS BIBLIOTECAS DE EXTRAÇÃO
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
# CONFIGURAÇÕES PRINCIPAIS
# ----------------------------------------------------------------------
# Caminho do PDF de entrada e do Excel de saída
# 🔧 Altere apenas os nomes abaixo conforme seus arquivos (sem caminhos absolutos)
NOME_ARQUIVO_PDF = "entrada.pdf"
NOME_ARQUIVO_EXCEL = "saida.xlsx"

# Nomes esperados das colunas da planilha
COLUNAS_NOMES = ["Nome", "Data", "Entr.Manha", "Saíd.Manha", "Entr.Tarde", "Said.Tarde", "Total"]
NUM_COLUNAS_ESPERADO = len(COLUNAS_NOMES)


# ----------------------------------------------------------------------
# FUNÇÃO: EXTRAÇÃO COM DOCLING
# ----------------------------------------------------------------------
def extrair_com_docling(pdf_path):
    """
    Tenta extrair tabelas usando a biblioteca Docling.
    Retorna um DataFrame pandas ou None em caso de falha.
    """
    if not DOCLING_AVAILABLE:
        print("Docling não disponível nesta instalação. Pulando Docling.")
        return None

    if not os.path.exists(pdf_path):
        print(f"Arquivo não encontrado: {pdf_path}")
        return None

    try:
        print("Tentando extrair com Docling (Document.from_file)...")
        doc = Document.from_file(pdf_path)

        if not getattr(doc, "tables", None):
            print("Docling não detectou tabelas.")
            return None

        df_final = pd.DataFrame()

        def row_contains_keywords(row):
            """Remove linhas de cabeçalho e rodapé baseadas em palavras-chave."""
            keywords_to_check = [
                "nome", "data", "entr.manha", "governo", "instituição", "página",
                "emissão", "estado de mato grosso", "relação de registro",
                "said.manha", "entr.tarde", "totalnome"
            ]
            try:
                for item in row:
                    item_str = str(item).lower()
                    if any(keyword in item_str for keyword in keywords_to_check):
                        return True
            except Exception:
                return True
            return False

        for i, tabela in enumerate(doc.tables):
            try:
                df_tabela = tabela.to_pandas(fill_na=True)
            except Exception as e:
                print(f"Falha ao converter tabela {i+1} para pandas: {e}. Pulando.")
                continue

            if df_tabela.empty:
                continue

            # Garante número correto de colunas
            if len(df_tabela.columns) > NUM_COLUNAS_ESPERADO:
                df_tabela = df_tabela.iloc[:, :NUM_COLUNAS_ESPERADO]
            elif len(df_tabela.columns) < NUM_COLUNAS_ESPERADO:
                for _ in range(NUM_COLUNAS_ESPERADO - len(df_tabela.columns)):
                    df_tabela[len(df_tabela.columns)] = np.nan

            df_tabela.columns = range(NUM_COLUNAS_ESPERADO)
            df_tabela = df_tabela[~df_tabela.apply(row_contains_keywords, axis=1)]

            df_final = pd.concat([df_final, df_tabela], ignore_index=True)

        if df_final.empty:
            print("Docling produziu DataFrame vazio após filtragem.")
            return None

        df_final.columns = COLUNAS_NOMES
        df_final = df_final.replace(r'^\s*$', '', regex=True).replace(['nan', 'None'], '', regex=True)

        # Limpeza de nomes
        def limpar_nome(texto):
            texto = str(texto)
            texto = re.sub(r'[\n;:]', ' ', texto)
            texto = re.sub(r'(\s+\w+)?\s+(Domingo|Sábado|Feriado|Folga|REGISTRO AUSENTE|FALTA INJUSTIFICADA|JORNADA INCOMPLETA)?\s*$',
                           '', texto, flags=re.IGNORECASE).strip()
            texto = re.sub(r'[^\w\sÁÀÃÂÉÈÊÍÌÓÒÕÔÚÙÇ\-]', ' ', texto)
            return re.sub(r'\s{2,}', ' ', texto).strip()

        df_final["Nome"] = df_final["Nome"].apply(limpar_nome)

        # Remove ruídos de rodapé
        def limpar_ruidos(texto):
            texto = str(texto)
            texto = re.sub(r"MT PARTICIPAÇÕES S\.A.*|PARQUE NOVO MT.*|Instituição.*|Governo.*|Página.*|Emissão.*|Estado de Mato Grosso.*", "", texto)
            return re.sub(r"\s{2,}", " ", texto).strip()

        for col in COLUNAS_NOMES:
            df_final[col] = df_final[col].apply(limpar_ruidos)

        # Filtra linhas válidas
        df_final = df_final[df_final['Data'].str.match(r'\d{2}/\d{2}/\d{4}', na=False)]
        df_final.reset_index(drop=True, inplace=True)

        print("Extração com Docling concluída com sucesso.")
        return df_final

    except Exception as e:
        print(f"Erro crítico no Docling: {e}")
        return None


# ----------------------------------------------------------------------
# FUNÇÃO: EXTRAÇÃO ALTERNATIVA COM PyMuPDF
# ----------------------------------------------------------------------
def extrair_com_pymupdf(pdf_path):
    """
    Tenta extrair texto do PDF usando PyMuPDF e organiza os dados via regex.
    Retorna um DataFrame pandas ou None em caso de falha.
    """
    if not PYMUPDF_AVAILABLE:
        print("PyMuPDF não está instalado. Instale com 'pip install pymupdf'.")
        return None

    if not os.path.exists(pdf_path):
        print(f"Arquivo não encontrado: {pdf_path}")
        return None

    try:
        print("Tentando extrair texto com PyMuPDF...")
        doc = fitz.open(pdf_path)
        textos = [page.get_text("text") for page in doc]
        doc.close()

        text = "\n".join(textos)
        text = text.replace("\t", " ")
        text = re.sub(r'\n{2,}', '\n', text)

        # Remove cabeçalhos padrão do documento
        text = re.sub(r'MT PARTICIPAÇÕES S\.A.*?Relação de registro no período .*?\n', '', text, flags=re.DOTALL)
        text = re.sub(r'Entr\.Manha Saíd\.Manha Entr\.Tarde Said\.Tarde TotalNome Data', '', text, flags=re.IGNORECASE)

        registros = []
        for m in re.finditer(r'\d{2}/\d{2}/\d{4}', text):
            date_str = m.group(0)
            snippet = text[max(0, m.start() - 400):m.end()]
            times = re.findall(r'\d{2}:\d{2}:\d{2}', snippet)[-5:]
            while len(times) < 5:
                times.append('00:00:00')

            first_time_match = re.search(r'\d{2}:\d{2}:\d{2}', snippet)
            nome_cand = snippet[:first_time_match.start()].strip() if first_time_match else snippet.strip()
            nome_cand = re.sub(r'[^\w\sÁÀÃÂÉÈÊÍÌÓÒÕÔÚÙÇ\-]', ' ', nome_cand)
            nome_cand = re.sub(r'\s{2,}', ' ', nome_cand).strip()

            registros.append({
                "Nome": nome_cand,
                "Data": date_str,
                "Entr.Manha": times[0],
                "Saíd.Manha": times[1],
                "Entr.Tarde": times[2],
                "Said.Tarde": times[3],
                "Total": times[4]
            })

        if not registros:
            print("Nenhum registro encontrado via PyMuPDF.")
            return None

        df = pd.DataFrame(registros)[COLUNAS_NOMES]
        df = df.replace(r'^\s*$', '', regex=True)
        df["Nome"] = df["Nome"].apply(lambda x: re.sub(r'\s{2,}', ' ', str(x)).strip())

        print(f"Fallback PyMuPDF produziu {len(df)} registros.")
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
# EXECUÇÃO
# ----------------------------------------------------------------------
if __name__ == "__main__":
    df_result = extrair_tabelas(NOME_ARQUIVO_PDF)

    if df_result is None or df_result.empty:
        print("-" * 40)
        print("❌ Falha na extração. O DataFrame final está vazio ou o arquivo não foi encontrado.")
    else:
        print("-" * 40)
        print(f"✅ Total de registros extraídos: {len(df_result)}")

        try:
            df_result.to_excel(NOME_ARQUIVO_EXCEL, index=False)
            print(f"📁 Planilha salva com sucesso: {NOME_ARQUIVO_EXCEL}")
        except Exception as e:
            print(f"❌ Erro ao salvar Excel: {e}")
