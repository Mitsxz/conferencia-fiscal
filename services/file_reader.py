import os
import pandas as pd
import pdfplumber

class FileReadError(Exception):
    pass

def ler_arquivo_tabular(caminho_arquivo: str) -> pd.DataFrame:
    if not os.path.exists(caminho_arquivo):
        raise FileReadError(f"Arquivo não encontrado: {caminho_arquivo}")

    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    try:
        if extensao in ['.xlsx', '.xls']:
            df = pd.read_excel(caminho_arquivo, dtype=str)
        elif extensao == '.csv':
            df = None
            for encoding in ['utf-8-sig', 'utf-8', 'latin-1']:
                for sep in [';', ',']:
                    try:
                        temp_df = pd.read_csv(caminho_arquivo, sep=sep, encoding=encoding, dtype=str)
                        if len(temp_df.columns) > 1:
                            df = temp_df
                            break
                    except Exception:
                        continue
                if df is not None:
                    break
            if df is None:
                df = pd.read_csv(caminho_arquivo, sep=None, engine='python', dtype=str)
        elif extensao == '.pdf':
            linhas_extraidas = []
            cabecalho = None
            with pdfplumber.open(caminho_arquivo) as pdf:
                for pagina in pdf.pages:
                    tabelas = pagina.extract_tables()
                    for tabela in tabelas:
                        for linha in tabela:
                            linha_limpa = [str(cel).replace('\n', ' ').strip() if cel is not None else '' for cel in linha]
                            if not any(linha_limpa):
                                continue
                            if cabecalho is None:
                                cabecalho = linha_limpa
                            elif linha_limpa != cabecalho and len(linha_limpa) == len(cabecalho):
                                linhas_extraidas.append(linha_limpa)
            if not cabecalho or not linhas_extraidas:
                raise FileReadError("Nenhuma tabela estruturada foi encontrada dentro do PDF.")
            df = pd.DataFrame(linhas_extraidas, columns=cabecalho)
        else:
            raise FileReadError(f"Formato não suportado ({extensao}). Envie .xlsx, .xls, .csv ou .pdf.")

        df.columns = [str(col).strip() for col in df.columns]
        df.dropna(how='all', inplace=True)

        if df.empty:
            raise FileReadError("O arquivo enviado está vazio ou não possui linhas válidas.")

        return df.fillna('')

    except FileReadError:
        raise
    except Exception as e:
        raise FileReadError(f"Falha ao ler o arquivo '{os.path.basename(caminho_arquivo)}': {str(e)}")

def obter_colunas_arquivo(caminho_arquivo: str) -> list:
    df = ler_arquivo_tabular(caminho_arquivo)
    return list(df.columns)
