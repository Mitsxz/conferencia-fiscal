import os
import re
import csv
import openpyxl
from pypdf import PdfReader

class FileReadError(Exception):
    pass

def ler_arquivo_tabular(caminho_arquivo: str) -> list[dict]:
    if not os.path.exists(caminho_arquivo):
        raise FileReadError(f"Arquivo não encontrado: {caminho_arquivo}")

    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    registros = []

    try:
        if extensao == '.xlsx':
            wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
            ws = wb.active
            linhas = list(ws.iter_rows(values_only=True))
            if len(linhas) < 2:
                raise FileReadError("A planilha Excel está vazia ou sem dados.")
            cabecalho = [str(c).strip() if c is not None else f"col_{i}" for i, c in enumerate(linhas[0])]
            for row in linhas[1:]:
                if not any(row):
                    continue
                registros.append({cabecalho[i]: (str(row[i]).strip() if row[i] is not None else '') for i in range(len(cabecalho))})

        elif extensao == '.csv':
            conteudo = None
            for encoding in ['utf-8-sig', 'utf-8', 'latin-1']:
                try:
                    with open(caminho_arquivo, 'r', encoding=encoding) as f:
                        conteudo = f.read()
                    break
                except UnicodeDecodeError:
                    continue
            if not conteudo:
                raise FileReadError("Não foi possível decodificar o arquivo CSV.")

            sep = ';' if conteudo.count(';') >= conteudo.count(',') else ','
            leitor = csv.DictReader(conteudo.splitlines(), delimiter=sep)
            for row in leitor:
                if any(row.values()):
                    registros.append({str(k).strip(): (str(v).strip() if v is not None else '') for k, v in row.items() if k})

        elif extensao == '.pdf':
            reader = PdfReader(caminho_arquivo)
            texto_completo = ""
            for pagina in reader.pages:
                texto_ext = pagina.extract_text()
                if texto_ext:
                    texto_completo += texto_ext + "\n"

            for linha in texto_completo.splitlines():
                linha_limpa = linha.strip()
                match = re.search(r'(\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2})\D+(\d{1,8})\D+(\d{4})\D+([\d\.,]+)', linha_limpa)
                if match:
                    data_raw, nota_val, cfop_val, valor_str = match.groups()
                    registros.append({
                        "data": data_raw,
                        "nota": nota_val,
                        "cfop": cfop_val,
                        "valor": valor_str
                    })
            if not registros:
                raise FileReadError("Nenhum lançamento fiscal (Data, Nota, CFOP, Valor) foi encontrado no PDF.")
        else:
            raise FileReadError(f"Formato não suportado ({extensao}). Envie .xlsx, .csv ou .pdf.")

        if not registros:
            raise FileReadError("O arquivo enviado está vazio ou não possui linhas válidas.")

        return registros

    except FileReadError:
        raise
    except Exception as e:
        raise FileReadError(f"Falha ao ler o arquivo '{os.path.basename(caminho_arquivo)}': {str(e)}")
