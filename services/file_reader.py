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

            # Procura automaticamente em qual linha está o cabeçalho real (até a linha 15)
            idx_cabecalho = 0
            palavras_chave_cab = ['cfop', 'código', 'codigo', 'nota', 'nf', 'valor', 'total', 'contabil', 'contábil', 'data', 'emissao', 'emissão']
            for idx, row in enumerate(linhas[:15]):
                texto_linha = " ".join([str(c).lower() for c in row if c is not None])
                if any(p in texto_linha for p in palavras_chave_cab):
                    idx_cabecalho = idx
                    break

            cabecalho = [str(c).strip() if c is not None else f"col_{i}" for i, c in enumerate(linhas[idx_cabecalho])]
            for row in linhas[idx_cabecalho + 1:]:
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
            # Lê de trás para frente ou as últimas 15 páginas (onde fica o Resumo de Entradas/Saídas)
            paginas_alvo = reader.pages[-15:] if len(reader.pages) > 15 else reader.pages
            texto_paginas = []
            for pagina in paginas_alvo:
                txt = pagina.extract_text()
                if txt:
                    texto_paginas.append(txt)

            texto_completo = "\n".join(texto_paginas)
            linhas = texto_completo.splitlines()

            # Tenta localizar o início do quadro de Resumo
            indice_inicio = 0
            palavras_resumo = ['RESUMO', 'DEMONSTRATIVO', 'TOTALIZACAO', 'TOTALIZAÇÃO', 'POR CFOP', 'CODIGO FISCAL', 'CÓDIGO FISCAL']
            for idx, l in enumerate(linhas):
                l_up = l.upper()
                if any(p in l_up for p in palavras_resumo):
                    indice_inicio = idx  # Pega a última ocorrência (geralmente no fim do livro)

            linhas_analise = linhas[indice_inicio:] if indice_inicio > 0 else linhas

            # 1ª Tentativa: Extrair linhas de Resumo por CFOP (Ex: "1.102 ... 15.420,50" ou "5102 ... 15.420,50")
            # Captura CFOP iniciando em 1, 2, 3, 5, 6 ou 7 (com ou sem ponto: 1102 ou 1.102) seguido de valor monetário
            padrao_resumo_cfop = re.compile(r'(?:^|\s)([123567]\.?\d{3})(?:\s+|\D+?)(\d{1,3}(?:\.\d{3})*,\d{2})')

            cfops_encontrados = {}
            for linha in linhas_analise:
                linha_limpa = linha.strip()
                # Ignora linhas de totais gerais para não duplicar
                if any(t in linha_limpa.upper() for t in ['TOTAL GERAL', 'SUBTOTAL', 'TOTAL DAS', 'TOTAL DOS']):
                    continue

                match = padrao_resumo_cfop.search(linha_limpa)
                if match:
                    cfop_raw, valor_str = match.groups()
                    cfop_limpo = cfop_raw.replace('.', '')
                    # Ignora anos (ex: 2024, 2025, 2026) caso apareçam no cabeçalho
                    if cfop_limpo.startswith('202'):
                        continue
                    cfops_encontrados[cfop_limpo] = {
                        "cfop": cfop_limpo,
                        "valor": valor_str
                    }

            if cfops_encontrados:
                registros = list(cfops_encontrados.values())
            else:
                # 2ª Tentativa (fallback): Caso seja um PDF simples nota a nota
                for linha in linhas:
                    match_nota = re.search(r'(\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2})\D+(\d{1,8})\D+(\d{4})\D+([\d\.,]+)', linha.strip())
                    if match_nota:
                        data_raw, nota_val, cfop_val, valor_str = match_nota.groups()
                        registros.append({
                            "data": data_raw,
                            "nota": nota_val,
                            "cfop": cfop_val,
                            "valor": valor_str
                        })

            if not registros:
                raise FileReadError("Não foi possível localizar o quadro de Resumo por CFOP no final do PDF.")
        else:
            raise FileReadError(f"Formato não suportado ({extensao}). Envie .xlsx, .csv ou .pdf.")

        if not registros:
            raise FileReadError("O arquivo enviado está vazio ou não possui linhas válidas.")

        return registros

    except FileReadError:
        raise
    except Exception as e:
        raise FileReadError(f"Falha ao ler o arquivo '{os.path.basename(caminho_arquivo)}': {str(e)}")
