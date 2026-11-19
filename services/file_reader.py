import os
import re
import csv
import openpyxl
from pypdf import PdfReader

class FileReadError(Exception):
    pass

def ler_arquivo_tabular(caminho_arquivo: str, modo_operacao: str = 'resumo_cfop') -> list[dict]:
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

            idx_cabecalho = 0
            palavras_chave_cab = ['cfop', 'código', 'codigo', 'nota', 'nf', 'valor', 'total', 'contabil', 'contábil', 'data', 'emissao', 'emissão']
            for idx, row in enumerate(linhas[:25]):
                texto_linha = " ".join([str(c).lower() for c in row if c is not None])
                if sum(1 for p in palavras_chave_cab if p in texto_linha) >= 2 or 'cfop' in texto_linha or 'código' in texto_linha or 'codigo' in texto_linha:
                    idx_cabecalho = idx
                    break

            cabecalho = [str(c).strip() if c is not None and str(c).strip() != '' else f"col_{i}" for i, c in enumerate(linhas[idx_cabecalho])]
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
            linhas_csv = [l for l in conteudo.splitlines() if l.strip()]

            idx_cab = 0
            for idx, l in enumerate(linhas_csv[:25]):
                if any(p in l.lower() for p in ['cfop', 'código', 'codigo', 'nota', 'valor', 'total', 'contabil', 'contábil']):
                    idx_cab = idx
                    break

            leitor = csv.DictReader(linhas_csv[idx_cab:], delimiter=sep)
            for row in leitor:
                if any(row.values()):
                    registros.append({str(k).strip(): (str(v).strip() if v is not None else '') for k, v in row.items() if k})

        elif extensao == '.pdf':
            reader = PdfReader(caminho_arquivo)

            if modo_operacao == 'resumo_cfop':
                paginas_alvo = reader.pages[-15:] if len(reader.pages) > 15 else reader.pages
                texto_completo = "\n".join([p.extract_text() or '' for p in paginas_alvo])
                linhas = texto_completo.splitlines()

                indice_inicio = 0
                palavras_resumo = ['RESUMO', 'DEMONSTRATIVO', 'TOTALIZACAO', 'TOTALIZAÇÃO', 'POR CFOP', 'CODIGO FISCAL', 'CÓDIGO FISCAL']
                for idx, l in enumerate(linhas):
                    if any(p in l.upper() for p in palavras_resumo):
                        indice_inicio = idx

                linhas_analise = linhas[indice_inicio:] if indice_inicio > 0 else linhas
                padrao_resumo_cfop = re.compile(r'(?:^|\s)([123567][\.\-]?\d{3})(?:\s+|\D+?)(\d{1,3}(?:\.\d{3})*,\d{2})')

                cfops_encontrados = {}
                for linha in linhas_analise:
                    linha_limpa = linha.strip()
                    if any(t in linha_limpa.upper() for t in ['TOTAL GERAL', 'SUBTOTAL', 'TOTAL DAS', 'TOTAL DOS']):
                        continue
                    match = padrao_resumo_cfop.search(linha_limpa)
                    if match:
                        cfop_raw, valor_str = match.groups()
                        cfop_limpo = re.sub(r'\D', '', cfop_raw)
                        if cfop_limpo.startswith('202'):
                            continue
                        cfops_encontrados[cfop_limpo] = {
                            "cfop": cfop_limpo,
                            "valor": valor_str
                        }

                if not cfops_encontrados:
                    linhas_validas = [l.strip() for l in linhas_analise if l.strip()]
                    amostra = " | ".join(linhas_validas[:30])
                    raise FileReadError(f"Resumo não encontrado. Texto lido do fim do PDF: {amostra[:500]}")
                registros = list(cfops_encontrados.values())

            else:
                texto_completo = "\n".join([p.extract_text() or '' for p in reader.pages])
                linhas = texto_completo.splitlines()

                # Aceita Código do Emitente opcional entre a Data de Emissão e a UF (ex: 29/10/2026 111 RJ 14.300,00 2-919)
                padrao_modelo_p1 = re.compile(
                    r'(\d{2}/\d{2}/\d{4})'                    # 1. Data de entrada (ex: 29/10/2026)
                    r'.*?'                                    # Espécie / Série / valores intermediários
                    r'(\d{1,9})\s*'                           # 2. Número da Nota (ex: 2809 ou 036)
                    r'(\d{2}/\d{2}/\d{4})\s+'                 # 3. Data de emissão (ex: 29/10/2026)
                    r'(?:\d+\s+)?'                            # Código do emitente opcional (ex: 111)
                    r'[A-Z]{2}\s+'                            # UF (ex: RJ, SP, DF)
                    r'(\d{1,3}(?:\.\d{3})*,\d{2})\s+'         # 4. Valor Contábil (ex: 14.300,00)
                    r'([123567][\.\-]?\d{3})'                 # 5. CFOP (ex: 2-919, 2-102)
                )

                for idx, linha in enumerate(linhas):
                    bloco = linha.strip()
                    if idx + 1 < len(linhas) and not re.match(r'^\d{2}/\d{2}/\d{4}', linhas[idx + 1].strip()):
                        bloco += " " + linhas[idx + 1].strip()

                    match = padrao_modelo_p1.search(bloco)
                    if match:
                        data_ent, nota_raw, data_emi, valor_str, cfop_raw = match.groups()
                        nota_limpa = nota_raw.lstrip('0') or '0'
                        cfop_limpo = re.sub(r'\D', '', cfop_raw)

                        registros.append({
                            "data": data_ent,
                            "nota": nota_limpa,
                            "cfop": cfop_limpo,
                            "valor": valor_str
                        })

                if not registros:
                    linhas_validas = [l.strip() for l in linhas if l.strip()]
                    amostra = " | ".join(linhas_validas[25:65])
                    raise FileReadError(f"Nenhuma nota capturada. Amostra: {amostra[:500]}")
        else:
            raise FileReadError(f"Formato não suportado ({extensao}). Envie .xlsx, .csv ou .pdf.")

        if not registros:
            raise FileReadError("O arquivo enviado está vazio ou não possui linhas válidas.")

        return registros

    except FileReadError:
        raise
    except Exception as e:
        raise FileReadError(f"Falha ao ler o arquivo '{os.path.basename(caminho_arquivo)}': {str(e)}")
