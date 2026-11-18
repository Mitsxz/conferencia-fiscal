import re
from datetime import datetime
import pandas as pd

CAMPOS_NUMERICOS = {'quantidade', 'valor_total', 'base_calculo', 'icms', 'ipi', 'pis', 'cofins'}
CAMPOS_DATA = {'data_emissao'}
CAMPOS_DOCUMENTO = {'cnpj'}
CAMPOS_NF = {'numero_nf', 'serie'}

def normalizar_cnpj(valor: str) -> str:
    if not valor:
        return ''
    apenas_digitos = re.sub(r'\D', '', str(valor))
    return apenas_digitos.zfill(14) if len(apenas_digitos) in [11, 14] else apenas_digitos

def normalizar_numero_nf(valor: str) -> str:
    if not valor:
        return ''
    limpo = str(valor).strip()
    return limpo.lstrip('0') or '0' if limpo.isdigit() else limpo.upper()

def normalizar_data(valor: str) -> str:
    if not valor:
        return ''
    texto = str(valor).strip().split(' ')[0]
    formatos = ['%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%d/%m/%y']
    for fmt in formatos:
        try:
            return datetime.strptime(texto, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return texto

def normalizar_valor_monetario(valor) -> float:
    if valor is None or str(valor).strip() == '':
        return 0.0
    if isinstance(valor, (int, float)):
        return round(float(valor), 4)

    texto = str(valor).strip()
    texto = re.sub(r'[^\d.,-]', '', texto)
    if not texto:
        return 0.0

    if ',' in texto and '.' in texto:
        if texto.rfind(',') > texto.rfind('.'):
            texto = texto.replace('.', '').replace(',', '.')
        else:
            texto = texto.replace(',', '')
    elif ',' in texto:
        texto = texto.replace(',', '.')

    try:
        return round(float(texto), 4)
    except ValueError:
        return 0.0

def normalizar_texto(valor: str) -> str:
    if not valor:
        return ''
    return re.sub(r'\s+', ' ', str(valor).strip()).upper()

def normalizar_dataframe(df: pd.DataFrame, mapeamento_colunas: dict) -> pd.DataFrame:
    dados_normalizados = {}

    for campo_padrao, coluna_origem in mapeamento_colunas.items():
        if not coluna_origem or coluna_origem not in df.columns:
            continue

        serie = df[coluna_origem]

        if campo_padrao in CAMPOS_DOCUMENTO:
            dados_normalizados[campo_padrao] = serie.apply(normalizar_cnpj)
        elif campo_padrao in CAMPOS_NF:
            dados_normalizados[campo_padrao] = serie.apply(normalizar_numero_nf)
        elif campo_padrao in CAMPOS_DATA:
            dados_normalizados[campo_padrao] = serie.apply(normalizar_data)
        elif campo_padrao in CAMPOS_NUMERICOS:
            dados_normalizados[campo_padrao] = serie.apply(normalizar_valor_monetario)
        else:
            dados_normalizados[campo_padrao] = serie.apply(normalizar_texto)

    return pd.DataFrame(dados_normalizados)
