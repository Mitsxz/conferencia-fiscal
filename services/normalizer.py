import re
from datetime import datetime

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

    texto = re.sub(r'[^\d.,-]', '', str(valor).strip())
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

def normalizar_registros(registros: list[dict], mapeamento_colunas: dict) -> list[dict]:
    resultado = []
    for row in registros:
        novo = {}
        for campo_padrao, coluna_origem in mapeamento_colunas.items():
            if not coluna_origem or coluna_origem not in row:
                continue
            val = row[coluna_origem]
            if campo_padrao in CAMPOS_DOCUMENTO:
                novo[campo_padrao] = normalizar_cnpj(val)
            elif campo_padrao in CAMPOS_NF:
                novo[campo_padrao] = normalizar_numero_nf(val)
            elif campo_padrao in CAMPOS_DATA:
                novo[campo_padrao] = normalizar_data(val)
            elif campo_padrao in CAMPOS_NUMERICOS:
                novo[campo_padrao] = normalizar_valor_monetario(val)
            else:
                novo[campo_padrao] = normalizar_texto(val)
        resultado.append(novo)
    return resultado
