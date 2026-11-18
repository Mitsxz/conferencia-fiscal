import pandas as pd
from services.normalizer import CAMPOS_NUMERICOS

def construir_chave(row: pd.Series, campos_chave: list) -> str:
    partes = []
    for campo in campos_chave:
        val = str(row.get(campo, '')).strip()
        partes.append(f"{campo}:{val}")
    return " | ".join(partes)

def executar_cruzamento(
    df_livro: pd.DataFrame,
    df_planilhas: pd.DataFrame,
    campos_chave: list,
    campos_comparacao: list,
    tolerancia: float = 0.01
) -> dict:
    df_livro = df_livro.copy()
    df_planilhas = df_planilhas.copy()

    df_livro['_chave'] = df_livro.apply(lambda r: construir_chave(r, campos_chave), axis=1)
    df_planilhas['_chave'] = df_planilhas.apply(lambda r: construir_chave(r, campos_chave), axis=1)

    resultados = []
    contadores = {
        'total_analisados': 0,
        'correspondencias': 0,
        'divergencias': 0,
        'ausentes_planilha': 0,
        'ausentes_livro': 0,
        'duplicidades': 0
    }

    dup_planilha = df_planilhas[df_planilhas.duplicated(subset=['_chave'], keep=False)]
    chaves_dup_planilha = set(dup_planilha['_chave'].unique())

    dup_livro = df_livro[df_livro.duplicated(subset=['_chave'], keep=False)]
    chaves_dup_livro = set(dup_livro['_chave'].unique())

    for chave in chaves_dup_planilha:
        linhas_dup = df_planilhas[df_planilhas['_chave'] == chave].drop(columns=['_chave']).to_dict(orient='records')
        contadores['duplicidades'] += 1
        resultados.append({
            'chave_identificacao': chave,
            'categoria_resultado': 'DUPLICIDADE',
            'tipo_divergencia': 'DUPLICIDADE_PLANILHA',
            'dados_livro': None,
            'dados_planilha': linhas_dup,
            'diferencas': [{'campo': 'Chave', 'livro': '-', 'planilha': f'{len(linhas_dup)} ocorrências', 'diferenca': 'Registro duplicado nas planilhas'}]
        })

    for chave in chaves_dup_livro:
        linhas_dup = df_livro[df_livro['_chave'] == chave].drop(columns=['_chave']).to_dict(orient='records')
        contadores['duplicidades'] += 1
        resultados.append({
            'chave_identificacao': chave,
            'categoria_resultado': 'DUPLICIDADE',
            'tipo_divergencia': 'DUPLICIDADE_LIVRO',
            'dados_livro': linhas_dup,
            'dados_planilha': None,
            'diferencas': [{'campo': 'Chave', 'livro': f'{len(linhas_dup)} ocorrências', 'planilha': '-', 'diferenca': 'Registro duplicado no livro fiscal'}]
        })

    livro_unico = df_livro.drop_duplicates(subset=['_chave'], keep='first').set_index('_chave')
    planilha_unica = df_planilhas.drop_duplicates(subset=['_chave'], keep='first').set_index('_chave')

    todas_chaves = list(dict.fromkeys(list(livro_unico.index) + list(planilha_unica.index)))
    contadores['total_analisados'] = len(todas_chaves)

    for chave in todas_chaves:
        no_livro = chave in livro_unico.index
        na_planilha = chave in planilha_unica.index

        if no_livro and not na_planilha:
            reg_livro = livro_unico.loc[chave].to_dict()
            contadores['ausentes_planilha'] += 1
            resultados.append({
                'chave_identificacao': chave,
                'categoria_resultado': 'AUSENTE_PLANILHA',
                'tipo_divergencia': 'NAO_ENCONTRADO_PLANILHA',
                'dados_livro': reg_livro,
                'dados_planilha': None,
                'diferencas': [{'campo': 'Existência', 'livro': 'Presente', 'planilha': 'Ausente', 'diferenca': 'Consta no Livro Fiscal, mas não nas Planilhas'}]
            })
            continue

        if na_planilha and not no_livro:
            reg_planilha = planilha_unica.loc[chave].to_dict()
            contadores['ausentes_livro'] += 1
            resultados.append({
                'chave_identificacao': chave,
                'categoria_resultado': 'AUSENTE_LIVRO',
                'tipo_divergencia': 'NAO_ENCONTRADO_LIVRO',
                'dados_livro': None,
                'dados_planilha': reg_planilha,
                'diferencas': [{'campo': 'Existência', 'livro': 'Ausente', 'planilha': 'Presente', 'diferenca': 'Consta nas Planilhas, mas não no Livro Fiscal'}]
            })
            continue

        reg_livro = livro_unico.loc[chave].to_dict()
        reg_planilha = planilha_unica.loc[chave].to_dict()
        diferencas_encontradas = []

        for campo in campos_comparacao:
            if campo not in reg_livro or campo not in reg_planilha:
                continue

            val_l = reg_livro[campo]
            val_p = reg_planilha[campo]

            if campo in CAMPOS_NUMERICOS:
                diff = round(float(val_l) - float(val_p), 4)
                if abs(diff) > tolerancia:
                    diferencas_encontradas.append({
                        'campo': campo.upper(),
                        'livro': float(val_l),
                        'planilha': float(val_p),
                        'diferenca': diff
                    })
            else:
                if str(val_l).strip() != str(val_p).strip():
                    diferencas_encontradas.append({
                        'campo': campo.upper(),
                        'livro': str(val_l),
                        'planilha': str(val_p),
                        'diferenca': 'Texto/Data diferente'
                    })

        if diferencas_encontradas:
            contadores['divergencias'] += 1
            tipos = ', '.join([d['campo'] for d in diferencas_encontradas])
            resultados.append({
                'chave_identificacao': chave,
                'categoria_resultado': 'DIVERGENCIA',
                'tipo_divergencia': tipos,
                'dados_livro': reg_livro,
                'dados_planilha': reg_planilha,
                'diferencas': diferencas_encontradas
            })
        else:
            contadores['correspondencias'] += 1
            resultados.append({
                'chave_identificacao': chave,
                'categoria_resultado': 'CORRESPONDENCIA',
                'tipo_divergencia': None,
                'dados_livro': reg_livro,
                'dados_planilha': reg_planilha,
                'diferencas': []
            })

    total = contadores['total_analisados']
    contadores['percentual_correspondencia'] = round((contadores['correspondencias'] / total * 100), 2) if total > 0 else 0.0

    return {
        'resumo': contadores,
        'itens': resultados
    }
