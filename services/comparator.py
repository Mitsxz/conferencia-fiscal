from collections import defaultdict
from services.normalizer import CAMPOS_NUMERICOS

def construir_chave(row: dict, campos_chave: list) -> str:
    return " | ".join([f"{campo}:{str(row.get(campo, '')).strip()}" for campo in campos_chave])

def executar_cruzamento(
    registros_livro: list[dict],
    registros_planilha: list[dict],
    campos_chave: list,
    campos_comparacao: list,
    tolerancia: float = 0.01
) -> dict:
    agrupado_livro = defaultdict(list)
    agrupado_planilha = defaultdict(list)

    for r in registros_livro:
        agrupado_livro[construir_chave(r, campos_chave)].append(r)

    for r in registros_planilha:
        agrupado_planilha[construir_chave(r, campos_chave)].append(r)

    resultados = []
    contadores = {
        'total_analisados': 0,
        'correspondencias': 0,
        'divergencias': 0,
        'ausentes_planilha': 0,
        'ausentes_livro': 0,
        'duplicidades': 0
    }

    for chave, linhas in agrupado_planilha.items():
        if len(linhas) > 1:
            contadores['duplicidades'] += 1
            resultados.append({
                'chave_identificacao': chave,
                'categoria_resultado': 'DUPLICIDADE',
                'tipo_divergencia': 'DUPLICIDADE_PLANILHA',
                'dados_livro': None,
                'dados_planilha': linhas,
                'diferencas': [{'campo': 'Chave', 'livro': '-', 'planilha': f'{len(linhas)} ocorrências', 'diferenca': 'Registro duplicado nas planilhas'}]
            })

    for chave, linhas in agrupado_livro.items():
        if len(linhas) > 1:
            contadores['duplicidades'] += 1
            resultados.append({
                'chave_identificacao': chave,
                'categoria_resultado': 'DUPLICIDADE',
                'tipo_divergencia': 'DUPLICIDADE_LIVRO',
                'dados_livro': linhas,
                'dados_planilha': None,
                'diferencas': [{'campo': 'Chave', 'livro': f'{len(linhas)} ocorrências', 'planilha': '-', 'diferenca': 'Registro duplicado no livro fiscal'}]
            })

    todas_chaves = list(dict.fromkeys(list(agrupado_livro.keys()) + list(agrupado_planilha.keys())))
    contadores['total_analisados'] = len(todas_chaves)

    for chave in todas_chaves:
        no_livro = chave in agrupado_livro
        na_planilha = chave in agrupado_planilha

        if no_livro and not na_planilha:
            contadores['ausentes_planilha'] += 1
            resultados.append({
                'chave_identificacao': chave,
                'categoria_resultado': 'AUSENTE_PLANILHA',
                'tipo_divergencia': 'NAO_ENCONTRADO_PLANILHA',
                'dados_livro': agrupado_livro[chave][0],
                'dados_planilha': None,
                'diferencas': [{'campo': 'Existência', 'livro': 'Presente', 'planilha': 'Ausente', 'diferenca': 'Consta no Livro Fiscal, mas não nas Planilhas'}]
            })
            continue

        if na_planilha and not no_livro:
            contadores['ausentes_livro'] += 1
            resultados.append({
                'chave_identificacao': chave,
                'categoria_resultado': 'AUSENTE_LIVRO',
                'tipo_divergencia': 'NAO_ENCONTRADO_LIVRO',
                'dados_livro': None,
                'dados_planilha': agrupado_planilha[chave][0],
                'diferencas': [{'campo': 'Existência', 'livro': 'Ausente', 'planilha': 'Presente', 'diferenca': 'Consta nas Planilhas, mas não no Livro Fiscal'}]
            })
            continue

        reg_livro = agrupado_livro[chave][0]
        reg_planilha = agrupado_planilha[chave][0]
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
            resultados.append({
                'chave_identificacao': chave,
                'categoria_resultado': 'DIVERGENCIA',
                'tipo_divergencia': ', '.join([d['campo'] for d in diferencas_encontradas]),
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

    return {'resumo': contadores, 'itens': resultados}
