import os
import re
import json
import uuid
from collections import defaultdict
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify
from werkzeug.utils import secure_filename

from config import Config
from database import init_db, get_db_connection
from services.file_reader import ler_arquivo_tabular, FileReadError
from services.normalizer import normalizar_registros
from services.comparator import executar_cruzamento
from services.report_generator import gerar_relatorio_excel, gerar_relatorio_pdf

app = Flask(__name__)
app.config.from_object(Config)

Config.init_app()
init_db()


def formatar_brl(valor) -> str:
    try:
        if isinstance(valor, str):
            v = valor.strip().replace('R$', '').strip()
            if ',' in v and '.' in v:
                v = v.replace('.', '').replace(',', '.')
            elif ',' in v:
                v = v.replace(',', '.')
            num = float(v)
        else:
            num = float(valor or 0.0)
        return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(valor)


def detectar_mapeamento_colunas(colunas: list) -> dict:
    mapeamento = {}
    for col in colunas:
        c = col.lower().strip()
        if any(k in c for k in ['data', 'emissao', 'emissão', 'dt', 'entrada', 'saida', 'saída']) and 'data_emissao' not in mapeamento:
            mapeamento['data_emissao'] = col
        elif any(k in c for k in ['nota', 'nf', 'numero', 'número', 'doc', 'nfe', 'nfc']) and 'numero_nf' not in mapeamento:
            mapeamento['numero_nf'] = col
        elif (c in ['cfop', 'código', 'codigo', 'cod', 'cód', 'cod. fiscal', 'cód. fiscal'] or 'cfop' in c) and 'cfop' not in mapeamento:
            mapeamento['cfop'] = col
        elif any(k in c for k in ['valor', 'total', 'vlr', 'liquido', 'líquido', 'contabil', 'contábil', 'mercadoria']) and 'valor_total' not in mapeamento:
            mapeamento['valor_total'] = col
    return mapeamento


def agrupar_por_cfop(registros: list[dict]) -> dict[str, float]:
    somas = defaultdict(float)
    for r in registros:
        cfop = re.sub(r'\D', '', str(r.get('cfop', '')).strip())
        if cfop:
            somas[cfop] = round(somas[cfop] + float(r.get('valor_total', 0.0) or 0.0), 2)
    return dict(somas)


def agrupar_por_nota_cfop(registros: list[dict]) -> list[dict]:
    agrupados = {}
    for r in registros:
        nota = str(r.get('numero_nf', '')).strip().lstrip('0') or '0'
        cfop = re.sub(r'\D', '', str(r.get('cfop', '')).strip())
        chave = (nota, cfop)
        if chave not in agrupados:
            agrupados[chave] = {
                'data_emissao': r.get('data_emissao', ''),
                'numero_nf': nota,
                'cfop': cfop,
                'valor_total': 0.0
            }
        agrupados[chave]['valor_total'] = round(
            agrupados[chave]['valor_total'] + float(r.get('valor_total', 0.0) or 0.0), 2
        )
    return list(agrupados.values())


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_template('index.html', resultados=[], mensagem=None)

    modo_operacao = request.form.get('modo_operacao', 'resumo_cfop')
    arquivo_livro = request.files.get('livro_fiscal')
    arquivo_planilha = request.files.get('planilha_fiscal')

    if not arquivo_livro or not arquivo_planilha or not arquivo_livro.filename or not arquivo_planilha.filename:
        return render_template('index.html', resultados=[], mensagem="Envie ambos os arquivos.")

    try:
        Config.init_app()
        init_db()

        nome_livro_orig = secure_filename(arquivo_livro.filename)
        nome_plan_orig = secure_filename(arquivo_planilha.filename)
        prefixo = uuid.uuid4().hex[:8]

        caminho_livro = os.path.join(Config.UPLOAD_RAW_FOLDER, f"{prefixo}_{nome_livro_orig}")
        caminho_planilha = os.path.join(Config.UPLOAD_RAW_FOLDER, f"{prefixo}_{nome_plan_orig}")

        arquivo_livro.save(caminho_livro)
        arquivo_planilha.save(caminho_planilha)

        regs_livro_raw = ler_arquivo_tabular(caminho_livro, modo_operacao=modo_operacao)
        regs_plan_raw = ler_arquivo_tabular(caminho_planilha, modo_operacao=modo_operacao)

        map_livro = detectar_mapeamento_colunas(list(regs_livro_raw[0].keys()))
        map_plan = detectar_mapeamento_colunas(list(regs_plan_raw[0].keys()))

        regs_livro_norm = normalizar_registros(regs_livro_raw, map_livro)
        regs_plan_norm = normalizar_registros(regs_plan_raw, map_plan)

        filtro_data_inicio = request.form.get('data_inicio', '').strip()
        filtro_data_fim = request.form.get('data_fim', '').strip()
        filtro_nota = request.form.get('nota', '').strip().lstrip('0')
        filtro_cfop = re.sub(r'\D', '', request.form.get('cfop', '').strip())

        # =====================================================================
        # MODO 1: RESUMO CONSOLIDADO POR CFOP
        # =====================================================================
        if modo_operacao == 'resumo_cfop':
            if 'cfop' not in map_livro or 'cfop' not in map_plan:
                raise FileReadError("Coluna de CFOP não encontrada em um dos arquivos.")

            somas_livro = agrupar_por_cfop(regs_livro_norm)
            somas_plan = agrupar_por_cfop(regs_plan_norm)
            todos_cfops = sorted(set(somas_livro.keys()) | set(somas_plan.keys()))

            resultados_tabela = []
            itens_db_para_salvar = []
            total_geral_livro = 0.0
            total_geral_plan = 0.0

            for cfop in todos_cfops:
                if filtro_cfop and filtro_cfop != cfop:
                    continue

                v_livro = somas_livro.get(cfop, 0.0)
                v_plan = somas_plan.get(cfop, 0.0)
                dif = round(v_livro - v_plan, 2)

                total_geral_livro = round(total_geral_livro + v_livro, 2)
                total_geral_plan = round(total_geral_plan + v_plan, 2)

                if abs(dif) <= 0.05:
                    detalhe_str = f"Livro: R$ {formatar_brl(v_livro)} | Planilha: R$ {formatar_brl(v_plan)} (Bateu 100%)"
                else:
                    detalhe_str = f"Livro: R$ {formatar_brl(v_livro)} | Planilha: R$ {formatar_brl(v_plan)} | Dif: R$ {formatar_brl(dif)}"

                resultados_tabela.append({
                    'data': 'Período',
                    'nota': 'Resumo CFOP',
                    'cfop': cfop,
                    'valor': formatar_brl(v_livro),
                    'status': 'OK',
                    'detalhe': detalhe_str
                })

                itens_db_para_salvar.append({
                    'chave_identificacao': f"cfop={cfop}",
                    'categoria_resultado': 'CORRESPONDENCIA',
                    'tipo_divergencia': None,
                    'dados_livro': {'cfop': cfop, 'valor_total': v_livro},
                    'dados_planilha': {'cfop': cfop, 'valor_total': v_plan},
                    'diferencas': [] if abs(dif) <= 0.05 else [{'campo': 'valor_total', 'livro': v_livro, 'planilha': v_plan, 'diferenca': dif}]
                })

            dif_geral = round(total_geral_livro - total_geral_plan, 2)
            resultados_tabela.append({
                'data': 'TOTAL GERAL',
                'nota': 'Soma Todos CFOPs',
                'cfop': 'TOTAL',
                'valor': formatar_brl(total_geral_livro),
                'status': 'OK',
                'detalhe': f"Total Livro: R$ {formatar_brl(total_geral_livro)} | Total Planilha: R$ {formatar_brl(total_geral_plan)} | Dif: R$ {formatar_brl(dif_geral)}"
            })

            titulo = f"Resumo CFOP - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            qtd_cfops = len(itens_db_para_salvar)

            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO conferencias (
                    titulo, arquivo_livro_original, arquivo_livro_salvo,
                    arquivos_planilhas_json, config_cruzamento_json, tolerancia_valor,
                    total_analisados, total_correspondencias, total_divergencias,
                    total_ausentes_planilha, total_ausentes_livro, total_duplicidades,
                    percentual_correspondencia
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                titulo, nome_livro_orig, caminho_livro, json.dumps([nome_plan_orig]),
                json.dumps({'modo': 'resumo_cfop'}), 0.05,
                qtd_cfops, qtd_cfops, 0, 0, 0, 0, 100.0
            ))
            conferencia_id = cursor.lastrowid

            for item in itens_db_para_salvar:
                cursor.execute('''
                    INSERT INTO conferencia_itens (
                        conferencia_id, chave_identificacao, categoria_resultado,
                        tipo_divergencia, dados_livro_json, dados_planilha_json, diferencas_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    conferencia_id, item['chave_identificacao'], item['categoria_resultado'],
                    item['tipo_divergencia'], json.dumps(item['dados_livro']),
                    json.dumps(item['dados_planilha']), json.dumps(item['diferencas'])
                ))
            conn.commit()
            conn.close()

            msg = (
                f"Resumo #{conferencia_id} gerado com sucesso! "
                f"CFOPs consolidados: {qtd_cfops} | "
                f"Total Livro: R$ {formatar_brl(total_geral_livro)} | "
                f"Total Planilha: R$ {formatar_brl(total_geral_plan)}"
            )
            return render_template('index.html', resultados=resultados_tabela, mensagem=msg, conferencia_id=conferencia_id)

        # =====================================================================
        # MODO 2: DETALHADO POR NOTA FISCAL (CRUZAMENTO ITEM A ITEM)
        # =====================================================================
        if 'numero_nf' not in map_livro or 'numero_nf' not in map_plan:
            raise FileReadError("Coluna de Nota não encontrada em um dos arquivos.")

        regs_livro_final = agrupar_por_nota_cfop(regs_livro_norm)
        regs_plan_final = agrupar_por_nota_cfop(regs_plan_norm)
        campos_chave = ['numero_nf']
        if 'cfop' in map_livro and 'cfop' in map_plan:
            campos_chave.append('cfop')
        campos_comparacao = [
            c for c in ['valor_total', 'data_emissao', 'cfop']
            if c in map_livro and c in map_plan and c not in campos_chave
        ]

        resultado_cruzamento = executar_cruzamento(
            registros_livro=regs_livro_final,
            registros_planilha=regs_plan_final,
            campos_chave=campos_chave,
            campos_comparacao=campos_comparacao,
            tolerancia=0.05
        )

        resumo = resultado_cruzamento['resumo']
        itens = resultado_cruzamento['itens']
        titulo = f"Conferência (Detalhado por Nota) - {datetime.now().strftime('%d/%m/%Y %H:%M')}"

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO conferencias (
                titulo, arquivo_livro_original, arquivo_livro_salvo,
                arquivos_planilhas_json, config_cruzamento_json, tolerancia_valor,
                total_analisados, total_correspondencias, total_divergencias,
                total_ausentes_planilha, total_ausentes_livro, total_duplicidades,
                percentual_correspondencia
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            titulo, nome_livro_orig, caminho_livro, json.dumps([nome_plan_orig]),
            json.dumps({'modo': modo_operacao, 'chaves': campos_chave, 'comparacao': campos_comparacao}),
            0.05, resumo['total_analisados'], resumo['correspondencias'], resumo['divergencias'],
            resumo['ausentes_planilha'], resumo['ausentes_livro'], resumo['duplicidades'],
            resumo['percentual_correspondencia']
        ))
        conferencia_id = cursor.lastrowid

        for item in itens:
            cursor.execute('''
                INSERT INTO conferencia_itens (
                    conferencia_id, chave_identificacao, categoria_resultado,
                    tipo_divergencia, dados_livro_json, dados_planilha_json, diferencas_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                conferencia_id, item['chave_identificacao'], item['categoria_resultado'],
                item['tipo_divergencia'],
                json.dumps(item['dados_livro']) if item['dados_livro'] else None,
                json.dumps(item['dados_planilha']) if item['dados_planilha'] else None,
                json.dumps(item['diferencas'])
            ))

        conn.commit()
        conn.close()

        resultados_tabela = []
        for item in itens:
            base = item['dados_livro'] if isinstance(item['dados_livro'], dict) else (
                item['dados_planilha'] if isinstance(item['dados_planilha'], dict) else {}
            )
            if isinstance(item['dados_livro'], list) and item['dados_livro']:
                base = item['dados_livro'][0]
            elif isinstance(item['dados_planilha'], list) and item['dados_planilha']:
                base = item['dados_planilha'][0]

            data_item = str(base.get('data_emissao', ''))
            nota_item = str(base.get('numero_nf', ''))
            cfop_item = re.sub(r'\D', '', str(base.get('cfop', '')))
            valor_item = float(base.get('valor_total', 0.0) or 0.0)
            valor_fmt = formatar_brl(valor_item)

            if filtro_data_inicio and data_item and data_item < filtro_data_inicio:
                continue
            if filtro_data_fim and data_item and data_item > filtro_data_fim:
                continue
            if filtro_nota and filtro_nota not in nota_item:
                continue
            if filtro_cfop and filtro_cfop != cfop_item:
                continue

            diffs = item.get('diferencas', [])
            campos_divergentes = {d.get('campo') for d in diffs}

            if item['categoria_resultado'] == 'CORRESPONDENCIA':
                status_txt = 'OK'
            elif item['categoria_resultado'] == 'DIVERGENCIA':
                if campos_divergentes == {'data_emissao'}:
                    status_txt = 'DIVERGENCIA_DATA'
                else:
                    status_txt = 'DIVERGENCIA'
            else:
                status_txt = item['categoria_resultado']

            partes_detalhe = []
