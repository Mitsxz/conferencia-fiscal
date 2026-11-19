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
    termos_impostos = [
        'icms', 'pis', 'cofins', 'ipi', 'iss', 'irrf', 'csll', 'inss',
        'ibs', 'cbs', 'is', 'st', 'subst', 'base', 'bc', 'aliquota', 'alíquota',
        'difal', 'fcp', 'tributo', 'imposto', 'credito', 'crédito', 'cobr',
        'frete', 'seguro', 'desconto', 'outras', 'isenta', 'tributada'
    ]
    nomes_prioritarios_valor = [
        'valor contábil', 'valor contabil', 'vlr. contábil', 'vlr contábil', 'vlr contabil',
        'valor total', 'vlr. total', 'vlr total', 'total da nota', 'total nota',
        'valor da nota', 'valor nota', 'valor líquido', 'valor liquido',
        'valor dos produtos', 'valor produtos', 'valor mercadoria', 'valor merc.', 'valor'
    ]

    for col in colunas:
        c = col.lower().strip()
        if any(k in c for k in ['data', 'emissao', 'emissão', 'dt', 'entrada', 'saida', 'saída']) and 'data_emissao' not in mapeamento:
            mapeamento['data_emissao'] = col
        elif any(k in c for k in ['nota', 'nf', 'numero', 'número', 'doc', 'nfe', 'nfc']) and 'numero_nf' not in mapeamento:
            mapeamento['numero_nf'] = col
        elif (c in ['cfop', 'código', 'codigo', 'cod', 'cód', 'cod. fiscal', 'cód. fiscal'] or 'cfop' in c) and 'cfop' not in mapeamento:
            mapeamento['cfop'] = col

    # 1ª tentativa: busca por nomes prioritários do valor da nota
    for alvo in nomes_prioritarios_valor:
        for col in colunas:
            c = col.lower().strip()
            if c == alvo or (alvo in c and not any(imp in c for imp in termos_impostos)):
                mapeamento['valor_total'] = col
                break
        if 'valor_total' in mapeamento:
            break

    # 2ª tentativa: busca genérica ignorando impostos e colunas auxiliares
    if 'valor_total' not in mapeamento:
        for col in colunas:
            c = col.lower().strip()
            if (
                any(k in c for k in ['valor', 'total', 'vlr', 'liquido', 'líquido', 'contabil', 'contábil', 'merc'])
                and not any(imp in c for imp in termos_impostos)
            ):
                mapeamento['valor_total'] = col
                break

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

        if modo_operacao == 'resumo_cfop':
            if 'cfop' not in map_livro or 'cfop' not in map_plan:
                raise FileReadError("Coluna de CFOP não encontrada em um dos arquivos.")

            somas_livro = agrupar_por_cfop(regs_livro_norm)
            somas_plan = agrupar_por_cfop(regs_plan_norm)
            todos_cfops = sorted(set(somas_livro.keys()) | set(somas_plan.keys()))

            resultados_tabela = []
            itens_db = []
            tot_livro = 0.0
            tot_plan = 0.0

            for cfop in todos_cfops:
                if filtro_cfop and filtro_cfop != cfop:
                    continue
                v_l = somas_livro.get(cfop, 0.0)
                v_p = somas_plan.get(cfop, 0.0)
                dif = round(v_l - v_p, 2)
                tot_livro = round(tot_livro + v_l, 2)
                tot_plan = round(tot_plan + v_p, 2)

                if abs(dif) <= 0.05:
                    det = f"Livro: R$ {formatar_brl(v_l)} | Planilha: R$ {formatar_brl(v_p)} (Bateu 100%)"
                else:
                    det = f"Livro: R$ {formatar_brl(v_l)} | Planilha: R$ {formatar_brl(v_p)} | Dif: R$ {formatar_brl(dif)}"

                resultados_tabela.append({
                    'data': 'Período', 'nota': 'Resumo CFOP', 'cfop': cfop,
                    'valor': formatar_brl(v_l), 'status': 'OK', 'detalhe': det
                })
                itens_db.append({
                    'chave_identificacao': f"cfop={cfop}",
                    'categoria_resultado': 'CORRESPONDENCIA',
                    'tipo_divergencia': None,
                    'dados_livro': {'cfop': cfop, 'valor_total': v_l},
                    'dados_planilha': {'cfop': cfop, 'valor_total': v_p},
                    'diferencas': [] if abs(dif) <= 0.05 else [{'campo': 'valor_total', 'livro': v_l, 'planilha': v_p, 'diferenca': dif}]
                })

            dif_g = round(tot_livro - tot_plan, 2)
            resultados_tabela.append({
                'data': 'TOTAL GERAL', 'nota': 'Soma Todos CFOPs', 'cfop': 'TOTAL',
                'valor': formatar_brl(tot_livro), 'status': 'OK',
                'detalhe': f"Total Livro: R$ {formatar_brl(tot_livro)} | Total Planilha: R$ {formatar_brl(tot_plan)} | Dif: R$ {formatar_brl(dif_g)}"
            })

            titulo = f"Resumo CFOP - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            qtd = len(itens_db)

            conn = get_db_connection()
            cur = conn.cursor()
            cur.execute('''
                INSERT INTO conferencias (
                    titulo, arquivo_livro_original, arquivo_livro_salvo,
                    arquivos_planilhas_json, config_cruzamento_json, tolerancia_valor,
                    total_analisados, total_correspondencias, total_divergencias,
                    total_ausentes_planilha, total_ausentes_livro, total_duplicidades,
                    percentual_correspondencia
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                titulo, nome_livro_orig, caminho_livro, json.dumps([nome_plan_orig]),
                json.dumps({'modo': 'resumo_cfop'}), 0.05, qtd, qtd, 0, 0, 0, 0, 100.0
            ))
            conf_id = cur.lastrowid

            for it in itens_db:
                cur.execute('''
                    INSERT INTO conferencia_itens (
                        conferencia_id, chave_identificacao, categoria_resultado,
                        tipo_divergencia, dados_livro_json, dados_planilha_json, diferencas_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    conf_id, it['chave_identificacao'], it['categoria_resultado'],
                    it['tipo_divergencia'], json.dumps(it['dados_livro']),
                    json.dumps(it['dados_planilha']), json.dumps(it['diferencas'])
                ))
            conn.commit()
            conn.close()

            msg = f"Resumo #{conf_id} concluído! CFOPs: {qtd} | Livro: R$ {formatar_brl(tot_livro)} | Planilha: R$ {formatar_brl(tot_plan)}"
            return render_template('index.html', resultados=resultados_tabela, mensagem=msg, conferencia_id=conf_id)

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

        res_cruz = executar_cruzamento(
            registros_livro=regs_livro_final,
            registros_planilha=regs_plan_final,
            campos_chave=campos_chave,
            campos_comparacao=campos_comparacao,
            tolerancia=0.05
        )

        resumo = res_cruz['resumo']
        itens = res_cruz['itens']
        titulo = f"Conferência (Detalhado por Nota) - {datetime.now().strftime('%d/%m/%Y %H:%M')}"

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute('''
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
        conf_id = cur.lastrowid

        for item in itens:
            cur.execute('''
                INSERT INTO conferencia_itens (
                    conferencia_id, chave_identificacao, categoria_resultado,
                    tipo_divergencia, dados_livro_json, dados_planilha_json, diferencas_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                conf_id, item['chave_identificacao'], item['categoria_resultado'],
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
            valor_fmt = formatar_brl(float(base.get('valor_total', 0.0) or 0.0))

            if filtro_data_inicio and data_item and data_item < filtro_data_inicio:
                continue
            if filtro_data_fim and data_item and data_item > filtro_data_fim:
                continue
            if filtro_nota and filtro_nota not in nota_item:
                continue
            if filtro_cfop and filtro_cfop != cfop_item:
                continue

            diffs = item.get('diferencas', [])
            tem_erro_valor = any(
                'valor' in str(d.get('campo', '')).lower() for d in diffs
            )

            if item['categoria_resultado'] == 'CORRESPONDENCIA':
                status_txt = 'OK'
            elif item['categoria_resultado'] == 'DIVERGENCIA':
                status_txt = 'DIVERGENCIA' if tem_erro_valor else 'DIVERGENCIA_DATA'
            else:
                status_txt = item['categoria_resultado']

            partes = []
            for d in diffs:
                if d['campo'] == 'valor_total':
                    partes.append(f"Valor: Livro (R$ {formatar_brl(d['livro'])}) vs Planilha (R$ {formatar_brl(d['planilha'])}) | Dif: R$ {formatar_brl(d['diferenca'])}")
                elif d['campo'] == 'data_emissao':
                    partes.append(f"Data: Livro ({d['livro']}) vs Planilha ({d['planilha']})")
                else:
                    partes.append(f"{d['campo']}: Livro ({d['livro']}) vs Planilha ({d['planilha']}) | Dif: {d['diferenca']}")

            resultados_tabela.append({
                'data': data_item or 'Período', 'nota': nota_item or '-',
                'cfop': cfop_item or '-', 'valor': valor_fmt,
                'status': status_txt, 'detalhe': "; ".join(partes) or f"Total confere (R$ {valor_fmt})"
            })

        msg = f"Conferência #{conf_id} concluída! Analisados: {resumo['total_analisados']} | OK: {resumo['correspondencias']} | Divergências: {resumo['divergencias']}"
        return render_template('index.html', resultados=resultados_tabela, mensagem=msg, conferencia_id=conf_id)

    except FileReadError as e:
        return render_template('index.html', resultados=[], mensagem=f"Erro ao ler arquivo: {str(e)}")
    except Exception as e:
        return render_template('index.html', resultados=[], mensagem=f"Erro inesperado: {str(e)}")


@app.route('/resultado/<int:conferencia_id>')
def ver_resultado(conferencia_id):
    conn = get_db_connection()
    conf = conn.execute('SELECT * FROM conferencias WHERE id = ?', (conferencia_id,)).fetchone()
    if not conf:
        conn.close()
        return redirect(url_for('index'))
    itens_db = conn.execute('SELECT * FROM conferencia_itens WHERE conferencia_id = ?', (conferencia_id,)).fetchall()
    conn.close()
    itens = [dict(r, diferencas=json.loads(r['diferencas_json']) if r['diferencas_json'] else []) for r in itens_db]
    return render_template('resultado.html', conf=dict(conf), itens=itens)


@app.route('/relatorio/<int:conferencia_id>/<tipo>')
def exportar_relatorio(conferencia_id, tipo):
    conn = get_db_connection()
    conf = conn.execute('SELECT * FROM conferencias WHERE id = ?', (conferencia_id,)).fetchone()
    if not conf:
        conn.close()
        return "Conferência não encontrada", 404
    itens_db = conn.execute('SELECT * FROM conferencia_itens WHERE conferencia_id = ?', (conferencia_id,)).fetchall()
    conn.close()
    resumo = {
        'total_analisados': conf['total_analisados'], 'correspondencias': conf['total_correspondencias'],
        'divergencias': conf['total_divergencias'], 'ausentes_planilha': conf['total_ausentes_planilha'],
        'ausentes_livro': conf['total_ausentes_livro'], 'duplicidades': conf['total_duplicidades'],
        'percentual_correspondencia': conf['percentual_correspondencia'],
    }
    itens = [dict(r, diferencas=json.loads(r['diferencas_json']) if r['diferencas_json'] else []) for r in itens_db]
    Config.init_app()
    if tipo == 'excel':
        caminho = gerar_relatorio_excel(conferencia_id, conf['titulo'], resumo, itens)
    elif tipo == 'pdf':
        caminho = gerar_relatorio_pdf(conferencia_id, conf['titulo'], resumo, itens)
    else:
        return "Tipo inválido", 400
    return send_file(caminho, as_attachment=True)


@app.route('/api/item/<int:item_id>/status', methods=['POST'])
def atualizar_status_item(item_id):
    dados = request.get_json(silent=True) or {}
    novo_status = dados.get('status', 'PENDENTE')
    conn = get_db_connection()
    conn.execute('UPDATE conferencia_itens SET status_revisao = ? WHERE id = ?', (novo_status, item_id))
    conn.commit()
    conn.close()
    return jsonify({'ok': True, 'status': novo_status})


if __name__ == '__main__':
    app.run(debug=True)
