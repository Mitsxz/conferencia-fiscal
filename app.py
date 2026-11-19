import os
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


def detectar_mapeamento_colunas(colunas: list) -> dict:
    """Identifica automaticamente as colunas de Data, Nota, CFOP e Valor."""
    mapeamento = {}
    for col in colunas:
        c = col.lower().strip()
        if any(k in c for k in ['data', 'emissao', 'emissão', 'dt']) and 'data_emissao' not in mapeamento:
            mapeamento['data_emissao'] = col
        elif any(k in c for k in ['nota', 'nf', 'numero', 'número', 'doc']) and 'numero_nf' not in mapeamento:
            mapeamento['numero_nf'] = col
        elif 'cfop' in c and 'cfop' not in mapeamento:
            mapeamento['cfop'] = col
        elif any(k in c for k in ['valor', 'total', 'vlr', 'liquido', 'líquido', 'contabil', 'contábil']) and 'valor_total' not in mapeamento:
            mapeamento['valor_total'] = col
    return mapeamento


def agrupar_por_cfop(registros: list[dict]) -> list[dict]:
    """Soma os valores por CFOP para comparar com o Resumo do Livro Fiscal."""
    somas = defaultdict(float)
    for r in registros:
        cfop = str(r.get('cfop', '')).replace('.', '').strip()
        if cfop:
            somas[cfop] += float(r.get('valor_total', 0.0) or 0.0)
    return [{'cfop': k, 'valor_total': round(v, 2)} for k, v in somas.items()]


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_template('index.html', resultados=[], mensagem=None)

    arquivo_livro = request.files.get('livro_fiscal')
    arquivo_planilha = request.files.get('planilha_fiscal')

    if not arquivo_livro or not arquivo_planilha or not arquivo_livro.filename or not arquivo_planilha.filename:
        return render_template('index.html', resultados=[], mensagem="Envie ambos os arquivos (Livro Fiscal e Planilha Fiscal).")

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

        regs_livro_raw = ler_arquivo_tabular(caminho_livro)
        regs_plan_raw = ler_arquivo_tabular(caminho_planilha)

        map_livro = detectar_mapeamento_colunas(list(regs_livro_raw[0].keys()))
        map_plan = detectar_mapeamento_colunas(list(regs_plan_raw[0].keys()))

        regs_livro_norm = normalizar_registros(regs_livro_raw, map_livro)
        regs_plan_norm = normalizar_registros(regs_plan_raw, map_plan)

        # Filtra a planilha antes de agrupar, caso o usuário tenha informado Data Inicial/Final ou Nota
        filtro_data_inicio = request.form.get('data_inicio', '').strip()
        filtro_data_fim = request.form.get('data_fim', '').strip()
        filtro_nota = request.form.get('nota', '').strip().lstrip('0')
        filtro_cfop = request.form.get('cfop', '').strip().replace('.', '')

        if filtro_data_inicio or filtro_data_fim or filtro_nota:
            regs_plan_filtrados = []
            for r in regs_plan_norm:
                d = str(r.get('data_emissao', ''))
                n = str(r.get('numero_nf', ''))
                if filtro_data_inicio and d and d < filtro_data_inicio:
                    continue
                if filtro_data_fim and d and d > filtro_data_fim:
                    continue
                if filtro_nota and filtro_nota not in n:
                    continue
                regs_plan_filtrados.append(r)
            if regs_plan_filtrados:
                regs_plan_norm = regs_plan_filtrados

        # Verifica se o Livro veio em formato de Resumo (tem CFOP, mas não tem numero_nf)
        modo_resumo_cfop = ('numero_nf' not in map_livro) and ('cfop' in map_livro)

        if modo_resumo_cfop:
            if 'cfop' not in map_plan:
                raise FileReadError("O Livro Fiscal está em formato de Resumo por CFOP, mas a Planilha não possui coluna 'CFOP'.")
            regs_livro_final = agrupar_por_cfop(regs_livro_norm)
            regs_plan_final = agrupar_por_cfop(regs_plan_norm)
            campos_chave = ['cfop']
            campos_comparacao = ['valor_total']
        else:
            if 'numero_nf' not in map_livro or 'numero_nf' not in map_plan:
                raise FileReadError("Não foi possível identificar a coluna de Número da Nota ou CFOP nos arquivos.")
            regs_livro_final = regs_livro_norm
            regs_plan_final = regs_plan_norm
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

        # Salva a conferência no banco de dados
        titulo = f"Conferência {'(Resumo CFOP) ' if modo_resumo_cfop else ''}{datetime.now().strftime('%d/%m/%Y %H:%M')}"
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
            titulo,
            nome_livro_orig,
            caminho_livro,
            json.dumps([nome_plan_orig]),
            json.dumps({'chaves': campos_chave, 'comparacao': campos_comparacao}),
            0.05,
            resumo['total_analisados'],
            resumo['correspondencias'],
            resumo['divergencias'],
            resumo['ausentes_planilha'],
            resumo['ausentes_livro'],
            resumo['duplicidades'],
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
                conferencia_id,
                item['chave_identificacao'],
                item['categoria_resultado'],
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
            nota_item = str(base.get('numero_nf', 'Resumo CFOP' if modo_resumo_cfop else ''))
            cfop_item = str(base.get('cfop', ''))
            valor_item = float(base.get('valor_total', 0.0) or 0.0)

            if filtro_cfop and filtro_cfop != cfop_item:
                continue

            status_txt = 'OK' if item['categoria_resultado'] == 'CORRESPONDENCIA' else item['categoria_resultado']
            detalhes_txt = "; ".join([
                f"{d['campo']}: Livro (R$ {d['livro']}) vs Planilha (R$ {d['planilha']}) | Dif: R$ {d['diferenca']}"
                for d in item.get('diferencas', [])
            ]) or f"Total confere (Livro: R$ {valor_item:.2f})"

            resultados_tabela.append({
                'data': data_item or 'Período Total',
                'nota': nota_item or '-',
                'cfop': cfop_item or '-',
                'valor': valor_item,
                'status': status_txt,
                'detalhe': detalhes_txt
            })

        msg = (
            f"Conferência #{conferencia_id} concluída! "
            f"CFOPs/Itens Analisados: {resumo['total_analisados']} | "
            f"OK: {resumo['correspondencias']} | "
            f"Divergências: {resumo['divergencias']}"
        )

        return render_template('index.html', resultados=resultados_tabela, mensagem=msg, conferencia_id=conferencia_id)

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

    itens = []
    for row in itens_db:
        item = dict(row)
        item['diferencas'] = json.loads(item['diferencas_json']) if item['diferencas_json'] else []
        itens.append(item)

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
        'total_analisados': conf['total_analisados'],
        'correspondencias': conf['total_correspondencias'],
        'divergencias': conf['total_divergencias'],
        'ausentes_planilha': conf['total_ausentes_planilha'],
        'ausentes_livro': conf['total_ausentes_livro'],
        'duplicidades': conf['total_duplicidades'],
        'percentual_correspondencia': conf['percentual_correspondencia'],
    }

    itens = []
    for row in itens_db:
        item = dict(row)
        item['diferencas'] = json.loads(item['diferencas_json']) if item['diferencas_json'] else []
        itens.append(item)

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
