import os
import json
import uuid
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
from werkzeug.utils import secure_filename

from config import Config
from database import init_db, get_db_connection
from services.file_reader import ler_arquivo_tabular, obter_colunas_arquivo, FileReadError
from services.normalizer import normalizar_dataframe
from services.comparator import executar_cruzamento
from services.report_generator import gerar_relatorio_excel, gerar_relatorio_pdf

app = Flask(__name__)
app.config.from_object(Config)
Config.init_app()
init_db()

def arquivo_permitido(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS

@app.route('/')
def index():
    conn = get_db_connection()
    historico = conn.execute('SELECT * FROM conferencias ORDER BY criado_em DESC LIMIT 20').fetchall()
    conn.close()
    return render_template('index.html', historico=historico)

@app.route('/api/inspecionar-arquivos', methods=['POST'])
def inspecionar_arquivos():
    if 'livro_fiscal' not in request.files:
        return jsonify({'erro': 'Envie o arquivo do Livro Fiscal.'}), 400

    livro = request.files['livro_fiscal']
    planilhas = request.files.getlist('planilhas')

    if not livro or livro.filename == '' or not arquivo_permitido(livro.filename):
        return jsonify({'erro': 'Arquivo do Livro Fiscal inválido (.xlsx, .xls ou .csv).'}), 400

    if not planilhas or all(p.filename == '' for p in planilhas):
        return jsonify({'erro': 'Envie pelo menos uma planilha para comparação.'}), 400

    try:
        ext_livro = os.path.splitext(livro.filename)[1].lower()
        nome_salvo_livro = f"{uuid.uuid4()}{ext_livro}"
        caminho_livro = os.path.join(Config.UPLOAD_RAW_FOLDER, nome_salvo_livro)
        livro.save(caminho_livro)
        colunas_livro = obter_colunas_arquivo(caminho_livro)

        info_planilhas = []
        colunas_planilhas_unificadas = []

        for p in planilhas:
            if p.filename == '':
                continue
            if not arquivo_permitido(p.filename):
                return jsonify({'erro': f'Formato inválido na planilha: {p.filename}'}), 400

            ext_p = os.path.splitext(p.filename)[1].lower()
            nome_salvo_p = f"{uuid.uuid4()}{ext_p}"
            caminho_p = os.path.join(Config.UPLOAD_RAW_FOLDER, nome_salvo_p)
            p.save(caminho_p)

            cols = obter_colunas_arquivo(caminho_p)
            for c in cols:
                if c not in colunas_planilhas_unificadas:
                    colunas_planilhas_unificadas.append(c)

            info_planilhas.append({
                'original': secure_filename(p.filename),
                'salvo': nome_salvo_p,
                'colunas': cols
            })

        return jsonify({
            'livro': {
                'original': secure_filename(livro.filename),
                'salvo': nome_salvo_livro,
                'colunas': colunas_livro
            },
            'planilhas': info_planilhas,
            'colunas_planilhas': colunas_planilhas_unificadas
        })

    except FileReadError as e:
        return jsonify({'erro': str(e)}), 400
    except Exception as e:
        return jsonify({'erro': f'Erro inesperado ao processar arquivos: {str(e)}'}), 500

@app.route('/api/processar', methods=['POST'])
def processar_conferencia():
    dados = request.json
    titulo = dados.get('titulo') or 'Conferência Fiscal'
    livro_info = dados.get('livro')
    planilhas_info = dados.get('planilhas')
    mapeamento_livro = dados.get('mapeamento_livro', {})
    mapeamento_planilha = dados.get('mapeamento_planilha', {})
    campos_chave = dados.get('campos_chave', [])
    campos_comparacao = dados.get('campos_comparacao', [])
    tolerancia = float(dados.get('tolerancia', 0.01))

    if not campos_chave:
        return jsonify({'erro': 'Selecione pelo menos um campo como Chave de Cruzamento.'}), 400

    try:
        caminho_livro = os.path.join(Config.UPLOAD_RAW_FOLDER, livro_info['salvo'])
        df_livro_bruto = ler_arquivo_tabular(caminho_livro)
        df_livro_norm = normalizar_dataframe(df_livro_bruto, mapeamento_livro)

        dfs_planilhas = []
        for p in planilhas_info:
            caminho_p = os.path.join(Config.UPLOAD_RAW_FOLDER, p['salvo'])
            df_p_bruto = ler_arquivo_tabular(caminho_p)
            df_p_norm = normalizar_dataframe(df_p_bruto, mapeamento_planilha)
            dfs_planilhas.append(df_p_norm)

        df_planilhas_total = pd.concat(dfs_planilhas, ignore_index=True)

        resultado = executar_cruzamento(
            df_livro=df_livro_norm,
            df_planilhas=df_planilhas_total,
            campos_chave=campos_chave,
            campos_comparacao=campos_comparacao,
            tolerancia=tolerancia
        )

        resumo = resultado['resumo']
        itens = resultado['itens']

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
            livro_info['original'],
            livro_info['salvo'],
            json.dumps(planilhas_info),
            json.dumps({'chave': campos_chave, 'comparacao': campos_comparacao}),
            tolerancia,
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

        return jsonify({'conferencia_id': conferencia_id})

    except Exception as e:
        return jsonify({'erro': f'Erro durante o cruzamento: {str(e)}'}), 500

@app.route('/conferencia/<int:conferencia_id>')
def ver_resultado(conferencia_id):
    conn = get_db_connection()
    conf = conn.execute('SELECT * FROM conferencias WHERE id = ?', (conferencia_id,)).fetchone()
    if not conf:
        conn.close()
        return redirect(url_for('index'))

    itens_rows = conn.execute('SELECT * FROM conferencia_itens WHERE conferencia_id = ?', (conferencia_id,)).fetchall()
    conn.close()

    itens = []
    for r in itens_rows:
        itens.append({
            'id': r['id'],
            'chave_identificacao': r['chave_identificacao'],
            'categoria_resultado': r['categoria_resultado'],
            'tipo_divergencia': r['tipo_divergencia'],
            'dados_livro': json.loads(r['dados_livro_json']) if r['dados_livro_json'] else None,
            'dados_planilha': json.loads(r['dados_planilha_json']) if r['dados_planilha_json'] else None,
            'diferencas': json.loads(r['diferencas_json']) if r['diferencas_json'] else [],
            'status_revisao': r['status_revisao']
        })

    return render_template('resultado.html', conf=conf, itens=itens)

@app.route('/api/item/<int:item_id>/status', methods=['POST'])
def atualizar_status_item(item_id):
    novo_status = request.json.get('status')
    conn = get_db_connection()
    conn.execute('UPDATE conferencia_itens SET status_revisao = ? WHERE id = ?', (novo_status, item_id))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})

@app.route('/relatorio/<int:conferencia_id>/<formato>')
def baixar_relatorio(conferencia_id, formato):
    conn = get_db_connection()
    conf = conn.execute('SELECT * FROM conferencias WHERE id = ?', (conferencia_id,)).fetchone()
    itens_rows = conn.execute('SELECT * FROM conferencia_itens WHERE conferencia_id = ?', (conferencia_id,)).fetchall()
    conn.close()

    resumo = {
        'total_analisados': conf['total_analisados'],
        'correspondencias': conf['total_correspondencias'],
        'divergencias': conf['total_divergencias'],
        'ausentes_planilha': conf['total_ausentes_planilha'],
        'ausentes_livro': conf['total_ausentes_livro'],
        'duplicidades': conf['total_duplicidades'],
        'percentual_correspondencia': conf['percentual_correspondencia']
    }

    itens = [{
        'chave_identificacao': r['chave_identificacao'],
        'categoria_resultado': r['categoria_resultado'],
        'tipo_divergencia': r['tipo_divergencia'],
        'diferencas': json.loads(r['diferencas_json']) if r['diferencas_json'] else [],
        'status_revisao': r['status_revisao']
    } for r in itens_rows]

    if formato == 'excel':
        caminho = gerar_relatorio_excel(conferencia_id, conf['titulo'], resumo, itens)
    else:
        caminho = gerar_relatorio_pdf(conferencia_id, conf['titulo'], resumo, itens)

    return send_file(caminho, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
