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


def detectar_mapeamento_colunas(colunas: list) -> dict:
    """Identifica automaticamente as colunas de Data, Nota, CFOP (incluindo 'Código') e Valor."""
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


def agrupar_por_cfop(registros: list[dict]) -> list[dict]:
    """Consolida e soma os valores por CFOP (Modo 1: Resumo por CFOP)."""
    somas = defaultdict(float)
    for r in registros:
        cfop = re.sub(r'\D', '', str(r.get('cfop', '')).strip())
        if cfop:
            somas[cfop] += float(r.get('valor_total', 0.0) or 0.0)
    return [{'cfop': k, 'valor_total': round(v, 2)} for k, v in somas.items()]


def agrupar_por_nota_cfop(registros: list[dict]) -> list[dict]:
    """Soma lançamentos repetidos da mesma Nota Fiscal + CFOP para evitar falsos alertas de duplicidade (Modo 2)."""
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
            if 'cfop' not in map_livro:
                raise FileReadError(f"Coluna de CFOP não encontrada no Livro. Colunas lidas: {list(regs_livro_raw[0].keys())}")
            if 'cfop' not in map_plan:
                raise FileReadError(f"Coluna de CFOP ('Código') não encontrada na Planilha. Colunas lidas: {list(regs_plan_raw[0].keys())}")

            regs_livro_final = agrupar_por_cfop(regs_livro_norm)
            regs_plan_final = agrupar_por_cfop(regs_plan_norm)
            campos_chave = ['cfop']
            campos_comparacao = ['valor_total']
        else:
            if 'numero_nf' not in map_livro:
                raise FileReadError(f"Coluna de Nota não encontrada no Livro. Colunas lidas: {list(regs_livro_raw[0].keys())}")
            if 'numero_nf' not in map_plan:
                raise FileReadError(f"Coluna de Nota não encontrada na Planilha. Colunas lidas: {list(regs_plan_raw[0].keys())}")

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

        label_modo = "Resumo CFOP" if modo_operacao == 'resumo_cfop' else "Detalhado por Nota"
        titulo = f"Conferência ({label_modo}) - {datetime.now().strftime('%d/%m/%Y %H:%M')}"

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
            json.dumps({'modo': modo_operacao, 'chaves': campos_chave, 'comparacao
