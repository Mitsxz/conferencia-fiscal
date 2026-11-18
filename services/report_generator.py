import os
import pandas as pd
from fpdf import FPDF
from config import Config

def gerar_relatorio_excel(conferencia_id: int, titulo: str, resumo: dict, itens: list) -> str:
    nome_arquivo = f"Relatorio_Conferencia_{conferencia_id}.xlsx"
    caminho_saida = os.path.join(Config.REPORTS_FOLDER, nome_arquivo)

    df_resumo = pd.DataFrame([
        {'Indicador': 'Título da Conferência', 'Valor': titulo},
        {'Indicador': 'Total de Registros Analisados', 'Valor': resumo['total_analisados']},
        {'Indicador': 'Correspondências Exatas', 'Valor': resumo['correspondencias']},
        {'Indicador': 'Divergências Encontradas', 'Valor': resumo['divergencias']},
        {'Indicador': 'Ausentes nas Planilhas (Só no Livro)', 'Valor': resumo['ausentes_planilha']},
        {'Indicador': 'Ausentes no Livro (Só nas Planilhas)', 'Valor': resumo['ausentes_livro']},
        {'Indicador': 'Possíveis Duplicidades', 'Valor': resumo['duplicidades']},
        {'Indicador': 'Percentual de Correspondência (%)', 'Valor': f"{resumo['percentual_correspondencia']}%"},
    ])

    linhas_divergencias = []
    linhas_ausentes = []
    linhas_correspondencias = []

    for item in itens:
        cat = item['categoria_resultado']
        chave = item['chave_identificacao']
        diffs = item.get('diferencas') or []
        detalhe_diff = "; ".join([f"{d['campo']}: Livro({d['livro']}) vs Planilha({d['planilha']}) -> {d['diferenca']}" for d in diffs])

        base_row = {
            'Chave Identificação': chave,
            'Categoria': cat,
            'Tipo Divergência': item.get('tipo_divergencia') or 'OK',
            'Detalhamento': detalhe_diff or 'Dados idênticos',
            'Status Revisão': item.get('status_revisao', 'PENDENTE')
        }

        if cat in ['DIVERGENCIA', 'DUPLICIDADE']:
            linhas_divergencias.append(base_row)
        elif cat in ['AUSENTE_PLANILHA', 'AUSENTE_LIVRO']:
            linhas_ausentes.append(base_row)
        else:
            linhas_correspondencias.append(base_row)

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
        pd.DataFrame(linhas_divergencias).to_excel(writer, sheet_name='Divergencias_e_Duplicidades', index=False)
        pd.DataFrame(linhas_ausentes).to_excel(writer, sheet_name='Registros_Ausentes', index=False)
        pd.DataFrame(linhas_correspondencias).to_excel(writer, sheet_name='Correspondencias_OK', index=False)

    return caminho_saida

def gerar_relatorio_pdf(conferencia_id: int, titulo: str, resumo: dict, itens: list) -> str:
    nome_arquivo = f"Relatorio_Conferencia_{conferencia_id}.pdf"
    caminho_saida = os.path.join(Config.REPORTS_FOLDER, nome_arquivo)

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "Relatorio de Conferencia e Comparacao Fiscal", ln=True)

    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 8, f"Conferencia #{conferencia_id}: {titulo}".encode('latin-1', 'replace').decode('latin-1'), ln=True)
    pdf.ln(4)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "1. Resumo Executivo", ln=True)
    pdf.set_font("Helvetica", "", 10)
    metricas = [
        f"Total Analisados: {resumo['total_analisados']}",
        f"Correspondencias: {resumo['correspondencias']} ({resumo['percentual_correspondencia']}%)",
        f"Divergencias: {resumo['divergencias']}",
        f"Ausentes nas Planilhas: {resumo['ausentes_planilha']} | Ausentes no Livro: {resumo['ausentes_livro']}",
        f"Duplicidades: {resumo['duplicidades']}"
    ]
    for m in metricas:
        pdf.cell(0, 6, m, ln=True)

    pdf.ln(6)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "2. Detalhamento de Divergencias e Ausencias", ln=True)
    pdf.set_font("Helvetica", "", 9)

    problemas = [i for i in itens if i['categoria_resultado'] != 'CORRESPONDENCIA']
    if not problemas:
        pdf.cell(0, 7, "Nenhuma divergencia ou ausencia encontrada.", ln=True)
    else:
        for p in problemas[:100]:
            texto = f"[{p['categoria_resultado']}] {p['chave_identificacao']} -> {p.get('tipo_divergencia', '')}"
            pdf.multi_cell(0, 6, texto.encode('latin-1', 'replace').decode('latin-1'))

    pdf.output(caminho_saida)
    return caminho_saida
