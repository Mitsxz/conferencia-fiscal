import os
import openpyxl
from fpdf import FPDF
from config import Config

def gerar_relatorio_excel(conferencia_id: int, titulo: str, resumo: dict, itens: list) -> str:
    nome_arquivo = f"Relatorio_Conferencia_{conferencia_id}.xlsx"
    caminho_saida = os.path.join(Config.REPORTS_FOLDER, nome_arquivo)

    wb = openpyxl.Workbook()
    ws_resumo = wb.active
    ws_resumo.title = "Resumo"
    ws_resumo.append(["Indicador", "Valor"])
    ws_resumo.append(["Título da Conferência", titulo])
    ws_resumo.append(["Total de Registros Analisados", resumo['total_analisados']])
    ws_resumo.append(["Correspondências Exatas", resumo['correspondencias']])
    ws_resumo.append(["Divergências Encontradas", resumo['divergencias']])
    ws_resumo.append(["Ausentes nas Planilhas", resumo['ausentes_planilha']])
    ws_resumo.append(["Ausentes no Livro", resumo['ausentes_livro']])
    ws_resumo.append(["Possíveis Duplicidades", resumo['duplicidades']])
    ws_resumo.append(["Percentual de Correspondência (%)", f"{resumo['percentual_correspondencia']}%"])

    ws_div = wb.create_sheet("Divergencias_e_Duplicidades")
    ws_aus = wb.create_sheet("Registros_Ausentes")
    ws_ok = wb.create_sheet("Correspondencias_OK")

    cabecalho = ["Chave Identificação", "Categoria", "Tipo Divergência", "Detalhamento", "Status Revisão"]
    for ws in (ws_div, ws_aus, ws_ok):
        ws.append(cabecalho)

    for item in itens:
        cat = item['categoria_resultado']
        diffs = item.get('diferencas') or []
        detalhe_diff = "; ".join([f"{d['campo']}: Livro({d['livro']}) vs Planilha({d['planilha']}) -> {d['diferenca']}" for d in diffs])
        linha = [
            item['chave_identificacao'],
            cat,
            item.get('tipo_divergencia') or 'OK',
            detalhe_diff or 'Dados idênticos',
            item.get('status_revisao', 'PENDENTE')
        ]
        if cat in ['DIVERGENCIA', 'DUPLICIDADE']:
            ws_div.append(linha)
        elif cat in ['AUSENTE_PLANILHA', 'AUSENTE_LIVRO']:
            ws_aus.append(linha)
        else:
            ws_ok.append(linha)

    wb.save(caminho_saida)
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
    for m in [
        f"Total Analisados: {resumo['total_analisados']}",
        f"Correspondencias: {resumo['correspondencias']} ({resumo['percentual_correspondencia']}%)",
        f"Divergencias: {resumo['divergencias']}",
        f"Ausentes nas Planilhas: {resumo['ausentes_planilha']} | Ausentes no Livro: {resumo['ausentes_livro']}",
        f"Duplicidades: {resumo['duplicidades']}"
    ]:
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
