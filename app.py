from flask import Flask, render_template, request
import csv
import openpyxl
from pypdf import PdfReader
import re

app = Flask(__name__)

def processar_arquivo(file):
    dados = []
    if not file or file.filename == '':
        return dados
    
    filename = file.filename.lower()
    
    try:
        if filename.endswith('.csv') or filename.endswith('.txt'):
            decoded_file = file.read().decode('utf-8', errors='ignore').splitlines()
            reader = csv.DictReader(decoded_file)
            for row in reader:
                row_limpa = {k.strip().lower(): v for k, v in row.items() if k}
                nota_str = str(row_limpa.get("nota", "")).strip().lstrip('0') # Remove zeros à esquerda para alinhar
                if nota_str:
                    dados.append({
                        "data": str(row_limpa.get("data", "")).strip(),
                        "nota": nota_str,
                        "cfop": str(row_limpa.get("cfop", "")).strip(),
                        "valor": float(str(row_limpa.get("valor", 0)).replace("R$", "").replace(".", "").replace(",", ".").strip() or 0)
                    })
                
        elif filename.endswith('.xlsx'):
            wb = openpyxl.load_workbook(file, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if any(row) and len(row) >= 4:
                    nota_str = str(row[1]).strip().lstrip('0') if row[1] else ""
                    if nota_str:
                        val_raw = str(row[3]).replace("R$", "").replace(".", "").replace(",", ".").strip() if row[3] else "0"
                        dados.append({
                            "data": str(row[0]).strip() if row[0] else "",
                            "nota": nota_str,
                            "cfop": str(row[2]).strip() if row[2] else "",
                            "valor": float(val_raw or 0)
                        })
                    
        elif filename.endswith('.pdf'):
            reader = PdfReader(file)
            texto_completo = ""
            for pagina in reader.pages:
                texto_ext = pagina.extract_text()
                if texto_ext:
                    texto_completo += texto_ext + "\n"
            
            # Extração genérica inteligente de linhas do PDF
            linhas = texto_completo.splitlines()
            for linha in linhas:
                match = re.search(r'(\d{4}-\d{2}-\d{2})\s+(\d+)\s+(\d{4})\s+([\d\.,]+)', linha)
                if match:
                    data_val, nota_val, cfop_val, valor_str = match.groups()
                    nota_limpa = nota_val.strip().lstrip('0')
                    valor_val = float(valor_str.replace('.', '').replace(',', '.'))
                    dados.append({
                        "data": data_val,
                        "nota": nota_limpa,
                        "cfop": cfop_val.strip(),
                        "valor": valor_val
                    })

    except Exception as e:
        print(f"Erro ao processar arquivo: {e}")
        
    return dados

@app.route("/", methods=["GET", "POST"])
def index():
    resultados = []
    mensagem = None

    if request.method == "POST":
        livro_file = request.files.get("livro_fiscal")
        planilha_file = request.files.get("planilha_fiscal")
        
        data_inicio = request.form.get("data_inicio")
        data_fim = request.form.get("data_fim")
        nota_filtro = request.form.get("nota")
        cfop_filtro = request.form.get("cfop")

        livro_dados = processar_arquivo(livro_file)
        planilha_dados = processar_arquivo(planilha_file)

        livro_dict = {item["nota"]: item for item in livro_dados if item["nota"]}
        planilha_dict = {item["nota"]: item for item in planilha_dados if item["nota"]}

        todas_notas = set(livro_dict.keys()).union(set(planilha_dict.keys()))

        for nota in todas_notas:
            item_livro = livro_dict.get(nota)
            item_planilha = planilha_dict.get(nota)

            status = "OK"
            detalhe = "Conferido com sucesso. Dados idênticos."

            if not item_livro:
                status = "Divergente"
                detalhe = f"Nota {nota} presente apenas na Planilha."
                item_ref = item_planilha
            elif not item_planilha:
                status = "Divergente"
                detalhe = f"Nota {nota} presente apenas no Livro Fiscal."
                item_ref = item_livro
            else:
                item_ref = item_livro
                # Comparação tolerante de valores (arredondando para 2 casas decimais)
                val_livro = round(float(item_livro["valor"]), 2)
                val_planilha = round(float(item_planilha["valor"]), 2)

                if val_livro != val_planilha:
                    status = "Divergente"
                    detalhe = f"Divergência de Valor: Livro (R$ {val_livro}) x Planilha (R$ {val_planilha})"
                elif item_livro["cfop"].strip() != item_planilha["cfop"].strip():
                    status = "Divergente"
                    detalhe = f"Divergência de CFOP: Livro ({item_livro['cfop']}) x Planilha ({item_planilha['cfop']})"

            # Aplicação dos filtros
            incluir = True
            if data_inicio and item_ref["data"] < data_inicio:
                incluir = False
            if data_fim and item_ref["data"] > data_fim:
                incluir = False
            if nota_filtro and nota_filtro.lower() not in nota.lower():
                incluir = False
            if cfop_filtro and cfop_filtro.lower() not in item_ref["cfop"].lower():
                incluir = False

            if incluir:
                resultados.append({
                    "data": item_ref["data"],
                    "nota": nota,
                    "cfop": item_ref["cfop"],
                    "valor": item_ref["valor"],
                    "status": status,
                    "detalhe": detalhe
                })
        
        mensagem = f"Análise finalizada! Livro Lido: {len(livro_dados)} itens | Planilha Lida: {len(planilha_dados)} itens."

    return render_template("index.html", resultados=resultados, mensagem=mensagem)

if __name__ == "__main__":
    app.run(debug=True)
