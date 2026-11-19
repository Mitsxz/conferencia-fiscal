from flask import Flask, render_template, request
import csv
import openpyxl
from pypdf import PdfReader
import re

app = Flask(__name__)

def limpar_nota(val):
    if not val:
        return ""
    # Converte para string, remove '.0' se vier do Excel e remove espaços/zeros à esquerda
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s.lstrip('0')

def limpar_data(val):
    if not val:
        return ""
    # Remove a parte da hora caso venha como '2026-06-01 00:00:00'
    return str(val).split()[0].strip()

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
                nota = limpar_nota(row_limpa.get("nota", ""))
                if nota:
                    dados.append({
                        "data": limpar_data(row_limpa.get("data", "")),
                        "nota": nota,
                        "cfop": str(row_limpa.get("cfop", "")).replace('.0', '').strip(),
                        "valor": float(str(row_limpa.get("valor", 0)).replace("R$", "").replace(".", "").replace(",", ".").strip() or 0)
                    })
                
        elif filename.endswith('.xlsx'):
            wb = openpyxl.load_workbook(file, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if any(row) and len(row) >= 4:
                    nota = limpar_nota(row[1])
                    if nota:
                        val_raw = str(row[3]).replace("R$", "").replace(".", "").replace(",", ".").strip() if row[3] else "0"
                        dados.append({
                            "data": limpar_data(row[0]),
                            "nota": nota,
                            "cfop": str(row[2]).replace('.0', '').strip() if row[2] else "",
                            "valor": float(val_raw or 0)
                        })
                    
        elif filename.endswith('.pdf'):
            reader = PdfReader(file)
            texto_completo = ""
            for pagina in reader.pages:
                texto_ext = pagina.extract_text()
                if texto_ext:
                    texto_completo += texto_ext + "\n"
            
            linhas = texto_completo.splitlines()
            for linha in linhas:
                match = re.search(r'(\d{4}-\d{2}-\d{2})\s+(\d+)\s+(\d{4})\s+([\d\.,]+)', linha)
                if match:
                    data_val, nota_val, cfop_val, valor_str = match.groups()
                    dados.append({
                        "data": limpar_data(data_val),
                        "nota": limpar_nota(nota_val),
                        "cfop": cfop_val.strip(),
                        "valor": float(valor_str.replace('.', '').replace(',', '.'))
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
            detalhe = "Conferido com sucesso. Valores e dados idênticos."

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
                val_livro = round(float(item_livro["valorn"]), 2) if "valorn" in item_livro else round(float(item_livro["valor"]), 2)
                val_planilha = round(float(item_planilha["valor"]), 2)

                if val_livro != val_planilha:
                    status = "Divergente"
                    detalhe = f"Divergência de Valor: Livro (R$ {val_livro}) x Planilha (R$ {val_planilha})"
                elif item_livro["cfop"] != item_planilha["cfop"]:
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
        
        mensagem = f"Conferência realizada com sucesso! Total de notas analisadas: {len(resultados)}"

    return render_template("index.html", resultados=resultados, mensagem=mensagem)

if __name__ == "__main__":
    app.run(debug=True)
