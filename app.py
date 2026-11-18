from flask import Flask, render_template, request
import csv
import openpyxl

app = Flask(__name__)

def processar_arquivo(file):
    dados = []
    if not file or file.filename == '':
        return dados
    
    filename = file.filename.lower()
    
    try:
        if filename.endswith('.csv'):
            decoded_file = file.read().decode('utf-8').splitlines()
            reader = csv.DictReader(decoded_file)
            for row in reader:
                dados.append({
                    "data": row.get("data", ""),
                    "nota": str(row.get("nota", "")).strip(),
                    "cfop": str(row.get("cfop", "")).strip(),
                    "valor": float(row.get("valor", 0) or 0)
                })
        elif filename.endswith('.xlsx'):
            wb = openpyxl.load_workbook(file)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if any(row) and len(row) >= 4:
                    dados.append({
                        "data": str(row[0]) if row[0] else "",
                        "nota": str(row[1]).strip() if row[1] else "",
                        "cfop": str(row[2]).strip() if row[2] else "",
                        "valor": float(row[3] or 0) if row[3] else 0.0
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

        livro_dict = {item["nota"]: item for item in livro_dados}
        planilha_dict = {item["nota"]: item for item in planilha_dados}

        todas_notas = set(livro_dict.keys()).union(set(planilha_dict.keys()))

        for nota in todas_notas:
            item_livro = livro_dict.get(nota)
            item_planilha = planilha_dict.get(nota)

            status = "OK"
            detalhe = "Conferido com sucesso."

            if not item_livro:
                status = "Divergente"
                detalhe = "Presente apenas na Planilha."
                item_ref = item_planilha
            elif not item_planilha:
                status = "Divergente"
                detalhe = "Presente apenas no Livro Fiscal."
                item_ref = item_livro
            else:
                item_ref = item_livro
                if float(item_livro["valor"]) != float(item_planilha["valor"]):
                    status = "Divergente"
                    detalhe = f"Divergência de Valor (Livro: {item_livro['valor']} x Planilha: {item_planilha['valor']})"
                elif item_livro["cfop"] != item_planilha["cfop"]:
                    status = "Divergente"
                    detalhe = f"Divergência de CFOP (Livro: {item_livro['cfop']} x Planilha: {item_planilha['cfop']})"

            # Aplicação dos 4 filtros
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
        
        mensagem = f"Cruzamento concluído! Total de registros analisados: {len(resultados)}"

    return render_template("index.html", resultados=resultados, mensagem=mensagem)

if __name__ == "__main__":
    app.run(debug=True)
