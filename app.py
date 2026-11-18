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
        # Se for enviado como CSV ou texto plano
        if filename.endswith('.csv') or filename.endswith('.txt'):
            decoded_file = file.read().decode('utf-8', errors='ignore').splitlines()
            reader = csv.DictReader(decoded_file)
            for row in reader:
                # Normaliza as chaves removendo espaços extras
                row_limpa = {k.strip().lower(): v for k, v in row.items() if k}
                dados.append({
                    "data": row_limpa.get("data", ""),
                    "nota": str(row_limpa.get("nota", "")).strip(),
                    "cfop": str(row_limpa.get("cfop", "")).strip(),
                    "valor": float(row_limpa.get("valor", 0) or 0)
                })
        # Se for enviado como Excel (.xlsx)
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
        print(f"Erro detalhado ao processar arquivo: {e}")
        
    return dados

@app.route("/", methods=["GET", "POST"])
def index():
    resultados = []
    mensagem = None

    if request.method == "POST":
        livro_file = request.files.get("livro_fiscal")
        planilha_file = request.files.get("planilha_fiscal")
        
        livro_dados = processar_arquivo(livro_file)
        planilha_dados = processar_arquivo(planilha_file)

        # Mensagem de diagnóstico para ver se os arquivos foram lidos
        if not livro_dados and not planilha_dados:
            mensagem = "⚠️ Nenhum dado foi lido dos arquivos. Verifique se o formato é CSV ou XLSX e se as colunas estão corretas."
        else:
            livro_dict = {item["nota"]: item for item in livro_dados if item["nota"]}
            planilha_dict = {item["nota"]: item for item in planilha_dados if item["nota"]}

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

                resultados.append({
                    "data": item_ref["data"],
                    "nota": nota,
                    "cfop": item_ref["cfop"],
                    "valor": item_ref["valor"],
                    "status": status,
                    "detalhe": detalhe
                })
            
            mensagem = f"✅ Conferência realizada! Livro: {len(livro_dados)} itens | Planilha: {len(planilha_dados)} itens."

    return render_template("index.html", resultados=resultados, mensagem=mensagem)

if __name__ == "__main__":
    app.run(debug=True)
