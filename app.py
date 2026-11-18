from flask import Flask, render_template, request, redirect, url_for
import os

app = Flask(__name__)

# Base de dados simulada ou em memória para testes rápidos no Vercel (ou adapte ao seu database.py)
# Certifique-se de que cada registro possui chaves correspondentes aos filtros:
# 'data', 'nota', 'cfop', 'valor'
DADOS_EXEMPLO = [
    {"id": 1, "data": "2026-06-01", "nota": "1001", "cfop": "5102", "valor": 150.00},
    {"id": 2, "data": "2026-06-02", "nota": "1002", "cfop": "6102", "valor": 320.50},
    {"id": 3, "data": "2026-06-03", "nota": "1003", "cfop": "5405", "valor": 89.90},
]

@app.route("/", methods=["GET", "POST"])
def index():
    resultados = DADOS_EXEMPLO
    
    # Captura dos filtros solicitados
    data_inicio = request.form.get("data_inicio") or request.args.get("data_inicio")
    data_fim = request.form.get("data_fim") or request.args.get("data_fim")
    nota_filtro = request.form.get("nota") or request.args.get("nota")
    cfop_filtro = request.form.get("cfop") or request.args.get("cfop")
    valor_min = request.form.get("valor_min") or request.args.get("valor_min")
    valor_max = request.form.get("valor_max") or request.args.get("valor_max")

    # Aplicação dos filtros restritos estritamente a: Datas, Notas, CFOPs e Valores
    if data_inicio:
        resultados = [r for r in resultados if r["data"] >= data_inicio]
    if data_fim:
        resultados = [r for r in resultados if r["data"] <= data_fim]
    if nota_filtro:
        resultados = [r for r in resultados if nota_filtro.lower() in str(r["nota"]).lower()]
    if cfop_filtro:
        resultados = [r for r in resultados if cfop_filtro.lower() in str(r["cfop"]).lower()]
    if valor_min:
        try:
            resultados = [r for r in resultados if float(r["valor"]) >= float(valor_min)]
        except ValueError:
            pass
    if valor_max:
        try:
            resultados = [r for r in resultados if float(r["valor"]) <= float(valor_max)]
        except ValueError:
            pass

    return render_template("index.html", resultados=resultados)

if __name__ == "__main__":
    app.run(debug=True)
