import os
import re
import io
import csv
import json
import ssl
import base64
import datetime
import unicodedata
import urllib.request
from flask import Flask, render_template, request, jsonify
import openpyxl
from pypdf import PdfReader

app = Flask(__name__)


DRIVE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxannlE00ind3R4pjqSfkAonVfo38D2jSmK_7q3RScaYGNxbgVh3vB34ikp9dbm2_hy/exec"
DRIVE_TOKEN = "fiscal2026"


def normalizar_texto(txt):
    if txt is None:
        return ""
    s = unicodedata.normalize("NFD", str(txt))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.upper().strip()


def limpar_numero_nota(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = re.sub(r"\D", "", s).lstrip("0")
    return digits or "0"


def limpar_cfop(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = re.sub(r"\D", "", s)
    return digits[:4] if len(digits) >= 4 else digits


def parse_valor_br(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return round(float(val), 2)
    s = str(val).strip()
    if not s or s == "-":
        return 0.0
    s = re.sub(r"[^\d,.-]", "", s)
    if not s:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return round(float(s), 2)
    except ValueError:
        return 0.0


def formatar_data_str(val):
    if val is None:
        return "-"
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return "-"
    m_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m_iso:
        return f"{m_iso.group(3)}/{m_iso.group(2)}/{m_iso.group(1)}"
    m_br = re.match(r"^(\d{2})[/.-](\d{2})[/.-](\d{4})", s)
    if m_br:
        return f"{m_br.group(1)}/{m_br.group(2)}/{m_br.group(3)}"
    return s


def extrair_cnpj_e_empresa(texto):
    if not texto:
        return "", "", ""
    m_cnpj = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    cnpj_fmt = m_cnpj.group(1) if m_cnpj else ""
    cnpj_limpo = re.sub(r"\D", "", cnpj_fmt)

    empresa_det = ""
    for linha in texto.splitlines()[:25]:
        l_norm = normalizar_texto(linha)
        if any(k in l_norm for k in ["LTDA", "EIRELI", "S/A", "S.A", "ME", "EPP", "COMERCIO", "INDUSTRIA", "SERVICOS", "HOSPITALAR", "MEDICO", "MEDICOS"]):
            empresa_det = re.sub(r"\s+", " ", l_norm).strip()
            break

    return cnpj_limpo, cnpj_fmt, empresa_det


# ==============================================================================
# EXTRAÇÃO DE LIVRO FISCAL (PDF) USANDO PYPDF
# ==============================================================================
def processar_pdf_livro(file_bytes, modo_operacao="notas"):
    registros = []
    texto_cabecalho = ""

    reader = PdfReader(io.BytesIO(file_bytes))
    for idx, page in enumerate(reader.pages):
        txt = page.extract_text() or ""
        if idx < 2:
            texto_cabecalho += txt + "\n"

        if modo_operacao == "resumo_cfop":
            for linha in txt.splitlines():
                l_clean = linha.strip()
                m_cfop = re.match(r"^([123567][.\-]?\d{3})\b(.*)$", l_clean)
                if m_cfop:
                    cfop = limpar_cfop(m_cfop.group(1))
                    valores = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}", m_cfop.group(2))
                    if valores:
                        val_contabil = parse_valor_br(valores[0])
                        if val_contabil > 0:
                            registros.append({
                                "data": "-",
                                "data_emissao": "-",
                                "nota": "RESUMO",
                                "cfop": cfop,
                                "valor": val_contabil
                            })
        else:
            data_atual = "-"
            for linha in txt.splitlines():
                l_clean = linha.strip()
                if not l_clean:
                    continue

                datas_linha = re.findall(r"\b(\d{2}/\d{2}/\d{4})\b", l_clean)
                if datas_linha:
                    data_atual = datas_linha[0]

                m_cfop = re.search(r"\b([123][.\-]?\d{3})\b", l_clean)
                valores_linha = re.findall(r"\b\d{1,3}(?:\.\d{3})*,\d{2}\b", l_clean)

                if m_cfop and valores_linha:
                    cfop = limpar_cfop(m_cfop.group(1))
                    l_up = normalizar_texto(l_clean)
                    if any(p in l_up for p in ["TOTAL", "SUBTOTAL", "TRANSPORTE", "RESUMO"]):
                        continue

                    antes_cfop = l_clean[:m_cfop.start()]
                    antes_sem_datas = re.sub(r"\d{2}/\d{2}/\d{4}", " ", antes_cfop)
                    nums_candidatos = re.findall(r"\b0*(\d{1,9})\b", antes_sem_datas)
                    nums_filtrados = [n for n in nums_candidatos if n not in ("55", "1", "01", "0")]
                    nota = (nums_filtrados[-1] if nums_filtrados else (nums_candidatos[-1] if nums_candidatos else "")).lstrip("0")

                    if nota and len(nota) <= 9:
                        val_contabil = parse_valor_br(valores_linha[0])
                        data_emissao = datas_linha[1] if len(datas_linha) > 1 else data_atual
                        registros.append({
                            "data": data_atual,
                            "data_emissao": data_emissao,
                            "nota": nota,
                            "cfop": cfop,
                            "valor": val_contabil
                        })

    cnpj_limpo, cnpj_fmt, empresa_det = extrair_cnpj_e_empresa(texto_cabecalho)
    return {
        "sucesso": True,
        "cnpj": cnpj_limpo,
        "cnpj_formatado": cnpj_fmt,
        "empresa_detectada": empresa_det,
        "cabecalho_texto": texto_cabecalho[:2500],
        "registros": registros
    }


# ==============================================================================
# EXTRAÇÃO DE PLANILHA (.XLSX, .CSV, .ODS) SEM PANDAS (OPENPYXL PURO)
# ==============================================================================
def encontrar_cabecalho_linhas(linhas):
    palavras_chave = ["NOTA", "NF", "NUMERO", "DOCUMENTO", "CFOP", "VALOR", "TOTAL", "CONTABIL", "DATA", "ENTRADA", "EMISSAO"]
    melhor_idx = 0
    melhor_score = -1

    for idx in range(min(25, len(linhas))):
        vals = [normalizar_texto(v) for v in linhas[idx]]
        score = sum(1 for v in vals if any(k in v for k in palavras_chave))
        if score > melhor_score:
            melhor_score = score
            melhor_idx = idx

    return melhor_idx if melhor_score >= 2 else 0


def mapear_indices_colunas(headers):
    idx_nota = None
    idx_cfop = None
    idx_valor = None
    idx_data = None
    idx_emissao = None

    for i, c in enumerate(headers):
        cn = normalizar_texto(c)
        if not cn:
            continue

        if idx_nota is None and any(k in cn for k in ["NOTA", "NFE", "NF-E", "NUM DOC", "NUMERO", "N. DOC", "DOCUMENTO", "DOC."]):
            if "CHAVE" not in cn and "SERIE" not in cn and "DATA" not in cn:
                idx_nota = i

        if idx_cfop is None and "CFOP" in cn:
            idx_cfop = i

        if idx_valor is None and any(k in cn for k in ["VALOR CONTABIL", "VLR CONTABIL", "VALOR TOTAL", "VLR TOTAL", "TOTAL NF", "VALOR NF", "VALOR DOCUMENTO", "VALOR"]):
            if not any(x in cn for x in ["ICMS", "IPI", "PIS", "COFINS", "BASE", "DESCONTO", "FRETE", "UNIT"]):
                idx_valor = i

        if idx_data is None and any(k in cn for k in ["DATA ENTRADA", "DT ENTRADA", "ENTRADA", "LANCAMENTO", "DATA"]):
            if "EMISSAO" not in cn:
                idx_data = i

        if idx_emissao is None and "EMISSAO" in cn:
            idx_emissao = i

    if idx_valor is None:
        for i, c in enumerate(headers):
            cn = normalizar_texto(c)
            if "VALOR" in cn or "TOTAL" in cn:
                idx_valor = i
                break

    if idx_data is None and idx_emissao is not None:
        idx_data = idx_emissao

    return idx_nota, idx_cfop, idx_valor, idx_data, idx_emissao


def ler_todas_abas_linhas(file_bytes, filename):
    ext = os.path.splitext(filename.lower())[1]
    abas = {}

    if ext == ".csv":
        texto = file_bytes.decode("utf-8", errors="replace")
        sep = ";" if texto.count(";") >= texto.count(",") else ","
        reader = csv.reader(io.StringIO(texto), delimiter=sep)
        abas["CSV"] = [list(row) for row in reader]
    elif ext == ".ods":
        from odf.opendocument import load as load_ods
        from odf.table import Table, TableRow, TableCell
        from odf.text import P
        doc = load_ods(io.BytesIO(file_bytes))
        for table in doc.spreadsheet.getElementsByType(Table):
            sheet_name = table.getAttribute("name") or "ODS"
            linhas = []
            for tr in table.getElementsByType(TableRow):
                row_vals = []
                for tc in tr.getElementsByType(TableCell):
                    rep = int(tc.getAttribute("numbercolumnsrepeated") or 1)
                    val = tc.getAttribute("value")
                    if val is None:
                        ps = tc.getElementsByType(P)
                        val = " ".join(str(p) for p in ps).strip()
                    for _ in range(min(rep, 50)):
                        row_vals.append(val)
                if any(v not in (None, "") for v in row_vals):
                    linhas.append(row_vals)
            abas[sheet_name] = linhas
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            abas[sheet_name] = [list(row) for row in ws.iter_rows(values_only=True)]
        wb.close()

    return abas


def processar_planilha_bytes(file_bytes, filename, modo_operacao="notas"):
    abas = ler_todas_abas_linhas(file_bytes, filename)
    registros = []
    texto_cab = ""

    for sheet_name, linhas in abas.items():
        if not linhas:
            continue

        idx_cab = encontrar_cabecalho_linhas(linhas)
        for r in range(idx_cab):
            linha_txt = " ".join(str(x) for x in linhas[r] if x is not None).strip()
            if linha_txt:
                texto_cab += linha_txt + "\n"

        headers = [str(x).strip() if x is not None else f"COL_{i}" for i, x in enumerate(linhas[idx_cab])]
        idx_nota, idx_cfop, idx_valor, idx_data, idx_emissao = mapear_indices_colunas(headers)

        if idx_valor is None or (idx_nota is None and idx_cfop is None):
            continue

        for row in linhas[idx_cab + 1:]:
            if not row:
                continue

            def pegar(idx):
                return row[idx] if (idx is not None and idx < len(row)) else None

            val_num = parse_valor_br(pegar(idx_valor))
            cfop_str = limpar_cfop(pegar(idx_cfop)) if idx_cfop is not None else ""
            nota_str = limpar_numero_nota(pegar(idx_nota)) if idx_nota is not None else "RESUMO"
            data_str = formatar_data_str(pegar(idx_data)) if idx_data is not None else "-"
            emissao_str = formatar_data_str(pegar(idx_emissao)) if idx_emissao is not None else data_str

            linha_concat = " ".join(normalizar_texto(v) for v in row if v is not None)
            if any(t in linha_concat for t in ["TOTAL GERAL", "SUBTOTAL", "TOTAIS"]):
                continue

            if modo_operacao == "resumo_cfop":
                if not cfop_str or val_num == 0:
                    continue
                registros.append({
                    "data": data_str,
                    "data_emissao": emissao_str,
                    "nota": "RESUMO",
                    "cfop": cfop_str,
                    "valor": val_num
                })
            else:
                if not nota_str or nota_str == "0":
                    continue
                registros.append({
                    "data": data_str,
                    "data_emissao": emissao_str,
                    "nota": nota_str,
                    "cfop": cfop_str,
                    "valor": val_num
                })

    cnpj_limpo, cnpj_fmt, empresa_det = extrair_cnpj_e_empresa(texto_cab)
    return {
        "sucesso": True,
        "cnpj": cnpj_limpo,
        "cnpj_formatado": cnpj_fmt,
        "empresa_detectada": empresa_det,
        "cabecalho_texto": texto_cab[:2500],
        "registros": registros
    }


# ==============================================================================
# ROTAS FLASK
# ==============================================================================
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/extrair", methods=["POST"])
def api_extrair():
    if "arquivo" not in request.files:
        return jsonify({"sucesso": False, "erro": "Nenhum arquivo enviado."}), 400

    arquivo = request.files["arquivo"]
    modo = request.form.get("modo_operacao", "notas")
    nome = arquivo.filename or ""
    ext = os.path.splitext(nome.lower())[1]

    try:
        file_bytes = arquivo.read()
        if ext == ".pdf":
            resultado = processar_pdf_livro(file_bytes, modo)
        elif ext in (".xlsx", ".xlsm", ".csv", ".ods"):
            resultado = processar_planilha_bytes(file_bytes, nome, modo)
        else:
            return jsonify({"sucesso": False, "erro": f"Formato não suportado: {ext}"}), 400

        return jsonify(resultado)
    except Exception as e:
        return jsonify({"sucesso": False, "erro": f"Erro ao processar {nome}: {str(e)}"}), 500


def _chamar_apps_script(payload_dict, timeout=45):
    if not DRIVE_APPS_SCRIPT_URL or "COLE_A_NOVA_URL_AQUI" in DRIVE_APPS_SCRIPT_URL:
        raise RuntimeError("Configure a URL do Google Apps Script na variável DRIVE_APPS_SCRIPT_URL do app.py.")

    ctx = ssl._create_unverified_context()
    payload_dict["token"] = DRIVE_TOKEN.strip()
    req = urllib.request.Request(
        DRIVE_APPS_SCRIPT_URL.strip(),
        data=json.dumps(payload_dict).encode("utf-8"),
        headers={"User-Agent": "Mozilla/5.0", "Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
        conteudo = resp.read().decode("utf-8")
        if conteudo.strip().startswith("<"):
            raise RuntimeError('O Google retornou HTML em vez de JSON. Verifique se a implantação está como "Qualquer pessoa".')
        return json.loads(conteudo)


@app.route("/api/drive-entradas", methods=["POST"])
def api_drive_entradas():
    body = request.get_json(silent=True) or {}
    try:
        dados = _chamar_apps_script({
            "acao": "buscar_entradas",
            "aliases": body.get("aliases", []),
            "ano": body.get("ano", ""),
            "mes": body.get("mes", 0)
        }, timeout=45)
        return jsonify(dados)
    except Exception as e:
        return jsonify({"sucesso": False, "erro": str(e)}), 500


@app.route("/api/drive-extrair-planilha", methods=["POST"])
def api_drive_extrair_planilha():
    body = request.get_json(silent=True) or {}
    file_id = body.get("fileId", "")
    nome = body.get("nome", "planilha_drive.xlsx")
    mime_type = body.get("mimeType", "")
    modo = body.get("modo_operacao", "notas")

    if not file_id:
        return jsonify({"sucesso": False, "erro": "fileId não informado."}), 400

    try:
        dados = _chamar_apps_script({
            "acao": "baixar_planilha",
            "fileId": file_id,
            "mimeType": mime_type
        }, timeout=60)

        if not dados.get("sucesso") or not dados.get("base64"):
            msg_erro = dados.get("erro") or f"Resposta sem base64 do Apps Script: {json.dumps(dados)[:300]}"
            return jsonify({"sucesso": False, "erro": msg_erro}), 500

        raw_bytes = base64.b64decode(dados["base64"])
        ext = dados.get("ext", ".xlsx")
        nome_final = nome if nome.lower().endswith((".xlsx", ".xlsm", ".csv", ".ods")) else (nome + ext)

        resultado = processar_planilha_bytes(raw_bytes, nome_final, modo)
        return jsonify(resultado)
    except Exception as e:
        return jsonify({"sucesso": False, "erro": f"Erro ao processar planilha do Drive: {str(e)}"}), 500


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
