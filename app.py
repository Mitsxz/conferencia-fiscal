elif filename.endswith('.pdf'):
    reader = PdfReader(file)
    texto_completo = ""
    for pagina in reader.pages:
        texto_ext = pagina.extract_text()
        if texto_ext:
            texto_completo += texto_ext + "\n"
    
    linhas = texto_completo.splitlines()
    capturando = False
    
    for linha in linhas:
        linha_limpa = linha.strip()
        
        # Opcional: Detecta se entrou em uma seção útil (se o seu PDF tiver termos como "Lançamentos", "Entradas" ou "Resumo")
        # Se preferir ler tudo buscando os padrões de nota, podemos remover o bloqueio de seção.
        
        # Procura por linhas que contenham o padrão fiscal (Data YYYY-MM-DD ou DD/MM/YYYY, Nota, CFOP e Valor)
        # Regex flexível para capturar a linha fiscal mesmo com textos ao redor
        match = re.search(r'(\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2})\D+(\d{1,8})\D+(\d{4})\D+([\d\.,]+)', linha_limpa)
        if match:
            data_raw, nota_val, cfop_val, valor_str = match.groups()
            
            # Padroniza a data para YYYY-MM-DD se vier em DD/MM/YYYY
            if '/' in data_raw:
                p = data_raw.split('/')
                data_val = f"{p[2]}-{p[1]}-{p[0]}"
            else:
                data_val = data_raw

            nota = limpar_nota(nota_val)
            cfop = limpar_cfop(cfop_val)
            valor = limpar_valor(valor_str)
            
            if nota and valor > 0:
                dados.append({
                    "data": limpar_data(data_val),
                    "nota": nota,
                    "cfop": cfop,
                    "valor": valor
                })
