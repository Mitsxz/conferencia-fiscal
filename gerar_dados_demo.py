import os
import pandas as pd

def gerar_arquivos_demonstracao():
    os.makedirs('dados_demonstracao', exist_ok=True)

    dados_livro = [
        {
            'Numero_NF': '000101', 'Serie': '1', 'CNPJ': '12.345.678/0001-90',
            'Data_Emissao': '05/11/2026', 'Cod_Produto': 'P001', 'Descricao': 'Chapa de Aço 2mm',
            'Quantidade': 10, 'Valor_Total': 1500.00, 'Base_Calculo': 1500.00,
            'ICMS': 270.00, 'IPI': 75.00, 'PIS': 24.75, 'COFINS': 114.00
        },
        {
            'Numero_NF': '000102', 'Serie': '1', 'CNPJ': '98.765.432/0001-10',
            'Data_Emissao': '06/11/2026', 'Cod_Produto': 'P002', 'Descricao': 'Parafuso Sextavado',
            'Quantidade': 500, 'Valor_Total': 850.00, 'Base_Calculo': 850.00,
            'ICMS': 153.00, 'IPI': 0.00, 'PIS': 14.03, 'COFINS': 64.60
        },
        {
            'Numero_NF': '000103', 'Serie': '1', 'CNPJ': '11.222.333/0001-44',
            'Data_Emissao': '08/11/2026', 'Cod_Produto': 'P003', 'Descricao': 'Tubo Industrial 3pol',
            'Quantidade': 20, 'Valor_Total': 3200.00, 'Base_Calculo': 3200.00,
            'ICMS': 576.00, 'IPI': 160.00, 'PIS': 52.80, 'COFINS': 243.20
        },
        {
            'Numero_NF': '000104', 'Serie': '1', 'CNPJ': '44.555.666/0001-77',
            'Data_Emissao': '10/11/2026', 'Cod_Produto': 'P004', 'Descricao': 'Válvula Hidráulica',
            'Quantidade': 2, 'Valor_Total': 940.00, 'Base_Calculo': 940.00,
            'ICMS': 169.20, 'IPI': 47.00, 'PIS': 15.51, 'COFINS': 71.44
        },
        {
            'Numero_NF': '000105', 'Serie': '2', 'CNPJ': '12.345.678/0001-90',
            'Data_Emissao': '12/11/2026', 'Cod_Produto': 'P001', 'Descricao': 'Chapa de Aço 2mm',
            'Quantidade': 5, 'Valor_Total': 750.00, 'Base_Calculo': 750.00,
            'ICMS': 135.00, 'IPI': 37.50, 'PIS': 12.38, 'COFINS': 57.00
        }
    ]

    dados_planilha_1 = [
        {
            'NF': '101', 'Serie': '1', 'CNPJ_Emitente': '12345678000190',
            'Emissao': '2026-11-05', 'Item': 'P001', 'Produto': 'Chapa de Aço 2mm',
            'Qtd': 10, 'Total': 'R$ 1.500,00', 'BC_ICMS': 1500.00,
            'Valor_ICMS': 270.00, 'Valor_IPI': 75.00, 'PIS': 24.75, 'COFINS': 114.00
        },
        {
            'NF': '102', 'Serie': '1', 'CNPJ_Emitente': '98765432000110',
            'Emissao': '2026-11-06', 'Item': 'P002', 'Produto': 'Parafuso Sextavado',
            'Qtd': 500, 'Total': 'R$ 890,00', 'BC_ICMS': 890.00,
            'Valor_ICMS': 160.20, 'Valor_IPI': 0.00, 'PIS': 14.03, 'COFINS': 64.60
        },
        {
            'NF': '103', 'Serie': '1', 'CNPJ_Emitente': '11222333000144',
            'Emissao': '2026-11-09', 'Item': 'P003', 'Produto': 'Tubo Industrial 3pol',
            'Qtd': 18, 'Total': 'R$ 3.200,00', 'BC_ICMS': 3200.00,
            'Valor_ICMS': 576.00, 'Valor_IPI': 160.00, 'PIS': 52.80, 'COFINS': 243.20
        }
    ]

    dados_planilha_2 = [
        {
            'NF': '000105', 'Serie': '2', 'CNPJ_Emitente': '12.345.678/0001-90',
            'Emissao': '12/11/2026', 'Item': 'P001', 'Produto': 'Chapa de Aço 2mm',
            'Qtd': 5, 'Total': 750.00, 'BC_ICMS': 750.00,
            'Valor_ICMS': 135.00, 'Valor_IPI': 37.50, 'PIS': 12.38, 'COFINS': 57.00
        },
        {
            'NF': '000105', 'Serie': '2', 'CNPJ_Emitente': '12.345.678/0001-90',
            'Emissao': '12/11/2026', 'Item': 'P001', 'Produto': 'Chapa de Aço 2mm',
            'Qtd': 5, 'Total': 750.00, 'BC_ICMS': 750.00,
            'Valor_ICMS': 135.00, 'Valor_IPI': 37.50, 'PIS': 12.38, 'COFINS': 57.00
        },
        {
            'NF': '000106', 'Serie': '1', 'CNPJ_Emitente': '77.888.999/0001-22',
            'Emissao': '15/11/2026', 'Item': 'P009', 'Produto': 'Flange Cego 4pol',
            'Qtd': 4, 'Total': 620.00, 'BC_ICMS': 620.00,
            'Valor_ICMS': 111.60, 'Valor_IPI': 0.00, 'PIS': 10.23, 'COFINS': 47.12
        }
    ]

    pd.DataFrame(dados_livro).to_excel('dados_demonstracao/DEMO_Livro_Fiscal_Nov2026.xlsx', index=False)
    pd.DataFrame(dados_planilha_1).to_excel('dados_demonstracao/DEMO_Planilha_Compras_Parte1.xlsx', index=False)
    pd.DataFrame(dados_planilha_2).to_csv('dados_demonstracao/DEMO_Planilha_Compras_Parte2.csv', sep=';', index=False, encoding='utf-8-sig')

if __name__ == '__main__':
    gerar_arquivos_demonstracao()
