import pandas as pd
import os
import sys

def formatar_valor(valor, tamanho):
    """
    Formata o valor para um tamanho fixo, preenchendo com espaços ou truncando.
    Garante que valores numéricos tenham 2 casas decimais, mas REMOVE a vírgula/ponto para o TXT.
    Exemplo: 1000.00 vira "100000" (preenchido com espaços até o tamanho).
    """
    if valor is None:
        return "".ljust(tamanho)
    
    s_valor = str(valor).strip()
    if s_valor.lower() in ['nan', 'none', 'null', '']:
        return "".ljust(tamanho)

    # Tenta reconhecer se é um número
    try:
        if isinstance(valor, str):
            if any(c.isalpha() for c in s_valor.replace('e', '').replace('E', '')):
                raise ValueError
            test_valor = s_valor.replace('.', '').replace(',', '.')
            v_float = float(test_valor)
        else:
            v_float = float(valor)
        
        # Formata com 2 casas decimais e remove o ponto (ex: 100.00 -> "10000")
        s_valor = f"{v_float:.2f}".replace('.', '').replace(',', '')
    except (ValueError, TypeError):
        pass
    
    if len(s_valor) > tamanho:
        return s_valor[:tamanho]
    
    return s_valor.ljust(tamanho)

def processar_excel(caminho_arquivo_excel, caminho_arquivo_txt):
    try:
        # Carregar a planilha 'preencher', sem cabeçalho
        df = pd.read_excel(caminho_arquivo_excel, sheet_name='preencher', header=None)

        # Obter os comprimentos dos campos da primeira linha (índice 0)
        comprimentos = df.iloc[0].tolist()
        
        # O processamento começa na linha 3 (índice 2)
        dados = df.iloc[2:]

        with open(caminho_arquivo_txt, 'w', encoding='utf-8') as arquivo_txt:
            for _, linha in dados.iterrows():
                # Linha 1: Colunas A até BF (índices 0 a 57)
                linha1_partes = []
                for i in range(0, 58):
                    valor = linha.iloc[i]
                    tamanho = int(comprimentos[i])
                    linha1_partes.append(formatar_valor(valor, tamanho))
                
                arquivo_txt.write("".join(linha1_partes) + '\n')

                # Linha 2: Colunas BG até o final (índices 58 a 83)
                linha2_partes = []
                for i in range(58, 84):
                    valor = linha.iloc[i]
                    tamanho = int(comprimentos[i])
                    linha2_partes.append(formatar_valor(valor, tamanho))
                
                arquivo_txt.write("".join(linha2_partes) + '\n')
        
        print(f"Sucesso! Arquivo '{caminho_arquivo_txt}' gerado.")
    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")

if __name__ == "__main__":
    # Procurar por arquivos .xlsx no diretório atual se nenhum for fornecido
    if len(sys.argv) > 1:
        caminho_excel = sys.argv[1]
    else:
        arquivos_excel = [f for f in os.listdir('.') if f.endswith('.xlsx')]
        if not arquivos_excel:
            print("Nenhum arquivo Excel (.xlsx) encontrado no diretório atual.")
            sys.exit(1)
        caminho_excel = arquivos_excel[0]
        print(f"Usando o arquivo: {caminho_excel}")

    nome_base = os.path.splitext(caminho_excel)[0]
    caminho_txt = f"{nome_base}_resultado.txt"
    
    processar_excel(caminho_excel, caminho_txt)
