# Conversor de Excel para TXT (Layout EFD0003)

Esta aplicação lê um arquivo Excel com o layout definido na aba "preencher" e gera um arquivo TXT formatado.

## Requisitos

- Python 3.x instalado
- Bibliotecas: `pandas` e `openpyxl`

## Como instalar as dependências

Abra o terminal ou prompt de comando e execute:

```bash
pip install pandas openpyxl
```

## Como usar

1. Coloque o arquivo Excel (`.xlsx`) na mesma pasta que o script `conversor_excel_txt.py`.
2. Execute o script:

```bash
python conversor_excel_txt.py
```

O script irá procurar automaticamente por qualquer arquivo `.xlsx` na pasta e gerar um arquivo `.txt` com o mesmo nome, adicionando o sufixo `_resultado`.

### Detalhes do Processamento

- O script lê a aba **"preencher"**.
- Os comprimentos dos campos são obtidos da **linha 1** da planilha.
- Os dados são lidos a partir da **linha 3**.
- Para cada linha do Excel, são geradas **2 linhas no TXT**:
  - A primeira linha começa com a **coluna A** (até a coluna BF).
  - A segunda linha começa com a **coluna BG** (até o final).
- Todos os campos são alinhados à esquerda e preenchidos com espaços à direita para respeitar o tamanho definido.
