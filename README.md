# Conversor de Planilha Excel para JSON

Este script Python converte uma planilha Excel (`.xlsx`) para um arquivo JSON. Ele lida com tipos de dados específicos, como datas, horas e strings contendo caracteres especiais, garantindo uma conversão limpa para o formato JSON.

## Dependências

- `openpyxl`: Para manipulação de arquivos Excel.
- `json`: Para codificação JSON.
- `re`: Para expressões regulares (usadas na limpeza de strings).
- `datetime`: Para manipulação de datas e horas.

Instale as dependências usando pip:

```bash
pip install openpyxl
```

## Como Usar

1. **Salve o script:** Salve o código Python como um arquivo `.py` (por exemplo, `converter.py`).
2. **Prepare sua planilha:** Certifique-se de que sua planilha Excel (`.xlsx`) esteja no mesmo diretório do script ou forneça o caminho correto para o arquivo. A primeira linha da planilha deve conter os cabeçalhos das colunas.
3. **Execute o script:** Execute o script a partir do seu terminal:

```bash
python converter.py
```

4. **Arquivo JSON gerado:** O script criará um arquivo chamado `dados.json` no mesmo diretório contendo os dados da planilha no formato JSON.

## Funcionamento do Código

1. **Importação de bibliotecas:** Importa as bibliotecas necessárias: `openpyxl`, `json`, `re` e `datetime`.
2. **Função `converter_valor`:** Esta função lida com a conversão de diferentes tipos de dados:
    - Datas e horas (`datetime`, `time`): Converte para o formato ISO 8601.
    - Strings: Remove caracteres especiais de controle de carriage return (`_x000D_` e quebras de linha `\n`) frequentemente encontrados em dados copiados de outras fontes.
    - Outros tipos: Retorna o valor sem modificações.
3. **Carregamento da planilha:** Carrega a planilha Excel usando `openpyxl.load_workbook`.
4. **Extração de cabeçalhos:** Lê a primeira linha da planilha para extrair os cabeçalhos das colunas.
5. **Extração de dados:** Itera pelas linhas da planilha (a partir da segunda linha) e extrai os valores das células. A função `converter_valor` é aplicada a cada valor para garantir a conversão correta.
6. **Criação do JSON:** Cria um dicionário Python para cada linha da planilha, usando os cabeçalhos como chaves e os valores convertidos como valores. Esses dicionários são então adicionados a uma lista.
7. **Escrita no arquivo:** Converte a lista de dicionários para JSON usando `json.dumps` com indentação para melhor legibilidade e codificação UTF-8 para suportar caracteres especiais. O JSON resultante é gravado em um arquivo chamado `dados.json`.


## Exemplo de Planilha

```
Nome | Idade | Data de Nascimento
-----|-------|------------------
João | 30    | 1993-01-15
Maria| 25    | 1998-05-20
```

## Exemplo de JSON gerado

```json
[
    {
        "Nome": "João",
        "Idade": 30,
        "Data de Nascimento": "1993-01-15 00:00:00"
    },
    {
        "Nome": "Maria",
        "Idade": 25,
        "Data de Nascimento": "1998-05-20 00:00:00"
    }
]
```

## Observações

- O script assume que a primeira linha da planilha contém os cabeçalhos das colunas.
- A codificação UTF-8 é usada para garantir que caracteres especiais sejam tratados corretamente.
- A função `converter_valor` pode ser estendida para lidar com outros tipos de dados ou necessidades específicas de limpeza de dados.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir um pull request com melhorias ou correções de bugs.
```
