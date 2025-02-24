# SPED Fiscal - Analisador e Gerador de Relatório BNDES
<br/>
Este é um script simples em Python que processa um arquivo SPED Fiscal, extraí informações relacionadas às notas fiscais e gera um relatório formatado no formato .xlsx para facilitar a análise e utilização de dados para o BNDES.
<br/>
<br/>

### Funcionalidades

Lê o arquivo SPED Fiscal no formato .txt (delimitado por |).
Extrai informações das notas fiscais, como:
Chave de Acesso
Código, Descrição, Valor, CST, CFOP
Utiliza a tabela de NCM (código de mercadoria) para incluir informações adicionais nas notas.
Filtra as notas de acordo com determinados critérios.
Gera um arquivo Excel (.xlsx) com as informações processadas, pronto para ser utilizado.
<br/>
<br/>

### Pré-requisitos:
Python 3.x;  
Bibliotecas: pandas, locale e tkinter;  
O arquivo SPED Fiscal deve ser no formato .txt e estar codificado em ISO-8859-1;
<br/>
<br/>

### Execução:
Execute o script Python no seu terminal ou ambiente de desenvolvimento.
O script irá pedir para selecionar o arquivo SPED Fiscal e processará os dados automaticamente.
O script irá gerar um arquivo Excel (.xlsx) com os resultados na mesma pasta onde o script foi executado.
<br/>
Exemplo de Saída
Um arquivo BNDES 23-02-2025.xlsx será gerado com as informações das notas fiscais que atendem aos critérios definidos (CST iniciando com '0', '4', ou '5') e seus respectivos NCMs.
<br/>
<br/>

### Dependências
Este projeto depende de algumas bibliotecas Python. Você pode instalá-las usando pip:
```bash
pip install pandas
```
As bibliotecas locale e tkinter já fazem parte da instalação padrão do Python.
