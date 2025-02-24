import pandas as pd
from tkinter.filedialog import askopenfilename
import locale


locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


arquivo_sped = askopenfilename(title='Selecione o arquivo SPED Fiscal')

with open(arquivo_sped, 'r', encoding='ISO-8859-1') as file:
    lines = file.readlines()


max_columns = 0
for line in lines:
    num_columns = len(line.split('|'))
    if num_columns > max_columns:
        max_columns = num_columns


colunas = [f'coluna_{i+1}' for i in range(max_columns)]


processed_lines = []

for line in lines:
    columns = line.strip().split('|')
    
    while len(columns) < max_columns:
        columns.append('')

    processed_lines.append(columns)

df = pd.DataFrame(processed_lines, columns=colunas)
df['coluna_2'] = df['coluna_2'].str.strip()

c200 = df.loc[df['coluna_2'].str.upper() == '0200']
c200_df = c200.iloc[:, :9]


with open(arquivo_sped, 'r', encoding='ISO-8859-1') as file:
    notas = []
    nota_atual = None
    nova_nota = bool

    linha = file.readline()
    dados = linha.strip().split('|')
    data = dados[4][2:]
    data = data[:2] + " " + data[2:]

    for line in file:
        campos = line.strip().split('|')

        if len(campos) < 2:
            continue  

        registro = campos[1]

        if registro == 'C190':
            continue
                
        if registro == 'C100':
            nota_atual = {
                'Chave de Acesso': campos[9],
            }
            nova_nota = True
        
        if not nova_nota:
            nota_atual = {
                'Chave de Acesso': nota_atual["Chave de Acesso"],
            }
        
        if registro == 'C170':
            nota_atual['Codigo'] = campos[3]
            nota_atual['Descricao'] = campos[4]
            nota_atual['Valor'] = campos[7]
            nota_atual['CST'] = campos[10]
            nota_atual['CFOP'] = campos[11]
            nova_nota = False
        
            if nota_atual is not None:
                notas.append(nota_atual)


df2 = pd.DataFrame(notas)

df2 = df2[df2['Chave de Acesso'].notna() & (df2['Chave de Acesso'] != '')]

df2 = df2[df2['CST'].str.startswith(('0', '4', '5'))]

df2['Valor'] = df2['Valor'].str.replace(',', '.').astype(float)

df2['Valor'] = df2['Valor'].apply(lambda x: locale.currency(x, grouping=True))


df_resultado = df2.merge(c200_df[['coluna_3','coluna_9']], left_on='Codigo', right_on="coluna_3", how='left')

df_resultado.drop(columns=['coluna_3'], inplace=True)

df_resultado.rename(columns={'coluna_9': 'NCM'}, inplace=True)

df_resultado.to_excel(f"BNDES {data}.xlsx", index=False, sheet_name=f"BNDES {data}")

