import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from banco_dados import data  # Importando os dados do banco de dados

# Função para carregar a planilha Excel
def carregar_planilha():
    root = Tk()
    root.withdraw()  # Esconde a janela principal do Tkinter
    caminho_arquivo = askopenfilename(
        title="Selecione a planilha Excel",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
    )
    if caminho_arquivo:
        try:
            df = pd.read_excel(caminho_arquivo)
            print("Planilha carregada com sucesso!")
            return df
        except Exception as e:
            print(f"Erro ao carregar a planilha: {e}")
            return None
    else:
        print("Nenhum arquivo foi selecionado.")
        return None

# Função para salvar a planilha modificada
def salvar_planilha(df):
    caminho_salvar = asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Planilhas Excel", "*.xlsx")],
        title="Salvar planilha como"
    )
    if caminho_salvar:
        try:
            df.to_excel(caminho_salvar, index=False)
            print(f"Planilha salva com sucesso em: {caminho_salvar}")
        except Exception as e:
            print(f"Erro ao salvar a planilha: {e}")
    else:
        print("Operação de salvamento cancelada.")

# Criando o DataFrame a partir dos dados do banco de dados
banco_dados = pd.DataFrame(data)

# Carregar a planilha
planilha = carregar_planilha()

if planilha is not None:
    # Verifica se a coluna 5 (índice 4) existe
    if planilha.shape[1] > 4:
        # Duplicar a coluna 5 e criar a nova coluna 'number'
        planilha['number'] = planilha.iloc[:, 4]

        # Manipula a coluna 5 para manter apenas o texto antes da primeira vírgula
        planilha.iloc[:, 4] = planilha.iloc[:, 4].str.split(',').str[0]

        # Manipula a coluna 'number' para manter o texto entre a primeira e a segunda vírgula
        planilha['number'] = planilha['number'].str.split(',').str[1]

        # Atualiza a coluna 5 com base no banco de dados usando merge
        planilha = planilha.merge(banco_dados[['CEP', 'Logradouro']], left_on='Zipcode/Postal code', right_on='CEP', how='left')
        planilha.iloc[:, 4] = planilha['Logradouro']
        planilha.drop(columns=['CEP', 'Logradouro'], inplace=True)

        # Criar a coluna "Pacotes na Parada" usando "Destination Address" e "number"
        planilha['Pacotes na Parada'] = planilha.groupby(['Destination Address', 'number'])['Sequence']\
            .transform(lambda x: ', '.join(map(str, x)))
        
        # Adicionar os valores da coluna 'number' na 'Destination Address' no formato correto
        planilha['Destination Address'] = planilha.apply(
            lambda row: f"{row['Destination Address']}, {row['number']}".strip() if pd.notnull(row['number']) else row['Destination Address'], 
            axis=1
        )

        # Excluir a coluna 'number'
        planilha.drop(columns=['number'], inplace=True)

        # Excluir a coluna "Sequence" e "Stop"
        planilha.drop(columns=['Sequence', 'Stop', 'SPX TN'], inplace=True)

        # Remover linhas duplicadas com base em "Destination Address" e "Pacotes na Parada"
        planilha.drop_duplicates(subset=['Destination Address', 'Pacotes na Parada'], inplace=True)

        print("Colunas manipuladas, valores adicionados a 'Destination Address' e duplicatas removidas com sucesso!")
        print(planilha.head())  # Exibe as primeiras linhas do dataframe para verificar
        
        # Chama a função para salvar a planilha modificada
        salvar_planilha(planilha)
    else:
        print("A planilha não possui colunas suficientes para as operações.")
else:
    print("Erro ao carregar a planilha.")
