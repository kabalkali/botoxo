import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from banco_dados import data  # Certifique-se de que esse módulo esteja corretamente implementado
import telebot
import os
from datetime import datetime

# Chave da API do Telegram
CHAVE_API = "7795566868:AAG5jU1tDM4DNop6m8oymVs3c8XoK4_v6bk"
bot = telebot.TeleBot(CHAVE_API)

# Função para carregar a planilha Excel a partir do arquivo recebido pelo Telegram
def carregar_planilha(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo)
        print("Planilha carregada com sucesso!")
        return df
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None

# Função para salvar a planilha modificada com nome personalizado
def salvar_planilha(df):
    # Obtém a data atual no formato desejado
    data_atual = datetime.now().strftime("%d-%m-%Y")
    # Nome do arquivo como "Toxo + data atual"
    caminho_salvar = f"Toxo {data_atual}.xlsx"
    try:
        df.to_excel(caminho_salvar, index=False)
        print(f"Planilha salva com sucesso em: {caminho_salvar}")
        return caminho_salvar
    except Exception as e:
        print(f"Erro ao salvar a planilha: {e}")
        return None

# Função para processar a planilha
def processar_planilha(caminho_arquivo):
    # Criando o DataFrame a partir dos dados do banco de dados
    banco_dados = pd.DataFrame(data)

    # Carregar a planilha
    planilha = carregar_planilha(caminho_arquivo)

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

            # Salvar a planilha modificada
            caminho_salvar = salvar_planilha(planilha)
            return caminho_salvar
        else:
            print("A planilha não possui colunas suficientes para as operações.")
            return None
    else:
        print("Erro ao carregar a planilha.")
        return None

# Função para lidar com o comando /Corrigir
@bot.message_handler(commands=["Corrigir"])
def opcao2(mensagem):
    bot.send_message(mensagem.chat.id, "Por favor, envie o arquivo da sua Rota Shopee para correção.")

# Função para lidar com arquivos enviados
@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        file_id = message.document.file_id
        file_info = bot.get_file(file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        # Salvar o arquivo temporariamente
        with open("received_file.xlsx", 'wb') as new_file:
            new_file.write(downloaded_file)

        # Processar o arquivo
        caminho_modificado = processar_planilha("received_file.xlsx")

        if caminho_modificado:
            # Enviar o arquivo modificado de volta para o usuário
            with open(caminho_modificado, 'rb') as doc:
                bot.send_document(message.chat.id, doc)
            os.remove(caminho_modificado)  # Remover o arquivo modificado após o envio
        else:
            bot.send_message(message.chat.id, "Erro ao processar a planilha. Verifique o formato do arquivo.")
        
        os.remove("received_file.xlsx")  # Remover o arquivo recebido após o processamento

    except Exception as e:
        bot.send_message(message.chat.id, f"Ocorreu um erro ao processar o arquivo: {e}")

# Funções adicionais do bot
@bot.message_handler(commands=["Doe"])
def opcao1(mensagem):
    bot.send_message(mensagem.chat.id, "Ajude a manter o projeto CPF: 153.714.787-01")

@bot.message_handler(commands=["Circuit"])
def opcao3(mensagem):
    bot.send_message(mensagem.chat.id, "Segue abaixo o link do canal para baixar o circuit premium: https://t.me/+FSRv_3UzjYQxMzgx")

# Função padrão para exibir as opções
def verificar(mensagem):
    return True

@bot.message_handler(func=verificar)
def responder(mensagem):
    texto = """
    Escolha uma opção para continuar (Clique no item):
     /Doe Mande um valorzinho pro café haha
     /Corrigir Aprimore sua Rota
     /Circuit Baixe o roteirizador
Responder qualquer outra coisa não vai funcionar, clique em uma das opções"""
    bot.reply_to(mensagem, texto)

# Iniciar o bot
bot.polling()
