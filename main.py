import pandas as pd
import win32com.client
from bs4 import BeautifulSoup
import os
from openpyxl import load_workbook
#from tkinter import *
#from tkinter import messagebox
#from tkinter import filedialog

class cor:
    RESET = '\033[0m'
    PRETO = '\033[30m'
    VERMELHO = '\033[31m'
    VERDE = '\033[32m'
    AMARELO = '\033[33m'
    AZUL = '\033[34m'
    MAGENTA = '\033[35m'
    CIANO = '\033[36m'
    BRANCO = '\033[37m'
    BRILHANTE_PRETO = '\033[90m'
    BRILHANTE_VERMELHO = '\033[91m'
    BRILHANTE_VERDE = '\033[92m'
    BRILHANTE_AMARELO = '\033[93m'
    BRILHANTE_AZUL = '\033[94m'
    BRILHANTE_MAGENTA = '\033[95m'
    BRILHANTE_CIANO = '\033[96m'
    BRILHANTE_BRANCO = '\033[97m'

"""def abrir_arquivo(label):
    # Abre a caixa de diálogo para seleção de arquivo
    arquivo = filedialog.askopenfilename(title="Selecione um arquivo", filetypes=[("Arquivos de Texto", "*.txt"), ("Todos os arquivos", "*.*")])

    # Atualiza o rótulo com o caminho do arquivo selecionado
    label.config(text=f"Arquivo selecionado: {arquivo}")

def executar_programa():
    janela = Tk()
    janela.title("Seleção de Arquivo")

    # Criar um rótulo na janela
    label = Label(janela, text="Nenhum arquivo selecionado")
    label.pack(pady=10)

    # Criar um botão na janela
    botao = Button(text="Selecionar Arquivo", command= abrir_arquivo(label))
    botao.pack(pady=20)

    # Iniciar o loop principal da aplicação
    janela.mainloop()
    """


def definir_estilo_tabela(tabela):
    estilo_cabecalho = {
        'selector': 'thead th',  # Seleciona todas as células do cabeçalho
        'props': [('background-color', 'black'), ('color', 'white')]  # Define o fundo preto e texto branco
    }
    tabela_html = (
        tabela.style
        .set_table_styles([estilo_cabecalho])
        .to_html(index=False, classes="table table-striped"))

    return tabela_html

def formatar_valor(x):
    valor_invertido = x * (-1)  # Inverte o sinal do valor da célula
    valor_formatado = '{:,.2f}'.format(valor_invertido)  # Formata o valor com separadores de milhar e duas casas decimais
    return valor_formatado

def ler_planilha():
    remetente = str(input('email:'))
    caminho_excel = str(input('caminho do arquivo:').strip('"'))
    caminho_excel_normalizado = os.path.normpath(caminho_excel)
    planilha = pd.read_excel(caminho_excel_normalizado, header=1)
    df = pd.DataFrame(planilha)
    grupos_email = df.groupby("E-mail", group_keys=True)

    for destinatario, grupo in grupos_email:
        df_grupo = grupo[["Reference", "Currency", "Amount SAP"]]

        # Formatação da coluna 'Amount SAP'
        df_grupo = df_grupo.assign(**{'Amount SAP': df_grupo['Amount SAP'].apply(formatar_valor)})

        data = grupo['Data'].iloc[0]  # Assume que 'Data' é a mesma para todas as linhas do grupo
        nome_empresa = grupo['Company'].iloc[0]  # Assume que 'Company' é a mesma para todas as linhas do grupo
        moeda = grupo['Currency'].iloc[0]  # Assume que 'Currency' é a mesma para todas as linhas do grupo
        valor = grupo['Amount SAP'].sum()

        valor_formatado = '{:,.2f}'.format(-valor)

        df_grupo = df_grupo.reset_index(drop=True)
        tabela_html = '''
                       <h4>Hello!</h4> 
                        <p>How are you?</p>

                        <p>We’d like to let you know that we are sending you today a wire transfer.</p>

                        Wire transfer date: {0}<br/>
                        To: {1}<br/>
                        Total Amount: {2} {3}<br/>

                        <p>Please see below the details of this payment:</p>
                        '''.format(data, nome_empresa, moeda, valor_formatado)
        tabela_html += definir_estilo_tabela(df_grupo)
        tabela_html += '''
                        <p>Best Regards</p>
                        '''

        try:
            enviar_emails(remetente, [destinatario], tabela_html)
            print(f'E-mail enviado com sucesso para {destinatario}')
        except Exception as e:
            print(f"Email '{destinatario}' da empresa '{nome_empresa}' não foi encontrado e/ou não existe!")


# Função para enviar e-mails
def enviar_emails(remetente, destinatarios, corpo):
    outlook = win32com.client.Dispatch("Outlook.Application")

    for destinatario in destinatarios:
        email = outlook.CreateItem(0)  # Criação do objeto de e-mail
        email.To = destinatario
        email.Subject = "Wire Transfer Details - Hi-Mix Eletrônicos"
        email.HTMLBody = corpo

        # Definindo o remetente
        email.SentOnBehalfOfName = remetente

        email.Send()

    return True

def bem_vindo():
    print(cor.AZUL + "Hi-mix Eletrônicos S/A" + cor.RESET)
    print(cor.AZUL + "Sistema para Mala Direta" + cor.RESET)

def mostrar_ao_usuario(tabela):
    soup = BeautifulSoup(tabela, 'html.parser')
    texto_amigavel = soup.get_text()
    print(texto_amigavel)
    resposta = str(input("Enviar email (S/N): "))
    while True:
        if(resposta == 'S'):
            return True
        if(resposta == 'N'):
            return False
        print('Valor não corresponde')
        resposta = str(input("Enviar email (S/N): "))



if __name__ == "__main__":
    bem_vindo()
    ler_planilha()
    input("\nAperte Enter para sair!")
    #executar_programa()

