import pandas as pd
import pyarrow
import win32com.client
from openpyxl import load_workbook
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog


def abrir_arquivo(label):
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


def ler_planilha():
    remetente = str(input('email:'))
    planilha = pd.read_excel(r'./planilha/planilha_teste.xlsx')
    df = pd.DataFrame(planilha)
    grupos_email = df.groupby("E-mail", group_keys=True)

    for destinatario, grupo in grupos_email:
        df_grupo = grupo[["Reference", "Currency", "Amount SAP"]]

        # Formatação da coluna 'Amount SAP'
        df_grupo = df_grupo.assign(**{'Amount SAP': df_grupo['Amount SAP'].apply(lambda x: '{:,.2f}'.format(x))})

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

        try:
            enviar_emails(remetente, [destinatario], tabela_html)
            print(f'E-mail enviado com sucesso para {destinatario}')
        except Exception as e:
            print(f"Falha ao enviar e-mail para {destinatario}. Erro: {str(e)}")


# Função para enviar e-mails
def enviar_emails(remetente, destinatarios, corpo):
    outlook = win32com.client.Dispatch("Outlook.Application")

    for destinatario in destinatarios:
        email = outlook.CreateItem(0)  # Criação do objeto de e-mail
        email.To = destinatario
        email.Subject = "Assunto do e-mail teste"
        email.HTMLBody = corpo

        # Definindo o remetente
        email.SentOnBehalfOfName = remetente

        email.Send()

    return True


if __name__ == "__main__":
    #ler_planilha()
    executar_programa()

