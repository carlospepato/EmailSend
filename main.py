import pandas as pd
import pyarrow
import win32com.client
from openpyxl import load_workbook


def ler_planilha():
    remetente = str(input('email:'))
    planilha = pd.read_excel(r'planilha_teste.xlsx')
    df = pd.DataFrame(planilha)
    grupos_email = df.groupby("E-mail")
    for email, grupo in grupos_email:
        df_grupo = grupo[["Reference", "Currency", "Amount SAP"]]
        print(email)

        # Convertendo o DataFrame para uma tabela HTML
        tabela_html = df_grupo.to_html(index=False, border=1, classes="table table-bordered table-striped")
        print(tabela_html)

        try:
            enviar_emails(remetente, email, tabela_html)
            print(f'E-mail enviado com sucesso para {email}')
        except Exception as e:
            print(f"Falha ao enviar e-mail para {email}. Erro: {str(e)}")

    return remetente, email, tabela_html



def obter_dados_email(planilha_excel):
    workbook = load_workbook(planilha_excel)
    sheet = workbook.active

    remetente = sheet.cell(row=2, column=1).value  # Assume que o remetente está na célula (0, 0)
    print(remetente)

    destinatarios = []
    for row in range(2, sheet.max_row + 1):  # Começa da segunda linha para evitar o remetente
        destinatario = sheet.cell(row=row, column=2).value
        destinatarios.append(destinatario)
        print(destinatario)

    return remetente, destinatarios


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
    planilha_excel = "C:\\Users\\carlos.pepato\\Desktop\\teste_envio_email.xlsx"  # Substitua pelo caminho da sua planilha
    remetente, destinatarios = obter_dados_email(planilha_excel)

    if remetente and destinatarios:
        if(enviar_emails(remetente, destinatarios, corpo='<h1>Teste</h1>')):
            print("E-mails enviados com sucesso!")
        else:
            print("Erro ao enviar o e-mail")
    else:
        if not remetente:
            print("Remetente não encontrado na planilha.")
        if not destinatarios:
            print("Destinatários não encontrados na planilha.")
    #ler_planilha()
