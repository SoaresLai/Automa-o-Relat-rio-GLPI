import win32com.client as win32


def enviar_relatorio(
    arquivo,
    destinatarios,
    assunto,
    corpo
):

    outlook = win32.Dispatch("Outlook.Application")

    mensagem = outlook.CreateItem(0)

    mensagem.To = ";".join(destinatarios)

    mensagem.Subject = assunto

    mensagem.Body = corpo

    mensagem.Attachments.Add(arquivo)

    mensagem.Send()

    print("E-mail enviado com sucesso.")