import win32com.client as win32
import base64

# Listas de contatos
contatos_GN1 = ["e-mail 1",
                "e-mail 2",
                "e-mail 3"]

# Lista com os anexos dos consumos de gás natural dos clientes
caminho_anexos = [
        r"C:\Users\00805129\OneDrive - NATURGY INFORMATICA S.A\Documentos\02 - Industrial\Qualidade GNV\CFQ - SPS\01 - JANEIRO\Diário\Consumo Ajinomoto Janeiro.xlsx",
        r"C:\Users\00805129\OneDrive - NATURGY INFORMATICA S.A\Documentos\02 - Industrial\Qualidade GNV\CFQ - SPS\01 - JANEIRO\Diário\Consumo CBA Janeiro.xlsx",
        r"C:\Users\00805129\OneDrive - NATURGY INFORMATICA S.A\Documentos\02 - Industrial\Qualidade GNV\CFQ - SPS\01 - JANEIRO\Diário\Consumo DeNora Janeiro.xlsx"
        ]

# Caminho da imagem da minha assinatura
caminho_imagem = "C:\\Users\\00805129\\OneDrive - NATURGY INFORMATICA S.A\\Escritorio\\Studying_Python\\E-mail_automation\\ass_email.png"

# Criando a integração com o Outlook
outlook = win32.Dispatch('outlook.application')

# Loop sobre os contatos
for i in range(len(contatos_GN1)):
    # Criar um e-mail
    email = outlook.CreateItem(0)

    # Configurar as informações do e-mail para o contato atual
    email.To = contatos_GN1[i]  # Substitua pelo contato atual
    email.Subject = f"E-mail automático do Python para {contatos_GN1[i]}"  # Assunto

    # Ler o conteúdo da imagem e convertê-lo para base64
    with open(caminho_imagem, "rb") as img_file:
        img_data = img_file.read()
        img_base64 = base64.b64encode(img_data).decode("utf-8")

    # Incorporar o anexo diretamente no HTMLBody
    email.HTMLBody = f"""
    <p>Boa tarde {contatos_GN1[i]},</p>

    <p>Anexado, segue atualização a respeito da composição físico-química.</p>

    <p>At.te,</p>
    <img src="data:image/png;base64,{img_base64}">
    """

    # Adicionar o anexo ao e-mail
    attachment = email.Attachments.Add(Source=caminho_anexos[i])

    # Enviar o e-mail
    email.Send()

"""
for i in range(len(contatos_GN1)):
    print(contatos_GN1[i], end=" / ")
"""

print("\nE-mails Enviados")
