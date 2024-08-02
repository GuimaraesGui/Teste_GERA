import win32com.client as win32
import base64

# Lista de contatos
contatos_GN = ["guimaraes_gui@outlook.com.br;tripaseca.kun@gmail.com;ep.0.season.0@gmail.com;vicktoriagranger15@gmail.com;equipeotaku.kun@gmail.com"]

contatos_GN1 = [["guimaraes_gui@outlook.com.br"],["tripaseca.kun@gmail.com"],
                ["ep.0.season.0@gmail.com"], ["vicktoriagranger15@gmail.com"],
                  ["equipeotaku.kun@gmail.com"]]

# Caminho da imagem
caminho_imagem = "C:\\Users\\00805129\\OneDrive - NATURGY INFORMATICA S.A\\Escritorio\\Studying_Python\\E-mail_automation\\ass_email.png"

# Criar a integração com o Outlook
outlook = win32.Dispatch('outlook.application')

# Criar um e-mail
email = outlook.CreateItem(0)

# Configurar as informações do e-mail
email.To = ";".join(contatos_GN)
email.Subject = "E-mail automático do Python"  # Assunto

# Ler o conteúdo da imagem e convertê-lo para base64
with open(caminho_imagem, "rb") as img_file:
    img_data = img_file.read()
    img_base64 = base64.b64encode(img_data).decode("utf-8")

# Incorporar a imagem diretamente no HTMLBody
email.HTMLBody = f"""
<p>Boa tarde a todos,</p>

<p>Anexado, segue atualização a respeito da composição físico-química.</p>

<p>Aqui está uma imagem incorporada:</p>

<p>At.te,</p>
<img src="data:image/png;base64,{img_base64}">
"""

# Enviar o e-mail
email.Send()

print("\nEmail Enviado")