import os
import win32com.client

# Configura el path donde quieres guardar los adjuntos
path_to_save = r"C:\Users\qband\OneDrive\Desktop\daily"

# Inicia sesión en Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Selecciona la carpeta de la bandeja de entrada
inbox = outlook.GetDefaultFolder(6)  # 6 es la bandeja de entrada

# Lista de correos electrónicos de los remitentes específicos
sender_emails = ["jcastrov@cosapi.com.pe", "kabads@cosapi.com.pe"]

# Recorre todos los mensajes en la bandeja de entrada
for message in inbox.Items:
    try:
        # Verifica si el mensaje es de tipo MailItem utilizando el valor numérico directamente
        if message.Class == 43:  # 43 corresponde a la constante olMailItem
            # Acceder al objeto Sender y luego a la propiedad Address
            sender = message.Sender
            sender_email_address = sender.Address if sender else None
            # Verifica si el correo electrónico del remitente está en la lista de remitentes específicos
            if sender_email_address and sender_email_address.lower() in (email.lower() for email in sender_emails):
                # Recorre todos los adjuntos en el mensaje actual
                for attachment in message.Attachments:
                    # Define el path completo para guardar el archivo
                    attachment_path = os.path.join(path_to_save, attachment.FileName)
                    # Guarda el adjunto en el path definido
                    attachment.SaveAsFile(attachment_path)
                    print(f"Archivo {attachment.FileName} guardado en {attachment_path}")
    except Exception as e:
        print(f"Se encontró un error procesando un mensaje: {e}")

# Nota: Si el script aún muestra errores, es posible que tengas que verificar los permisos de acceso al correo,
# la configuración de tu entorno de Python, o la presencia de otros problemas específicos de tu configuración de Outlook.
