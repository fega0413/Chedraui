from imap_tools import MailBox, AND
import re
import pyodbc
import smtplib
from email.message import EmailMessage

# Configura las credenciales y detalles del servidor IMAP y SMTP
IMAP_HOST = 'imap.noip.com'
SMTP_HOST = 'mail.noip.com'
SMTP_PORT = 587
USERNAME = 'afernandez@wellcom.biz'
PASSWORD = 'Example.0413'

# Configura la cadena de conexión para SQL Server
SQL_SERVER = '172.28.246.73'
DATABASE = 'lgtdb'
USERNAME_SQL = 'lgt'
PASSWORD_SQL = 'lgtbd2017!!'

# Función para conectarse a la base de datos y ejecutar el query
def update_order_status(order_ids):
    try:
        connection = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={SQL_SERVER};DATABASE={DATABASE};UID={USERNAME_SQL};PWD={PASSWORD_SQL}')
        cursor = connection.cursor()
        order_ids_str = "', '".join(order_ids)
        query = f"UPDATE TBL_LGT_OC SET ESTATUS_OC_ID = 90 WHERE OC_ID IN ('{order_ids_str}');"
        
        cursor.execute(query)
        connection.commit()
        
        print(f"Órdenes actualizadas correctamente: {order_ids}")
        return True
    except pyodbc.Error as e:
        print(f"Error en la base de datos: {e}")
        return False
    finally:
        if 'connection' in locals():
            connection.close()

# Expresión regular para encontrar OC_IDs (10 dígitos)
oc_pattern = re.compile(r'\b\d{10}\b')

# Actualizar la función send_confirmation_email para manejar el campo 'References'
# Función para enviar un correo de confirmación como respuesta en el mismo hilo
def send_confirmation_email(to_list, cc_list, subject, original_body, oc_ids, message_id, references):
    # Crear el mensaje
    msg = EmailMessage()
    msg['Subject'] = f"Re: {subject}"  # Re: para indicar que es una respuesta
    msg['From'] = USERNAME
    msg['To'] = ', '.join(to_list)
    if cc_list:
        msg['Cc'] = ', '.join(cc_list)
    
    # Añadir los encabezados para que sea parte del hilo de correo
    msg['In-Reply-To'] = message_id
    msg['References'] = references

    # Crear la versión en texto sin formato del correo
    plain_text_content = (f"Se han actualizado las siguientes órdenes de compra a 'canceladas': {', '.join(oc_ids)}\n\n"
                            "Este es un correo de confirmación automática.\n\n"
                            f"Mensaje original:\n{original_body}")

    # Crear la versión en HTML del correo (si el correo original estaba en HTML)
    html_content = f"""
    <html>
        <body>
            <p>Se han actualizado las siguientes órdenes de compra a 'canceladas': {', '.join(oc_ids)}</p>
            <p>Este es un correo de confirmación automática.</p>
            <hr>
            <p><strong>Mensaje original:</strong></p>
            {original_body}  <!-- Mantenemos el contenido original en HTML -->
        </body>
    </html>
    """

    # Adjuntar ambas versiones (texto sin formato y HTML) al mensaje
    msg.set_content(plain_text_content)  # Versión de texto plano
    msg.add_alternative(html_content, subtype='html')  # Versión HTML

    # Enviar el correo
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(USERNAME, PASSWORD)
            server.send_message(msg)
            print(f"Correo de confirmación enviado a: {to_list} y CC: {cc_list}")
    except Exception as e:
        print(f"Error al enviar el correo: {e}")


# Función principal para procesar correos y actualizar órdenes
def process_emails():
    with MailBox(IMAP_HOST).login(USERNAME, PASSWORD) as mailbox:
        # Buscar correos no leídos con asunto "ELIMINACION LEGADOS" o "CANCELACIÓN"
        for msg in mailbox.fetch(AND(seen=False), mark_seen=False):
            if 'ELIMINACION LEGADOS' in msg.subject or 'CANCELACIÓN' in msg.subject:
                email_body = msg.text
                oc_ids = oc_pattern.findall(email_body)
                
                if oc_ids:
                    print(f"Órdenes encontradas: {oc_ids}")
                    # Actualizar las órdenes de compra
                    if update_order_status(oc_ids):
                        # Extraer lista de destinatarios para responder a todos
                        to_list = [msg.from_]
                        cc_list = msg.cc or []
                        
                        # Obtener el Message-ID del correo original
                        message_id_tuple = msg.headers.get('message-id')
                        if isinstance(message_id_tuple, tuple):
                            message_id = message_id_tuple[0]  # Obtener el primer valor de la tupla
                        else:
                            message_id = message_id_tuple  # Si no es una tupla, úsalo directamente
                        
                        # Obtener el encabezado 'References' del correo original, si está disponible
                        references = msg.headers.get('references', '')
                        
                        # Si el 'References' está vacío, usar el 'Message-ID' original, si no, concatenar con él
                        if references:
                            references = f"{references} {message_id}"
                        else:
                            references = message_id
                        
                        # Enviar el correo de confirmación a todos los involucrados, en el mismo hilo
                        #send_confirmation_email(to_list, cc_list, msg.subject, email_body, oc_ids, message_id, references)


# Ejecutar la función principal
process_emails()