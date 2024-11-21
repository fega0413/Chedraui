from imap_tools import MailBox, AND
import pyodbc
from datetime import datetime

# Configuración de credenciales y detalles del servidor
HOST = 'imap.noip.com'
USERNAME = 'afernandez@wellcom.biz'
PASSWORD = 'Example.0413'

# Configuración de conexión a SQL Server
SQL_SERVER = '172.28.246.73'
DATABASE = 'lgtdb'
USERNAME_SQL = 'lgt'
PASSWORD_SQL = 'lgtbd2017!!'

DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y"]  # Formatos válidos de fecha

def connect_to_db():
    """
    Conecta a la base de datos SQL Server y devuelve la conexión.
    """
    connection_string = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={DATABASE};"
        f"UID={USERNAME_SQL};"
        f"PWD={PASSWORD_SQL}"
    )
    conn = pyodbc.connect(connection_string)
    print("Conexión a la base de datos establecida exitosamente.")
    return conn

def parse_date(date_string):
    """
    Intenta analizar una fecha utilizando varios formatos conocidos.
    Si el formato no es válido, lanza una excepción.
    """
    for date_format in DATE_FORMATS:
        try:
            return datetime.strptime(date_string, date_format).strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise ValueError(f"Formato de fecha inválido: {date_string}")

def update_orders_in_db(orders):
    """
    Actualiza las órdenes de compra en la base de datos SQL Server
    y muestra un resumen con el total de órdenes y cuántas se actualizaron.
    """
    conn = connect_to_db()
    cursor = conn.cursor()

    print("Actualizando órdenes en la base de datos...")
    updated_orders = []
    failed_orders = []  # Para registrar órdenes que fallaron

    for order in orders:
        order_number, start_date, end_date = order
        
        try:
            # Intentar analizar y formatear las fechas
            start_date_parsed = parse_date(start_date)
            end_date_parsed = parse_date(end_date) + " 23:59:59"  # Ajustar la hora de fin

            # Consultas SQL para actualizar
            sql_update_oc = f"""
            UPDATE TBL_LGT_OC
            SET FECHA_VIGENCIA_INI = '{start_date_parsed}',
                FECHA_VIGENCIA_FIN = '{end_date_parsed}'
            WHERE OC_ID = '{order_number}';
            """
            sql_update_lpn = f"""
            UPDATE TBL_LGT_LPN
            SET FECHA_VIGENCIA_INI = '{start_date_parsed}',
                FECHA_VIGENCIA_FIN = '{end_date_parsed}'
            WHERE OC_ID = '{order_number}';
            """

            # Ejecutar las consultas y verificar si se actualizó al menos una fila
            cursor.execute(sql_update_oc)
            if cursor.rowcount > 0:
                updated_orders.append(order_number)  # Registrar órdenes actualizadas

            cursor.execute(sql_update_lpn)

        except ValueError as e:
            print(f"Error al procesar la orden {order_number}: {e}")
            failed_orders.append(order_number)

    # Confirmar los cambios si hubo actualizaciones
    if updated_orders:
        conn.commit()
        print(f"Resumen de actualizaciones:")
        print(f" - Total de órdenes procesadas: {len(orders)}")
        print(f" - Total de órdenes actualizadas: {len(updated_orders)}")
        print(f" - Órdenes fallidas: {len(failed_orders)}")
        print("Órdenes actualizadas exitosamente:")
        for order in updated_orders:
            print(f" - OC_ID: {order}")
    else:
        conn.rollback()
        print("No se actualizaron órdenes.")

    if failed_orders:
        print("Órdenes con errores en el formato de fechas:")
        for order in failed_orders:
            print(f" - OC_ID: {order}")

    cursor.close()
    conn.close()

def extract_info_from_text(text):
    """
    Extrae la información de las órdenes de compra y sus fechas de vigencia 
    desde el contenido de texto plano del correo electrónico.
    """
    orders = []
    lines = text.splitlines()
    start_date = None
    end_date = None

    for i, line in enumerate(lines):
        line = line.strip()

        if line.startswith('FECHA_VIGENCIA_INICIO:'):
            if i + 1 < len(lines):
                start_date = lines[i + 1].strip()
        elif line.startswith('FECHA_VIGENCIA_FIN:'):
            if i + 1 < len(lines):
                end_date = lines[i + 1].strip()
        elif line.isdigit():
            order_number = line
            if start_date and end_date:
                orders.append((order_number, start_date, end_date))

    return orders

# Procesar correos
with MailBox(HOST).login(USERNAME, PASSWORD, initial_folder='INBOX') as mailbox:
    emails = mailbox.fetch(AND(seen=False, subject='SOLICITUD DE VIGENCIA'), mark_seen=False)

    for email in emails:
        print("Asunto:", email.subject)
        print("Remitente:", email.from_)

        text_content = email.text
        orders = extract_info_from_text(text_content)

        if orders:
            # Actualizar órdenes en la base de datos y mostrar el resumen
            update_orders_in_db(orders)
