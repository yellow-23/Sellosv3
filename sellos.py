#!/home/yellow/Escritorio/Sellos/venv/bin/python3
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from collections import defaultdict
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Seleccione el archivo Excel",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    return file_path

try:
    # Select and read Excel file
    filepath = select_excel_file()
    if not filepath:
        print("No se seleccionó ningún archivo. El programa se cerrará.")
        exit()
        
    print(f"Archivo seleccionado: {filepath}")

    # Read Excel with expected columns
    expected_columns = [
        'Area Datos', 'Título', 'InforReporte', 'Desarrollador', 
        'Sello', 'Data Owner', 
        'Data Steward 1', 'Data Steward 2', 'Data Steward 3',
        'Data Steward 4', 'Data Steward 5', 'Data Steward 6'
    ]

    df = pd.read_excel(filepath, engine='openpyxl')
    df['Sello'] = df['Sello'].fillna('')

    # Filter reports without Business seal
    df_filtered = df[~df['Sello'].str.contains('Negocio', na=False)]

    # Group by Area Datos (no renaming needed)
    reportes_por_dominio = defaultdict(list)
    for index, row in df_filtered.iterrows():
        dominio = row['Area Datos']
        reportes_por_dominio[dominio].append(row)

    # Group by Data Owner and Stewards
    reportes_por_owner = defaultdict(list)
    for index, row in df_filtered.iterrows():
        data_owner = row['Data Owner']
        reportes_por_owner[data_owner].append(row)

    # Crear DataFrame para Excel
    excel_data = []
    for _, row in df_filtered.iterrows():
        # Obtener todos los Data Stewards para esta fila
        steward_columns = [col for col in df.columns if 'Data Steward' in col]
        steward_data = {f'Data Steward {i+1}': row[col] for i, col in enumerate(steward_columns)}
        
        excel_data.append({
            'Dominio': row['Area Datos'],
            'Título': row['Título'],  # Cambiado de ID a Title
            'Área PBI': row['InforReporte'],
            'Responsable': row['Desarrollador'],
            'Sellos Actuales': row['Sello'],
            'Estado': 'Pendiente' if pd.isna(row['Sello']) or 'Negocio' not in str(row['Sello']) else 'Completo',
            'Data Owner': row['Data Owner'],
            **steward_data  # Agregar todos los Data Stewards al diccionario
        })

    # Convertir a DataFrame y escribir Excel
    df_output = pd.DataFrame(excel_data)
    
    # Reordenar las columnas para que los Data Stewards estén después del Data Owner
    column_order = ['Dominio', 'Título', 'Área PBI', 'Responsable', 'Sellos Actuales', 'Estado', 
                   'Data Owner'] + [f'Data Steward {i+1}' for i in range(len(steward_columns))]
    df_output = df_output[column_order]

    output_path = '/home/yellow/Escritorio/Sellos/Reportes con sello de negocio pendientes.xlsx'

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_output.to_excel(writer, sheet_name='Reportes Pendientes', index=False)
        worksheet = writer.sheets['Reportes Pendientes']
        
        # Ajustar ancho de columnas
        for idx, col in enumerate(df_output.columns):
            max_length = max(df_output[col].astype(str).apply(len).max(), len(col)) + 2
            worksheet.column_dimensions[chr(65 + idx)].width = max_length

    print(f"Archivo Excel generado exitosamente en: {output_path}")

    # Configuración del correo
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    sender_email = "datadriven@cmpc.cl"   #"datadriven@cmpc.cl"    #"cflores.practica@cmpc.com"
    app_password = "ccgu zixq lzme xmsr"   #"ccgu zixq lzme xmsr"  #"ywsb sfgz fmyf qdsg"

    def formatear_sellos(sello_texto):
        # Lista de sellos posibles
        sellos = []
        # Dividir el texto por punto y coma o coma si existe
        if isinstance(sello_texto, str):
            # Dividir por punto y coma o coma y eliminar espacios extras
            sello_lista = [s.strip() for s in sello_texto.replace(',', ';').split(';')]
            
            for sello in sello_lista:
                if 'Seguridad' in sello:
                    sellos.append('Seguridad')
                if 'Tecnico' in sello or 'Tecnologia' in sello or 'Tecnología' in sello:
                    sellos.append('Tecnología')
                if 'Negocio' in sello:
                    sellos.append('Negocio')

        # Eliminar duplicados y ordenar
        sellos = sorted(list(set(sellos)))
        return ' ; '.join(sellos) if sellos else 'Sin sellos'

    # Modificar la agrupación para asegurar que todos los reportes se incluyan
    reportes_agrupados = defaultdict(dict)
    for index, row in df_filtered.iterrows():
        dominio = row['Area Datos']
        data_owner = row['Data Owner']  # Nombre exacto como está en el Excel
        
        # Inicializar la lista de reportes para este dominio si no existe
        if dominio not in reportes_agrupados[data_owner]:
            reportes_agrupados[data_owner][dominio] = []
        
        # Agregar el reporte a la lista correspondiente
        reportes_agrupados[data_owner][dominio].append(row)

    def crear_contenido_html(reportes_por_dominio, owner_email):
        contenido_dominios = ""
        dominios_ordenados = sorted(reportes_por_dominio.keys())
        
        for dominio in dominios_ordenados:
            reportes = reportes_por_dominio[dominio]
            reportes_ordenados = sorted(reportes, key=lambda x: x['Título'])  # Ordenar por Title en lugar de ID
            
            tabla_reportes = """
            <table class="reporte-table">
                <tr>
                    <th>Área PBI</th>
                    <th>Título</th>
                    <th>Responsable</th>
                    <th>Sellos Actuales</th>
                </tr>
            """
            
            for reporte in reportes_ordenados:
                tabla_reportes += f"""
                <tr>
                    <td>{reporte['InforReporte']}</td>
                    <td>{reporte['Título']}</td>
                    <td>{reporte['Desarrollador']}</td>
                    <td>{formatear_sellos(str(reporte['Sello']))}</td>
                </tr>
                """
            
            tabla_reportes += "</table>"
            
            dominio_html = f"""
            <div class="dominio-section">
                <h3 class="dominio-title">Dominio: {dominio}</h3>
                {tabla_reportes}
            </div>
            """
            contenido_dominios += dominio_html 

        template_path = os.path.join(os.path.dirname(__file__), 'email_template.html')
        with open(template_path, 'r', encoding='utf-8') as file:
            template = file.read()

        template = template.replace('{contenido_reportes}', contenido_dominios)
        return template.replace('{owner_email}', owner_email)

    # Establecer una única conexión SMTP antes del bucle de envío
    try:
        print(f"Conectando a Gmail como {sender_email}...")
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, app_password)
        print("Autenticación exitosa!")

        # Preguntar si se desea enviar correo de prueba
        while True:
            test_option = input("\n¿Desea enviar un correo de prueba primero? (s/n): ").lower()
            if test_option in ['s', 'n']:
                break
            print("Por favor ingrese 's' para sí o 'n' para no.")

        # En la sección del correo de prueba
        if test_option == 's':
            carolina_email = "carolina.reydeduarte@cmpc.cl"
            print("\n=== ENVIANDO CORREO DE PRUEBA A CAROLINA ===")
            
            # Crear correo de prueba con TODOS los reportes
            test_msg = MIMEMultipart('alternative')
            test_msg['From'] = sender_email
            test_msg['To'] = carolina_email
            test_msg['Subject'] = "[PRUEBA] Reportes pendientes de asignación de sello de Negocio"

            # Crear versión HTML del contenido de prueba con TODOS los reportes
            html_content = crear_contenido_html(reportes_agrupados, carolina_email)
            test_msg.attach(MIMEText(html_content, 'html'))

            # Enviar correo de prueba
            server.send_message(test_msg, to_addrs=[carolina_email])
            print(f"Correo de prueba enviado a Carolina ({carolina_email})")
            
            # Mostrar opciones al usuario
            print("\n=== OPCIONES DE ENVÍO ===")
            print("1: Continuar con el envío a todos los Data Owners")
            print("2: Cancelar el envío")
            
            while True:
                try:
                    opcion = input("\nIngrese su opción (1 o 2): ")
                    if opcion == "1":
                        print("\n=== COMENZANDO ENVÍO MASIVO ===")
                        break
                    elif opcion == "2":
                        print("\nEnvío cancelado por el usuario")
                        server.quit()
                        exit()
                    else:
                        print("Opción no válida. Por favor ingrese 1 o 2.")
                except Exception as e:
                    print(f"Error en la entrada: {str(e)}")
                    print("Por favor ingrese 1 o 2.")
        else:
            print("\n=== COMENZANDO ENVÍO MASIVO ===")

        # Mostrar resumen antes de enviar
        print("\nSe enviarán correos a los siguientes Data Owners:")
        for owner in reportes_agrupados.keys():
            if pd.notna(owner) and '@' in str(owner):
                print(f"- {owner}")

        # Continuar con el envío normal a todos los destinatarios
        # En la sección de envío a Data Owners
        for owner, dominios in reportes_agrupados.items():
            if pd.notna(owner) and '@' in str(owner):
                msg = MIMEMultipart('alternative')
                msg['From'] = sender_email
                msg['To'] = owner
                
                # Preparar lista de destinatarios en copia
                cc_recipients = []
                
                # Buscar todos los Data Stewards correspondientes de múltiples columnas
                steward_columns = [col for col in df.columns if 'Data Steward' in col]
                for col in steward_columns:
                    steward_email = df[df['Data Owner'] == owner][col].iloc[0] if not df[df['Data Owner'] == owner].empty else None
                    if pd.notna(steward_email) and '@' in str(steward_email):
                        cc_recipients.append(steward_email)
                
                # Mantener también a Carolina en copia
                carolina_email = "carolina.reydeduarte@cmpc.cl"
                if carolina_email not in cc_recipients:
                    cc_recipients.append(carolina_email)

                # Eliminar duplicados manteniendo el orden
                cc_recipients = list(dict.fromkeys(cc_recipients))
                
                # Agregar todos los CC al mensaje si hay destinatarios
                if cc_recipients:
                    msg['Cc'] = ','.join(cc_recipients)
                
                msg['Subject'] = "Reportes pendientes de asignación de sello de Negocio"

                # Crear versión HTML del contenido
                html_content = crear_contenido_html(dominios, owner)
                msg.attach(MIMEText(html_content, 'html'))

                try:
                    # Obtener todos los destinatarios (To + Cc)
                    all_recipients = [owner] + cc_recipients
                    
                    server.send_message(msg, to_addrs=all_recipients)
                    print(f"Correo enviado exitosamente a {owner}")
                    if cc_recipients:
                        print(f"Con copia a: {', '.join(cc_recipients)}")
                except Exception as e:
                    print(f"Error al enviar correo a {owner}: {str(e)}")

        # Cerrar la conexión después de enviar todos los correos
        server.quit()
        print("Conexión SMTP cerrada exitosamente")

    except smtplib.SMTPAuthenticationError:
        print("Error de autenticación con Gmail")
    except Exception as e:
        print(f"Error inesperado en la conexión SMTP: {str(e)}")

except FileNotFoundError:
    print(f"Error: No se encontró el archivo en la ruta: {filepath}")
except Exception as e:
    print(f"Error: Ocurrió un error al procesar el archivo: {str(e)}")