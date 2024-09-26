import os
import re
import zipfile
import win32com.client
import datetime
import pymongo
import gridfs  # Importar GridFS
from PyPDF2 import PdfReader
import xml.etree.ElementTree as ET  # Para manejar XML

# Conexión a MongoDB
client = pymongo.MongoClient('mongodb://localhost:27017/')
db = client['Prueba']
pdfs_collection = db['Pdfs']
fs = gridfs.GridFS(db)  # Usamos GridFS para gestionar archivos grandes

# Ruta temporal para guardar los archivos adjuntos
temp_dir = r"C:\Users\USUARIO\Desktop\Temp"  # Cambia esta ruta si lo prefieres

# Asegúrate de que el directorio temporal exista
if not os.path.exists(temp_dir):
    os.makedirs(temp_dir)

# Conexión a Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 es la bandeja de entrada

# Intervalos de fechas
start_date = "22/09/2024"
end_date = "27/09/2024"

start_date_dt = datetime.datetime.strptime(start_date, "%d/%m/%Y")
end_date_dt = datetime.datetime.strptime(end_date, "%d/%m/%Y")
messages = inbox.Items
found_emails = 0

# Función para extraer atributos de un archivo PDF
def extraer_atributos_pdf(pdf_path):
    atributos = {
        "RUC": None,
        "Señor(es)": None,
        "Fecha de Emisión": None,
        "Descripción": None  # El valor de la descripción será extraído del XML o ZIP
    }
    
    # Abrir y leer el PDF
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        text = ""
        for page in reader.pages:
            text += page.extract_text()

        # Buscar los atributos dentro del texto
        ruc_match = re.search(r"RUC:\s*(\d+)", text)
        if ruc_match:
            atributos["RUC"] = ruc_match.group(1)
        
        cliente_match = re.search(r"Señor\(es\)\s*:\s*(.+)", text)
        if cliente_match:
            atributos["Señor(es)"] = cliente_match.group(1).strip()

        fecha_emision_match = re.search(r"Fecha de Emisión\s*:\s*(\d{2}/\d{2}/\d{4})", text)
        if fecha_emision_match:
            atributos["Fecha de Emisión"] = fecha_emision_match.group(1)

    return atributos

# Función para extraer atributos de un archivo XML
def extraer_descripcion_xml(xml_path):
    # Parsear el archivo XML y extraer la descripción
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Buscar la etiqueta <cbc:Description> y extraer el texto dentro de <![CDATA[]]>
    descripcion_element = root.find(".//{*}Description")
    if descripcion_element is not None and descripcion_element.text is not None:
        return descripcion_element.text.strip()

    return None

# Función para extraer archivos XML desde un archivo ZIP
def extraer_descripcion_desde_zip(zip_path):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for file_name in zip_ref.namelist():
            # Buscar archivos XML dentro del ZIP
            if file_name.lower().endswith('.xml'):
                # Extraer el archivo XML a un directorio temporal
                extracted_path = os.path.join(temp_dir, file_name)
                zip_ref.extract(file_name, temp_dir)
                print(f"Archivo XML extraído desde ZIP: {extracted_path}")
                
                # Extraer la descripción del XML
                descripcion = extraer_descripcion_xml(extracted_path)
                
                # Eliminar el archivo XML extraído después de leer
                os.remove(extracted_path)
                
                if descripcion:
                    return descripcion
    return None

# Función para guardar archivos PDF en GridFS y registrar los metadatos
def guardar_pdf_en_gridfs(file_path, atributos, asunto, received_time):
    # Abrir el archivo PDF en modo binario
    with open(file_path, "rb") as f:
        pdf_id = fs.put(f, filename=os.path.basename(file_path))  # Guardar en GridFS
    
    # Guardar los metadatos en la colección "Pdfs"
    pdfs_collection.insert_one({
        "gridfs_id": pdf_id,  # Referencia al archivo en GridFS
        "filename": os.path.basename(file_path),
        "asunto": asunto,
        "fecha_recepcion": received_time,
        "RUC": atributos.get("RUC"),
        "Señor(es)": atributos.get("Señor(es)"),
        "Fecha de Emisión": atributos.get("Fecha de Emisión"),
        "Descripción": atributos.get("Descripción")  # Guardar la descripción extraída del XML o ZIP
    })

# Iterar sobre cada correo
for message in messages:
    if message.ReceivedTime:
        received_time = message.ReceivedTime
        if received_time.tzinfo is not None:
            received_time = received_time.replace(tzinfo=None)
        # Filtrar correos dentro del intervalo de fechas
        if start_date_dt <= received_time <= end_date_dt:
            found_emails += 1
            print(f"Asunto: {message.Subject} - Fecha de recepción: {received_time}")
            
            # Variables para manejar el PDF y XML/ZIP del mismo correo
            pdf_file_path = None
            descripcion_xml = None

            # Si el correo tiene archivos adjuntos
            if message.Attachments.Count > 0:
                print(f"El correo tiene {message.Attachments.Count} archivo(s) adjunto(s)")
                for attachment in message.Attachments:
                    file_name = attachment.FileName.lower()
                    temp_file_path = os.path.join(temp_dir, file_name)

                    # Guardar el archivo temporalmente
                    attachment.SaveAsFile(temp_file_path)
                    print(f"Archivo guardado temporalmente: {temp_file_path}")

                    # Si es un archivo XML, extraer la descripción
                    if file_name.endswith('.xml'):
                        descripcion_xml = extraer_descripcion_xml(temp_file_path)
                        print(f"Descripción extraída del XML: {descripcion_xml}")
                        os.remove(temp_file_path)  # Eliminar el XML después de extraer la descripción

                    # Si es un archivo ZIP, extraer el XML y su descripción
                    elif file_name.endswith('.zip'):
                        descripcion_xml = extraer_descripcion_desde_zip(temp_file_path)
                        print(f"Descripción extraída del XML dentro del ZIP: {descripcion_xml}")
                        os.remove(temp_file_path)  # Eliminar el ZIP después de extraer el contenido

                    # Si es un archivo PDF, guardamos su ruta para procesarlo más tarde
                    elif file_name.endswith('.pdf'):
                        pdf_file_path = temp_file_path

                # Si se encontró un PDF y una descripción del XML o ZIP, procedemos a guardar el PDF
                if pdf_file_path:
                    # Extraer atributos del PDF
                    atributos = extraer_atributos_pdf(pdf_file_path)
                    
                    # Añadir la descripción extraída del XML o ZIP (si existe)
                    if descripcion_xml:
                        atributos["Descripción"] = descripcion_xml
                    
                    # Guardar el PDF y metadatos usando GridFS
                    guardar_pdf_en_gridfs(pdf_file_path, atributos, message.Subject, received_time)
                    print(f"Archivo PDF {os.path.basename(pdf_file_path)} guardado correctamente en MongoDB.")
                    
                    # Eliminar el archivo PDF temporal
                    os.remove(pdf_file_path)

if found_emails == 0:
    print("No se encontraron correos en el intervalo de fechas especificado.")

print("Proceso finalizado.")

