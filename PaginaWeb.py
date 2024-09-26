from flask import Flask, request, render_template_string, send_file
import pymongo
from bson.objectid import ObjectId
import gridfs
from io import BytesIO

app = Flask(__name__)

# Conexión a MongoDB
client = pymongo.MongoClient('mongodb://localhost:27017/')
db = client['Prueba']
pdfs_collection = db['Pdfs']
fs = gridfs.GridFS(db)  # Usamos GridFS para gestionar archivos grandes

# Página HTML en Flask
@app.route('/')
def index():
    return render_template_string('''
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Búsqueda de Archivos PDF</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                form { margin-bottom: 20px; }
                input, button { margin: 5px 0; }
                .result { border: 1px solid #ccc; padding: 10px; margin-bottom: 10px; }
            </style>
        </head>
        <body>

            <h1>Búsqueda de Archivos PDF</h1>

            <form action="/buscar" method="GET">
                <label for="ruc">RUC:</label><br>
                <input type="text" id="ruc" name="ruc"><br><br>

                <label for="senores">Señor(es):</label><br>
                <input type="text" id="senores" name="senores"><br><br>

                <label for="fecha_emision">Fecha de Emisión (dd/mm/yyyy):</label><br>
                <input type="text" id="fecha_emision" name="fecha_emision" placeholder="13/08/2024"><br><br>

                <label for="descripcion">Descripción:</label><br>
                <input type="text" id="descripcion" name="descripcion" placeholder="Parte de la descripción"><br><br>

                <button type="submit">Buscar</button>
            </form>

            <div id="results">
                <!-- Resultados de la búsqueda aparecerán aquí -->
            </div>

        </body>
        </html>
    ''')

# Endpoint para buscar PDFs en MongoDB
@app.route('/buscar', methods=['GET'])
def buscar():
    query = {}

    # Filtrar por RUC si se proporciona
    ruc = request.args.get('ruc')
    if ruc:
        query['RUC'] = ruc

    # Filtrar por Señor(es) si se proporciona
    senores = request.args.get('senores')
    if senores:
        query['Señor(es)'] = {'$regex': senores, '$options': 'i'}  # Búsqueda insensible a mayúsculas

    # Filtrar por Fecha de Emisión si se proporciona
    fecha_emision = request.args.get('fecha_emision')
    if fecha_emision:
        query['Fecha de Emisión'] = fecha_emision

    # Filtrar por Descripción si se proporciona (coincidencia parcial)
    descripcion = request.args.get('descripcion')
    if descripcion:
        query['Descripción'] = {'$regex': descripcion, '$options': 'i'}  # Búsqueda insensible a mayúsculas y parcial

    # Buscar en MongoDB
    resultados = list(pdfs_collection.find(query))

    return render_template_string('''
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Búsqueda de Archivos PDF</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                form { margin-bottom: 20px; }
                input, button { margin: 5px 0; }
                .result { border: 1px solid #ccc; padding: 10px; margin-bottom: 10px; }
            </style>
        </head>
        <body>

            <h1>Búsqueda de Archivos PDF</h1>

            <form action="/buscar" method="GET">
                <label for="ruc">RUC:</label><br>
                <input type="text" id="ruc" name="ruc"><br><br>

                <label for="senores">Señor(es):</label><br>
                <input type="text" id="senores" name="senores"><br><br>

                <label for="fecha_emision">Fecha de Emisión (dd/mm/yyyy):</label><br>
                <input type="text" id="fecha_emision" name="fecha_emision" placeholder="13/08/2024"><br><br>

                <label for="descripcion">Descripción:</label><br>
                <input type="text" id="descripcion" name="descripcion" placeholder="Parte de la descripción"><br><br>

                <button type="submit">Buscar</button>
            </form>

            <div id="results">
                {% if resultados %}
                    {% for result in resultados %}
                        <div class="result">
                            <p><strong>RUC:</strong> {{ result['RUC'] }}</p>
                            <p><strong>Señor(es):</strong> {{ result['Señor(es)'] }}</p>
                            <p><strong>Fecha de Emisión:</strong> {{ result['Fecha de Emisión'] }}</p>
                            <p><strong>Descripción:</strong> {{ result['Descripción'] }}</p>
                            <a href="/descargar/{{ result['gridfs_id'] }}">Descargar PDF</a>
                        </div>
                    {% endfor %}
                {% else %}
                    <p>No se encontraron resultados.</p>
                {% endif %}
            </div>

        </body>
        </html>
    ''', resultados=resultados)

# Endpoint para descargar archivos PDF desde GridFS
@app.route('/descargar/<id>')
def descargar(id):
    try:
        # Buscar el archivo por su gridfs_id
        pdf_file = fs.find_one({"_id": ObjectId(id)})
        
        if pdf_file:
            # Crear un flujo de bytes para enviar el archivo al cliente
            pdf_stream = BytesIO(pdf_file.read())
            pdf_stream.seek(0)
            return send_file(
                pdf_stream,
                as_attachment=True,
                download_name=f"{pdf_file.filename}",  # Nombre de archivo
                mimetype='application/pdf'
            )
        else:
            return "Archivo no encontrado", 404
    except Exception as e:
        return str(e), 500

if __name__ == '__main__':
    app.run(debug=True)
