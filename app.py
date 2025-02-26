from flask import Flask, request, render_template, send_file
import os
from procesar_inventario import procesar_inventario  # Importar la función de procesamiento

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Verificar si se cargó un archivo
        if "inventario_file" not in request.files:
            return "No se cargó ningún archivo", 400
        
        file = request.files["inventario_file"]
        if file.filename == "":
            return "Nombre de archivo inválido", 400
        
        # Guardar el archivo cargado
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        
        try:
            # Rutas de los archivos de referencia (pueden ser cargados o fijos)
            tabla_upc_path = "C:/DATA/Catalogo/TABLA UPC.xlsx"  # Puedes cambiarlo para cargarlo dinámicamente
            tiendas_path = "C:/DATA/Codigos/PropuestaInventarios/Tiendas M3.xlsx"  # Puedes cambiarlo para cargarlo dinámicamente

            # Procesar el archivo usando la función importada
            resultados = procesar_inventario(file_path, tabla_upc_path, tiendas_path, app.config['OUTPUT_FOLDER'])
            
            # Devolver el archivo generado para descargar (por ejemplo, la propuesta agrupada)
            return send_file(resultados["propuesta_agrupada"], as_attachment=True)
        
        except Exception as e:
            return f"Error al procesar el archivo: {str(e)}", 500
    
    # Mostrar el formulario de carga (GET)
    return render_template("index.html")

if __name__ == "__main__":
    # Crear las carpetas necesarias si no existen
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    if not os.path.exists(app.config['OUTPUT_FOLDER']):
        os.makedirs(app.config['OUTPUT_FOLDER'])
    
    # Iniciar la aplicación
    app.run(debug=False)