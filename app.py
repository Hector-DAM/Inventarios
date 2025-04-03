from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from generaPropuesta import generar_propuesta  # Importar la función de procesamiento

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'

# Cargar la tabla de UPC al inicio de la aplicación
upc_path = "./TABLA UPC.xlsx"  # Ruta del archivo de TABLA UPC
tabla_upc = pd.read_excel(upc_path)
tabla_upc['UPC'] = tabla_upc['UPC'].astype(str)
tabla_upc['UPC'] = tabla_upc['UPC'].str.replace(".0", "", regex=False)

tiendas_path = "./Tiendas M3.xlsx"  # Ruta del archivo de Tiendas M3
tiendas = pd.read_excel(tiendas_path)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == "POST":

        # Obtener la tienda seleccionada y el último barcode
        tienda_seleccionada = request.form["tienda"]
        ultimo_barcode = request.form["ultimo_barcode"]

        # Verificar si se cargó un archivo
        if "inventario_file" not in request.files:
            return "No se ha cargado ningún archivo de inventario.", 400
        
        file = request.files["inventario_file"]
        if file.filename == "":
            return "No se ha seleccionado ningún archivo.", 400
        
        # Guardar el archivo cargado
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)

        try:
            # Procesar el archivo usando la función importada
            zip_path = generar_propuesta(
                file_path,
                tabla_upc,
                app.config['OUTPUT_FOLDER'],
                tienda_seleccionada,
                ultimo_barcode
            )

            # Devolver el archivo ZIP generado
            return send_file(zip_path, as_attachment=True)
        
        except Exception as e:
            return f"Error al procesar el archivo: {str(e)}", 500
        
    # Mostrar el formulario de carga (GET)
    lista_tiendas = tiendas["Tienda"].unique().tolist()
    return render_template("index.html", tiendas=lista_tiendas)

# Asegurarse que las carpetas existen y correr la app
if __name__ == "__main__":
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    if not os.path.exists(app.config['OUTPUT_FOLDER']):
        os.makedirs(app.config['OUTPUT_FOLDER'])
    
    app.run(host="0.0.0.0", port=5000, debug=True)
