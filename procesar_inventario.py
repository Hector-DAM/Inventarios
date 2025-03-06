import pandas as pd
import os
import zipfile

def procesar_inventario(inventario_path, tabla_upc, tiendas, output_folder, tienda_seleccionada, ultimo_barcode):
    """
    Procesa el archivo de inventario y genera los archivos de salida.
    :param inventario_path: Ruta del archivo de inventario cargado.
    :param tabla_upc: DataFrame de la tabla de UPC (ya cargado).
    :param tiendas: DataFrame de la tabla de tiendas (ya cargado).
    :param output_folder: Carpeta donde se guardarán los archivos generados.
    :param tienda_seleccionada: Tienda seleccionada por el usuario.
    :param ultimo_barcode: Último barcode de la propuesta (puede estar vacío).
    :return: Ruta del archivo ZIP generado.
    """
    try:
        # Cargar el inventario de la semana
        inventario = pd.read_excel(inventario_path)
        inventario["UPC"] = inventario["UPC"].astype(str)
        inventario["UPC"] = inventario["UPC"].str.replace(".0", "")

        # Asegurar que no haya cantidades negativas en la columna AVAILABLE
        inventario["AVAILABLE"] = inventario["AVAILABLE"].apply(lambda x: max(x, 0))

        # Filtrar solo el inventario de tiendas
        inventario = inventario[inventario["WH"] == "XRS"]

        # Hacer un merge con la tabla de UPC
        inventario = pd.merge(inventario, tabla_upc, how="left", on="UPC")

        # Hacer un merge con la tabla de tiendas
        inventarioFinal = pd.merge(inventario, tiendas, how="left", on="STORE")

        # Crear la columna BARCODE y EstiloColor
        inventarioFinal["BARCODE"] = inventarioFinal["STYLE M3"] + inventarioFinal["Color Code"]
        inventarioFinal["EstiloColor"] = inventarioFinal["STYLE_y"] + "-" + inventarioFinal["Color Name"]

        # Filtrar la marca CALZANETTO
        inventarioFinal = inventarioFinal[inventarioFinal["Brand"] != "CALZANETTO"]

        # Filtrar por la tienda seleccionada
        inventarioFinal = inventarioFinal[inventarioFinal["Tienda"] == tienda_seleccionada]

        # Ordenar los barcodes de A a Z
        inventarioFinal = inventarioFinal.sort_values(by="BARCODE")

        # Si el último barcode está vacío, tomar el 25% superior
        if not ultimo_barcode:
            percent = 0.25
            num_registros = int(len(inventarioFinal) * percent)
            muestra = inventarioFinal.head(num_registros)
        else:
            # Filtrar por barcode (solo barcodes mayores al último barcode)
            muestra = inventarioFinal[inventarioFinal["BARCODE"] > ultimo_barcode]

        # Seleccionar las columnas necesarias
        muestra = muestra[["BARCODE", "Tienda", "UPC", "EstiloColor", "Size", "Brand", "AVAILABLE"]]

        # Crear la carpeta de salida si no existe
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Crear un archivo Excel con la propuesta para la tienda seleccionada
        propuesta_path = os.path.join(output_folder, f"Propuesta_{tienda_seleccionada}.xlsx")
        with pd.ExcelWriter(propuesta_path) as writer:
            df_propuesta = muestra.pivot_table(
                index=["Tienda", "BARCODE", "EstiloColor", "Brand"],
                columns="Size",
                values="AVAILABLE",
                fill_value=0
            ).reset_index()
            df_propuesta.columns = [f"Talla{col}" if str(col).isdigit() else col for col in df_propuesta.columns]
            df_propuesta.to_excel(writer, index=False)

        # Crear un archivo ZIP con el archivo generado
        zip_path = os.path.join(output_folder, f"Propuesta_{tienda_seleccionada}.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(propuesta_path, os.path.basename(propuesta_path))

        return zip_path  # Devolver la ruta del archivo ZIP

    except Exception as e:
        raise Exception(f"Error al procesar el archivo: {str(e)}")
