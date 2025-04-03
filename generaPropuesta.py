import pandas as pd
import os
import zipfile
from datetime import date

def generar_propuesta(inventario_path, tabla_upc, output_folder, tienda_seleccionada, ultimo_barcode):
    """
    Procesa el archivo de inventario y genera los archivos de salida.

    :param inventario_path: Ruta del archivo de inventario cargado.
    :param tabla_upc: DataFrame de la tabla de UPC (ya cargado).
    :param output_folder: Carpeta donde se guardarán los archivos generados.
    :param tienda_seleccionada: Tienda seleccionada por el usuario.
    :param ultimo_barcode: Último barcode de la propuesta (puede estar vacío).
    :return: Ruta del archivo ZIP generado.
    """
    hoy = date.today()
    fecha = hoy.strftime("%d-%m-%Y")

    try:
        # Cargar archivo de inventario
        df_inv = pd.read_excel(inventario_path)

        # Cambiamos el nombre de las columnas en el inventario
        df_inv = df_inv.rename(columns={"STYLE": "STYLE_INV"})

        # Convertir UPC a string para evitar problemas al unir tablas
        df_inv['UPC'] = df_inv['UPC'].astype(str).str.replace(".0", "", regex=False)
        tabla_upc['UPC'] = tabla_upc['UPC'].astype(str).str.replace(".0", "", regex=False)

        # Unir los datos con la tabla de UPC
        df_junto = pd.merge(df_inv, tabla_upc[["UPC", "Brand", "STYLE", "Color Name"]], on="UPC", how="left")

        # Crear la columna BARCODE combinando estilo y color
        df_junto["BARCODE"] = df_junto["STYLE_INV"].astype(str) + "-" + df_junto["COLOR_CODE"].astype(str)

        # Filtrar la marca CALZANETTO si existe la columna "Brand"
        if "Brand" in df_junto.columns:
            df_junto = df_junto[df_junto["Brand"] != "CALZANETTO"]

        # Filtrar por la tienda seleccionada
        if "STORE_NAME" in df_junto.columns:
            df_junto = df_junto[df_junto["STORE_NAME"] == tienda_seleccionada]
        else:
            raise ValueError("La columna 'STORE_NAME' no se encuentra en el archivo de inventario.")

        # Ordenar por BARCODE
        df_junto = df_junto.sort_values(by="BARCODE")

        # Seleccionar registros
        if not ultimo_barcode:
            # Si no hay último barcode, tomar el 25% de los registros
            num_registros = int(len(df_junto) * 0.25)
            muestra = df_junto.head(num_registros)
        else:
            # Si hay último barcode, filtrar por ese valor
            muestra = df_junto[df_junto["BARCODE"] == ultimo_barcode]

        # Seleccionar columnas necesarias
        columnas_necesarias = ["BARCODE", "STORE_NAME", "UPC", "STYLE", "Color Name", "SIZE_DESC", "Brand", "STORE_ON_HAND"]
        muestra = muestra[columnas_necesarias]

        # Crear carpeta de salida si no existe
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Generar archivo Excel de propuesta
        propuesta_path = os.path.join(output_folder, f"Propuesta_{tienda_seleccionada}_{fecha}.xlsx")
        with pd.ExcelWriter(propuesta_path) as writer:
            df_propuesta = muestra.pivot_table(
                index=["STORE_NAME", "BARCODE", "STYLE", "Color Name", "Brand"],
                columns=["SIZE_DESC"],
                values="STORE_ON_HAND",
                fill_value=0,
            ).reset_index()

            df_propuesta.columns = [f"Talla{col}" if isinstance(col, (int, float)) else col for col in df_propuesta.columns]
            df_propuesta.to_excel(writer, index=False)

        # Generar archivo Excel de cantidades por UPC
        upc_cantidades_path = os.path.join(output_folder, f"UPC_Cantidades_{tienda_seleccionada}_{fecha}.xlsx")
        upc_cantidades = muestra.groupby("UPC", as_index=False)["STORE_ON_HAND"].sum()
        upc_cantidades.to_excel(upc_cantidades_path, index=False)

        # Crear archivo ZIP con los dos archivos
        zip_path = os.path.join(output_folder, f"Propuesta_{tienda_seleccionada}_{fecha}.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(propuesta_path, os.path.basename(propuesta_path))
            zipf.write(upc_cantidades_path, os.path.basename(upc_cantidades_path))

        return zip_path

    except Exception as e:
        raise Exception(f"Error al procesar el archivo: {str(e)}")
