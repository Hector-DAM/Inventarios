import pandas as pd
import os

def procesar_inventario(inventario_path, tabla_upc_path, tiendas_path, output_folder):
    """
    Procesa el archivo de inventario y genera los archivos de salida.
    :param inventario_path: Ruta del archivo de inventario cargado.
    :param tabla_upc_path: Ruta del archivo de la tabla de UPC.
    :param tiendas_path: Ruta del archivo de información de tiendas.
    :param output_folder: Carpeta donde se guardarán los archivos generados.
    :return: Rutas de los archivos generados.
    """
    try:
        # Cargar la tabla de UPC
        tabla_upc = pd.read_excel(tabla_upc_path)
        tabla_upc['UPC'] = tabla_upc['UPC'].astype(str)
        tabla_upc['UPC'] = tabla_upc['UPC'].str.replace(".0", "")

        # Cargar el inventario de la semana
        inventario = pd.read_excel(inventario_path)
        inventario["UPC"] = inventario["UPC"].astype(str)
        inventario["UPC"] = inventario["UPC"].str.replace(".0", "")

        # Cargar la información de las tiendas
        tiendas = pd.read_excel(tiendas_path)

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

        # Seleccionar las columnas necesarias y ordenar
        inventarioFinal = inventarioFinal[["BARCODE", "Tienda", "UPC", "EstiloColor", "Size", "Brand", "AVAILABLE"]]
        inventarioFinal = inventarioFinal.sort_values(by=["BARCODE", "Tienda"])

        # Seleccionar el 25% superior del inventario ordenado
        percent = 0.25
        num_registros = int(len(inventarioFinal) * percent)
        muestra = inventarioFinal.head(num_registros)

        # Crear la carpeta de salida si no existe
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Crear un archivo Excel con una hoja por tienda (propuesta agrupada)
        propuesta_agrupada_path = os.path.join(output_folder, "PropuestaConteo_PorTienda_Ajustada.xlsx")
        with pd.ExcelWriter(propuesta_agrupada_path) as writer:
            for tienda in muestra["Tienda"].unique():
                df_tienda = muestra[muestra["Tienda"] == tienda]
                df_propuesta = df_tienda.pivot_table(
                    index=["Tienda", "BARCODE", "EstiloColor", "Brand"],
                    columns="Size",
                    values="AVAILABLE",
                    fill_value=0
                ).reset_index()
                df_propuesta.columns = [f"Talla{col}" if str(col).isdigit() else col for col in df_propuesta.columns]
                df_propuesta.to_excel(writer, sheet_name=tienda, index=False)

        # Crear un archivo Excel con solo los UPC (propuesta de UPC)
        upc_lista_path = os.path.join(output_folder, "Lista_UPC_Propuesta.xlsx")
        upc_lista = muestra[["UPC"]].drop_duplicates()
        upc_lista.to_excel(upc_lista_path, index=False)

        # Crear un archivo Excel con una hoja por tienda (propuesta de UPC por tienda con cantidades)
        upc_por_tienda_path = os.path.join(output_folder, "Lista_UPC_PorTienda_ConCantidades.xlsx")
        with pd.ExcelWriter(upc_por_tienda_path) as writer:
            for tienda in muestra["Tienda"].unique():
                df_tienda = muestra[muestra["Tienda"] == tienda]
                upc_tienda = df_tienda.groupby("UPC", as_index=False)["AVAILABLE"].sum()
                upc_tienda.to_excel(writer, sheet_name=tienda, index=False)

        return {
            "propuesta_agrupada": propuesta_agrupada_path,
            "upc_lista": upc_lista_path,
            "upc_por_tienda": upc_por_tienda_path
        }

    except Exception as e:
        raise Exception(f"Error al procesar el archivo: {str(e)}")