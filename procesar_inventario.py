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

        # Depuración: Verificar las primeras filas del DataFrame
        print("Primeras filas del archivo de inventario:")
        print(inventario.head())

        # Eliminar la primera fila (no contiene datos útiles)
        inventario = inventario.drop([0])

        # Convertir la segunda fila en los títulos de las columnas
        nuevas_columnas = inventario.iloc[0]
        inventario = pd.DataFrame(inventario[1:], columns=nuevas_columnas)

        # Eliminar la tercera fila (contiene subtotales)
        inventario = inventario.drop([1])

        # Resetear el índice
        inventario = inventario.reset_index(drop=True)

        # Depuración: Verificar las columnas y las primeras filas después del procesamiento
        print("Columnas después de procesar:", inventario.columns.tolist())
        print("Primeras filas después de procesar:")
        print(inventario.head())

        # Verificar si la columna 'UPC' existe
        if 'UPC' not in inventario.columns:
            raise ValueError("La columna 'UPC' no se encuentra en el archivo de inventario.")

        # Convertir la columna UPC a tipo string y eliminar ".0" si es necesario
        inventario["UPC"] = inventario["UPC"].astype(str)
        inventario["UPC"] = inventario["UPC"].str.replace(".0", "")

        # Asegurar que no haya cantidades negativas en la columna AVAILABLE
        if 'AVAILABLE' in inventario.columns:
            inventario["AVAILABLE"] = inventario["AVAILABLE"].apply(lambda x: max(x, 0))
        else:
            raise ValueError("La columna 'AVAILABLE' no se encuentra en el archivo de inventario.")

        # Filtrar solo el inventario de tiendas
        if 'WH' in inventario.columns:
            inventario = inventario[inventario["WH"] == "XRS"]
            print(f"Registros después de filtrar por WH == 'XRS': {len(inventario)}")
        else:
            raise ValueError("La columna 'WH' no se encuentra en el archivo de inventario.")

        # Hacer un merge con la tabla de UPC
        inventario = pd.merge(inventario, tabla_upc, how="left", on="UPC")
        print(f"Registros después de merge con tabla UPC: {len(inventario)}")

        # Hacer un merge con la tabla de tiendas
        inventarioFinal = pd.merge(inventario, tiendas, how="left", on="STORE")
        print(f"Registros después de merge con tabla tiendas: {len(inventarioFinal)}")

        # Crear la columna BARCODE y EstiloColor
        if 'STYLE M3' in inventarioFinal.columns and 'Color Code' in inventarioFinal.columns:
            inventarioFinal["BARCODE"] = inventarioFinal["STYLE M3"] + inventarioFinal["Color Code"]
        else:
            raise ValueError("Las columnas 'STYLE M3' o 'Color Code' no se encuentran en el archivo de inventario.")

        if 'STYLE_y' in inventarioFinal.columns and 'Color Name' in inventarioFinal.columns:
            inventarioFinal["EstiloColor"] = inventarioFinal["STYLE_y"] + "-" + inventarioFinal["Color Name"]
        else:
            raise ValueError("Las columnas 'STYLE_y' o 'Color Name' no se encuentran en el archivo de inventario.")

        # Filtrar la marca CALZANETTO
        if 'Brand' in inventarioFinal.columns:
            inventarioFinal = inventarioFinal[inventarioFinal["Brand"] != "CALZANETTO"]
            print(f"Registros después de filtrar por Brand != 'CALZANETTO': {len(inventarioFinal)}")
        else:
            raise ValueError("La columna 'Brand' no se encuentra en el archivo de inventario.")

        # Filtrar por la tienda seleccionada
        if 'Tienda' in inventarioFinal.columns:
            inventarioFinal = inventarioFinal[inventarioFinal["Tienda"] == tienda_seleccionada]
            print(f"Registros después de filtrar por Tienda == '{tienda_seleccionada}': {len(inventarioFinal)}")
        else:
            raise ValueError("La columna 'Tienda' no se encuentra en el archivo de inventario.")

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

        print(f"Registros en la muestra final: {len(muestra)}")

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

        # Crear un archivo Excel con la lista de UPC y cantidades
        upc_cantidades_path = os.path.join(output_folder, f"UPC_Cantidades_{tienda_seleccionada}.xlsx")
        upc_cantidades = muestra.groupby("UPC", as_index=False)["AVAILABLE"].sum()
        upc_cantidades.to_excel(upc_cantidades_path, index=False)

        # Crear un archivo ZIP con los archivos generados
        zip_path = os.path.join(output_folder, f"Propuesta_{tienda_seleccionada}.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(propuesta_path, os.path.basename(propuesta_path))
            zipf.write(upc_cantidades_path, os.path.basename(upc_cantidades_path))

        return zip_path  # Devolver la ruta del archivo ZIP

    except Exception as e:
        raise Exception(f"Error al procesar el archivo: {str(e)}")
