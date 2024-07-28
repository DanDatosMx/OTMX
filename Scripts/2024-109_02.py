# %%
#Importar librerías 

import pandas as pd
import os
from pathlib import Path
import openpyxl


PathC = "C:/OTMX"
NameItem = "2024-109"

# %%
xlsx_dir = Path(PathC) / "Outputs" / NameItem

# Define & create output directory
output_dir = Path(PathC) / "Outputs" / NameItem / 'Files'
output_dir.mkdir(parents=True, exist_ok=True)


# %%
File = "Base_Forvia.xlsx"

PathFile = xlsx_dir / File
df = pd.read_excel(PathFile)
df.columns


# %%
# Eliminar los espacios en blanco de la columna
df['SITIO'] = df['SITIO'].str.replace(' ', '')

# Obtener los valores únicos de la columna
valores_unicos = df['SITIO'].unique()


# %%
# Crear un archivo separado para cada valor único de "DISTRIBUIDOR"
for valor in valores_unicos:

    # Filtrar el DataFrame original por el valor único de "DISTRIBUIDOR"
    df_filtrado = df[df['SITIO'] == valor]

    # Guardar el DataFrame filtrado con la columna "MONTO" agregada en un nuevo archivo
    nombre_archivo = f'{valor}.xlsx'  # Puedes cambiar el formato del nombre del archivo si lo deseas

    df_filtrado.to_excel(output_dir / nombre_archivo, index=False)
    #Edición del archivo de excel
    wb = openpyxl.load_workbook(output_dir / nombre_archivo)
    sheet = wb.active
    ancho = 20
    # Recorre todas las columnas y establece el ancho deseado
    for columna in sheet.columns:
        sheet.column_dimensions[columna[0].column_letter].width = ancho
        #sheet.column_dimensions[columna[0].column_letter].fill = fill

    
    wb.save(output_dir / nombre_archivo)
    print('Archivo separado y generado...'   f'{valor}.xlsx')


    

print("Archivos separados y generados exitosamente.")


