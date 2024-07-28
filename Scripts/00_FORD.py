# %%
#Importar librerías 

import pandas as pd
import os
from pathlib import Path


PathC = "C:/OTMX"
NameItem = "00_FORD"

# %%
xlsx_dir = Path(PathC) / "Inputs"

# Define & create output directory
output_dir = Path(PathC) / "Outputs" / NameItem
output_dir.mkdir(parents=True, exist_ok=True)


# %%
File = "00_FORD.xlsx"

PathFile = xlsx_dir / File
df = pd.read_excel(PathFile)
df.columns


# %%
# Eliminar los espacios en blanco de la columna "DISTRIBUIDOR"
#df['AGENCIA'] = df['AGENCIA'].str.strip()

# Obtener los valores únicos de la columna "DISTRIBUIDOR"
valores_unicos = df['AGENCIA'].unique()


# %%
# Crear un archivo separado para cada valor único de "DISTRIBUIDOR"
for valor in valores_unicos:

    # Filtrar el DataFrame original por el valor único de "DISTRIBUIDOR"
    df_filtrado = df[df['AGENCIA'] == valor]
    
    # Agregar la columna "MONTO" al DataFrame filtrado
    df_filtrado['MONTO'] = df['MONTO']    # Reemplaza esto con tus propios valores de monto

    # Agregar una columna de conteo uno por uno en la primera columna
    df_filtrado.insert(0, 'EXHIBICION', range(1, len(df_filtrado) + 1))

    # Calcular la suma de la columna "MONTO"
    suma_monto = df_filtrado['MONTO'].sum()

    # Dar formato de número a la suma del MONTO
    suma_monto_formatted = '{:,.2f}'.format(suma_monto)

     # Agregar la suma del MONTO al final del DataFrame
    df_filtrado.loc[df_filtrado.shape[1], 'MONTO'] = suma_monto_formatted
    df_filtrado.loc[df_filtrado.shape[1], 'MES'] = 'TOTAL ANTICIPO'


    # Guardar el DataFrame filtrado con la columna "MONTO" agregada en un nuevo archivo
    nombre_archivo = f'{valor}.xlsx'  # Puedes cambiar el formato del nombre del archivo si lo deseas
    df_filtrado.to_excel(output_dir / nombre_archivo, index=False)
    

print("Archivos separados y generados exitosamente.")


