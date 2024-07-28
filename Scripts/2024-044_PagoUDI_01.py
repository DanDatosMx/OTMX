# %% [markdown]
# 2024-044
# Proceso de pago UDI | Paso 01

# %%
# LLamar bibliotecas
print('Trace | Import Library')
import pandas as pd
import numpy as np
from pathlib import Path
import os
from openpyxl import workbook,load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import glob


# %%
# Definir los directorios
PathC = "C:/OTMX"
NameItem = "2024-044"
Output_dir = Path(PathC) / "Outputs" / NameItem
Output_dir.mkdir(parents=True, exist_ok=True)

#Creamos la carpeta Temp para colocar los archivos modificados
#Temp_dir = Path(PathC) / "Inputs" / NameItem / "Temp"
#Temp_dir.mkdir(parents=True, exist_ok=True)

print(f'Define path: {PathC} | Name Item: {NameItem} ')

print('Successful folders creation')



# %% [markdown]
# Unificar archivos Atlas | Se requiere que sólo estén los puros archivos csv para unificar
# 

# %%
print('Trace | Unificar archivos de Atlas')
# Ruta de la carpeta que contiene los archivos Excel
PathAtlas = Path(PathC) / "Inputs" / NameItem / "Files Atlas"

# Obtener la lista de archivos en la carpeta
Inputs_f_Atlas = os.listdir(PathAtlas)
#list(PathAtlas.iterdir())

# Crear una lista vacía para almacenar los DataFrames individuales
dataframes = []



# %%
# Recorrer cada archivo en la lista
for archivo in Inputs_f_Atlas:
    # Verificar si el archivo es un archivo .csv
    if archivo.endswith(".csv"):
        # Construir la ruta completa del archivo
        ruta_archivo = os.path.join(PathAtlas, archivo)
        # Leer el archivo en un DataFrame de pandas
        df = pd.read_csv(ruta_archivo,index_col=0, encoding='latin-1',skiprows=3)
        # Obtener el número de filas en el archivo
        num_filas = df.shape[0]
        # Omitir la última fila
        df = df.iloc[:-2]
        #Eliminar cad archivo que ya se colocó en el DataFrame
        #os.remove(ruta_archivo)
        
        # Agregar el DataFrame a la lista
        dataframes.append(df)
     
        # Guardar el DataFrame modificado en el mismo archivo
        #df.to_csv(ruta_archivo, index=False)
        #df.to_csv(Temp_dir / archivo, index=False)


        #print(f"Archivo {archivo} editado y guardado exitosamente.")

# Combinar los DataFrames en uno solo
dfATLAS = pd.concat(dataframes, ignore_index=True)
dfATLAS.columns = dfATLAS.columns.str.strip()
dfATLAS.columns = dfATLAS.columns.str.replace(' ', '')



# %%
#Ordenar columnas
#dfATLAS = dfATLAS[['Ofna', 'Póliza', 'Rbo', 'Inicio','Término', 'Prima Neta','Prima Total', 'Asegurado', 'Agente UDI', 'N° Serie', 'Comisión']]


# %%

# Guardar los datos combinados en un nuevo archivo CSV
#datos_combinados.to_csv('datos_combinados.csv', index=False)
dfATLAS.to_excel(Path(PathC) / "Inputs" / NameItem / "EDC_ATLAS.xlsx", index=False)
        
print("Data cleaning process | Done")

# %%
# Recorrer cada archivo en la lista
#for archivo in Inputs_f_Atlas:
    # Verificar si el archivo es un archivo .csv
 #   if archivo.endswith(".csv"):
        # Construir la ruta completa del archivo
  #      ruta_archivo = os.path.join(PathAtlas, archivo)
        # Leer el archivo en un DataFrame de pandas
   #     df = pd.read_csv(ruta_archivo,index_col=0, encoding='latin-1',skiprows=3)
        # Obtener el número de filas en el archivo
    #    num_filas = df.shape[0]
        # Omitir la última fila
     #   df = df.iloc[:-2]
        #Eliminar cad archivo que ya se colocó en el DataFrame
        #os.remove(ruta_archivo)
        
     
        # Guardar el DataFrame modificado en el mismo archivo
        #df.to_csv(ruta_archivo, index=False)
      #  df.to_csv(Temp_dir / archivo, index=False)


       # print(f"Archivo {archivo} editado y guardado exitosamente.")

        
#print("Proceso limpieza de datos | Complete")


# %%
#os.chdir(Temp_dir)
#extension = 'csv'
#todos_los_archivos = [i for i in glob.glob('*.{}'.format(extension))]

#combina todos los archivos de la lista
#combinado_csv = pd.concat([pd.read_csv(f,encoding='latin-1') for f in todos_los_archivos ])
#Seleccionar ciertas columnas
#combinado_csv.columns


# %%

#Ordenar columnas
#dfATLAS=combinado_csv[['Ofna', 'PÃ³liza', 'Rbo', 'Inicio','TÃ©rmino', 'Prima Neta','Prima Total', 
 #                            'Asegurado', 'Agente UDI', 'NÂ° Serie', ' ComisiÃ³n', ' IVA ', ' Total']]


# %%
#exporta a xlsx
#combinado_csv.to_csv(Path(PathC) / "Inputs" / NameItem / "EDC_ATLAS.csv", index=False, encoding='utf-8-sig')
#combinado_csv.to_excel(Path(PathC) / "Inputs" / NameItem / "EDC_ATLAS.xlsx", index=False)


# %% [markdown]
# Edición de los archivos por aseguradora

# %%
#Ruta archivos aseguradora
PathInsurance = Path(PathC) / "Inputs" / NameItem


# %% [markdown]
# Catálogos

# %%
print('Trace | Catalogues Files')
#Carga de catálogo general
dfcat = pd.read_excel(Path(PathC) / "Catalogues" / "c_2024_044.xlsx", sheet_name='Catalogo')
dfcat.columns


# %%

#Carga de catálogo HDI
dfcatHDI = pd.read_excel(Path(PathC) / "Catalogues" / "c_2024_044.xlsx", sheet_name='c_HDI')
dfcatHDI = dfcatHDI.rename(columns={'Clave Perfil':'HDI'})
dfcatHDI = dfcatHDI.rename(columns={'Agencia':'Agencia Mazda'})
dfcat.columns = dfcat.columns.str.replace("_x", "")
dfcat.columns = dfcat.columns.str.replace("_y", "")

dfcatHDI["HDI"] = pd.to_numeric(dfcatHDI["HDI"], errors="coerce").astype(pd.Int64Dtype())


dfcatHDI.columns


# %% [markdown]
# ATLAS

# %%
print('Trace | Atlas Files')

# Leer base de datos
dfATLAS = pd.read_excel(PathInsurance / "EDC_ATLAS.xlsx")
# Eliminar los espacios en blanco en todas las celdas
dfATLAS = dfATLAS.applymap(lambda x: x.strip() if isinstance(x, str) else x)


# %%
# Insertar columnas
dfATLAS['Aseguradora'] = 'ATLAS'

#Renomrar columna para el cruce
dfATLAS = dfATLAS.rename(columns={'AgenteUDI':'ATLAS'})

#Hacer los cruces 
dfATLAS = pd.merge(dfATLAS,dfcat,on='ATLAS',how='left')

#Campos calculados
## Necesitamos combertir los la columna flotante con un número tipo float 
dfATLAS['Ofna'] = dfATLAS['Ofna'].apply(lambda x: x.strip() if isinstance(x, str) else x)

# Separar la columna por el guion medio
#dfATLAS[['Ini Vig', 'Fin Vig']] = dfATLAS['Periodo'].str.split('-', expand=True)
dfATLAS['Póliza'] = dfATLAS['Ofna'].apply(str) + '-' + dfATLAS['Póliza'].apply(str)



# %%
dfATLAS = dfATLAS[['Grupo','Agencia Mazda','Aseguradora','Póliza','Rbo','Asegurado','N°Serie','Inicio','Término','PrimaNeta','PrimaTotal','Comisión','IVA','Total','MARSH']]


# %%
#Remobrar para Layout completo
dfATLAS = dfATLAS.rename(columns={'Grupo':'Grupo',
                                  'Agencia Mazda':'Dealer',
                                  'Aseguradora':'Aseguradora',
                                  'Póliza':'Poliza',
                                  'Rbo':'Recibo',
                                  'Asegurado':'Asegurado',
                                  'N°Serie':'Serie',
                                  'Inicio':'Ini_Vig',
                                  'Término':'Fin_Vig',
                                  'PrimaNeta':'Prima_Neta',
                                  'PrimaTotal':'Prima_Total',
                                  'Comisión':'UDI_Neto',
                                  'IVA':'UDI_IVA',
                                  'Total':'UDI_Total',
                                  'MARSH':'Alias'                                                         
                                  })



dfATLAS.columns

# %% [markdown]
# CHUBB

# %%
print('Trace | CHUBB Files')

#Leer base de datos
dfCHUBB = pd.read_excel(PathInsurance / "EDC_CHUBB.xlsx")

#Insertar columnas
dfCHUBB['Aseguradora'] = 'CHUBB'

#Renomrar columna para el cruce
dfCHUBB = dfCHUBB.rename(columns={'Conducto':'CHUBB'})

#Hacer los cruces 
dfCHUBB = pd.merge(dfCHUBB,dfcat,on='CHUBB',how='left')

#Campos calculados
## Necesitamos combertir los la columna flotante con un número tipo float 
dfCHUBB['ClaveId'] = dfCHUBB['ClaveId'].apply(lambda x: x.strip() if isinstance(x, str) else x)


# Separar la columna por el guion medio
#dfCHUBB[['Ini Vig', 'Fin Vig']] = dfCHUBB['Periodo'].str.split('-', expand=True)
dfCHUBB['Póliza'] = dfCHUBB['ClaveId'].apply(str) + '-' + dfCHUBB['PolizaId'].apply(str)



# %%
# Separar la columna por el guion medio
#dfCHUBB[['Ini Vig', 'Fin Vig']] = dfCHUBB['Periodo'].str.split('-', expand=True)
dfCHUBB['Póliza'] = dfCHUBB['ClaveId'].apply(str) + '-' + dfCHUBB['PolizaId'].apply(str)

#Definir columnas
dfCHUBB.columns

dfCHUBB = dfCHUBB[['Grupo','Agencia Mazda','Aseguradora','Póliza','ReciboId','AseguradoNombre','Serie',
                   'InicioVigencia','FinVigencia','PrimaNeta','PrimaTotal','UDINeta','UDIIva','UDITotal','MARSH']]






# %%
#Remobrar para Layout completo
dfCHUBB = dfCHUBB.rename(columns={'Grupo':'Grupo',
                                  'Agencia Mazda':'Dealer',
                                  'Aseguradora':'Aseguradora',
                                  'Póliza':'Poliza',
                                  'ReciboId':'Recibo',
                                  'AseguradoNombre':'Asegurado',
                                  'Serie':'Serie',
                                  'InicioVigencia':'Ini_Vig',
                                  'FinVigencia':'Fin_Vig',
                                  'PrimaNeta':'Prima_Neta',
                                  'PrimaTotal':'Prima_Total',
                                  'UDINeta':'UDI_Neto',
                                  'UDIIva':'UDI_IVA',
                                  'UDITotal':'UDI_Total',
                                  'MARSH':'Alias'                                                         
                                  })

dfCHUBB.columns

# %% [markdown]
# GNP

# %%
print('Trace | GNP Files')

#Leer base de datos
dfGNP = pd.read_excel(PathInsurance / "EDC_GNP.xlsx")
#Insertar columnas calculadas
dfGNP['Aseguradora'] = 'GNP'
#Renomrar columna para el cruce
dfGNP = dfGNP.rename(columns={'Codigo_Intermediario':'GNP'})

#Hacer los cruces
dfGNP = pd.merge(dfGNP,dfcat,on='GNP',how='left')

#Definir columnas
dfGNP.columns


# %%

dfGNP = dfGNP[['Grupo','PTO. VTA','Aseguradora','Poliza','No_Fraccion','Asegurado','Serie',
          'Inicio_vigencia', 'Fin_vigencia','Prima_neta','Prima_total','Udi neto','Udi IVA', 'Udi total','MARSH']]



# %%
#Remobrar para Layout completo
dfGNP = dfGNP.rename(columns={'Grupo':'Grupo',
                              'PTO. VTA':'Dealer',
                              'Aseguradora':'Aseguradora',
                              'Poliza':'Poliza',
                              'No_Fraccion':'Recibo',
                              'Asegurado':'Asegurado',
                              'Serie':'Serie',
                              'Inicio_vigencia':'Ini_Vig',
                              'Fin_vigencia':'Fin_Vig',
                              'Prima_neta':'Prima_Neta',
                              'Prima_total':'Prima_Total',
                              'Udi neto':'UDI_Neto',
                              'Udi IVA':'UDI_IVA',
                              'Udi total':'UDI_Total',
                              'MARSH':'Alias'                            
                              })

dfGNP.columns


# %% [markdown]
# HDI CONTADO

# %%
print('Trace | HDI CONTADO Files')

#Leer base de datos
dfHDI_C = pd.read_excel(PathInsurance / "EDC_HDI_CONTADO.xlsx")
#Insertar columnas calculadas
dfHDI_C['Aseguradora'] = 'HDI CONTADO'
#dfHDI_C['Grupo'] = ''
#dfHDI_C['MARSH'] = ''

#Renomrar columna para el cruce
dfHDI_C = dfHDI_C.rename(columns={'Nip Agente':'HDI'})

#Hacer los cruces
dfHDI_C = pd.merge(dfHDI_C,dfcatHDI,on='HDI',how='left')


dfHDI_C.columns = dfHDI_C.columns.str.replace("_x", "")
dfHDI_C.columns = dfHDI_C.columns.str.replace("_y", "")

#Insertar columnas
IVA = 0.16
#Campos calculados
dfHDI_C['MontoAgencia'] = dfHDI_C['Prima Neta'] * 0.22
dfHDI_C['UDI IVA'] = dfHDI_C['MontoAgencia'] * IVA
dfHDI_C['UDI Total'] = dfHDI_C['MontoAgencia'] + dfHDI_C['UDI IVA']

##Se tiene que sacar los primero dígitos
dfHDI_C['Ofc'] = dfHDI_C['Oficina'].astype(str).str[:3]
## Necesitamos combertir los la columna flotante con un número tipo float 
dfHDI_C['Ofc'] = dfHDI_C['Ofc'].apply(lambda x: x.strip() if isinstance(x, str) else x)
dfHDI_C['Certificado'] = dfHDI_C['Certificado'].apply(lambda x: x.strip() if isinstance(x, str) else x)
## Separar la columna por el guion medio
dfHDI_C['Póliza'] = dfHDI_C['Ofc'].apply(str) + '-' + dfHDI_C['Póliza'].apply(str) + '-' + dfHDI_C['Certificado'].apply(str) 

#dfHDI_C.columns

#Hacer el cruce con el cat
#Hacer los cruces
dfHDI_C = pd.merge(dfHDI_C,dfcat,on='Agencia Mazda',how='left')

dfHDI_C.columns = dfHDI_C.columns.str.replace("_x", "")
dfHDI_C.columns = dfHDI_C.columns.str.replace("_y", "")

dfHDI_C.columns



# %%
#Seleccionar colummas
dfHDI_C = dfHDI_C[['Grupo', 'Agencia Mazda','Aseguradora','Póliza','Certificado','Asegurado','Serie',
                  'Inicio de Vigencia','Fin de Vigencia','Prima Neta','Prima Total','Monto UDI','UDI IVA','UDI Total','MARSH']]

# %%
#Remobrar para Layout completo
dfHDI_C = dfHDI_C.rename(columns={ 'Grupo':'Grupo',
                                'Agencia Mazda':'Dealer',
                                'Aseguradora':'Aseguradora',
                                'Póliza':'Poliza',
                                'Certificado':'Recibo',  
                                'Asegurado':'Asegurado',  
                                'Serie':'Serie',  
                                'Inicio de Vigencia':'Ini_Vig',
                                'Fin de Vigencia':'Fin_Vig',
                                'Prima Neta':'Prima_Neta',
                                'Prima Total':'Prima_Total',
                                'Monto UDI':'UDI_Neto',
                                'UDI IVA':'UDI_IVA',
                                'UDI Total':'UDI_Total',
                                'MARSH':'Alias'
                                  })

# %% [markdown]
# HDI FINANCIADO

# %%
print('Trace | HDI FINANCIADO Files')

#Leer base de datos
dfHDI_F = pd.read_excel(PathInsurance / "EDC_HDI_FINANCIADO.xlsx")
#Insertar columnas calculadas
dfHDI_F['Aseguradora'] = 'HDI FINANCIADO'
#dfHDI_F['Grupo'] = ''
#dfHDI_F['MARSH'] = ''


#Renomrar columna para el cruce
dfHDI_F = dfHDI_F.rename(columns={'Nip Agente':'HDI'})

#Hacer los cruces
dfHDI_F = pd.merge(dfHDI_F,dfcatHDI,on='HDI',how='left')


dfHDI_F.columns = dfHDI_F.columns.str.replace("_x", "")
dfHDI_F.columns = dfHDI_F.columns.str.replace("_y", "")

#Insertar columnas
IVA = 0.16
#Campos calculados
dfHDI_F['MontoAgencia'] = dfHDI_F['Prima Neta'] * 0.22
dfHDI_F['UDI IVA'] = dfHDI_F['MontoAgencia'] * IVA
dfHDI_F['UDI Total'] = dfHDI_F['MontoAgencia'] + dfHDI_F['UDI IVA']

##Se tiene que sacar los primero dígitos
dfHDI_F['Ofc'] = dfHDI_F['Oficina'].astype(str).str[:3]
## Necesitamos combertir los la columna flotante con un número tipo float 
dfHDI_F['Ofc'] = dfHDI_F['Ofc'].apply(lambda x: x.strip() if isinstance(x, str) else x)
dfHDI_F['Certificado'] = dfHDI_F['Certificado'].apply(lambda x: x.strip() if isinstance(x, str) else x)
## Separar la columna por el guion medio
dfHDI_F['Póliza'] = dfHDI_F['Ofc'].apply(str) + '-' + dfHDI_F['Póliza'].apply(str) + '-' + dfHDI_F['Certificado'].apply(str) 

#dfHDI_F.columns

#Hacer el cruce con el cat
#Hacer los cruces
dfHDI_F = pd.merge(dfHDI_F,dfcat,on='Agencia Mazda',how='left')

dfHDI_F.columns = dfHDI_F.columns.str.replace("_x", "")
dfHDI_F.columns = dfHDI_F.columns.str.replace("_y", "")

dfHDI_F.columns



# %%
#Seleccionar colummas
dfHDI_F = dfHDI_F[['Grupo', 'Agencia Mazda','Aseguradora','Póliza','Certificado','Asegurado','Serie',
                  'Inicio de Vigencia','Fin de Vigencia','Prima Neta','Prima Total','Monto UDI','UDI IVA','UDI Total','MARSH']]

# %%
#Remobrar para Layout completo
dfHDI_F = dfHDI_F.rename(columns={ 'Grupo':'Grupo',
                                'Agencia Mazda':'Dealer',
                                'Aseguradora':'Aseguradora',
                                'Póliza':'Poliza',
                                'Certificado':'Recibo',  
                                'Asegurado':'Asegurado',  
                                'Serie':'Serie',  
                                'Inicio de Vigencia':'Ini_Vig',
                                'Fin de Vigencia':'Fin_Vig',
                                'Prima Neta':'Prima_Neta',
                                'Prima Total':'Prima_Total',
                                'Monto UDI':'UDI_Neto',
                                'UDI IVA':'UDI_IVA',
                                'UDI Total':'UDI_Total',
                                'MARSH':'Alias'
                                  })

# %% [markdown]
# QUALITAS

# %%
print('Trace | QUALITAS Files')

#Leer base de datos
dfQUA = pd.read_excel(PathInsurance / "EDC_QUALITAS.xlsx")
#limpiar la base de datos

dfQUA = dfQUA.dropna(subset=['Unnamed: 8'])

dfQUA = dfQUA[dfQUA['Unnamed: 8'].str.strip() != ""]
dfQUA = dfQUA[dfQUA['Unnamed: 8'].str.strip() != "MONEDA NAL"]
dfQUA = dfQUA[dfQUA['Unnamed: 8'] != 0]

dfQUA.columns = dfQUA.iloc[1,]

dfQUA = dfQUA[dfQUA['AGENTE'].str.strip() != "AGENTE"]

dfQUA['AGENCIA'] = dfQUA['AGENTE'].where(dfQUA['AGENTE'].str.startswith('AGENCIA NUM:')).ffill()

dfQUA = dfQUA.dropna(subset=['POLIZA'])
dfQUA['AGENCIA'] = dfQUA['AGENCIA'].str.replace("AGENCIA NUM: ","")

#Renomrar columna para el cruce
dfQUA = dfQUA.rename(columns={'AGENCIA':'QUALITAS'})


# %%
#Insertar columnas calculadas
dfQUA['Aseguradora'] = 'QUALITAS'
dfQUA['Ini_Vig'] = ''
dfQUA['Fin_Vig'] = ''
dfQUA['Prima_Neta'] = dfQUA['IMPORTE']
dfQUA['Prima_Total'] = ''
dfQUA['UDI_Neto'] = dfQUA['HON.']
dfQUA['UDI_IVA'] = dfQUA['IVA_PAG']
dfQUA['UDI_Total'] = dfQUA['ABONO'] - dfQUA['CARGO']

#Insertar columnas calculadas
dfQUA['Aseguradora'] = 'QUALITAS'

#cambio col QUA
dfQUA['QUALITAS'] = dfQUA['QUALITAS'].astype('int64')

dfQUA.to_excel(Path(PathC) / "Outputs" / NameItem / "EDC_QUALITAS_V2.xlsx", index=False)



# %%
#Hacer los cruces
dfQUA = pd.merge(dfQUA,dfcat,on='QUALITAS',how='left')


# %%
dfQUA.columns
dfQUA = dfQUA[['Grupo','Agencia Mazda','Aseguradora','POLIZA','RECIBO','CONCEPTO','NIV',
          'Ini_Vig','Fin_Vig','Prima_Neta','Prima_Total','UDI_Neto','UDI_IVA','UDI_Total','MARSH']]



# %%
#Remobrar para Layout completo
dfQUA = dfQUA.rename(columns={  'Grupo':'Grupo',
                                'Agencia Mazda':'Dealer',
                                'Aseguradora':'Aseguradora',
                                'POLIZA':'Poliza',
                                'RECIBO':'Recibo',
                                'CONCEPTO':'Asegurado',
                                'NIV':'Serie',
                                'Inicio de Vigencia':'Ini_Vig',
                                'Fin de Vigencia':'Fin_Vig',
                                'Prima Neta':'Prima_Neta',
                                'Prima Total':'Prima_Total',
                                'Monto UDI':'UDI_Neto',
                                'UDI IVA':'UDI_IVA',
                                'UDI Total':'UDI_Total',
                                'MARSH':'Alias'
                              })

dfQUA.to_excel(Path(PathC) / "Outputs" / NameItem / "EDC_QUALITAS_V3.xlsx", index=False)

dfQUA.columns


# %% [markdown]
# Base de datos completas

# %%
print('Trace | Merged Files')

# Unir los DataFrames de las aseguradora en uno solo
df_merged = pd.concat([dfATLAS,dfCHUBB, dfGNP, dfHDI_C,dfHDI_F,dfQUA], ignore_index=True)

#Agregar columnas restantes
df_merged['Canal']=''
df_merged['Tipo Poliza']=''
df_merged['Limite']=''

#
df_merged.columns

df_merged = df_merged[['Grupo', 'Dealer', 'Aseguradora', 'Canal', 'Tipo Poliza','Poliza', 'Recibo', 'Asegurado',
       'Serie', 'Ini_Vig', 'Fin_Vig', 'Prima_Neta', 'Prima_Total', 'UDI_Neto',
       'UDI_IVA', 'UDI_Total', 'Limite','Alias']]




# %%
from datetime import date

# Supongamos que tienes un dataframe llamado "df" que deseas guardar en un archivo de Excel
# Puedes obtener la fecha actual utilizando la biblioteca datetime
fecha_actual = date.today().strftime("%d-%m-%Y")

# Luego, puedes utilizar el método to_excel() de Pandas para guardar el dataframe en un archivo de Excel
nombre_archivo = f"Files_{fecha_actual}.xlsx"
df_merged.to_excel(Path(PathC) / "Outputs" / NameItem / nombre_archivo, index=False)

# El dataframe se guardará en un archivo de Excel con el nombre que incluye la fecha actual


# %%
#exporta a csv
#combinado_csv.to_csv(Path(PathC) / "Outputs" / NameItem / "File_Atlas_All.csv", index=False, encoding='utf-8-sig')

#df_merged.to_excel(Path(PathC) / "Outputs" / NameItem / "Files_All.xlsx", index=False)

# %%
# Guardar el DataFrame unido en un archivo de Excel
print("Los DataFrames se han unido y la información se ha guardado en 'resultado.xlsx'.")



