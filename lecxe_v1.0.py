# -*- coding: utf-8 -*-
"""
Created on Mon Feb 26 17:54:11 2024

@author: jpoot
"""
import pandas as pd
import glob

def menu():
    
   while True:
        print("Bienvenido a lecxe")
        print("Para unificar documentos de excel, introduce 1")
        print("Para agregar una hoja a varios documentos de excel, introduce 2")
        print("Para salir: 6")
    
        userSelection = int(input("Por favor intruduce tu seleccion: "))
        
        if userSelection == 1:
            unificarDocumentos()
        elif userSelection == 6: 
            print("Cerrando el programa...")
            break
        else:
            print("No es una seleccion valida")
            menu()
    


# Directorio donde se encuentran los archivos
    
def unificarDocumentos():
    
    nombreHoja = input("Por favor, introduce el nombre de la hoja que quieres unificar: ")
    print("Recuerda que todos los documentos tienen que contener la misma hoja para poder ser unificados")
    nombreDocumento = input("Por favor, introduce el nombre que te gustaría que tenga tu documento consolidado: ")
    directorio = 'C:/Users/jpoot/Documents/Practicas/Unificar_sharedmail'
    
    # Patrón para buscar archivos específicos (por ejemplo, todos los archivos Excel)
    patron = '*.xlsx'
    
    # Obtener la lista de nombres de archivo que coinciden con el patrón
    archivos = glob.glob(directorio + '/' + patron)
    
    # Lista para almacenar los DataFrames de los archivos que contienen la hoja "Standard_Shared_Accounts"
    dataframes = []
    
    # Iterar sobre cada archivo
    for archivo in archivos:
        # Intentar leer el archivo y verificar si la hoja "Standard_Shared_Accounts" existe
        try:
            df = pd.read_excel(archivo, sheet_name = nombreHoja)
            dataframes.append(df)
        except Exception as e:
            print(f"No se pudo leer la hoja 'Standard_DLs' en el archivo {archivo}: {e}")
    
    # Combinar los DataFrames en uno solo si hay al menos uno
    if dataframes:
        combinados = pd.concat(dataframes, ignore_index=True)
        
        # Guardar el DataFrame combinado en un archivo Excel
        combinados.to_excel(nombreDocumento + ".xlsx", index=False)
    else:
        print("No se encontraron archivos con la hoja 'Progress_'.")


menu()