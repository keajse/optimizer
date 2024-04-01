import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
#import pyexcel
import pandas as pd
import numpy as np
import os
import datetime


# Fecha de expiración del programa
fecha_expiracion = datetime.date(2024, 6, 30)

# Verificar la fecha de expiración
fecha_actual = datetime.date.today()
if fecha_actual > fecha_expiracion:
    print("El programa ha expirado.")
    messagebox.showinfo("Error", "El programa ha expirado")
    # Aquí podrías mostrar un mensaje de error y salir del programa
    exit()

tiempo = (fecha_actual - fecha_expiracion).days
messageFin = f'El programa expirará en {tiempo} días'
messagebox.showinfo("Cuenta regresiva", messageFin)

def validar_numero(input):
    if input.isdigit():
        return True
    elif input == "":
        return True
    else:
        return False

def obtener_datos():
    global dato1, dato2, descuento
    
    dato1 = int(entrada1.get()) if entrada1.get().isdigit() else 0
    
    if dato1 == 0:
        dato1 = 6000  # Definir valor por defecto para Dato 1

    dato2 = int(entrada2.get()) if entrada2.get().isdigit() else 0
    if dato2 == 0:
        dato2 = 6500  # Definir valor por defecto para Dato 2

    descuento = int(entrada_descuento.get() if entrada_descuento.get().isdigit() else 0)

    #print("Dato 1:", dato1)
    #print("Dato 2:", dato2)
    #print("Ruta del archivo:", ruta_archivo)
    ventana.destroy()  # Cerrar la ventana al obtener datos

def seleccionar_archivo():
    global ruta_archivo
    ruta_archivo = filedialog.askopenfilename()
    ruta_archivo_entry.delete(0, tk.END)
    ruta_archivo_entry.insert(0, ruta_archivo)

# Crear la ventana
ventana = tk.Tk()
ventana.title("Optimizador de Aluminio")
ventana.minsize(width=300, height=180)
ventana.maxsize(width=350, height=200)

# Crear etiquetas y entradas para los datos
etiqueta1 = tk.Label(ventana, text="Perfil tradicional:")
etiqueta1.grid(row=0, column=0, padx=5, pady=5)

validacion_dato1 = ventana.register(validar_numero)
entrada1 = tk.Entry(ventana, validate="key", validatecommand=(validacion_dato1, '%P'))
entrada1.grid(row=0, column=1, padx=5, pady=5)

etiqueta2 = tk.Label(ventana, text="Perfil Europeo:")
etiqueta2.grid(row=1, column=0, padx=5, pady=5)

validacion_dato2 = ventana.register(validar_numero)
entrada2 = tk.Entry(ventana, validate="key", validatecommand=(validacion_dato2, '%P'))
entrada2.grid(row=1, column=1, padx=5, pady=5)

etiqueta_descuento = tk.Label(ventana, text="Descuento por corte (mm):")
etiqueta_descuento.grid(row=2, column=0, padx=5, pady=5)

entrada_descuento = tk.Entry(ventana, validate="key", validatecommand=(etiqueta_descuento, '%P'))
entrada_descuento.grid(row=2, column=1, padx=5, pady=5)

etiqueta_archivo = tk.Label(ventana, text="Ruta del archivo:")
etiqueta_archivo.grid(row=3, column=0, padx=5, pady=5)

ruta_archivo_entry = tk.Entry(ventana)
ruta_archivo_entry.grid(row=3, column=1, padx=5, pady=5)

# Botón para seleccionar el archivo
boton_archivo = tk.Button(ventana, text="Seleccionar Archivo", command=seleccionar_archivo)
boton_archivo.grid(row=4, column=0, padx=5, pady=5, sticky="ew")  # Centrar y expandir horizontalmente
boton_archivo.config(width=20)  # Ajustar el ancho del botón

# Botón para obtener los datos
boton_obtener = tk.Button(ventana, text="Obtener Datos", command=obtener_datos)
boton_obtener.grid(row=4, column=1, padx=5, pady=5, sticky="ew")  # Centrar y expandir horizontalmente
boton_obtener.config(width=20)  # Ajustar el ancho del botón

# Label como pie de página
pie_de_pagina = tk.Label(ventana, text="""Realizado por Laura Juliana Serrano García - MaKeajse 2024.
Recuerda: 'Hacer de lo Ordinario, algo Extraordinario'.
""", bg="lightgray", fg="black")
pie_de_pagina.grid(row=6, columnspan=2, padx=5, pady=5, sticky="ew")

# Variable para almacenar la ruta del archivo seleccionado
ruta_archivo = ""


# Ejecutar el bucle principal
ventana.mainloop()


def optimizar():
    if ruta_archivo:
        # Leer los datos del archivo Excel seleccionado
        df_init = pd.read_excel(ruta_archivo)
        
        #Renombrar el nombre de las columnas para evitar errores en la ejecución del código, ya que es obligatorio los siguientes nombres.
        df_init.rename(columns={df_init.columns[0]: 'Perfil'}, inplace=True )
        df_init.rename(columns={df_init.columns[1]: 'Medida Corte'}, inplace=True )
        df_init.rename(columns={df_init.columns[2]: 'Cant.'}, inplace=True )

        #Agregar columna que extraiga el primer caracter de la columna perfil, se tiene en cuenta que los perfiles tradicionales empiezan por una letra mientras que el europeo por número.
        df_init['Empieza'] = df_init['Perfil'].apply(lambda x: x[0] if isinstance(x, str) and len(x)>0 else None)
        df_init['MedidaAumentada'] = df_init['Medida Corte'] + descuento

        #Organizar los valores primero por perfiles A - Z y después por medidas de Mayor a Menor
        df = df_init.sort_values(by=['Perfil','MedidaAumentada'], ascending=[True, False])

        #Se imprime en consola la nueva dataFrame para verificar que se estén aplicando correctamente los criterios.
        #print(df)

        #Obtener referencias únicas por perfil
        referencias_unicas = np.array(df['Perfil'].unique())

        #Inicialización
        resultados = []
        sobrante= []
        hojaCorte = []
        

        #Ciclo para recorrer cada una de las filas que componen el array de las referencias únicas
        for idx, val in np.ndenumerate(referencias_unicas):
            
            MedidaTotal = 0
            totalPerfiles = 0
            sobr = 0   
            compr = 0        

            #Ciclo para recorrer las filas de la DataFrame 
            for indice, fila in df.iterrows():     
        

                # Verificamos la condición presente en la columna 'Condicion'
                if fila['Perfil'] == val:

                    if(fila['Empieza'] == None):
                        unidad = dato2
                    else:
                        unidad = dato1

                    #print(unidad)                   


                    #Con este ciclo se busca tener en cuenta el número de cortes por medidas y perfil.
                    for i in range(fila['Cant.']):

                        sobrante = sorted(sobrante, key=lambda x: (str(x['Perfil']), x['Sobrante']))   
                        #print(sobrante)
                        element_remove = None

                        for elem in sobrante:
                            if elem['Perfil'] == val:
                                if (fila['MedidaAumentada']) <= elem['Sobrante']:
                                    element_remove = elem

                                    compr = elem['Sobrante'] - fila['MedidaAumentada']

                                    if compr > 0:
                                        sobrante.append({
                                            'NumPerfil': elem['NumPerfil'],
                                            'Perfil': val,
                                            'Sobrante': compr
                                        }
                                        )
                                    hojaCorte.append({
                                        'NumPerfil': elem['NumPerfil'],
                                        'Perfil': val,
                                        'Medida Corte': fila['Medida Corte'],
                                        'Medida Aumentada': fila['MedidaAumentada']
                                    }
                                    )
                                    break
                                    
                                    
                        
                        if element_remove:
                            sobrante.remove(element_remove)

                        else:                                
                    
                            #print(f'{i}, {fila['Cant.']} - {unidad} vs {fila['Medida']}')   
                            
                            if((MedidaTotal + fila['MedidaAumentada']) <= unidad):
                                # Si la condición se cumple, con relación al nombre del perfil
                                MedidaTotal += fila['MedidaAumentada']
                                hojaCorte.append({
                                    'NumPerfil': totalPerfiles + 1,
                                    'Perfil': val,
                                    'Medida Corte': fila['Medida Corte'],
                                    'Medida Aumentada': fila['MedidaAumentada']
                                }
                                )

                            else:
                                totalPerfiles += 1
                                sobr = unidad - MedidaTotal
                                #print(f'{val} = {sobr}')
                                MedidaTotal = fila['MedidaAumentada']

                                if sobr > 0:
                                    sobrante.append({
                                        'NumPerfil': totalPerfiles,
                                        'Perfil': val,
                                        'Sobrante': sobr
                                    }
                                    )
                                
                                hojaCorte.append({
                                    'NumPerfil': totalPerfiles + 1,
                                    'Perfil': val,
                                    'Medida Corte': fila['Medida Corte'],
                                    'Medida Aumentada': fila['MedidaAumentada']
                                }
                                )

            if MedidaTotal > 0:
                sobr = unidad - MedidaTotal
                if sobr > 0:
                    sobrante.append({
                        'NumPerfil': totalPerfiles+1,
                        'Perfil': val,
                        'Sobrante': sobr
                    }
                    )
                #print(f'{val} = {sobr}')

        
            totalPerfiles += 1     

            resultados.append({
                'referencia_unica': val,  # Suponiendo que tienes una columna de referencia única en tu DataFrame original
                'Perfiles Aprox': totalPerfiles
            })                



            #print(f'Perfil: {val} = Total aprox de perfiles:  {totalPerfiles}')

        nuevo_df = pd.DataFrame(resultados)

        sobr_df = pd.DataFrame(sobrante)
        sobr_df = sobr_df.sort_values(by=['Perfil','NumPerfil'], ascending=[True, True])

        hc_df = pd.DataFrame(hojaCorte)
        hc_df = hc_df.sort_values(by=['Perfil','NumPerfil'], ascending=[True, True])
      
        #Nombre de las hojas del libro
        sheet_names = ['Perfiles_Aprox', 'Medidas_Sobrantes', 'HojaCorte']

        directorio = os.path.dirname(ruta_archivo)
        nomArchivoDestino = 'Resumen_Optimizado.xlsx'

        rutaDestino = os.path.join(directorio, nomArchivoDestino)

        #Nombre del nuevo libro
        excel_writer = pd.ExcelWriter(rutaDestino, engine='xlsxwriter')


        nuevo_df.to_excel(excel_writer, sheet_name=sheet_names[0], index=False)
        sobr_df.to_excel(excel_writer, sheet_name=sheet_names[1], index=False)
        hc_df.to_excel(excel_writer, sheet_name=sheet_names[2], index=False)

        excel_writer.close()

         # Mostrar un mensaje de confirmación
        messagebox.showinfo("Proceso completado", "El proceso de exportación ha finalizado.")




# Llamar a la función para abrir el explorador de archivos
optimizar()