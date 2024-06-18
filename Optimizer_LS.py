import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
#import pyexcel
import pandas as pd
import numpy as np
import os
import datetime


# Fecha de expiración del programa
fecha_expiracion = datetime.date(2024, 8, 21)

# Verificar la fecha de expiración
fecha_actual = datetime.date.today()
if fecha_actual > fecha_expiracion:
    #print("Paz y bien. El programa ha expirado.")
    messagebox.showinfo("Error", "Paz y bien. El programa ha expirado. Bendiciones")
    # Aquí podrías mostrar un mensaje de error y salir del programa
    exit()

tiempo = (fecha_actual - fecha_expiracion).days
messageFin = f'Paz y bien. El programa expirará en {tiempo} días. Bendiciones'
messagebox.showinfo("Cuenta regresiva", messageFin)

def mostrar_indicaciones():
    indicaciones = """
    Indicaciones para el desarrollo del programa:

    1. Ingrese el valor para el "Perfil tradicional" y "Perfil Europeo" en los campos correspondientes. 
        *Tener en cuenta si el campo se deja vacío por defecto se tomará 6000 para el perfil tradicional y 6500 para perfil europeo.
        *En las referencias se tomarán como europeas aquellas que son de caracter numérico o que empiezan por '3' y tradicional las que son alfanuméricas o de tipo texto.


    2. Las medidas se deben ingresar en milímetros (mm) al igual que el descuento por corte.


    3. Seleccione un archivo haciendo clic en "Seleccionar Archivo".
        Recuerda el archivo a cargar debe contar con las siguientes especificaciones:
            *Archivo en Excel que contenga sólo una hoja.
            *La hoja debe contener cincos columnas sin formato en el siguiente orden perfil, cantidad, medida corte, Nombre Ref. e Item. Si se desea se puede descargar el modelo, se puede trabajar sobre ese archivo.
            *Los campos numéricos deben ser enteros, por favor tener cuidado al momento de subir los datos.


    4. Una vez cargados los datos, haga clic en "Optimizar" para procesar y generar el archivo optimizado.


    5. El archivo optimizado se guardará en el mismo directorio que el archivo seleccionado, con el nombre "Resumen_Optimizado.xlsx".


    6. Los botones de color azul serán de apoyo mientras que los verdes son de caracter obligatorio para la ejecución del programa.


    ¡Recuerda hacer de lo ordinario algo extraordinario!
    """
    messagebox.showinfo("Indicaciones", indicaciones)

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

    ventana.destroy()  # Cerrar la ventana al obtener datos

def seleccionar_archivo():
    global ruta_archivo
    ruta_archivo = filedialog.askopenfilename()
    ruta_archivo_entry.delete(0, tk.END)
    ruta_archivo_entry.insert(0, ruta_archivo)

def descargar_modelo():
    # Define the column names
    column_names = ['Perfil', 'Cant.', 'Medida Corte', 'Nombre Ref.', 'Item']

    # Create an empty DataFrame with the specified columns
    template_df = pd.DataFrame(columns=column_names)

    # Path to save the template file
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))

    # If user selected a save location
    if save_path:
        # Save the DataFrame to the selected location
        try:
            template_df.to_excel(save_path, index=False)
            messagebox.showinfo("Descarga completa", "El modelo de archivo se ha descargado exitosamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo completar la descarga: {str(e)}")


def optimizar():

    obtener_datos()

    if not ruta_archivo:
        messagebox.showerror("Error", "Por favor seleccione un archivo.")
        return
    
    try:

        # Leer los datos del archivo Excel seleccionado
        df_init = pd.read_excel(ruta_archivo)

        #Validar si el número de columnas corresponde a las que necesita el programa para su ejecución.
        if len(df_init.columns) != 5:
            messagebox.showerror("Error", """
            El archivo debe tener exactamente 5 columnas.
            ---------------------------------------------------------------
            |Perfil | Cant. | Medida Corte | Nombre Ref.| Item |
            ---------------------------------------------------------------
            Por favor revisar el libro de Excel, organizarlo y volver a subirlo. Muchas gracias, bendiciones.
            """)
            return
        

        #Valida que los campos de la medida de corte y cantidad contengan valores enteros.
        if not df_init.iloc[:, 1].apply(lambda x: isinstance(x, int)).all() or \
           not df_init.iloc[:, 2].apply(lambda x: isinstance(x, int)).all():
            messagebox.showerror("Error", """Las columnas de Medida Corte y Cantidad deben contener valores enteros. Por favor revisar el archivo y volver a subirlo.
                                 
                                 Muchas gracias. Bendiciones.""")
            return
        
        #Renombrar el nombre de las columnas para evitar errores en la ejecución del código, ya que es obligatorio los siguientes nombres.
        df_init.rename(columns={df_init.columns[0]: 'Perfil'}, inplace=True )        
        df_init.rename(columns={df_init.columns[1]: 'Cant.'}, inplace=True )
        df_init.rename(columns={df_init.columns[2]: 'Medida Corte'}, inplace=True )
        df_init.rename(columns={df_init.columns[3]: 'Nombre Ref.'}, inplace=True)
        df_init.rename(columns={df_init.columns[4]: 'Item'}, inplace=True)

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

                    if(fila['Empieza'] == None or fila['Empieza'] == '3'):
                        unidad = dato2
                    else:
                        unidad = dato1

                    #Con este ciclo se busca tener en cuenta el número de cortes por medidas y perfil.
                    for i in range(fila['Cant.']):

                        sobrante = sorted(sobrante, key=lambda x: (str(x['Perfil']), x['Sobrante']))   
                        element_remove = None

                        for elem in sobrante:
                            if elem['Perfil'] == val:
                                if (fila['MedidaAumentada']) <= elem['Sobrante']:
                                    element_remove = elem

                                    compr = elem['Sobrante'] - fila['MedidaAumentada']

                                    if compr > 0:
                                        sobrante.append({
                                            'Longitud de perfil': elem['Longitud de perfil'],
                                            'NumPerfil': elem['NumPerfil'],
                                            'Perfil': val,
                                            'Sobrante': compr
                                        }
                                        )

                                    if(elem['NumPerfil'] < 10):
                                        numPerfil = "0" + str(elem['NumPerfil'])
                                    else:
                                        numPerfil = elem['NumPerfil']

                                    hojaCorte.append({
                                        'Longitud de perfil': unidad,
                                        'Medida Aumentada': fila['MedidaAumentada'],
                                        'Perfil': val,
                                        'NumPerfil': "Perfil #: " + str(numPerfil),
                                        'Nombre Ref.': fila['Nombre Ref.'],
                                        'item': "Item #: " + str(fila['Item']),
                                        'Medida Corte': fila['Medida Corte'],
                                        'Cant.': 1                                        
                                        
                                    }
                                    )
                                    break
                                    
                                    
                        
                        if element_remove:
                            sobrante.remove(element_remove)

                        else:                                
                            
                            if((MedidaTotal + fila['MedidaAumentada']) <= unidad):

                                # Si la condición se cumple, con relación al nombre del perfil
                                MedidaTotal += fila['MedidaAumentada']

                                if(totalPerfiles+1 < 10):
                                    numPerfil = "0" + str(totalPerfiles+1)
                                else:
                                    numPerfil = totalPerfiles + 1

                                hojaCorte.append({
                                    'Longitud de perfil': unidad,
                                    'Medida Aumentada': fila['MedidaAumentada'],                                    
                                    'Perfil': val,
                                    'NumPerfil': "Perfil #: " + str(numPerfil),
                                    'Nombre Ref.': fila['Nombre Ref.'],
                                    'item': "Item #: " + str(fila['Item']),
                                    'Medida Corte': fila['Medida Corte'],
                                    'Cant.': 1                                
                                    
                                }
                                )

                            else:
                                totalPerfiles += 1
                                
                                sobr = unidad - MedidaTotal
                                #print(f'{val} = {sobr}')
                                MedidaTotal = fila['MedidaAumentada']

                                if(totalPerfiles+1 < 10):
                                    numPerfil = "0" + str(totalPerfiles+1)
                                else:
                                    numPerfil = totalPerfiles + 1

                                if sobr > 0:
                                    sobrante.append({
                                        'Longitud de perfil': unidad,
                                        'NumPerfil': totalPerfiles,
                                        'Perfil': val,
                                        'Sobrante': sobr
                                    }
                                    )
                                
                                hojaCorte.append({
                                    'Longitud de perfil': unidad,
                                    'Medida Aumentada': fila['MedidaAumentada'],                                    
                                    'Perfil': val,
                                    'NumPerfil': "Perfil #: " + str(numPerfil),
                                    'Nombre Ref.': fila['Nombre Ref.'],
                                    'item': "Item #: " + str(fila['Item']),
                                    'Medida Corte': fila['Medida Corte'],
                                    'Cant.': 1                                  
                                    
                                }
                                )

            if MedidaTotal > 0:
                sobr = unidad - MedidaTotal
                if sobr > 0:
                    sobrante.append({
                        'Longitud de perfil': unidad,
                        'NumPerfil': totalPerfiles+1,
                        'Perfil': val,
                        'Sobrante': sobr
                    }
                    )

        
            totalPerfiles += 1     

            resultados.append({
                'referencia_unica': val,  # Suponiendo que tienes una columna de referencia única en tu DataFrame original
                'Perfiles Aprox': totalPerfiles
            })                

        nuevo_df = pd.DataFrame(resultados)

        sobr_df = pd.DataFrame(sobrante)
        sobr_df = sobr_df.sort_values(by=['Perfil','NumPerfil'], ascending=[True, True])

        hc_df = pd.DataFrame(hojaCorte)
        hc_df = hc_df.sort_values(by=['Perfil','NumPerfil', 'Medida Corte'], ascending=[True, True, False])
      
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
        messagebox.showinfo("Proceso completado", "El proceso de exportación ha finalizado. Bendiciones")

    except FileNotFoundError:
        messagebox.showerror("Error", "El archivo seleccionado no existe.")
    except pd.errors.EmptyDataError:
        messagebox.showerror("Error", "El archivo seleccionado está vacío.")
    except pd.errors.ParserError:
        messagebox.showerror("Error", "Error al procesar el archivo Excel.")
    except Exception as e:
        messagebox.showerror("Error", f"Se ha producido un error: {str(e)}")




# Crear la ventana
ventana = tk.Tk()
ventana.title("Optimizador de Aluminio - Optimizer_LS")
ventana.minsize(width=300, height=220)
ventana.maxsize(width=350, height=240)

# Crear etiquetas y entradas para los datos
etiqueta1 = tk.Label(ventana, text="Perfil tradicional (mm):")
etiqueta1.grid(row=0, column=0, padx=5, pady=5)

validacion_dato1 = ventana.register(validar_numero)
entrada1 = tk.Entry(ventana, validate="key", validatecommand=(validacion_dato1, '%P'))
entrada1.grid(row=0, column=1, padx=5, pady=5)

etiqueta2 = tk.Label(ventana, text="Perfil Europeo (mm):")
etiqueta2.grid(row=1, column=0, padx=5, pady=5)

validacion_dato2 = ventana.register(validar_numero)
entrada2 = tk.Entry(ventana, validate="key", validatecommand=(validacion_dato2, '%P'))
entrada2.grid(row=1, column=1, padx=5, pady=5)

etiqueta_descuento = tk.Label(ventana, text="Descuento por corte (mm):")
etiqueta_descuento.grid(row=2, column=0, padx=5, pady=5)

validacion_descuento = ventana.register(validar_numero)
entrada_descuento = tk.Entry(ventana, validate="key", validatecommand=(validacion_descuento, '%P'))
entrada_descuento.grid(row=2, column=1, padx=5, pady=5)

etiqueta_archivo = tk.Label(ventana, text="Ruta del archivo:")
etiqueta_archivo.grid(row=3, column=0, padx=5, pady=5)

ruta_archivo_entry = tk.Entry(ventana)
ruta_archivo_entry.grid(row=3, column=1, padx=5, pady=5)

# Botón para seleccionar el archivo
boton_archivo = tk.Button(ventana, text="Seleccionar Archivo", command=seleccionar_archivo, bg="#8BDD9E")
boton_archivo.grid(row=4, column=0, padx=5, pady=5, sticky="ew")  # Centrar y expandir horizontalmente
boton_archivo.config(width=20)  # Ajustar el ancho del botón

# Botón para obtener los datos
boton_obtener = tk.Button(ventana, text="Optimizar", command=optimizar, bg="#8BDD9E")
boton_obtener.grid(row=4, column=1, padx=5, pady=5, sticky="ew")  # Centrar y expandir horizontalmente
boton_obtener.config(width=20)  # Ajustar el ancho del botón


boton_descargar = tk.Button(ventana, text="Descargar Modelo", command=descargar_modelo, bg="#A5EAF3")
boton_descargar.grid(row=5, column=0, padx=5, pady=5, sticky="ew")
boton_descargar.config(width=20)

boton_indicaciones = tk.Button(ventana, text="Indicaciones", command=mostrar_indicaciones, bg="#A5EAF3")
boton_indicaciones.grid(row=5, column=1, padx=5, pady=5, sticky="ew")
boton_indicaciones.config(width=20)



# Label como pie de página
pie_de_pagina = tk.Label(ventana, text="""Realizado por Laura Juliana Serrano García - MaKeajse 2024.
Recuerda: 'Hacer de lo Ordinario, algo Extraordinario'.
""", bg="lightgray", fg="black")
pie_de_pagina.grid(row=7, columnspan=2, padx=5, pady=5, sticky="ew")

# Variable para almacenar la ruta del archivo seleccionado
ruta_archivo = ""


# Ejecutar el bucle principal
ventana.mainloop()
