import tkinter as tk
from tkinter import ttk
import pyodbc
from datetime import datetime, timedelta
import random
import tkinter.messagebox as messagebox
from decimal import Decimal
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import os
import subprocess


# Conexión a la base de datos
def obtener_conexion():
    
    conn = pyodbc.connect(
        'DRIVER={SQL Server};'
        'SERVER=192.168.200.35;'
        'DATABASE=CONTHIDRA;'
        'UID=logic;'
        'PWD=#Obelix*.99'
    )
    return conn

# Lista de colores suaves para las filas
COLORES = ['#FFDDC1', '#C1E1C5', '#C1D4FF', '#FFE1C1', '#FFD1DC', 
           '#D1C1FF', '#E1C1FF', '#C1FFD1', '#FFC1E1', '#FFF1C1', 
           '#C1FFF4', '#F4C1FF', '#E1FFD1', '#C1FFF1', '#FFF4C1']


# Crear tabla si no existe
def crear_tabla_si_no_existe(cursor):
    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FabricacionTemporal]') AND type in (N'U'))
        BEGIN
            CREATE TABLE FabricacionTemporal (
                GrupoID INT,
                Operario NVARCHAR(50),
                Orden NVARCHAR(50),
                Inicio DATETIME,
                Color NVARCHAR(20),
                Descripcion NVARCHAR(250)
            )
        END
    """)

# Crear una función para alternar colores en el Treeview
def alternar_colores(treeview, color1, color2):
    for index, item in enumerate(treeview.get_children()):
        if index % 2 == 0:
            treeview.item(item, tags=('color1',))
        else:
            treeview.item(item, tags=('color2',))

    treeview.tag_configure('color1', background=color1)  
    treeview.tag_configure('color2', background=color2)  

# Definir colores para las filas
COLOR1 = '#F0F0F0'  
COLOR2 = '#FFFFFF'  


# Consulta de operarios
def cargar_operarios():
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    # Consulta para obtener operarios que no están en FabricacionTemporal
    cursor.execute("""
        SELECT Operario, NombreOperario 
        FROM operarios 
        WHERE codigoempresa = 2 
        AND FechaBaja IS NULL
        AND Operario NOT IN (SELECT DISTINCT Operario FROM FabricacionTemporal)
    """)
    
    tree_operarios.delete(*tree_operarios.get_children())  
    for row in cursor.fetchall():
        operario_id = str(row[0]).strip('(),')  
        tree_operarios.insert("", tk.END, values=(operario_id, row[1]))  
    
    alternar_colores(tree_operarios, COLOR1, COLOR2)
    
    conn.close()

# Consulta de ordenes
def cargar_ordenes():
    # Leer las series a filtrar desde un archivo txt
    series_a_filtrar = []
    with open('series_a_filtrar.txt', 'r') as file:
        line = file.readline().strip()
        if line.startswith("Series a Filtrar ="):
            # Extraer las series de la línea de texto
            series_a_filtrar = line.split('=')[1].strip()
            # Quitar corchetes y separar por comas
            series_a_filtrar = [serie.strip() for serie in series_a_filtrar.strip('[]').split(',')]

    # Convertir la lista de series en un string para la consulta SQL
    series_str = ",".join([f"'{serie}'" for serie in series_a_filtrar])

    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT Ejerciciofabricacion, SerieFabricacion, NumeroFabricacion, CodigoArticulo, 
        CASE 
            WHEN (TipoFabricacion = 'A') THEN CAST(UnidadesAFabricar AS INT) 
            ELSE (SELECT CAST(SUM(unidadesafabricar) AS INT) 
                  FROM ordenestrabajo 
                  WHERE codigoempresa = OrdenesFabricacion.codigoempresa 
                    AND EjercicioFabricacion = OrdenesFabricacion.EjercicioFabricacion
                    AND SerieFabricacion = OrdenesFabricacion.SerieFabricacion 
                    AND NumeroFabricacion = OrdenesFabricacion.NumeroFabricacion
                    AND ordenestrabajo.NivelCompuesto = 90)
        END AS Unidades,
        DescripcionArticulo
        FROM OrdenesFabricacion 
        WHERE codigoempresa = 2 
        AND EstadoOF = 1 
        AND VMostrarOrden = -1 
        AND SerieFabricacion IN ({series_str})
        AND CAST(EjercicioFabricacion AS VARCHAR) + '/' + SerieFabricacion + '/' + CAST(NumeroFabricacion AS VARCHAR) 
            NOT IN (SELECT Orden FROM FabricacionTemporal)
    """)

    # Limpiar y llenar la tabla
    tree_ordenes.delete(*tree_ordenes.get_children())  
    for row in cursor.fetchall():
        # Reemplazar valores vacíos o None por 'N/A'
        row = ['N/A' if (r is None or r == '') else r for r in row]

        # Limpiar comillas dobles dentro de las descripciones
        row = [r.replace('"', "'") if isinstance(r, str) else r for r in row]

        # Insertar en el Treeview
        tree_ordenes.insert('', tk.END, values=row)
    
    alternar_colores(tree_ordenes, COLOR1, COLOR2)
    
    conn.close()

# Obtener color dentro de los disponibles que no se repita
def obtener_color_disponible():
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    # Obtener los colores ya existentes en la tabla FabricacionTemporal
    cursor.execute("SELECT DISTINCT Color FROM FabricacionTemporal")
    colores_existentes = {row[0] for row in cursor.fetchall()}  
    
    # Buscar un color que no esté en la lista de colores existentes
    color_disponible = None
    for color in COLORES:
        if color not in colores_existentes:
            color_disponible = color
            break
    
    # Si todos los colores ya están en uso, se elige uno al azar (como fallback)
    if color_disponible is None:
        color_disponible = random.choice(COLORES)
    
    conn.close()
    return color_disponible

# Iniciar fabricación
def iniciar_fabricacion():

    selected_operarios = [tree_operarios.item(item, 'values') for item in tree_operarios.selection()]
    selected_ordenes = [tree_ordenes.item(item, 'values') for item in tree_ordenes.selection()]
    
    if not selected_operarios or not selected_ordenes:
        mostrar_aviso("Selecciona al menos un operario y una orden")
        return
    
    conn = obtener_conexion()
    cursor = conn.cursor()

    # Crear la tabla si no existe
    crear_tabla_si_no_existe(cursor)
    
    # Crear un identificador de grupo para las fabricaciones iniciadas juntas
    grupo_id = random.randint(1, 10000)
    
    # Asignar un color suave para este grupo
    color_fondo = obtener_color_disponible()
    
    # Insertar fabricaciones y eliminarlas de los grids de selección
    for operario in selected_operarios:
        for orden in selected_ordenes:
            inicio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Limpiar el valor del operario para que sea solo el número
            operario_id = operario[0].strip('(),')

            # Verificar si la serie es "N/A" y cambiarla a una cadena vacía
            serie = orden[1] if orden[1] != "N/A" else ""  

            cursor.execute("""
                INSERT INTO FabricacionTemporal (GrupoID, Operario, Orden, Inicio, Color , Descripcion)
                VALUES (?, ?, ?, ?, ?, ?)
            """, grupo_id, operario_id, orden[0]+"/"+serie+"/"+orden[2], inicio, color_fondo, orden[5])

            # Insertar en el grid de En curso
            tree_en_curso.insert("", tk.END, values=(operario[1], orden[2], inicio))
    
    # Eliminar los operarios y órdenes seleccionados de los grids
    for item in tree_operarios.selection():
        tree_operarios.delete(item)
    for item in tree_ordenes.selection():
        tree_ordenes.delete(item)
    
    conn.commit()
    conn.close()

    cargar_fabricaciones_en_curso()


# Cargar fabricaciones en curso (incluye colores)
def cargar_fabricaciones_en_curso():
    # Guardar la selección actual
    selected_items = tree_en_curso.selection()
    selected_indices = [tree_en_curso.index(item) for item in selected_items]
    try:
        conn = obtener_conexion()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT o.NombreOperario, f.Orden, f.Inicio, f.Color , f.Descripcion
            FROM FabricacionTemporal f
            JOIN operarios o ON f.Operario = o.Operario
            WHERE o.codigoempresa = 2 
            ORDER BY f.Inicio
        """)
        tree_en_curso.delete(*tree_en_curso.get_children())  # Limpiar antes de cargar
        for row in cursor.fetchall():
            operario, orden, inicio, color_fondo, descripcion = row

            # Verificar si inicio es un objeto datetime o una cadena
            if isinstance(inicio, datetime):
                # Si es un objeto datetime, formatearlo directamente
                inicio_formateado = inicio.strftime("%d/%m/%Y %H:%M:%S")
            else:
                # Si es una cadena, convertirlo usando strptime
                inicio_dt = datetime.strptime(inicio, "%Y-%m-%d %H:%M:%S")
                inicio_formateado = inicio_dt.strftime("%d/%m/%Y %H:%M:%S")
            
            tiempo_transcurrido = ""
            item_id = tree_en_curso.insert("", tk.END, values=(operario, orden, inicio,tiempo_transcurrido, descripcion))
            tree_en_curso.item(item_id, tags=(color_fondo,))  
            tree_en_curso.tag_configure(color_fondo, background=color_fondo)  
    except Exception as e:
        mostrar_aviso(f"Error al cargar fabricaciones en curso: {e}")
    finally:
        conn.close()
     # Volver a seleccionar los elementos que estaban seleccionados
    for index in selected_indices:
        if index < len(tree_en_curso.get_children()):  # Verifica que el índice sea válido
            tree_en_curso.selection_set(tree_en_curso.get_children()[index])



# Actualizar los contadores de tiempo en el grid
def actualizar_tiempos():
    for item in tree_en_curso.get_children():
        inicio_str = tree_en_curso.item(item, 'values')[2]  
        
        try:
            # Intentar parsear la cadena como una fecha completa
            inicio = datetime.strptime(inicio_str, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            # Si no es un formato de fecha válido, saltar este registro
            continue
        
        # Calcular el tiempo transcurrido
        ahora = datetime.now()
        tiempo_transcurrido = ahora - inicio
        
        # Formatear el tiempo transcurrido (HH:MM:SS)
        horas, resto = divmod(tiempo_transcurrido.total_seconds(), 3600)
        minutos, segundos = divmod(resto, 60)
        tiempo_formateado = f"{int(horas):02}:{int(minutos):02}:{int(segundos):02}"
        
        # Actualizar una nueva columna con el tiempo transcurrido
        tree_en_curso.set(item, column="Tiempo Transcurrido", value=tiempo_formateado)
    
    # Repetir cada segundo
    root.after(1000, actualizar_tiempos)


# Variable global para controlar el estado del diálogo
dialogo_abierto = False

def mostrar_mensaje():

    global dialogo_abierto
    dialogo_abierto = True  # Marca que el diálogo está abierto


    # Crear ventana de diálogo personalizada
    dialogo = tk.Toplevel(root)
    dialogo.title("Confirmación")
    dialogo.resizable(False, False)

    # Calcular la posición de la ventana de diálogo
    ventana_ancho = 300
    ventana_alto = 180

    # Centrar la ventana de diálogo en la ventana principal
    x = root.winfo_x() + (root.winfo_width() // 2) - (ventana_ancho // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (ventana_alto // 2)
    dialogo.geometry(f"{ventana_ancho}x{ventana_alto}+{x}+{y}")  

    # Mensaje en la ventana
    label = tk.Label(dialogo, text="¿Finalizar orden de fabricacion?", font=("Arial", 12))
    label.pack(pady=20)

    # Función para cerrar la ventana sin hacer nada
    def cerrar_sin_accion():
        global dialogo_abierto
        dialogo_abierto = False
        dialogo.destroy()

    # Configurar el botón de cerrar
    dialogo.protocol("WM_DELETE_WINDOW", cerrar_sin_accion)    

    # Botones "Sí", "No" y "Cancelar"
    boton_si = tk.Button(dialogo, text="Sí", command=lambda: [terminar_fabricacion(), dialogo.destroy()],
                         bg="#4CAF50", fg="white", width=5, font=("Arial", 10))
    boton_si.pack(side="left", padx=15, pady=10)

    boton_no = tk.Button(dialogo, text="No", command=lambda: [terminar_fabricacion2(), dialogo.destroy()],
                         bg="#f44336", fg="white", width=5, font=("Arial", 10))
    boton_no.pack(side="left", padx=15, pady=10)

    boton_cancelar = tk.Button(dialogo, text="Cancelar", command=cerrar_sin_accion,
                                bg="#2196F3", fg="white", width=10, font=("Arial", 10))
    boton_cancelar.pack(side="right", padx=10, pady=10)

    # Mantener la ventana principal oculta y el diálogo abierto
    root.wait_window(dialogo)

# Terminar fabricación
def terminar_fabricacion():
    selected_curso = [tree_en_curso.item(item, 'values') for item in tree_en_curso.selection()]
    if not selected_curso:
        mostrar_aviso("Selecciona una fabricación en curso para finalizar")
        return
    

    # Obtener el tiempo de inicio de la primera fabricación seleccionada
    inicio_fabricacion = tree_en_curso.item(tree_en_curso.selection()[0])['values'][2]
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute("""
        Select * from FabricacionTemporal
        WHERE Inicio = ?
    """, (inicio_fabricacion))
    
    for row in cursor.fetchall():
        # Actualizar la fabricación en la base de datos
        actualizar_fabricacion(row[2] ,row[3],row[1])
    conn.close()

    

    # Convertir la cadena de fecha y hora a un objeto datetime para la comparación
    inicio_fabricacion_dt = datetime.strptime(inicio_fabricacion, "%Y-%m-%d %H:%M:%S")

    conn = obtener_conexion()
    cursor = conn.cursor()
    # Eliminar todas las fabricaciones que tengan el mismo tiempo de inicio
    cursor.execute("""
        DELETE FROM FabricacionTemporal
        WHERE Inicio = ?
    """, (inicio_fabricacion_dt,))

    # Eliminar todas las fabricaciones del grid
    tree_en_curso.delete(*tree_en_curso.get_children())  

    conn.commit()
    conn.close()
    
    # Recargar los operarios y las órdenes después de terminar la fabricación
    cargar_operarios()
    cargar_ordenes()
    cargar_fabricaciones_en_curso()

def terminar_fabricacion2():
    selected_curso = [tree_en_curso.item(item, 'values') for item in tree_en_curso.selection()]
    if not selected_curso:
        mostrar_aviso("Selecciona una fabricación en curso para finalizar")
        return
    

    # Obtener el tiempo de inicio de la primera fabricación seleccionada
    inicio_fabricacion = tree_en_curso.item(tree_en_curso.selection()[0])['values'][2]
    conn = obtener_conexion()
    cursor = conn.cursor()
    cursor.execute("""
        Select * from FabricacionTemporal
        WHERE Inicio = ?
    """, (inicio_fabricacion))
    
    for row in cursor.fetchall():
        # Actualizar la fabricación en la base de datos
        actualizar_fabricacion2(row[2] ,row[3],row[1])
    conn.close()

    

    # Convertir la cadena de fecha y hora a un objeto datetime para la comparación
    inicio_fabricacion_dt = datetime.strptime(inicio_fabricacion, "%Y-%m-%d %H:%M:%S")

    conn = obtener_conexion()
    cursor = conn.cursor()
    # Eliminar todas las fabricaciones que tengan el mismo tiempo de inicio
    cursor.execute("""
        DELETE FROM FabricacionTemporal
        WHERE Inicio = ?
    """, (inicio_fabricacion_dt,))

    # Eliminar todas las fabricaciones del grid
    tree_en_curso.delete(*tree_en_curso.get_children())  

    conn.commit()
    conn.close()
    
    # Recargar los operarios y las órdenes después de terminar la fabricación
    cargar_operarios()
    cargar_ordenes()
    cargar_fabricaciones_en_curso()

def actualizar_fabricacion(orden, inicio_fabricacion, Operario):
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    partes = orden.split("/")
    año, serie, numero = partes[0], partes[1], partes[2]

    # Obtener el TipoFabricacion de OrdenesFabricacion
    cursor.execute("""
        SELECT TipoFabricacion
        FROM OrdenesFabricacion
        WHERE EjercicioFabricacion = ? AND SerieFabricacion = ? AND NumeroFabricacion = ?
    """, (año, serie, numero))
    resultado = cursor.fetchone()

    print("0*******************************") 
    print(año)
    print(serie)
    print(numero)
    print("0*******************************") 

    if resultado:
        tipo_fabricacion = resultado[0]
        fecha_final = datetime.now()

        if tipo_fabricacion == 'A':
            print("Entramos a tipo_fabricacion A")
            cursor.execute("""
                SELECT TOP 1 EjercicioTrabajo, NumeroTrabajo 
                FROM OrdenesTrabajo
                WHERE EjercicioTrabajo = ? AND SerieFabricacion = ? AND NumeroFabricacion = ? 
                      AND codigoempresa = 2
                ORDER BY NumeroTrabajo ASC
            """, (año, serie, numero))

            print("*******************************") 
            print(año)
            print(serie)
            print(numero)
            print("*******************************") 

            orden_trabajo = cursor.fetchone()
            if orden_trabajo:
                Ejerciciotrabajo, Numerotrabajo = orden_trabajo
                insertar_incidencias(Ejerciciotrabajo, Numerotrabajo, inicio_fabricacion, fecha_final, Operario)
        
        elif tipo_fabricacion == 'P':
            print("Entramos a tipo_fabricacion P")
            cursor.execute("""
                SELECT EjercicioTrabajo, NumeroTrabajo, 
                       (SELECT TOP 1 TiempoUnFabricacion * 1440 
                        FROM operacionesot
                        WHERE ejerciciotrabajo = OrdenesTrabajo.ejerciciotrabajo
                          AND numerotrabajo = OrdenesTrabajo.numerotrabajo) AS TiempoUnFabricacionMinutos,
                       UnidadesAFabricar 
                FROM OrdenesTrabajo
                WHERE EjercicioTrabajo = ? AND SerieFabricacion = ? AND NumeroFabricacion = ?
                      AND NivelCompuesto = 90 AND codigoempresa = 2
            """, (año, serie, numero))
            
            ordenes_trabajo = cursor.fetchall()
            tiempo_ficticio_total = 0

            for orden in ordenes_trabajo:
                _, _, TiempoUnFabricacionMinutos, UnidadesAFabricar = orden
                print("TiempoUnFabricacionMinutos:", TiempoUnFabricacionMinutos)
                tiempo_ficticio_total += TiempoUnFabricacionMinutos * UnidadesAFabricar

            tiempo_real_total = (fecha_final - inicio_fabricacion).total_seconds() / 60  
            factor_proporcional = Decimal(tiempo_real_total) / Decimal(tiempo_ficticio_total) if tiempo_ficticio_total > 0 else Decimal(1)

            tiempo_inicio_orden = inicio_fabricacion
            
            for orden in ordenes_trabajo:
                Ejerciciotrabajo, Numerotrabajo, TiempoUnFabricacionMinutos, UnidadesAFabricar = orden
                tiempo_fabricacion_real = TiempoUnFabricacionMinutos * UnidadesAFabricar * factor_proporcional

                # Convertir tiempo_fabricacion_real a float antes de usarlo en timedelta
                tiempo_final_orden = tiempo_inicio_orden + timedelta(minutes=float(tiempo_fabricacion_real))
                
                # Convertir horas a formato Sage antes de insertar
                insertar_incidencias(Ejerciciotrabajo, Numerotrabajo, tiempo_inicio_orden, tiempo_final_orden, Operario)
                
                # Actualizar el inicio de la siguiente orden al final de la actual
                tiempo_inicio_orden = tiempo_final_orden  

    conn.close()
    conn = obtener_conexion()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE OrdenesFabricacion
        SET VMostrarOrden = '0'
        WHERE EjercicioFabricacion = ? AND SerieFabricacion = ? AND NumeroFabricacion = ?
    """, (año, serie, numero))

    conn.commit()
    conn.close()
    print(f"Fabricación {orden} actualizada correctamente.")
    
    

def actualizar_fabricacion2(orden, inicio_fabricacion, Operario):
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    partes = orden.split("/")
    año, serie, numero = partes[0], partes[1], partes[2]

    # Obtener el TipoFabricacion de OrdenesFabricacion
    cursor.execute("""
        SELECT TipoFabricacion
        FROM OrdenesFabricacion
        WHERE EjercicioFabricacion = ? AND SerieFabricacion = ? AND NumeroFabricacion = ?
    """, (año, serie, numero))
    resultado = cursor.fetchone()

    print("0*******************************") 
    print(año)
    print(serie)
    print(numero)
    print("0*******************************") 

    if resultado:
        tipo_fabricacion = resultado[0]
        fecha_final = datetime.now()

        if tipo_fabricacion == 'A':
            print("Entramos a tipo_fabricacion A")
            cursor.execute("""
                SELECT TOP 1 EjercicioTrabajo, NumeroTrabajo 
                FROM OrdenesTrabajo
                WHERE EjercicioTrabajo = ? AND SerieFabricacion = ? AND NumeroFabricacion = ? 
                      AND codigoempresa = 2
                ORDER BY NumeroTrabajo ASC
            """, (año, serie, numero))

            print("*******************************") 
            print(año)
            print(serie)
            print(numero)
            print("*******************************") 

            orden_trabajo = cursor.fetchone()
            if orden_trabajo:
                Ejerciciotrabajo, Numerotrabajo = orden_trabajo
                insertar_incidencias(Ejerciciotrabajo, Numerotrabajo, inicio_fabricacion, fecha_final, Operario)
        
        elif tipo_fabricacion == 'P':
            cursor.execute("""
                SELECT EjercicioTrabajo, NumeroTrabajo, 
                       (SELECT TOP 1 TiempoUnFabricacion * 1440 
                        FROM operacionesot
                        WHERE ejerciciotrabajo = OrdenesTrabajo.ejerciciotrabajo
                          AND numerotrabajo = OrdenesTrabajo.numerotrabajo) AS TiempoUnFabricacionMinutos,
                       UnidadesAFabricar 
                FROM OrdenesTrabajo
                WHERE EjercicioTrabajo = ? AND SerieFabricacion = ? AND NumeroFabricacion = ?
                      AND NivelCompuesto = 90 AND codigoempresa = 2
            """, (año, serie, numero))
            
            ordenes_trabajo = cursor.fetchall()
            tiempo_ficticio_total = 0

            for orden in ordenes_trabajo:
                _, _, TiempoUnFabricacionMinutos, UnidadesAFabricar = orden
                tiempo_ficticio_total += TiempoUnFabricacionMinutos * UnidadesAFabricar

            tiempo_real_total = (fecha_final - inicio_fabricacion).total_seconds() / 60  
            factor_proporcional = Decimal(tiempo_real_total) / Decimal(tiempo_ficticio_total) if tiempo_ficticio_total > 0 else Decimal(1)

            tiempo_inicio_orden = inicio_fabricacion
            
            for orden in ordenes_trabajo:
                Ejerciciotrabajo, Numerotrabajo, TiempoUnFabricacionMinutos, UnidadesAFabricar = orden
                tiempo_fabricacion_real = TiempoUnFabricacionMinutos * UnidadesAFabricar * factor_proporcional

                # Convertir tiempo_fabricacion_real a float antes de usarlo en timedelta
                tiempo_final_orden = tiempo_inicio_orden + timedelta(minutes=float(tiempo_fabricacion_real))
                
                # Convertir horas a formato Sage antes de insertar
                insertar_incidencias(Ejerciciotrabajo, Numerotrabajo, tiempo_inicio_orden, tiempo_final_orden, Operario)
                
                # Actualizar el inicio de la siguiente orden al final de la actual
                tiempo_inicio_orden = tiempo_final_orden  

    conn.close()
    print(f"Fabricación {orden} actualizada correctamente.")


def insertar_incidencias(Ejerciciotrabajo, Numerotrabajo, FechaInicio, FechaFinal, Operario):
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    incidencia = 3
    codigoempresa = 2

    print("Fecha de inicio:", str(FechaInicio))
    print("Fecha de finalización:", str(FechaFinal))

    # Calcular las horas, minutos y segundos de inicio y final en formato decimal (1 = 24 horas)
    HoraInicio = (FechaInicio.hour / 24) + (FechaInicio.minute / 1440) + (FechaInicio.second / 86400)
    Horafinal = (FechaFinal.hour / 24) + (FechaFinal.minute / 1440) + (FechaFinal.second / 86400)

    # Calcular la duración en segundos
    duracion_segundos = (FechaFinal - FechaInicio).total_seconds()
    # Convertir a minutos y segundos
    minutos = int(duracion_segundos // 60)
    segundos = int(duracion_segundos % 60)
    # Formato para mostrar
    duracion_formateada = f"{minutos:02}:{segundos:02}"
    
    print(f"Duración calculada: {duracion_formateada}")

    # Establecer las horas de FechaInicio y FechaFinal en 00:00:00 para la inserción de fechas
    FechaInicioInsert = FechaInicio.replace(hour=0, minute=0, second=0, microsecond=0)
    FechaFinalInsert = FechaFinal.replace(hour=0, minute=0, second=0, microsecond=0)

    # Insertar los datos en la base de datos
    cursor.execute("""
        INSERT INTO incidencias (incidencia, codigoempresa, Ejerciciotrabajo, Numerotrabajo,
                                 FechaInicio, FechaFinal, HoraInicio, Horafinal, Operario, Operacion, Orden)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (incidencia, codigoempresa, Ejerciciotrabajo, Numerotrabajo, 
          FechaInicioInsert, FechaFinalInsert, HoraInicio, Horafinal, Operario, "MONTAJE", "1"))

    conn.commit()
    conn.close()
    mostrar_aviso("Incidencia insertada Correctamente")

# Función para ajustar el ancho de las columnas al contenido
def ajustar_columnas(tree):
    for col in tree["columns"]:
        max_width = max([len(str(tree.set(item, col))) for item in tree.get_children()]) * 15
        tree.column(col, width=max(max_width, 100))  

# Evento para manejar la selección de operarios con solo clic
def seleccionar_operario(event):
    item = tree_operarios.identify_row(event.y)
    if item:
        if item in tree_operarios.selection():
            tree_operarios.selection_remove(item)  
        else:
            tree_operarios.selection_add(item)  
    return "break"  

# Evento para manejar la selección de órdenes con solo clic
def seleccionar_orden(event):
    item = tree_ordenes.identify_row(event.y)
    if item:
        if item in tree_ordenes.selection():
            tree_ordenes.selection_remove(item)  
        else:
            tree_ordenes.selection_add(item)  
    return "break"  

# Función para filtrar las órdenes de fabricación por DescripcionArticulo
def filtrar_ordenes():
    texto_busqueda = search_entry_ordenes.get().lower()  
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    # Consulta para buscar órdenes cuya DescripcionArticulo contenga el texto ingresado
    cursor.execute("""
        SELECT Ejerciciofabricacion, SerieFabricacion, NumeroFabricacion, CodigoArticulo, DescripcionArticulo
        FROM OrdenesFabricacion 
        WHERE codigoempresa = 2 
        AND EstadoOF = 1 
        AND VMostrarOrden = -1 
        AND DescripcionArticulo LIKE ?
        AND cast(EjercicioFabricacion as varchar)+'/'+SerieFabricacion+'/'+cast(NumeroFabricacion as varchar) 
        NOT IN (SELECT Orden FROM FabricacionTemporal)
    """, f"%{texto_busqueda}%")
    
    tree_ordenes.delete(*tree_ordenes.get_children())  
    for row in cursor.fetchall():
        row = ['N/A' if (r is None or r == '') else r for r in row]  
        row = [r.replace('"', "'") if isinstance(r, str) else r for r in row]  
        tree_ordenes.insert('', tk.END, values=row)  
    
    alternar_colores(tree_ordenes, COLOR1, COLOR2)
    conn.close()

# Función para filtrar las órdenes en curso por Descripcion
def filtrar_en_curso():
    texto_busqueda = search_entry_en_curso.get().lower()  
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    # Consulta para buscar fabricaciones en curso cuya Descripcion contenga el texto ingresado
    cursor.execute("""
        SELECT o.NombreOperario, f.Orden, f.Inicio, f.Color, f.Descripcion
        FROM FabricacionTemporal f
        JOIN operarios o ON f.Operario = o.Operario
        WHERE o.codigoempresa = 2 
        AND f.Descripcion LIKE ?
        ORDER BY f.Inicio
    """, f"%{texto_busqueda}%")
    
    tree_en_curso.delete(*tree_en_curso.get_children())  
    for row in cursor.fetchall():
        operario, orden, inicio, color_fondo, descripcion = row
        tiempo_transcurrido = ""
        item_id = tree_en_curso.insert("", tk.END, values=(operario, orden, inicio, tiempo_transcurrido, descripcion))
        tree_en_curso.item(item_id, tags=(color_fondo,))  
        tree_en_curso.tag_configure(color_fondo, background=color_fondo) 
    
    conn.close()

def mostrar_aviso(mensaje):
    aviso_label.config(text=mensaje)
    root.after(20000, lambda: aviso_label.config(text=""))

def actualizar_botones_scroll():
    # Verificar si está en la parte superior o inferior para ocultar botones
    if tree_ordenes.yview()[0] <= 0:
        boton_scroll_arriba.pack_forget()
    else:
        boton_scroll_arriba.pack()

    if tree_ordenes.yview()[1] >= 1:
        boton_scroll_abajo.pack_forget()
    else:
        boton_scroll_abajo.pack()

def scroll_arriba(event=None):
    tree_ordenes.yview_scroll(-1, "units")
    actualizar_botones_scroll()

def scroll_abajo(event=None):
    tree_ordenes.yview_scroll(1, "units")
    actualizar_botones_scroll()

# Configurar la ventana principal
root = tk.Tk()
root.title("Gestión de Operarios y Órdenes")
root.geometry("1366x768")
root.configure(bg='#C2C8D4')  

# Estilo de la fuente
estilo = ttk.Style()
estilo.configure("Treeview", font=("Open Sans", 11), rowheight=22)  
estilo.configure("Treeview.Heading", font=("Open Sans", 11, "bold"), foreground='#333333', background='#D5DBDB')

# Frame principal para columnas
frame_principal = tk.Frame(root, bg='#C2C8D4')
frame_principal.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Grid de Operarios
frame_operarios = tk.Frame(frame_principal, bg='#C2C8D4')
frame_operarios.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 20))  

# Cambiar el estilo de la etiqueta de Operarios
label_operarios = tk.Label(frame_operarios, text="Operarios", bg='#C2C8D4', font=("Arial Black", 19, "bold"), fg="#000000")  
label_operarios.pack(pady=5)

tree_operarios = ttk.Treeview(frame_operarios, columns=("Operario", "NombreOperario"), show="headings", height=15)
tree_operarios.heading("Operario", text="Operario")
tree_operarios.heading("NombreOperario", text="Nombre Operario")
tree_operarios.pack(fill=tk.BOTH, expand=True)

# Asignar ancho fijo a las columnas del grid de operarios
tree_operarios.column("Operario", width=100, anchor="center")  
tree_operarios.column("NombreOperario", width=200, anchor="center")  

# Grid de Órdenes y Órdenes en Curso
frame_ordenes_y_en_curso = tk.Frame(frame_principal, bg='#C2C8D4')
frame_ordenes_y_en_curso.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Grid de Órdenes
frame_ordenes = tk.Frame(frame_ordenes_y_en_curso, bg='#C2C8D4')
frame_ordenes.pack(fill=tk.BOTH, expand=True)

# Cambiar el estilo de la etiqueta de Órdenes de Fabricación
label_ordenes = tk.Label(frame_ordenes, text="Órdenes de Fabricación", bg='#C2C8D4', font=("Arial Black", 19, "bold"), fg="#000000")  
label_ordenes.pack(pady=5)

# Frame para el campo de búsqueda de órdenes
frame_busqueda_ordenes = tk.Frame(frame_ordenes, bg='#C2C8D4')
frame_busqueda_ordenes.pack(pady=5)

search_entry_ordenes = tk.Entry(frame_busqueda_ordenes, width=30)  
search_entry_ordenes.pack(side=tk.LEFT, padx=(0, 10))  

search_entry_ordenes.bind("<Return>", lambda event: filtrar_ordenes())
btn_buscar_ordenes = tk.Button(frame_busqueda_ordenes, text="Buscar", command=filtrar_ordenes) 
btn_buscar_ordenes.pack(side=tk.LEFT)

# Frame para el Treeview y botones de desplazamiento
frame_treeview_ordenes = tk.Frame(frame_ordenes, bg='#C2C8D4')
frame_treeview_ordenes.pack(fill=tk.BOTH, expand=True)

tree_ordenes = ttk.Treeview(frame_treeview_ordenes, columns=("Ejercicio", "Serie", "Número", "Código Artículo","Unidades", "Descripción"), show="headings", height=8)
tree_ordenes.heading("Ejercicio", text="")
tree_ordenes.heading("Serie", text="Serie")
tree_ordenes.heading("Número", text="Número")
tree_ordenes.heading("Código Artículo", text="Código Artículo")
tree_ordenes.heading("Unidades", text="Unidades")
tree_ordenes.heading("Descripción", text="Descripción")
tree_ordenes.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Asignar ancho fijo a las columnas del grid de órdenes
tree_ordenes.column("Ejercicio", width=0, anchor="center",stretch=tk.NO)
tree_ordenes.column("Serie", width=60, anchor="center")
tree_ordenes.column("Número", width=60, anchor="center")
tree_ordenes.column("Código Artículo", width=100, anchor="center")
tree_ordenes.column("Unidades", width=60, anchor="center")
tree_ordenes.column("Descripción", width=300)

# Crear frame para los botones de scroll a la derecha de la tabla
frame_botones_scroll = tk.Frame(frame_treeview_ordenes, bg='#C2C8D4')
frame_botones_scroll.pack(side=tk.RIGHT, fill=tk.Y)

# Frame para el botón de desplazamiento hacia arriba
frame_scroll_arriba = tk.Frame(frame_botones_scroll, bg='#C2C8D4')
frame_scroll_arriba.pack(side=tk.TOP, pady=(10, 0), padx=(10, 0))  

boton_scroll_arriba = tk.Button(frame_scroll_arriba, text="▲", command=lambda: scroll_arriba(), width=3)
boton_scroll_arriba.pack()

# Frame para el botón de desplazamiento hacia abajo
frame_scroll_abajo = tk.Frame(frame_botones_scroll, bg='#C2C8D4')
frame_scroll_abajo.pack(side=tk.BOTTOM, pady=(0, 10),padx=(10, 0))  

boton_scroll_abajo = tk.Button(frame_scroll_abajo, text="▼", command=lambda: scroll_abajo(), width=3)
boton_scroll_abajo.pack()

# Actualizar visibilidad de botones de scroll al inicio
actualizar_botones_scroll()

# Vincular eventos de cambio de posición para mostrar/ocultar los botones de desplazamiento
tree_ordenes.bind("<Configure>", lambda event: actualizar_botones_scroll())
tree_ordenes.bind("<MouseWheel>", lambda event: actualizar_botones_scroll())


# Botón Info para Órdenes de Fabricación en la parte superior derecha
btn_info_ordenes = tk.Button(frame_ordenes, text="Info", command=lambda: info())
btn_info_ordenes.pack(pady=5, anchor="ne")  # Colocar en la parte superior derecha

# Grid de En curso
frame_en_curso = tk.Frame(frame_ordenes_y_en_curso, bg='#C2C8D4')
frame_en_curso.pack(fill=tk.BOTH, expand=True)

# Cambiar el estilo de la etiqueta de Órdenes en Curso
label_en_curso = tk.Label(frame_en_curso, text="Órdenes en Curso", bg='#C2C8D4', font=("Arial Black", 19, "bold"), fg="#000000")  
label_en_curso.pack(pady=5)

# Frame para el campo de búsqueda de órdenes en curso
frame_busqueda_en_curso = tk.Frame(frame_en_curso, bg='#C2C8D4')
frame_busqueda_en_curso.pack(pady=5)

search_entry_en_curso = tk.Entry(frame_busqueda_en_curso, width=30)  
search_entry_en_curso.pack(side=tk.LEFT, padx=(0, 10))  

search_entry_en_curso.bind("<Return>", lambda event: filtrar_en_curso())
btn_buscar_en_curso = tk.Button(frame_busqueda_en_curso, text="Buscar", command=filtrar_en_curso)  
btn_buscar_en_curso.pack(side=tk.LEFT)

tree_en_curso = ttk.Treeview(frame_en_curso, columns=("Operario", "Orden", "Inicio", "Tiempo Transcurrido","Descripcion"), show="headings", height=8)
tree_en_curso.heading("Operario", text="Operario")
tree_en_curso.heading("Orden", text="Fabricación")
tree_en_curso.heading("Inicio", text="Inicio")
tree_en_curso.heading("Tiempo Transcurrido", text="Tiempo")
tree_en_curso.heading("Descripcion", text="Descripcion") 
tree_en_curso.pack(fill=tk.BOTH, expand=True)

# Asignar ancho fijo a las columnas del grid de en curso
tree_en_curso.column("Operario", width=100, anchor="center")
tree_en_curso.column("Orden", width=80, anchor="center")
tree_en_curso.column("Inicio", width=100, anchor="center")
tree_en_curso.column("Tiempo Transcurrido", width=60, anchor="center")
tree_en_curso.column("Descripcion", width=350)

# Botón Info para Órdenes en Curso en la parte superior derecha
btn_info_en_curso = tk.Button(frame_en_curso, text="Info", command=lambda: info2())
btn_info_en_curso.pack(pady=5, anchor="ne")  # Colocar en la parte superior derecha

# Botones debajo de la sección de Operarios
frame_botones = tk.Frame(root, bg='#C2C8D4')
frame_botones.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(0, 10))  

btn_iniciar = tk.Button(frame_botones, text="Iniciar Fabricación", command=iniciar_fabricacion, bg='#5DADE2', fg='white', font=("Open Sans", 9))
btn_iniciar.pack(side=tk.LEFT, padx=10)

btn_terminar = tk.Button(frame_botones, text="Terminar Fabricación", command=mostrar_mensaje, bg='#E74C3C', fg='white', font=("Open Sans", 9))
btn_terminar.pack(side=tk.LEFT, padx=10)

# Crear un label para mostrar mensajes de aviso
aviso_label = tk.Label(frame_botones, text="", font=("Arial", 12), bg='#C2C8D4', fg="blue")
aviso_label.pack(pady=20)

# Asociar los eventos a los Treeview correspondientes
tree_operarios.bind('<Button-1>', seleccionar_operario)
tree_ordenes.bind('<Button-1>', seleccionar_orden)


# Cargar datos en los grids
cargar_operarios()
cargar_ordenes()
# cargar_fabricaciones_en_curso()

# Iniciar el contador de tiempos
actualizar_tiempos()

# Al final de tu configuración, añade la llamada a la función periódica
def ejecutar_carga_periodica():
    if not dialogo_abierto:
        cargar_fabricaciones_en_curso()
        cargar_ordenes()   
    root.after(20000, ejecutar_carga_periodica)  

# Iniciar la carga periódica
ejecutar_carga_periodica()

def ejecutar_impresion():
    conn = obtener_conexion()
    cursor = conn.cursor()


     # Definir la consulta SQL
    query = """
    SELECT 
        (SELECT Vvv FROM articulos WHERE codigoempresa=ArticulosSeries.codigoempresa AND codigoarticulo=ArticulosSeries.codigoarticulo) AS VV,
        (SELECT VFilas FROM articulos WHERE codigoempresa=ArticulosSeries.codigoempresa AND codigoarticulo=ArticulosSeries.codigoarticulo) AS Filas,
        codigoarticulo,
        NumeroSerieLc,
        YEAR(FechaInicial) AS Año
    FROM ArticulosSeries
    WHERE codigoempresa=2
      AND statusdisponible=-1
      AND VStatusImpreso=0
    """
    
    # Ejecutar la consulta
    cursor.execute(query)
    results = cursor.fetchall()
    

    # Verificar si hay resultados
    if not results:
        print("No hay registros que procesar.")
        return
    
    # Crear la carpeta "etiquetas" si no existe
    os.makedirs("etiquetas", exist_ok=True)
    
    # Iterar sobre cada registro y crear un PDF individual
    for row in results:
        VV, Filas, codigoarticulo, NumeroSerieLc, Año = row
        
        # Crear el nombre del PDF usando Año y CodigoArticulo, dentro de la carpeta "etiquetas"
        pdf_path = os.path.join("etiquetas", f"{Año}_{codigoarticulo}.pdf")
        
        # Crear el PDF en tamaño pequeño
        c = canvas.Canvas(pdf_path, pagesize=(70 * mm, 40 * mm))
        
        # Agregar imagen del encabezado completa
        header_image_path = "logo.jpg"  # Reemplaza por la ruta de tu imagen de encabezado completa
        c.drawImage(header_image_path, 5 * mm, 28 * mm, width=60 * mm, height=12 * mm)  # Ajusta la posición y tamaño según sea necesario

        # Material PPR, VV y Filas
        c.setFont("Helvetica", 6)
        c.drawString(5 * mm, 24 * mm, "MATERIAL PPR")
        
        # Mostrar valores de VV y Filas sin cuadros
        c.setFont("Helvetica-Bold", 8)
        c.drawString(25 * mm, 24 * mm, f"{VV}")  # Valor VV
        c.drawString(34 * mm, 24 * mm, "V.V.")    # Texto "V.V."
        c.drawString(44 * mm, 24 * mm, f"{Filas}")  # Valor Filas
        c.drawString(52 * mm, 24 * mm, "FILAS")    # Texto "FILAS"

        # Agregar líneas de tabla para el resto de los campos, con separación entre "Código" y "Año/Serie"
        # Línea alrededor de "Código"
        c.setFont("Helvetica", 6)
        c.drawString(5 * mm, 18 * mm, "CODIGO")
        c.rect(20 * mm, 16 * mm, 45 * mm, 6 * mm)  # Caja alrededor de código
        c.setFont("Helvetica-Bold", 8)
        c.drawString(22 * mm, 18.5 * mm, f"{codigoarticulo}")
        
        # Espacio de separación entre "Código" y "Año/Serie"
        
        # Línea alrededor de "Año/Serie"
        c.setFont("Helvetica", 6)
        c.drawString(5 * mm, 10.5 * mm, "AÑO/SERIE")
        c.rect(20 * mm, 8 * mm, 45 * mm, 6 * mm)  # Caja alrededor de año y serie
        c.setFont("Helvetica-Bold", 8)
        c.drawString(22 * mm, 10.5 * mm, f"{Año} / {NumeroSerieLc}")
        
        # Pie de página
        c.setFont("Helvetica", 5)
        c.drawString(18 * mm, 3 * mm, "FABRICACION SEGUN NORMA UNE 53943")
        
        # Guardar el PDF
        c.save()
        print(f"PDF generado exitosamente: {pdf_path}")

        # Actualizar el estado de impresión en la base de datos
        update_query = """
        UPDATE ArticulosSeries
        SET VStatusImpreso = -1
        WHERE codigoempresa = 2 AND codigoarticulo = ? AND NumeroSerieLc = ?
        """
        cursor.execute(update_query, (codigoarticulo, NumeroSerieLc))
        conn.commit()  # Asegúrate de confirmar los cambios

        # # Enviar a la impresora virtual PDF
        # try:
        #     # Comando para imprimir en Windows
        #     subprocess.run(['print', '/d:"Microsoft Print to PDF"', pdf_path], check=True)
        #     print(f"PDF enviado a la impresora virtual: {pdf_path}")
        # except Exception as e:
        #     print(f"Error al intentar imprimir el PDF: {e}")

    conn.close()

    root.after(10000, ejecutar_impresion)  

# Iniciar la carga periódica
ejecutar_impresion()




# Funciones placeholders para los botones de Info
def info():
    selected_items = tree_ordenes.selection()

    if not selected_items:
        mostrar_aviso("Por favor, selecciona una orden primero.")
        return
    elif len(selected_items) > 1:
        mostrar_aviso("Por favor, selecciona solo una orden.")
        return

    selected_orden = tree_ordenes.item(tree_ordenes.selection()[0], 'values')
    
    if selected_orden:
        ejercicioFab = selected_orden[0]
        serieFab = selected_orden[1]
        numeroFab = selected_orden[2]

    ventana_info = tk.Toplevel()
    ventana_info.title("Detalles de Orden")
    ventana_info.geometry("1366x700")
    ventana_info.configure(bg='#C2C8D4')
    
    ventana_info.transient(root)
    ventana_info.grab_set()


    # Calcular la posición de la ventana de diálogo
    ventana_ancho = 1366
    ventana_alto = 700

    # Centrar la ventana de diálogo en la ventana principal
    x = root.winfo_x() + (root.winfo_width() // 2) - (ventana_ancho // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (ventana_alto // 2)
    ventana_info.geometry(f"{ventana_ancho}x{ventana_alto}+{x}+{y}")  
    
    frame_info_principal = tk.Frame(ventana_info, bg='#C2C8D4')
    frame_info_principal.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Frame para Órdenes de Trabajo
    frame_ordenes_trabajo = tk.Frame(frame_info_principal, bg='#C2C8D4')
    frame_ordenes_trabajo.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=(0, 20))

    label_ordenes_trabajo = tk.Label(frame_ordenes_trabajo, text="Órdenes de Trabajo", bg='#C2C8D4', font=("Arial Black", 15), fg="#000000")
    label_ordenes_trabajo.pack(pady=5)

    tree_ordenes_trabajo = ttk.Treeview(frame_ordenes_trabajo, columns=("Nivel", "EjercicioTrabajo", "NumeroTrabajo", "CodigoArticulo", "DescripcionArticulo", "UnidadesAFabricar"), show="headings", height=10)
    tree_ordenes_trabajo.heading("Nivel", text="Nivel")
    tree_ordenes_trabajo.heading("EjercicioTrabajo", text="Ejercicio Trabajo")
    tree_ordenes_trabajo.heading("NumeroTrabajo", text="Número Trabajo")
    tree_ordenes_trabajo.heading("CodigoArticulo", text="Código Artículo")
    tree_ordenes_trabajo.heading("DescripcionArticulo", text="Descripción ")
    tree_ordenes_trabajo.heading("UnidadesAFabricar", text="Unidades")
    tree_ordenes_trabajo.pack(fill=tk.BOTH, expand=True)

    # Ajuste de ancho de columnas
    tree_ordenes_trabajo.column("Nivel", width=60, anchor="center")
    tree_ordenes_trabajo.column("EjercicioTrabajo", width=100, anchor="center")
    tree_ordenes_trabajo.column("NumeroTrabajo", width=100, anchor="center")
    tree_ordenes_trabajo.column("CodigoArticulo", width=100, anchor="center")
    tree_ordenes_trabajo.column("DescripcionArticulo", width=300, anchor="w")  # Columna más ancha
    tree_ordenes_trabajo.column("UnidadesAFabricar", width=100, anchor="center")

    # Ajuste de ancho de columnas para una mejor visualización
    for col in tree_ordenes_trabajo["columns"]:
        tree_ordenes_trabajo.column(col, anchor="center")

    # Frame para Consumo
    frame_consumo = tk.Frame(frame_info_principal, bg='#C2C8D4')
    frame_consumo.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=(0, 20))

    label_consumo = tk.Label(frame_consumo, text="Consumo", bg='#C2C8D4', font=("Arial Black", 15), fg="#000000")
    label_consumo.pack(pady=5)

    tree_consumo = ttk.Treeview(frame_consumo, columns=("Orden", "ArticuloComponente", "DescripcionArticulo", "UnidadesComponente"), show="headings", height=10)
    tree_consumo.heading("Orden", text="Orden")
    tree_consumo.heading("ArticuloComponente", text="Artículo Componente")
    tree_consumo.heading("DescripcionArticulo", text="Descripción del Artículo")
    tree_consumo.heading("UnidadesComponente", text="Unidades Componente")
    tree_consumo.pack(fill=tk.BOTH, expand=True)

    # Ajuste de ancho de columnas para consumo
    tree_consumo.column("Orden", width=100, anchor="center")
    tree_consumo.column("ArticuloComponente", width=150, anchor="center")
    tree_consumo.column("DescripcionArticulo", width=300, anchor="w")  # Columna más ancha
    tree_consumo.column("UnidadesComponente", width=100, anchor="center")

    # Ajuste de ancho de columnas para una mejor visualización
    for col in tree_consumo["columns"]:
        tree_consumo.column(col, anchor="center")
    
    # Llamada a cargar_ordenes_trabajo
    cargar_ordenes_trabajo(tree_ordenes_trabajo, ejercicioFab, serieFab, numeroFab)
    
    # Asociar el evento de selección
    tree_ordenes_trabajo.bind("<<TreeviewSelect>>", lambda event: actualizar_consumo(tree_ordenes_trabajo, tree_consumo))
    # Botón "Volver"
    btn_volver = tk.Button(frame_info_principal, text="Volver", command=ventana_info.destroy)
    btn_volver.pack(side=tk.BOTTOM, pady=(10, 0))  # Agregar margen superior



def cargar_ordenes_trabajo(tree_ordenes_trabajo,ejercicio_fabricacion, serie_fabricacion, numero_fabricacion):
    # Conectarse a la base de datos
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    # Ejecutar la consulta SQL con los parámetros proporcionados
    cursor.execute("""
        SELECT NivelCompuesto, EjercicioTrabajo, NumeroTrabajo, codigoarticulo, 
               DescripcionArticulo, UnidadesAFabricar 
        FROM Ordenestrabajo
        WHERE codigoempresa = 2
        AND EjercicioFabricacion = ?
        AND SerieFabricacion = ?
        AND NumeroFabricacion = ?
        ORDER BY NivelCompuesto DESC
    """, (ejercicio_fabricacion, serie_fabricacion, numero_fabricacion))
    
    # Limpiar los datos actuales en el Treeview de Órdenes de Trabajo
    tree_ordenes_trabajo.delete(*tree_ordenes_trabajo.get_children())
    
    # Insertar los datos obtenidos en el Treeview
    for row in cursor.fetchall():
        nivel_compuesto = str(row[0]).strip()
        ejercicio_trabajo = str(row[1]).strip()
        numero_trabajo = str(row[2]).strip()
        codigo_articulo = str(row[3]).strip()
        descripcion_articulo = str(row[4]).strip()
        unidades_a_fabricar = str(row[5]).strip()
        
        # Insertar cada fila en el Treeview de Órdenes de Trabajo
        tree_ordenes_trabajo.insert("", tk.END, values=(
            nivel_compuesto, ejercicio_trabajo, numero_trabajo, 
            codigo_articulo, descripcion_articulo, unidades_a_fabricar
        ))
    
    # Aplicar alternancia de colores en las filas del Treeview
    alternar_colores(tree_ordenes_trabajo, COLOR1, COLOR2)
    
    # Cerrar la conexión a la base de datos
    conn.close()

def actualizar_consumo(tree_ordenes_trabajo, tree_consumo):
    # Obtener la selección actual
    selected_item = tree_ordenes_trabajo.selection()
    if not selected_item:
        return

    selected_values = tree_ordenes_trabajo.item(selected_item[0], 'values')
    ejercicio_trabajo = selected_values[1]  # EjercicioTrabajo
    numero_trabajo = selected_values[2]      # NumeroTrabajo

    # Conectarse a la base de datos para cargar los datos de consumo
    conn = obtener_conexion()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT Orden, ArticuloComponente, DescripcionArticulo, UnidadesComponente 
        FROM ConsumosOT 
        WHERE codigoempresa = 2 
        AND EjercicioTrabajo = ?  
        AND NumeroTrabajo = ? 
    """, (ejercicio_trabajo, numero_trabajo))

    # Limpiar los datos actuales en el Treeview de Consumo
    tree_consumo.delete(*tree_consumo.get_children())

    for row in cursor.fetchall():
        orden = str(row[0]).strip()
        articulocomponente = str(row[1]).strip()
        descripcion_articulo = str(row[2]).strip()
        unidadesComponente = str(row[3]).strip()
        
        
        # Insertar cada fila en el Treeview de Órdenes de Trabajo
        tree_consumo.insert("", tk.END, values=(
            orden, articulocomponente, descripcion_articulo, 
            unidadesComponente, 
        ))

     # Aplicar alternancia de colores en las filas del Treeview
    alternar_colores(tree_consumo, COLOR1, COLOR2)

    conn.close()

def info2():
    selected_items = tree_en_curso.selection()

    if not selected_items:
        mostrar_aviso("Por favor, selecciona una orden primero.")
        return
    elif len(selected_items) > 1:
        mostrar_aviso("Por favor, selecciona solo una orden.")
        return

    selected_orden = tree_en_curso.item(selected_items[0], 'values')
    
    if selected_orden:
        # Separar el registro en ejercicioFab, serieFab y numeroFab
        ejercicioFab, serieFab, numeroFab = selected_orden[1].split('/')

    ventana_info = tk.Toplevel()
    ventana_info.title("Detalles de Orden")
    ventana_info.geometry("1366x700")
    ventana_info.configure(bg='#C2C8D4')
    
    ventana_info.transient(root)
    ventana_info.grab_set()

    # Calcular la posición de la ventana de diálogo
    ventana_ancho = 1366
    ventana_alto = 700

    # Centrar la ventana de diálogo en la ventana principal
    x = root.winfo_x() + (root.winfo_width() // 2) - (ventana_ancho // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (ventana_alto // 2)
    ventana_info.geometry(f"{ventana_ancho}x{ventana_alto}+{x}+{y}")  
    
    frame_info_principal = tk.Frame(ventana_info, bg='#C2C8D4')
    frame_info_principal.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Frame para Órdenes de Trabajo
    frame_ordenes_trabajo = tk.Frame(frame_info_principal, bg='#C2C8D4')
    frame_ordenes_trabajo.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=(0, 20))

    label_ordenes_trabajo = tk.Label(frame_ordenes_trabajo, text="Órdenes de Trabajo", bg='#C2C8D4', font=("Arial Black", 15), fg="#000000")
    label_ordenes_trabajo.pack(pady=5)

    tree_ordenes_trabajo = ttk.Treeview(frame_ordenes_trabajo, columns=("Nivel", "EjercicioTrabajo", "NumeroTrabajo", "CodigoArticulo", "DescripcionArticulo", "UnidadesAFabricar"), show="headings", height=10)
    tree_ordenes_trabajo.heading("Nivel", text="Nivel")
    tree_ordenes_trabajo.heading("EjercicioTrabajo", text="Ejercicio Trabajo")
    tree_ordenes_trabajo.heading("NumeroTrabajo", text="Número Trabajo")
    tree_ordenes_trabajo.heading("CodigoArticulo", text="Código Artículo")
    tree_ordenes_trabajo.heading("DescripcionArticulo", text="Descripción ")
    tree_ordenes_trabajo.heading("UnidadesAFabricar", text="Unidades")
    tree_ordenes_trabajo.pack(fill=tk.BOTH, expand=True)

    # Ajuste de ancho de columnas
    tree_ordenes_trabajo.column("Nivel", width=60, anchor="center")
    tree_ordenes_trabajo.column("EjercicioTrabajo", width=100, anchor="center")
    tree_ordenes_trabajo.column("NumeroTrabajo", width=100, anchor="center")
    tree_ordenes_trabajo.column("CodigoArticulo", width=100, anchor="center")
    tree_ordenes_trabajo.column("DescripcionArticulo", width=300, anchor="w")  # Columna más ancha
    tree_ordenes_trabajo.column("UnidadesAFabricar", width=100, anchor="center")

    # Frame para Consumo
    frame_consumo = tk.Frame(frame_info_principal, bg='#C2C8D4')
    frame_consumo.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=(0, 20))

    label_consumo = tk.Label(frame_consumo, text="Consumo", bg='#C2C8D4', font=("Arial Black", 15), fg="#000000")
    label_consumo.pack(pady=5)

    tree_consumo = ttk.Treeview(frame_consumo, columns=("Orden", "ArticuloComponente", "DescripcionArticulo", "UnidadesComponente"), show="headings", height=10)
    tree_consumo.heading("Orden", text="Orden")
    tree_consumo.heading("ArticuloComponente", text="Artículo Componente")
    tree_consumo.heading("DescripcionArticulo", text="Descripción del Artículo")
    tree_consumo.heading("UnidadesComponente", text="Unidades Componente")
    tree_consumo.pack(fill=tk.BOTH, expand=True)

    # Ajuste de ancho de columnas para consumo
    tree_consumo.column("Orden", width=100, anchor="center")
    tree_consumo.column("ArticuloComponente", width=150, anchor="center")
    tree_consumo.column("DescripcionArticulo", width=300, anchor="w")  # Columna más ancha
    tree_consumo.column("UnidadesComponente", width=100, anchor="center")

    # Llamada a cargar_ordenes_trabajo
    cargar_ordenes_trabajo(tree_ordenes_trabajo, ejercicioFab, serieFab, numeroFab)

    # Asociar el evento de selección
    tree_ordenes_trabajo.bind("<<TreeviewSelect>>", lambda event: actualizar_consumo(tree_ordenes_trabajo, tree_consumo))

    # Botón "Volver"
    btn_volver = tk.Button(frame_info_principal, text="Volver", command=ventana_info.destroy)
    btn_volver.pack(side=tk.BOTTOM, pady=(10, 0))  # Agregar margen superior


root.mainloop()
