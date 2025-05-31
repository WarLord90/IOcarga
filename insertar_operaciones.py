import pandas as pd
import pyodbc
from datetime import datetime
import calendar
from conexion import conectar

ruta_excel = r"C:\CargaIO\carga.xlsx"
try:
    df = pd.read_excel(ruta_excel,engine='openpyxl')
    print("Archivo cargado correctamente")
    #print((df.head))
    # Conexión a la base de datos
    conn = conectar()
    cursor = conn.cursor()
    
    for index, row in df.iterrows():
        if row['ESTATUS CORTO'] != 'Declinado' and row['ESTATUS CORTO'] != 'No adjudicado' and row['ESTATUS CORTO'] != '':
            print(f"Procesando fila {index +1} -> Línea de negocio: {row['ESTADO']}")

            insert_query = """
                INSERT INTO IOP.PARTICIPACION_PROSPECIONES (PAP_TES_ID, PAP_FECHA_REGISTRO, PAP_USU_ID)
                VALUES (?, GETDATE(), ?);
                SELECT SCOPE_IDENTITY();
            """
            # Valores fijos por ahora
            pap_tes_id = 499  # BACKLOG
            pap_usu_id = 3    # Usuario de carga masiva

            cursor.execute(insert_query,pap_tes_id,pap_usu_id)
            cursor.nextset()
            ultimo_id = cursor.fetchone()[0]
            print (f'El último PAP_ID insertado es: {ultimo_id}')

            conn.commit()                
            break

    cursor.close()
    conn.close()

except Exception as e:
    print("Error al abrir el Excel:", e)