import pandas as pd
import pyodbc
from datetime import datetime
import calendar
from conexion import conectar
import os

def insert_and_get_id(cursor, query, params, descripcion="", fila_excel=None):
    try:
        cursor.execute(query, params)
        cursor.nextset()
        result = cursor.fetchone()
        new_id = result[0] if result else None
        #escribir_log(f"{descripcion} insertado con ID: {new_id}")
        return new_id
    except Exception as e:
        if fila_excel is not None:
            escribir_log(f"Error al insertar {descripcion} en fila Excel {fila_excel}: {e}")            
        else:
            escribir_log(f"Error al insertar {descripcion}: {e}")
        return None
    
def buscar_id_por_like(cursor, query, valor_busqueda, descripcion="", fila_excel=None):
    try:
        if pd.isna(valor_busqueda) or valor_busqueda.strip() == "":
            if fila_excel is not None:
                escribir_log(f"{descripcion} está vacío o nulo. (fila Excel {fila_excel})")
            else:
                escribir_log(f"{descripcion} está vacío o nulo.")
            return None

        valor_param = f"%{valor_busqueda.strip()}%"
        cursor.execute(query, (valor_param,))
        result = cursor.fetchone()

        if result:
            return result[0]
        else:
            if fila_excel is not None:
                escribir_log(f"{descripcion} no encontrado para '{valor_param}' (fila Excel {fila_excel})")
            else:
                escribir_log(f"{descripcion} no encontrado para '{valor_param}'")
            return None

    except Exception as e:
        if fila_excel is not None:
            escribir_log(f"Error al buscar {descripcion} (fila Excel {fila_excel}): {e}")
        else:
            escribir_log(f"Error al buscar {descripcion}: {e}")
        return None

def buscar_director_por_iniciales(cursor, iniciales, rol="DIRECTOR COMERCIAL"):
    if not isinstance(iniciales, str) or len(iniciales.strip()) != 2:
        escribir_log(f"Iniciales inválidas para DIRECTOR COMERCIAL: {iniciales}")
        return None

    try:
        iniciales = iniciales.strip().upper()
        inicial_nombre = iniciales[0] + "%"
        inicial_apellido = iniciales[1] + "%"
        query = """
            SELECT USU_ID FROM VST_USUARIOS 
            WHERE ROL_DESCRIPCION = ?
            AND USU_NOMBRE LIKE ? 
            AND USU_APELLIDO_PATERNO LIKE ?
        """
        cursor.execute(query, (rol, inicial_nombre, inicial_apellido))
        result = cursor.fetchone()
        if result:
            #escribir_log(f"{rol} encontrado: {result[0]}")
            return result[0]
        else:
            escribir_log(f"{rol} no encontrado con iniciales: {iniciales}")
            return None
    except Exception as e:
        escribir_log(f"Error al buscar {rol}: {e}")
        return None

def obtener_fecha_estimada_cierre(mes_str, año=2024):
    try:
        if not isinstance(mes_str, str):
            return None
        mapeo_meses = {
            "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
            "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12
        }
        mes = mapeo_meses.get(mes_str.strip().upper())
        if not mes:
            return None
        ultimo_dia = calendar.monthrange(año, mes)[1]
        return datetime(año, mes, ultimo_dia)
    except Exception as e:
        escribir_log(f"Error al calcular fecha estimada de cierre: {e}")
        return None

def escribir_log(mensaje):
    fecha = datetime.now().strftime("%Y-%m-%d")
    nombre_archivo = f"log_{fecha}.txt"
    ruta_log = os.path.join(os.getcwd(), nombre_archivo)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(ruta_log, "a", encoding="utf-8") as f:
        if "comienza ejecución" in mensaje.lower():
            f.write(f"\n[{timestamp}] ===== COMIENZA EJECUCIÓN =====\n")
            f.write(f"{'=' * 60}\n")
        f.write(f"[{timestamp}] {mensaje}\n")


ruta_excel = r"C:\CargasIO\carga.xlsx"

registros_insertados = 0
registros_omitidos = 0
errores = 0

df = pd.read_excel(ruta_excel,engine='openpyxl')
#escribir_log("Archivo cargado correctamente")
#print((df.head))
# Conexión a la base de datos
conn = conectar()
cursor = conn.cursor()

for index, row in df.iterrows():
    fila_excel = index 
    try: 
        if row['ESTATUS CORTO'] != 'Declinado' and row['ESTATUS CORTO'] != 'No adjudicado' and row['ESTATUS CORTO'] != '':
            escribir_log(f"Procesando fila {fila_excel}")

            # Valores fijos
            usuario_id_carga = 3
            pap_tes_id = 499

            pap_id = insert_and_get_id(cursor, """
                INSERT INTO IOP.PARTICIPACION_PROSPECIONES (PAP_TES_ID, PAP_FECHA_REGISTRO, PAP_USU_ID)
                VALUES (?, GETDATE(), ?); SELECT SCOPE_IDENTITY();
            """, (pap_tes_id, usuario_id_carga), "PARTICIPACION_PROSPECIONES", fila_excel=fila_excel)

            elp_id = insert_and_get_id(cursor, """
                INSERT INTO IOP.EJECUTIVOS_LICITACIONES_PROSPECIONES (ELP_FECHA_REGISTRO, ELP_USU_ID)
                VALUES (GETDATE(), ?); SELECT SCOPE_IDENTITY();
            """, (usuario_id_carga,), "EJECUTIVOS_LICITACIONES_PROSPECIONES", fila_excel=fila_excel)

            brp_id = insert_and_get_id(cursor, """
                INSERT INTO IOP.BASES_PROSPECIONES (BRP_FECHA_PUBLICACION_BASES,BRP_FECHA_REGISTRO,BRP_USU_ID)
                VALUES (GETDATE(), GETDATE(), ?); 
                SELECT SCOPE_IDENTITY();
            """, (usuario_id_carga,), "BASES_PROSPECIONES", fila_excel=fila_excel)


            cpp_id = insert_and_get_id(cursor, """
                INSERT INTO IOP.CEDULAS_PADRON_PROVEEDORES_PROSPECIONES (CPP_FECHA_REGISTRO, CPP_USU_ID)
                VALUES (GETDATE(), ?); SELECT SCOPE_IDENTITY();
            """, (usuario_id_carga,), "CEDULAS_PADRON_PROVEEDORES_PROSPECIONES", fila_excel=fila_excel)

            jpo_id = insert_and_get_id(cursor, """
                INSERT INTO IOP.JUNTA_PRE_OPERACION_PROSPECIONES (JPO_FECHA_REGISTRO, JPO_USU_ID)
                VALUES (GETDATE(), ?); SELECT SCOPE_IDENTITY();
            """, (usuario_id_carga,), "JUNTA_PRE_OPERACION_PROSPECIONES", fila_excel=fila_excel)

            jap_id = insert_and_get_id(cursor, """
                INSERT INTO IOP.JUNTA_ACLARACIONES_PROSPECIONES (JAP_FECHA_REGISTRO, JAP_USU_ID)
                VALUES (GETDATE(), ?); SELECT SCOPE_IDENTITY();
            """, (usuario_id_carga,), "JUNTA_ACLARACIONES_PROSPECIONES", fila_excel=fila_excel)

            scr_id = insert_and_get_id(cursor, """
                INSERT INTO IOP.SOSTENIMIENTOS_CERTIFICADOS_PROSPECIONES (SCR_FECHA_REGISTRO, SCR_USU_ID)
                VALUES (GETDATE(), ?); SELECT SCOPE_IDENTITY();
            """, (usuario_id_carga,), "SOSTENIMIENTOS_CERTIFICADOS_PROSPECIONES", fila_excel=fila_excel)


            # Se obtienen valores para realizar el insert en la tabla de OPERACIONES_COMERCIALES
            emp_id = buscar_id_por_like(cursor,
                "SELECT EMP_ID FROM EMPRESAS WHERE EMP_DESCRIPCION LIKE ?",
                row.get("EMPRESA", ""),
                "EMPRESA",
                fila_excel=fila_excel)
            if emp_id is None:
                registros_omitidos += 1
                continue

            cli_id_comercial = buscar_id_por_like(cursor,
                "SELECT USU_ID FROM VST_USUARIOS WHERE (LTRIM(RTRIM(USU_NOMBRE)) + ' ' + LTRIM(RTRIM(USU_APELLIDO_PATERNO))) LIKE ?",
                row.get("EJECUTIVO COMERCIAL", ""),
                "EJECUTIVO COMERCIAL",
                fila_excel=fila_excel)
            if cli_id_comercial is None:
                registros_omitidos += 1
                continue

            cli_id_prospecto = buscar_id_por_like(cursor,
                "SELECT CLI_ID FROM CLIENTES WHERE CLI_NOMBRE LIKE ?",
                row.get("NOMBRE DEL PROSPECTO", ""),
                "NOMBRE DEL PROSPECTO",
                fila_excel=fila_excel)
            if cli_id_prospecto is None:
                registros_omitidos += 1
                continue


            director_usu_id = buscar_director_por_iniciales(cursor, row.get("DIRECTOR COMERCIAL", ""))
            if director_usu_id is None:
                registros_omitidos += 1
                continue

            query_productos_prospecciones = """
                SELECT PRP_ID FROM IOP.PRODUCTOS_PROSPECIONES 
                WHERE PRP_DESCRIPCION LIKE ?
            """
            linea_negocio = row.get('Linea de Negocio LumoSys', '')
            prp_id = buscar_id_por_like(cursor, query_productos_prospecciones, linea_negocio, "Línea de negocio",fila_excel=fila_excel)
            if prp_id is None:
                registros_omitidos += 1
                continue

            query_tipos_clientes_ordenes = """
                SELECT TNO_ID FROM IOP.TIPOS_CLIENTES_ORDENES_COMPRAS_ENTREGAS 
                WHERE TNO_DESCRIPCION LIKE ?
            """
            sector = row.get('SECTOR', '')
            tno_id = buscar_id_por_like(cursor, query_tipos_clientes_ordenes, sector, "Sector",fila_excel=fila_excel)
            if tno_id is None:
                registros_omitidos += 1
                continue

            # Se realiza el insert a la tabla de Operaciones Comerciales            
            fecha_estimada_cierre = obtener_fecha_estimada_cierre(row.get("MES DE CIERRE (FALLO)", ""))
            if fecha_estimada_cierre is None:
                escribir_log("Mes de cierre inválido o no reconocido, se omite la fila.")
                registros_omitidos += 1
                continue

            valor_contrato = row['V. CONTRATO I.V.A. INCLUIDO']
            valor_contrato = float(valor_contrato) if pd.notnull(valor_contrato) else 0
            plazo_meses = row.get("PLAZO (MESES)","")

            opp_id = insert_and_get_id(cursor,"""
                INSERT INTO IOP.OPERACIONES_COMERCIALES (
                OPP_NOMBRE,
                OPP_FOP_ID,
                OPP_EMP_ID,
                OPP_CLI_ID_COMERCIAL,
                OPP_PRP_ID,
                OPP_CLI_ID_PROSPECTO,
                OPP_TES_ID,
                OPP_TRO_ID,
                OPP_PLAZO_MESES,                       
                OPP_VALOR_CONTRATO,
                OPP_FECHA_REGISTRO,
                OPP_USU_ID,
                OPP_PAP_ID,
                OPP_ELP_ID,
                OPP_BRP_ID,
                OPP_CPP_ID,
                OPP_JPO_ID,
                OPP_JAP_ID,
                OPP_SCR_ID,
                OPP_TNO_ID,
                OPP_DIRECTOR_USU_ID,
                OPP_FECHA_ESTIMADO_CIERRE
            )
            VALUES (
                ?,  -- OPP_NOMBRE
                ?,  -- OPP_FOP_ID
                ?,  -- OPP_EMP_ID
                ?,  -- OPP_CLI_ID_COMERCIAL
                ?,  -- OPP_PRP_ID
                ?,  -- OPP_CLI_ID_PROSPECTO
                ?,  -- OPP_TES_ID
                ?,  -- OPP_TRO_ID
                ?,  -- OPP_PLAZO_MESES
                ?,  -- OPP_VALOR_CONTRATO
                GETDATE(),  -- OPP_FECHA_REGISTRO
                ?,  -- OPP_USU_ID
                ?,  -- OPP_PAP_ID
                ?,  -- OPP_ELP_ID
                ?,  -- OPP_BRP_ID
                ?,  -- OPP_CPP_ID
                ?,  -- OPP_JPO_ID
                ?,  -- OPP_JAP_ID
                ?,  -- OPP_SCR_ID
                ?,  -- OPP_TNO_ID
                ?,  -- OPP_DIRECTOR_USU_ID
                ?   -- OPP_FECHA_ESTIMADO_CIERRE
            );SELECT SCOPE_IDENTITY();
            """, (
                'OPERACIÓN 1',
                3,
                emp_id,
                cli_id_comercial,
                prp_id,
                cli_id_prospecto,
                517,
                tno_id,
                plazo_meses,
                valor_contrato,
                3,
                pap_id,
                elp_id,
                brp_id,
                cpp_id,
                jpo_id,
                jap_id,
                scr_id,
                tno_id,
                director_usu_id,
                fecha_estimada_cierre
            ),"OPERACIONES_COMERCIALES", fila_excel=fila_excel)


            # Se realiza el insert a la tabla de Bienes Activos Prospecciones
            mes_entrega = obtener_fecha_estimada_cierre(row.get("MES DE ENTREGA", ""))
            if mes_entrega is None:
                escribir_log("Mes de entrega inválido o no reconocido, se omite la fila.")
                registros_omitidos += 1
                continue

            valor_activo = row['V. CONTRATO I.V.A. INCLUIDO']
            valor_activo = float(valor_activo) if pd.notnull(valor_activo) else 0

            cantidad = row.get("# BIENES", "")
            descripcion = row.get("DESCRIPCIÓN DE LOS BIENES", "")

            if pd.isna(cantidad) or not str(cantidad).strip() or pd.isna(descripcion) or not str(descripcion).strip():
                escribir_log(f"No hay bienes en la fila {fila_excel}, se omite la fila completa.")
                registros_omitidos += 1
                continue

            elementoscant = str(cantidad).split("|")
            elementosdesc = str(descripcion).split("|")

            for cant, desc in zip(elementoscant,elementosdesc):
                
                bia_id = insert_and_get_id(cursor,"""
                    INSERT INTO IOP.BIENES_ACTIVOS_PROSPECIONES (
                        BIA_OPP_ID,
                        BIA_TBT_ID,
                        BIA_NO_BIENES,
                        BIA_DESCRIPCION_BIENES,
                        BIA_VALOR_ACTIVO_CON_IVA,
                        BIA_VALOR_PARTIDA_ACTIVO_CON_IVA,
                        BIA_FECHA_ENTREGA,
                        BIA_TBI_ID,
                        BIA_EMP_ID_ACTIVO,
                        BIA_FECHA_REGISTRO,
                        BIA_USU_ID
                    )
                    VALUES (
                        ?,    -- BIA_OPP_ID
                        ?,    -- BIA_TBT_ID
                        ?,    -- BIA_NO_BIENES
                        ?,    -- BIA_DESCRIPCION_BIENES
                        ?,    -- BIA_VALOR_ACTIVO_CON_IVA
                        ?,    -- BIA_VALOR_PARTIDA_ACTIVO_CON_IVA
                        ?,    -- BIA_FECHA_ENTREGA
                        ?,    -- BIA_TBI_ID
                        ?,    -- BIA_EMP_ID_ACTIVO                                       
                        GETDATE(), -- BIA_FECHA_REGISTRO
                        ?     -- BIA_USU_ID
                    );SELECT SCOPE_IDENTITY();
                """, (
                    opp_id,
                    2,
                    str(cant).strip(),
                    str(desc).strip(),
                    valor_activo,
                    valor_activo,
                    mes_entrega,
                    1,
                    emp_id,                
                    3
                ),"BIENES_ACTIVOS_PROSPECCIONES", fila_excel=fila_excel)
                valor_activo = 0

            
            comentarios = row.get("COMENTARIOS", "")
            if str(comentarios).strip():
                nombre_completo_ejecutivo = row.get("EJECUTIVO COMERCIAL", "")
                fecha_estimada_flag = 1 if fecha_estimada_cierre else 0  # bit: 1 si existe, 0 si no

                coe_id = insert_and_get_id(cursor, """
                    INSERT INTO IOP.COMENTARIOS_PROSPECIONES (
                        COE_OPP_ID,
                        COE_DESCRIPCION,
                        COE_FECHA_REGISTRO,
                        COE_USU_ID,
                        COE_NOMBRE_COMPLETO,
                        COE_FECHA_ESTIMADA_CIERRE,
                        COE_ADJUDICADA,
                        COE_CANCELADA
                    )
                    VALUES (?, ?, GETDATE(), ?, ?, ?, ?, ?);
                    SELECT SCOPE_IDENTITY();
                """, (
                    opp_id,
                    comentarios,
                    3,
                    nombre_completo_ejecutivo,
                    fecha_estimada_flag,
                    0,  # adjudicada
                    0   # cancelada
                ), "COMENTARIOS_PROSPECIONES", fila_excel=fila_excel)
            else:
                escribir_log(f"No hay comentario en la fila {fila_excel}")                

            comentarios14abril = row.get("COMENTARIOS AL 14 DE ABRIL 2024", "")
            if pd.notna(comentarios14abril) and str(comentarios14abril).strip():
                nombre_completo_ejecutivo = row.get("EJECUTIVO COMERCIAL", "")
                fecha_estimada_flag = 1 if fecha_estimada_cierre else 0  # bit: 1 si existe, 0 si no

                coe_id = insert_and_get_id(cursor, """
                    INSERT INTO IOP.COMENTARIOS_PROSPECIONES (
                        COE_OPP_ID,
                        COE_DESCRIPCION,
                        COE_FECHA_REGISTRO,
                        COE_USU_ID,
                        COE_NOMBRE_COMPLETO,
                        COE_FECHA_ESTIMADA_CIERRE,
                        COE_ADJUDICADA,
                        COE_CANCELADA
                    )
                    VALUES (?, ?, GETDATE(), ?, ?, ?, ?, ?);
                    SELECT SCOPE_IDENTITY();
                """, (
                    opp_id,
                    comentarios14abril,
                    3,
                    nombre_completo_ejecutivo,
                    fecha_estimada_flag,
                    0,  # adjudicada
                    0   # cancelada
                ), "COMENTARIOS_PROSPECIONES", fila_excel=fila_excel)                
            else:
                escribir_log(f"No hay comentario al 14 de abril en la fila {fila_excel}")

            conn.commit()
            registros_insertados += 1
            escribir_log("Todo insertado correctamente")

    except Exception as e:
        errores += 1
        escribir_log(f"Error en fila {index + 1}: {e}")
        escribir_log(f"Contenido de la fila con error: {row.to_dict()}")
        conn.rollback()
        registros_omitidos += 1
        continue


escribir_log(f"Resumen final:")
escribir_log(f"Registros insertados correctamente: {registros_insertados}")
escribir_log(f"Registros omitidos por validación: {registros_omitidos}")

cursor.close()
conn.close()
