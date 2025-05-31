import pandas as pd
import pyodbc
from datetime import datetime
import calendar
from conexion import conectar

# Mapeo de meses a números
MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12
}

class OperacionProcessor:
    def __init__(self, ruta_excel):
        self.ruta_excel = ruta_excel
        self.conn = conectar()
        self.cursor = self.conn.cursor()

    def obtener_ultimo_dia_mes(self, mes_texto):
        mes_limpio = str(mes_texto).strip().lower()
        numero_mes = MESES.get(mes_limpio)
        if numero_mes:
            ultimo_dia = calendar.monthrange(2024, numero_mes)[1]
            return datetime(2024, numero_mes, ultimo_dia)
        return None

    def procesar_filas(self):
        try:
            df = pd.read_excel(self.ruta_excel)
            for index, row in df.iterrows():
                fecha_cierre = self.obtener_ultimo_dia_mes(row.get("MES DE CIERRE (FALLO)"))

                if not fecha_cierre:
                    print(f"⚠️ Fila {index + 2} omitida por mes inválido")
                    continue

                pap_id = self.insertar_basica("PARTICIPACION_PROSPECIONES")
                elp_id = self.insertar_basica("EJECUTIVOS_LICITACIONES_PROSPECIONES")
                brp_id = self.insertar_basica("BASES_PROSPECIONES")
                cpp_id = self.insertar_basica("CEDULAS_PADRON_PROVEEDORES_PROSPECIONES")
                jpo_id = self.insertar_basica("JUNTA_PRE_OPERACION_PROSPECIONES")
                jap_id = self.insertar_basica("JUNTA_ACLARACIONES_PROSPECIONES")
                scr_id = self.insertar_basica("SOSTENIMIENTOS_CERTIFICADOS_PROSPECIONES")

                self.cursor.execute("""
                    INSERT INTO IOP.OPERACIONES_COMERCIALES (
                        OPP_NOMBRE, OPP_FOP_ID, OPP_CLI_ID_COMERCIAL,
                        OPP_EMP_ID, OPP_PRP_ID, OPP_CLI_ID_PROSPECTO,
                        OPP_TES_ID, OPP_VALOR_CONTRATO, OPP_FECHA_ESTIMADO_CIERRE,
                        OPP_USU_ID, OPP_PAP_ID, OPP_ELP_ID, OPP_BRP_ID, OPP_CPP_ID,
                        OPP_JPO_ID, OPP_JAP_ID, OPP_SCR_ID
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    f"Operacion fila {index+2}", 3, 1,
                    1, 1, 1,
                    517, row.get("V. CONTRATO I.V.A. INCLUIDO"), fecha_cierre,
                    3, pap_id, elp_id, brp_id, cpp_id, jpo_id, jap_id, scr_id
                ))
                self.conn.commit()
                print(f"✅ Fila {index + 2} insertada correctamente")

        except Exception as e:
            print("❌ Error procesando archivo:", e)
        finally:
            self.cursor.close()
            self.conn.close()

    def insertar_basica(self, tabla):
        campos = {
            "PARTICIPACION_PROSPECIONES": ("INSERT INTO IOP.PARTICIPACION_PROSPECIONES (PAP_TES_ID, PAP_FECHA_REGISTRO, PAP_USU_ID) VALUES (?, GETDATE(), ?)", 517),
            "EJECUTIVOS_LICITACIONES_PROSPECIONES": ("INSERT INTO IOP.EJECUTIVOS_LICITACIONES_PROSPECIONES (ELP_FECHA_REGISTRO, ELP_USU_ID) VALUES (GETDATE(), ?)", None),
            "BASES_PROSPECIONES": ("INSERT INTO IOP.BASES_PROSPECIONES (BRP_FECHA_REGISTRO, BRP_USU_ID) VALUES (GETDATE(), ?)", None),
            "CEDULAS_PADRON_PROVEEDORES_PROSPECIONES": ("INSERT INTO IOP.CEDULAS_PADRON_PROVEEDORES_PROSPECIONES (CPP_FECHA_REGISTRO, CPP_USU_ID) VALUES (GETDATE(), ?)", None),
            "JUNTA_PRE_OPERACION_PROSPECIONES": ("INSERT INTO IOP.JUNTA_PRE_OPERACION_PROSPECIONES (JPO_FECHA_REGISTRO, JPO_USU_ID) VALUES (GETDATE(), ?)", None),
            "JUNTA_ACLARACIONES_PROSPECIONES": ("INSERT INTO IOP.JUNTA_ACLARACIONES_PROSPECIONES (JAP_FECHA_REGISTRO, JAP_USU_ID) VALUES (GETDATE(), ?)", None),
            "SOSTENIMIENTOS_CERTIFICADOS_PROSPECIONES": ("INSERT INTO IOP.SOSTENIMIENTOS_CERTIFICADOS_PROSPECIONES (SCR_FECHA_REGISTRO, SCR_USU_ID) VALUES (GETDATE(), ?)", None)
        }
        sql, fijo = campos[tabla]
        if fijo is not None:
            self.cursor.execute(sql, fijo, 3)
        else:
            self.cursor.execute(sql, 3)
        return self.cursor.execute("SELECT SCOPE_IDENTITY()").fetchval()

if __name__ == "__main__":
    proc = OperacionProcessor(r"C:\CargasIO\carga.xlsx")
    proc.procesar_filas()
