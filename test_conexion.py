from conexion import conectar

conn = conectar()

if conn:
    print("✅ Conexión exitosa a SQL Server")
    conn.close()
else:
    print("❌ No se pudo conectar a la base de datos")
