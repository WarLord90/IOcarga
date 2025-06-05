# Imagen base con Python
FROM python:3.11-slim

# Instala dependencias del sistema necesarias para pyodbc
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    unixodbc \
    unixodbc-dev \
    freetds-dev \
    freetds-bin \
    build-essential \
    libssl-dev \
    libffi-dev \
    libpq-dev \
    curl \
    gnupg && \
    rm -rf /var/lib/apt/lists/*

# Instala ODBC Driver 17 for SQL Server
RUN curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add - && \
    curl https://packages.microsoft.com/config/debian/10/prod.list > /etc/apt/sources.list.d/mssql-release.list && \
    apt-get update && ACCEPT_EULA=Y apt-get install -y msodbcsql17

# Establece el directorio de trabajo
WORKDIR /app

# Copia los archivos del proyecto
COPY . .

# Instala paquetes de Python
RUN pip install --no-cache-dir -r requirements.txt

# Comando para ejecutar tu script
CMD ["python", "insertar_operaciones.py"]
