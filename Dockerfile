FROM python:3.12-slim

# Establecer el directorio de trabajo
WORKDIR /app

# Copiar archivos del proyecto
COPY . .

# Instalar herramientas necesarias y drivers ODBC
RUN apt-get update && \
    apt-get install -y gcc curl gnupg unixodbc-dev && \
    curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add - && \
    curl https://packages.microsoft.com/config/debian/10/prod.list > /etc/apt/sources.list.d/mssql-release.list && \
    apt-get update && \
    ACCEPT_EULA=Y apt-get install -y msodbcsql17

# Instalar pipenv
RUN pip install pipenv

# Instalar dependencias desde Pipfile.lock
RUN pipenv install --system --deploy

# Ejecutar tu script principal (ajusta si el nombre es diferente)
CMD ["python", "insertar_operaciones.py"]
