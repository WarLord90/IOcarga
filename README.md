# 🐍 Proyecto de Carga Automatizada con Python + Docker

Este proyecto automatiza la carga de datos desde un archivo Excel a una base de datos SQL Server.  
Está desarrollado en **Python**, empacado en **Docker**, y se ejecuta fácilmente con un solo clic desde un archivo `.bat`.

---

## 📦 Contenido del proyecto

📁 Ejemplo/
├── cargaio.tar ← Imagen Docker exportada (opcional, no incluida en Git)
├── insertar_operaciones.py ← Script principal (dentro de la imagen)
├── Dockerfile ← Para construir la imagen Docker
├── Pipfile / Pipfile.lock ← Dependencias de Python
├── .dockerignore ← Archivos que no deben ir en la imagen
├── .env.example ← Plantilla del archivo de conexión
├── carga.xlsx ← Archivo de entrada (no subir a Git)
├── output/ ← Carpeta de salida para logs
├── run.bat ← Ejecuta todo con un doble clic


---

## 🚀 Cómo usar el proyecto

### 1. 🧱 Requisitos

- Tener instalado [Docker Desktop](https://www.docker.com/products/docker-desktop)
- Tener un archivo `.env` con tus credenciales

### 2. ✏️ Preparar la carpeta de ejecución

Debes tener los siguientes archivos en la misma carpeta:

- `.env` (lo puedes copiar desde `.env.example`)
- `carga.xlsx` (tu archivo de datos)
- Carpeta vacía llamada `output\` (donde se generará el log)

### 3. 🖱 Ejecutar

Solo da **doble clic en `run.bat`**  
Y listo: se conecta, procesa, y genera el log en la carpeta `output`.

---

## ⚙️ ¿Qué hace internamente?

- Monta tu `.env`, tu Excel y tu carpeta `output` dentro del contenedor
- Ejecuta el código `insertar_operaciones.py` (dentro de la imagen)
- Escribe un archivo de log como: `output/log_YYYY-MM-DD.txt`

---

## 📂 `.env.example` – Estructura esperada

```env
DB_SERVER=urlhost
DB_NAME=dbname
DB_USER=mi_usuario
DB_PASS=mi_contraseña_segura

🐳 ¿Cómo construir la imagen desde cero?
Si prefieres construir la imagen tú mismo en lugar de usar cargaio.tar:

docker build -t cargaio-py312 .

Y luego la ejecutas con:
docker run --rm `
  -v "${PWD}\.env:/app/.env" `
  -v "${PWD}\carga.xlsx:/app/carga.xlsx" `
  -v "${PWD}\output:/app/output" `
  cargaio-py312

🔒 Seguridad
El archivo .env nunca se incluye en la imagen
El código Python está empaquetado dentro de Docker
Solo se exponen los datos necesarios

✅ Créditos
Este desarrollo fue realizado por Javier Jaimes como su primer proyecto profesional en Python 💪
Con pandas, pyodbc, Docker y mucha persistencia.

---

✅ ¡Ahora sí! Solo pega esto en tu `README.md`, sube al repo, y está **listo para entregar, clonar o documentar tu éxito**. ¿Te preparo ahora un `run.bat` final que haga incluso el `docker load` si el `.tar` está en la carpeta?
