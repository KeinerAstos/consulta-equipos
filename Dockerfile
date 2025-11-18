# 1. Imagen base de Python. Usamos una versión 'slim' para ser ligeros.
FROM python:3.11-slim

# 2. Instalar el navegador Chromium y dependencias del sistema.
# Esto se ejecuta con permisos de root DENTRO del proceso de construcción de Docker.
RUN apt-get update && \
    apt-get install -y chromium ffmpeg libsm6 libxext6 && \
    rm -rf /var/lib/apt/lists/*

# 3. Establecer el directorio de trabajo dentro del contenedor
WORKDIR /app

# 4. Copiar los archivos de requerimientos e instalar dependencias de Python
# Esto asegura que tus librerías (Selenium, Flask, etc.) estén instaladas.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 5. Copiar el resto del código de tu proyecto
COPY . .

# 6. Definir el comando de inicio para la aplicación web (Flask con Gunicorn)
CMD gunicorn index:app --bind 0.0.0.0:$PORT