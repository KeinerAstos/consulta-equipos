#!/usr/bin/env bash
set -e

echo "==== INSTALANDO CHROME ===="

CHROME_DIR="/opt/render/project/src/.chrome"
mkdir -p $CHROME_DIR
cd $CHROME_DIR

# Descargar Chrome (última versión estable que coincide con Chromedriver)
curl -L -o chrome.zip https://storage.googleapis.com/chrome-for-testing-public/131.0.6778.85/linux64/chrome-linux64.zip
unzip -o chrome.zip

# Descargar Chromedriver
curl -L -o chromedriver.zip https://storage.googleapis.com/chrome-for-testing-public/131.0.6778.85/linux64/chromedriver-linux64.zip
unzip -o chromedriver.zip

# Permisos de ejecución
chmod +x chrome-linux64/chrome
chmod +x chromedriver-linux64/chromedriver

echo "Chrome y ChromeDriver instalados en:"
echo "  - $CHROME_DIR/chrome-linux64/chrome"
echo "  - $CHROME_DIR/chromedriver-linux64/chromedriver"

echo "==== INSTALANDO DEPENDENCIAS PYTHON ===="
pip install --upgrade pip
pip install -r requirements.txt

echo "==== INSTALACIÓN COMPLETA ===="
