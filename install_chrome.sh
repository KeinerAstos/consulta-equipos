#!/usr/bin/env bash
set -e

echo "=== INSTALANDO GOOGLE CHROME ==="

# Instalar dependencias mínimas
apt-get update
apt-get install -y wget gnupg unzip

# Descargar e instalar Google Chrome estable
wget -q https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
apt-get install -y ./google-chrome-stable_current_amd64.deb

echo "=== INSTALANDO CHROMEDRIVER ==="

CHROME_VERSION=$(google-chrome --version | awk '{print $3}' | cut -d. -f1)
echo "Versión de Chrome detectada: $CHROME_VERSION"

# Descargar chromedriver EXACTO para tu versión de Chrome
wget -q "https://chromedriver.storage.googleapis.com/$CHROME_VERSION.0.0/chromedriver_linux64.zip" -O chromedriver.zip
unzip -o chromedriver.zip
chmod +x chromedriver
mv chromedriver /usr/local/bin/chromedriver

echo "=== VERIFICANDO ==="
ls -la /usr/bin/google-chrome
ls -la /usr/local/bin/chromedriver

echo "=== INSTALACIÓN COMPLETA ==="