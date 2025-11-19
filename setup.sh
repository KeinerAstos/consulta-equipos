#!/usr/bin/env bash
set -e

echo "==== INSTALANDO CHROME ===="

mkdir -p $HOME/.local/bin/chrome
mkdir -p $HOME/.local/bin/chromedriver

# Descargar Chrome
curl -L -o chrome.zip https://storage.googleapis.com/chrome-for-testing-public/130.0.6723.0/linux64/chrome-linux64.zip
unzip -o chrome.zip -d $HOME/.local/bin/chrome

# Descargar Chromedriver
curl -L -o chromedriver.zip https://storage.googleapis.com/chrome-for-testing-public/130.0.6723.0/linux64/chromedriver-linux64.zip
unzip -o chromedriver.zip -d $HOME/.local/bin/chromedriver

# Permisos
chmod +x $HOME/.local/bin/chrome/chrome-linux64/chrome
chmod +x $HOME/.local/bin/chromedriver/chromedriver-linux64/chromedriver

# Export PATH
export PATH="$HOME/.local/bin:$PATH"

echo "==== INSTALANDO DEPENDENCIAS PYTHON ===="
pip install -r requirements.txt

echo "==== INSTALACIÓN COMPLETA ===="
