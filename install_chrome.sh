#!/usr/bin/env bash
set -e

echo "=== INSTALANDO CHROME EN RUNTIME ==="

mkdir -p $HOME/.local/bin/chrome
mkdir -p $HOME/.local/bin/chromedriver

# Descargar Chrome
wget -q https://storage.googleapis.com/chrome-for-testing-public/131.0.6778.108/linux64/chrome-linux64.zip
unzip -o -q chrome-linux64.zip -d $HOME/.local/bin/chrome

# Descargar Chromedriver
wget -q https://storage.googleapis.com/chrome-for-testing-public/131.0.6778.108/linux64/chromedriver-linux64.zip
unzip -o -q chromedriver-linux64.zip -d $HOME/.local/bin/chromedriver

chmod +x $HOME/.local/bin/chrome/chrome-linux64/chrome
chmod +x $HOME/.local/bin/chromedriver/chromedriver-linux64/chromedriver

echo "=== Chrome instalado en runtime ==="
