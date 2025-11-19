#!/usr/bin/env bash
set -e

mkdir -p $HOME/.local/bin/chrome
mkdir -p $HOME/.local/bin/chromedriver

curl -Lo chrome.zip https://storage.googleapis.com/chrome-for-testing-public/130.0.6723.0/linux64/chrome-linux64.zip
unzip chrome.zip -d $HOME/.local/bin/chrome

curl -Lo chromedriver.zip https://storage.googleapis.com/chrome-for-testing-public/130.0.6723.0/linux64/chromedriver-linux64.zip
unzip chromedriver.zip -d $HOME/.local/bin/chromedriver

chmod +x $HOME/.local/bin/chrome/chrome-linux64/chrome
chmod +x $HOME/.local/bin/chromedriver/chromedriver-linux64/chromedriver

export PATH="$HOME/.local/bin:$PATH"

ln -sf /opt/render/project/src/.venv/bin/gunicorn $HOME/.local/bin/gunicorn

pip install -r requirements.txt
