@echo off
title Arquivo Emails - Prospectiva
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo ERRO: Python nao encontrado. Instale a partir de https://www.python.org/downloads/
  echo e marque a opcao "Add python.exe to PATH" durante a instalacao.
  pause
  exit /b 1
)
python -c "import flask, win32com" >nul 2>nul
if errorlevel 1 (
  echo A instalar dependencias Python pela primeira vez...
  python -m pip install -r requirements.txt
)
python app.py
pause
