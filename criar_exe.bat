@echo off
title Criar ArquivoEmails.exe - Prospectiva
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo ERRO: Python nao encontrado. Instale a partir de https://www.python.org/downloads/
  echo e marque a opcao "Add python.exe to PATH" durante a instalacao.
  pause
  exit /b 1
)

echo ============================================================
echo   Arquivo Emails - criacao do executavel (dist\ArquivoEmails.exe)
echo ============================================================
echo.
echo [1/2] A instalar/verificar dependencias...
python -m pip install --quiet -r requirements.txt pywebview pyinstaller
if errorlevel 1 (
  echo ERRO na instalacao das dependencias. Verifique a ligacao a internet.
  pause
  exit /b 1
)

echo [2/2] A compilar (pode demorar alguns minutos)...
python -m PyInstaller --onefile --noconsole --name ArquivoEmails ^
  --icon "resources\icone.ico" ^
  --add-data "templates;templates" ^
  --add-data "resources;resources" ^
  --hidden-import win32com --hidden-import win32com.client ^
  --hidden-import pythoncom --hidden-import pywintypes ^
  launcher.py
if errorlevel 1 (
  echo.
  echo ERRO na compilacao — ver mensagens acima.
  pause
  exit /b 1
)

echo.
if exist dist\ArquivoEmails.exe (
  echo ============================================================
  echo   CONCLUIDO:  dist\ArquivoEmails.exe
  echo   Pode copiar o .exe para onde quiser; nao precisa de mais
  echo   nada para correr (nem de Python instalado).
  echo ============================================================
) else (
  echo ERRO: o executavel nao foi criado — ver mensagens acima.
)
pause
