@echo off
title Preparar pasta de envio - Arquivo Emails
cd /d "%~dp0"

if not exist dist\ArquivoEmails.exe (
  echo ERRO: dist\ArquivoEmails.exe nao existe.
  echo Corra primeiro o criar_exe.bat para compilar a versao actual.
  pause
  exit /b 1
)

echo A preparar a pasta de envio...
if exist prog\ArquivoEmails rmdir /s /q prog\ArquivoEmails
mkdir prog\ArquivoEmails
copy /y dist\ArquivoEmails.exe prog\ArquivoEmails\ >nul
copy /y prog\LEIA-ME.txt prog\ArquivoEmails\ >nul

echo A criar o zip...
if exist prog\ArquivoEmails.zip del prog\ArquivoEmails.zip
powershell -NoProfile -Command "Compress-Archive -Path 'prog\ArquivoEmails\*' -DestinationPath 'prog\ArquivoEmails.zip'"

echo.
if exist prog\ArquivoEmails.zip (
  echo ============================================================
  echo   CONCLUIDO:
  echo     prog\ArquivoEmails\        pasta pronta a copiar/partilhar
  echo     prog\ArquivoEmails.zip     zip pronto a enviar por e-mail
  echo   Conteudo: ArquivoEmails.exe + LEIA-ME.txt
  echo ============================================================
) else (
  echo ERRO ao criar o zip — pode enviar a pasta prog\ArquivoEmails.
)
pause
