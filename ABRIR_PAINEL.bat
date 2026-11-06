@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Abrindo o painel no navegador... (deixe esta janela aberta)
python painel.py
if errorlevel 9009 echo Python nao encontrado. Instale em python.org marcando "Add python.exe to PATH".
pause
