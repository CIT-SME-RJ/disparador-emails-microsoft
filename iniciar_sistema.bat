@echo off
:: ========================================================
:: O BLOCO ABAIXO IMPEDE A TELA DE FECHAR SOZINHA (ESCUDO)
if "%~1"=="--rodando" goto INICIO
cmd /k ""%~f0" --rodando"
exit /b
:: ========================================================

:INICIO
title Inicializador do Disparador SME
color 0A

echo ==========================================
echo    INICIANDO DISPARADOR DE E-MAILS
echo ==========================================
echo.

echo [Passo 1 de 4] Verificando Python...
python --version >nul 2>&1
if errorlevel 1 goto ERRO_PYTHON
echo Python detectado!
echo.

echo [Passo 2 de 4] Verificando ambiente virtual (venv)...
if exist "venv\Scripts\activate.bat" goto AMBIENTE_PRONTO

echo Criando as pastas do sistema (isso pode demorar uns segundos)...
python -m venv venv

:AMBIENTE_PRONTO
echo Ambiente pronto!
echo.

echo [Passo 3 de 4] Instalando pacotes (Aguarde, baixa da internet)...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
echo.

echo [Passo 4 de 4] Iniciando a interface visual...
echo Abrindo o Outlook (Classic)...
start "" outlook.exe

if not exist "app_streamlit.py" goto ERRO_APP

echo Abrindo o aplicativo no navegador...
streamlit run app_streamlit.py
goto FIM

:ERRO_PYTHON
color 0C
echo.
echo ERRO CRITICO: Python nao encontrado!
# trocar frase para buscar atendimento
echo Por favor, instale o Python e marque a caixa "Add Python to PATH". 
goto FIM

:ERRO_APP
color 0C
echo.
echo ERRO CRITICO: O arquivo app_streamlit.py nao foi encontrado na pasta!
goto FIM

:FIM
echo.
echo ==========================================
echo O processo terminou ou foi interrompido.
echo Pode fechar esta tela preta no "X" lá em cima.
echo ==========================================