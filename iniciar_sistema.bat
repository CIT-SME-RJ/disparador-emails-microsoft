@echo off
:: ========================================================
:: INICIALIZADOR DO DISPARADOR DE E-MAILS
:: ========================================================

cd /d "%~dp0"

:: ========================================================
:: O BLOCO ABAIXO IMPEDE A TELA DE FECHAR SOZINHA
:: ========================================================
if "%~1"=="--rodando" goto INICIO
cmd /k ""%~f0" --rodando"
exit /b

:INICIO
title Inicializador do Disparador SME
color 0A

set "PASTA_SISTEMA=sistema"
set "PASTA_VENV=%PASTA_SISTEMA%\venv"
set "ARQUIVO_LOG=%PASTA_SISTEMA%\instalacao_pacotes.log"

echo ==========================================
echo    INICIANDO DISPARADOR DE E-MAILS
echo ==========================================
echo.
echo IMPORTANTE:
echo Enquanto esta janela estiver aberta, o sistema fica ativo.
echo Se fechar esta janela, o aplicativo tambem sera encerrado.
echo.

echo [Passo 1 de 5] Verificando arquivos principais...

if not exist "%PASTA_SISTEMA%\app_streamlit.py" goto ERRO_APP
if not exist "%PASTA_SISTEMA%\requirements.txt" goto ERRO_REQUIREMENTS

echo Arquivos principais encontrados.
echo.

echo [Passo 2 de 5] Verificando Python...

python --version >nul 2>&1
if errorlevel 1 goto ERRO_PYTHON

echo Python detectado.
echo.

echo [Passo 3 de 5] Verificando ambiente virtual...

if exist "%PASTA_VENV%\Scripts\activate.bat" goto AMBIENTE_PRONTO

echo Ambiente virtual nao encontrado.
echo Criando ambiente virtual local...
python -m venv "%PASTA_VENV%" > "%ARQUIVO_LOG%" 2>&1

if errorlevel 1 goto ERRO_VENV

:AMBIENTE_PRONTO
echo Ambiente virtual pronto.
echo.

echo [Passo 4 de 5] Instalando ou verificando pacotes...
echo Aguarde. Os detalhes tecnicos estao sendo salvos em:
echo %ARQUIVO_LOG%
echo.

call "%PASTA_VENV%\Scripts\activate.bat"

python -m pip install --upgrade pip >> "%ARQUIVO_LOG%" 2>&1
if errorlevel 1 goto ERRO_PACOTES

pip install -r "%PASTA_SISTEMA%\requirements.txt" >> "%ARQUIVO_LOG%" 2>&1
if errorlevel 1 goto ERRO_PACOTES

echo Pacotes verificados.
echo.

echo [Passo 5 de 5] Iniciando a interface visual...

echo Abrindo o Outlook Classic...
start "" outlook.exe

echo.
echo Abrindo o aplicativo no navegador...
echo.
echo ATENCAO:
echo Nao feche esta janela enquanto estiver usando o sistema.
echo Para encerrar o sistema, feche esta janela ou pressione CTRL+C.
echo.

streamlit run "%PASTA_SISTEMA%\app_streamlit.py"

goto FIM

:ERRO_PYTHON
color 0C
echo.
echo ERRO: Python nao encontrado.
echo.
echo Procure o suporte tecnico ou a pessoa responsavel pelo sistema.
echo O Python precisa estar instalado no computador e disponivel no PATH.
echo Ao instalar, marque a opcao: Add Python to PATH.
echo.
goto FIM

:ERRO_APP
color 0C
echo.
echo ERRO: O arquivo sistema\app_streamlit.py nao foi encontrado.
echo.
echo Verifique se a pasta sistema existe e se os arquivos do sistema estao dentro dela.
echo.
goto FIM

:ERRO_REQUIREMENTS
color 0C
echo.
echo ERRO: O arquivo sistema\requirements.txt nao foi encontrado.
echo.
echo Esse arquivo e necessario para instalar os pacotes do sistema.
echo.
goto FIM

:ERRO_VENV
color 0C
echo.
echo ERRO: Nao foi possivel criar o ambiente virtual.
echo.
echo Verifique o arquivo de log:
echo %ARQUIVO_LOG%
echo.
echo Se o problema continuar, procure o suporte tecnico.
echo.
goto FIM

:ERRO_PACOTES
color 0C
echo.
echo ERRO: Nao foi possivel instalar ou atualizar os pacotes.
echo.
echo Verifique a conexao com a internet e tente novamente.
echo.
echo Detalhes tecnicos foram salvos em:
echo %ARQUIVO_LOG%
echo.
echo Se o problema continuar, procure o suporte tecnico.
echo.
goto FIM

:FIM
echo.
echo ==========================================
echo O sistema foi encerrado ou interrompido.
echo Agora voce pode fechar esta janela.
echo ==========================================
echo.