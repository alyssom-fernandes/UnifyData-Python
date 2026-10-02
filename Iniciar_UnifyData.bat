@echo off
chcp 65001 >nul
title UnifyData
color 0B
cd /d "%~dp0"
echo ========================================================
echo        INICIANDO O UNIFYDATA
echo        Aguarde alguns segundos...
echo ========================================================
echo.

:: Procura o Python (comando "python" ou o inicializador "py" do Windows)
set "PY=python"
python --version >nul 2>nul || set "PY=py"
%PY% --version >nul 2>nul
if errorlevel 1 (
    echo O Python não foi encontrado. Instale o Python 3.10 ou mais novo em python.org
    echo e marque a opção "Add Python to PATH" durante a instalação.
    pause
    exit /b 1
)

:: Evita a pergunta de e-mail do Streamlit na primeira execução
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    mkdir "%USERPROFILE%\.streamlit" 2>nul
    >"%USERPROFILE%\.streamlit\credentials.toml" echo [general]
    >>"%USERPROFILE%\.streamlit\credentials.toml" echo email = ""
)

echo [1/2] Verificando as dependências (só demora na primeira vez)...
%PY% -m pip install -r requirements.txt --quiet --disable-pip-version-check
if errorlevel 1 (
    echo.
    echo Não foi possível instalar as dependências. Confira a conexão com a internet e tente de novo.
    pause
    exit /b 1
)
echo.

echo [2/2] Abrindo o UnifyData no navegador...
%PY% -m streamlit run unifydata.py
if errorlevel 1 pause
