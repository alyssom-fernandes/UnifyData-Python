@echo off
title UnifyData - Inicializador
color 0B
echo ========================================================
echo        INICIANDO O SISTEMA UNIFYDATA
echo        Por favor, aguarde alguns segundos...
echo ========================================================
echo.

:: Verifica se as bibliotecas estao instaladas e instala se necessario
echo [1/2] Verificando e instalando dependencias (se necessario)...
pip install -r requirements.txt --quiet
echo.

:: Executa a aplicacao Streamlit
echo [2/2] Abrindo a aplicacao no seu navegador...
python -m streamlit run unifydata.py

:: Pausa apenas se der algum erro
pause
