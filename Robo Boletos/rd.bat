@echo off
echo =====================================
echo Iniciando o script de login...
echo =====================================

REM (Ativando o ambiente virtual, se houver - opcional)
REM call venv\Scripts\activate

REM Executa o script login.py
python login.py

echo =====================================
echo Script finalizado.
echo Pressione qualquer tecla para sair.
pause > nul