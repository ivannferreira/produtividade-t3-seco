@echo off
setlocal
set "GIT=C:\Users\ivferreira\AppData\Local\Programs\Git\cmd\git.exe"
set "REPO=C:\Users\ivferreira\Documents\GitHub\produtividade-t3-seco"
set "GIT_TERMINAL_PROMPT=0"
set "GCM_INTERACTIVE=never"

cd /d "%REPO%" || exit /b 1

"%GIT%" branch --show-current >".git\branch-publicacao.tmp"
if errorlevel 1 exit /b 1
set /p "BRANCH="<".git\branch-publicacao.tmp"
if not "%BRANCH%"=="main" exit /b 1

echo Gerando dados do painel individual...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%REPO%\gerar_dados_individual.ps1"
if errorlevel 1 (
    echo ERRO ao gerar dados_produtividade_individual.json
    exit /b 1
)

call :publicar >>".git\publicacao.log" 2>&1
set "RESULTADO=%ERRORLEVEL%"
exit /b %RESULTADO%

:publicar
echo Publicacao: %date% %time%

REM Mantem compatibilidade com o painel antigo, caso produtividade_t3.js seja atualizado por outra rotina.
"%GIT%" add -- dados_produtividade_individual.json
if exist produtividade_t3.js "%GIT%" add -- produtividade_t3.js

"%GIT%" diff --cached --quiet
if not errorlevel 1 goto enviar

"%GIT%" commit -m "Atualizar produtividade individual"
if errorlevel 1 exit /b 1

:enviar
"%GIT%" push origin main
exit /b %ERRORLEVEL%
