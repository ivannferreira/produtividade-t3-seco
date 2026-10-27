@echo off
setlocal
set "GIT=C:\Users\ivferreira\AppData\Local\Programs\Git\cmd\git.exe"
set "GIT_TERMINAL_PROMPT=0"
set "GCM_INTERACTIVE=never"

cd /d "C:\Users\ivferreira\Documents\GitHub\produtividade-t3-seco" || exit /b 1

"%GIT%" branch --show-current >".git\branch-publicacao.tmp"
if errorlevel 1 exit /b 1
set /p "BRANCH="<".git\branch-publicacao.tmp"
if not "%BRANCH%"=="main" exit /b 1

call :publicar >>".git\publicacao.log" 2>&1
set "RESULTADO=%ERRORLEVEL%"
rmdir ".git\publicacao-em-andamento"
exit /b %RESULTADO%

:publicar
echo Publicacao: %date% %time%
"%GIT%" diff --quiet HEAD -- produtividade_t3.js
if errorlevel 2 exit /b 1
if not errorlevel 1 goto enviar

"%GIT%" add -- produtividade_t3.js
if errorlevel 1 exit /b 1
"%GIT%" commit --only -m "Atualizar produtividade" -- produtividade_t3.js
if errorlevel 1 exit /b 1

:enviar
"%GIT%" push origin main
exit /b %ERRORLEVEL%