@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

title STEP DASHBOARD - PUBLICADOR AUTOMATICO PARA GITHUB

set "PROJECT_DIR=C:\Users\douglas.tabella\Downloads\Step Dashboard completo com login e app"
set "REPO_URL=https://github.com/stepmanus-boop/Step-dashboard-completo.git"
set "BRANCH=main"

echo ======================================================
echo   STEP DASHBOARD - PUBLICADOR AUTOMATICO PARA GITHUB
echo ======================================================
echo.
echo Pasta do projeto: %PROJECT_DIR%
echo Repositorio: %REPO_URL%
echo Branch: %BRANCH%
echo.

if not exist "%PROJECT_DIR%" (
    echo Pasta do projeto nao encontrada:
    echo %PROJECT_DIR%
    pause
    exit /b 1
)

cd /d "%PROJECT_DIR%" || (
    echo Nao foi possivel acessar a pasta do projeto.
    pause
    exit /b 1
)

echo [1/10] Verificando Git...
git --version >nul 2>&1
if errorlevel 1 (
    echo Git nao encontrado. Instale o Git for Windows e tente novamente.
    pause
    exit /b 1
)

echo [2/10] Inicializando repositorio se necessario...
if not exist ".git" (
    git init || goto :git_error
)

echo [3/10] Ajustando branch principal...
git branch -M %BRANCH% || goto :git_error

echo [4/10] Configurando remote origin...
git remote get-url origin >nul 2>&1
if errorlevel 1 (
    git remote add origin "%REPO_URL%" || goto :git_error
) else (
    git remote set-url origin "%REPO_URL%" || goto :git_error
)

echo [5/10] Buscando atualizacoes remotas...
git fetch origin %BRANCH% || goto :git_error

echo [6/10] Fazendo stage dos arquivos do projeto...
git add -A || goto :git_error

echo [7/10] Removendo publicadores da stage...
git reset HEAD -- "subir_step_dashboard_completo_auto.bat" >nul 2>&1
git reset HEAD -- "subir_github_step_dashboard_completo.bat" >nul 2>&1
git reset HEAD -- "subir_github_step_dashboard_completo.exe" >nul 2>&1
git reset HEAD -- "subir_github_step_dashboard_completo.cmd" >nul 2>&1

for /f %%i in ('git status --porcelain ^| find /c /v ""') do set CHANGES=%%i

echo [8/10] Criando commit se houver alteracoes locais...
if "%CHANGES%"=="0" (
    echo Nenhuma alteracao local nova para commit.
) else (
    git commit -m "chore: atualiza dashboard step" || goto :git_error
)

echo [9/10] Atualizando com o remoto sem sobrescrever...
git pull --rebase origin %BRANCH%
if errorlevel 1 (
    echo.
    echo ======================================================
    echo   CONFLITO NO REBASE
    echo ======================================================
    echo O repositorio remoto tem alteracoes que conflitaram com as locais.
    echo Resolva os conflitos, execute:
    echo   git add .
    echo   git rebase --continue
    echo Depois rode este arquivo novamente.
    echo.
    pause
    exit /b 1
)

echo [10/10] Enviando para o GitHub...
git push -u origin %BRANCH% || goto :git_error

echo.
echo ======================================================
echo   ENVIO CONCLUIDO COM SUCESSO
echo ======================================================
echo.
pause
exit /b 0

:git_error
echo.
echo ======================================================
echo   ERRO AO EXECUTAR COMANDO GIT
echo ======================================================
echo.
echo Se pedir autenticacao, entre com sua conta/token do GitHub.
echo.
pause
exit /b 1
