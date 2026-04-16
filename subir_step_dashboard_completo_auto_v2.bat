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

echo [1/11] Verificando Git...
git --version >nul 2>&1
if errorlevel 1 (
    echo Git nao encontrado. Instale o Git for Windows e tente novamente.
    pause
    exit /b 1
)

echo [2/11] Inicializando repositorio se necessario...
if not exist ".git" (
    git init || goto :git_error
)

echo [3/11] Ajustando branch principal...
git branch -M %BRANCH% || goto :git_error

echo [4/11] Configurando remote origin...
git remote get-url origin >nul 2>&1
if errorlevel 1 (
    git remote add origin "%REPO_URL%" || goto :git_error
) else (
    git remote set-url origin "%REPO_URL%" || goto :git_error
)

echo [5/11] Buscando atualizacoes remotas...
git fetch origin %BRANCH% || goto :git_error

echo [6/11] Fazendo stage dos arquivos do projeto...
git add -A || goto :git_error

echo [7/11] Removendo publicadores da stage e restaurando localmente...
git reset HEAD -- "subir_step_dashboard_completo_auto.bat" >nul 2>&1
git reset HEAD -- "subir_step_dashboard_completo_auto_corrigido.bat" >nul 2>&1
git reset HEAD -- "subir_github_step_dashboard_completo.bat" >nul 2>&1
git reset HEAD -- "subir_github_step_dashboard_completo.exe" >nul 2>&1
git reset HEAD -- "subir_github_step_dashboard_completo.cmd" >nul 2>&1

git restore --source=HEAD -- "subir_step_dashboard_completo_auto.bat" >nul 2>&1
git restore --source=HEAD -- "subir_step_dashboard_completo_auto_corrigido.bat" >nul 2>&1
git restore --source=HEAD -- "subir_github_step_dashboard_completo.bat" >nul 2>&1
git restore --source=HEAD -- "subir_github_step_dashboard_completo.exe" >nul 2>&1
git restore --source=HEAD -- "subir_github_step_dashboard_completo.cmd" >nul 2>&1

echo [8/11] Criando commit se houver alteracoes locais...
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "chore: atualiza dashboard step" || goto :git_error
) else (
    echo Nenhuma alteracao local nova para commit.
)

echo [9/11] Atualizando com o remoto sem sobrescrever...
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

echo [10/11] Enviando para o GitHub...
git push -u origin %BRANCH% || goto :git_error

echo [11/11] Finalizado.
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
