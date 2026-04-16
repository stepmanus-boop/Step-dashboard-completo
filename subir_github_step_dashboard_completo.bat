@echo off
setlocal

echo ==========================================
echo   ENVIO DO PROJETO PARA O GITHUB
echo ==========================================
echo.

cd /d "%~dp0"

git init
git branch -M main
git remote remove origin 2>nul
git remote add origin https://github.com/stepmanus-boop/Step-dashboard-completo.git

git add .
git commit -m "Atualizacao do projeto"
git push -u origin main

echo.
echo Concluido.
pause