@echo off
echo Démarrage du serveur pour l'add-in Excel...
cd /d "%~dp0"
call npm install
echo.
echo Serveur en cours de démarrage sur le port 5890...
echo NE PAS FERMER CETTE FENETRE
echo.
node server.js
pause 