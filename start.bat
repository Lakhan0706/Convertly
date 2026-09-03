@echo off
setlocal
where node >nul 2>nul || (echo Node.js is required. Install Node.js 18+ and try again.&pause&exit /b 1)
if not exist node_modules (echo Installing dependencies...&call npm install)
echo Starting Convertly on http://localhost:8080
call npm start
pause
