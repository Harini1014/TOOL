@echo off
echo Starting Publishing QA Validation Tool...
echo.

echo [1/2] Starting Backend (FastAPI)...
start "QA Backend" cmd /k "cd backend && pip install -r requirements.txt && python main.py"

timeout /t 3 /nobreak >nul

echo [2/2] Starting Frontend (React)...
start "QA Frontend" cmd /k "cd frontend && npm install && npm run dev"

timeout /t 5 /nobreak >nul

echo.
echo Both servers are starting...
echo Backend  : https://qa-tool-1oh2.onrender.com
echo Frontend : http://localhost:5173
echo.
echo Open your browser at http://localhost:5173
pause
