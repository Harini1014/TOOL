#!/bin/bash
echo "Starting Publishing QA Validation Tool..."
echo ""

echo "[1/2] Starting Backend (FastAPI)..."
cd backend
pip install -r requirements.txt -q
python main.py &
BACKEND_PID=$!
cd ..

echo "Backend PID: $BACKEND_PID"
sleep 2

echo "[2/2] Starting Frontend (React)..."
cd frontend
npm install --silent
npm run dev &
FRONTEND_PID=$!
cd ..

echo ""
echo "✅ Both servers are running:"
echo "   Backend  → https://qa-tool-1oh2.onrender.com//"
echo "   Frontend → https://qa-tool-1oh2.onrender.com"
echo ""
echo "Open https://qa-tool-1oh2.onrender.com in your browser."
echo "Press Ctrl+C to stop both servers."

wait $BACKEND_PID $FRONTEND_PID
