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

echo "✅ Both servers are running:"
echo "   Backend  → https://tool-2-3w1t.onrender.com/"
echo "   Frontend → https://tool-3-vctq.onrender.com
echo ""
echo "Open https://tool-3-vctq.onrender.com in your browser."
echo "Press Ctrl+C to stop both servers."

wait $BACKEND_PID $FRONTEND_PID
