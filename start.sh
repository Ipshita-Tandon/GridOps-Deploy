#!/bin/bash

# Set default ports if not provided
export FRONTEND_PORT=${FRONTEND_PORT:-3000}
export WOPI_PORT=${WOPI_PORT:-3001}
export BACKEND_PORT=${BACKEND_PORT:-5000}

# Install dependencies if needed
if [ ! -d "frontend/node_modules" ]; then
  echo "Installing frontend dependencies..."
  cd frontend && npm install && cd ..
fi

if [ ! -d "frontend/server/node_modules" ]; then
  echo "Installing WOPI server dependencies..."
  cd frontend/server && npm install && cd ../..
fi

if [ ! -d "backend/venv" ]; then
  echo "Setting up Python virtual environment..."
  cd backend && python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt && cd ..
fi

# Start all servers concurrently
npx concurrently \
  "cd frontend && npm run dev -- --port $FRONTEND_PORT" \
  "cd frontend/server && node wopi-server.js --port $WOPI_PORT" \
  "cd backend && source venv/bin/activate && python3 api_server.py --port $BACKEND_PORT" \
  "cd frontend && ngrok http $WOPI_PORT --log=stdout"
