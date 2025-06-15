#!/bin/bash

# Start all servers concurrently
npx concurrently \
  "cd frontend && npm run dev" \
  "cd frontend/server && node wopi-server.js" \
  "cd backend && python3 api_server.py" \
  "cd frontend && ngrok http 3001"
