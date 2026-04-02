#!/bin/bash
echo "🚀 INITIATING FLEET COMMAND..."

# Kill previous ghosts
fuser -k 3000/tcp > /dev/null 2>&1 || true
fuser -k 8000/tcp > /dev/null 2>&1 || true

# Start Brain in the background
cd services/brain
python3 main.py &
BRAIN_PID=$!
cd ../..

# Start Dashboard in the foreground
cd apps/dashboard
npm run dev
