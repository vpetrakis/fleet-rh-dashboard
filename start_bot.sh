#!/usr/bin/env bash
# ═══════════════════════════════════════════════════════════════════════════
#  Fleet Command Nexus — start_bot.sh
#  Works in GitHub Codespaces and local Linux / macOS environments.
#  Run from the project root:  bash start_bot.sh
# ═══════════════════════════════════════════════════════════════════════════
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BRAIN="$ROOT/services/brain"
DASH="$ROOT/apps/dashboard"

CYAN="\033[0;36m"; YELLOW="\033[1;33m"; GREEN="\033[0;32m"; RED="\033[0;31m"; RESET="\033[0m"
log()  { echo -e "${CYAN}[nexus]${RESET} $*"; }
warn() { echo -e "${YELLOW}[nexus]${RESET} $*"; }
ok()   { echo -e "${GREEN}[nexus]${RESET} $*"; }
err()  { echo -e "${RED}[nexus]${RESET} $*"; }

# ── 1. Free the ports ────────────────────────────────────────────────────────
log "Freeing ports 3000 and 8000 …"
for port in 3000 8000; do
  pid=$(lsof -ti tcp:"$port" 2>/dev/null || true)
  [ -n "$pid" ] && kill -9 $pid 2>/dev/null && warn "Killed PID $pid on :$port" || true
done

# ── 2. Python backend ────────────────────────────────────────────────────────
log "Checking Python dependencies …"
cd "$BRAIN"
pip install fastapi uvicorn python-multipart antiword --break-system-packages -q 2>/dev/null \
  || pip install fastapi uvicorn python-multipart -q

log "Starting Python backend on :8000 …"
python3 main.py &
PYTHON_PID=$!

# Wait until the health endpoint responds
TRIES=0
until curl -sf http://127.0.0.1:8000/api/health > /dev/null 2>&1; do
  sleep 1
  TRIES=$((TRIES+1))
  if [ $TRIES -ge 15 ]; then
    err "Python backend did not start after 15 s — check for errors above."
    kill "$PYTHON_PID" 2>/dev/null || true
    exit 1
  fi
done
ok "Backend ready  →  http://127.0.0.1:8000"

# ── 3. Next.js frontend ──────────────────────────────────────────────────────
cd "$DASH"
if [ ! -d "node_modules" ]; then
  log "Installing Node.js dependencies (first run) …"
  npm install
fi

# Write .env.local so Next.js API routes know where the backend is
cat > .env.local <<'ENV'
BACKEND_URL=http://127.0.0.1:8000
ENV

log "Starting Next.js frontend on :3000 …"
npm run dev -- --hostname 0.0.0.0 --port 3000 &
NEXT_PID=$!
cd "$ROOT"

# ── 4. GitHub Codespaces port visibility hint ────────────────────────────────
echo ""
ok "═══════════════════════════════════════════════════════════════"
ok "  Fleet Command Nexus is running!"
echo ""
ok "  Frontend  →  http://localhost:3000  (expose this port)"
ok "  Backend   →  http://localhost:8000  (keep private)"
echo ""
if [ -n "${CODESPACE_NAME:-}" ]; then
  warn "GitHub Codespaces detected."
  warn "Go to the PORTS tab and set port 3000 to Visibility: Public."
  warn "Port 8000 should remain Private (it's proxied via Next.js)."
fi
ok "═══════════════════════════════════════════════════════════════"
echo ""
log "Press Ctrl+C to stop both servers."

# ── 5. Clean shutdown ────────────────────────────────────────────────────────
cleanup() {
  log "Shutting down …"
  kill "$PYTHON_PID" "$NEXT_PID" 2>/dev/null || true
  ok "Stopped."
}
trap cleanup EXIT INT TERM
wait "$PYTHON_PID" "$NEXT_PID"