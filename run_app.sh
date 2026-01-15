#!/bin/bash
set -e

APP_DIR="/Users/augustasalisauskas/.gemini/antigravity/scratch/excel_transfer_app"
PORT=8501
URL="http://localhost:${PORT}"

cd "$APP_DIR"
source "$APP_DIR/venv/bin/activate"

# paleidžiam streamlit fone
streamlit run app.py --server.port $PORT --server.headless true --browser.gatherUsageStats false >/tmp/excel_transfer_app.log 2>&1 &
STREAMLIT_PID=$!

# palaukiam kol serveris pakils
for i in {1..50}; do
  if nc -z localhost $PORT 2>/dev/null; then
    break
  fi
  sleep 0.2
done

# atidarom atskirą naršyklės langą (ne tabą)
open -n "$URL"

# LAUKIAM kol naršyklės langas bus uždarytas
# (tikrinam ar dar yra aktyvi jungtis)
while lsof -i :$PORT >/dev/null 2>&1; do
  sleep 1
done

# kai langas uždarytas – stabdom streamlit
kill $STREAMLIT_PID 2>/dev/null || true