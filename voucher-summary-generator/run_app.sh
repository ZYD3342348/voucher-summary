#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

PORT="${PORT:-8501}"
ADDR="${ADDR:-127.0.0.1}"

echo "启动总台收入工作台..."
echo "界面主题：中式美学（custom_styles.css）"
echo "URL: http://${ADDR}:${PORT}"

python3 -m streamlit run app.py \
  --server.port "${PORT}" \
  --server.address "${ADDR}" \
  --server.headless true \
  --browser.gatherUsageStats false
