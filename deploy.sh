#!/bin/bash
# Build (or rebuild) the image and (re)start the container on UnRAID.
# Run from the repo folder:  bash deploy.sh
set -euo pipefail

NAME="inventory-dashboard"
HOST_PORT="${HOST_PORT:-8095}"
DATA_DIR="${DATA_DIR:-/mnt/user/Nextcloud/richamc01/files/Documents/Personal Assets/Home Network/Inventory}"
RUN_AS="${RUN_AS:-99:100}"

cd "$(dirname "$0")"

if [ ! -r "$DATA_DIR/Hardware Inventory.xlsx" ]; then
  echo "ERROR: cannot read '$DATA_DIR/Hardware Inventory.xlsx'" >&2
  exit 1
fi

echo "== Building image"
docker build -t "$NAME:latest" .

echo "== Checking the container user ($RUN_AS) can read the workbook"
docker run --rm --user "$RUN_AS" -v "$DATA_DIR:/data:ro" "$NAME:latest" \
  python -c "open('/data/Hardware Inventory.xlsx','rb').read(4); print('workbook readable')"

echo "== Replacing container"
docker rm -f "$NAME" >/dev/null 2>&1 || true
docker run -d \
  --name "$NAME" \
  --restart unless-stopped \
  --user "$RUN_AS" \
  -p "$HOST_PORT:8000" \
  -v "$DATA_DIR:/data:ro" \
  "$NAME:latest"

echo "== Waiting for it to answer"
for _ in $(seq 1 20); do
  if docker exec "$NAME" python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/api/status', timeout=3)" 2>/dev/null; then
    docker exec "$NAME" python -c "import json,urllib.request; s=json.load(urllib.request.urlopen('http://127.0.0.1:8000/api/status')); print('status:', s)"
    echo "== Done. Open http://<this-server-ip>:$HOST_PORT"
    exit 0
  fi
  sleep 1
done
echo "ERROR: container did not respond. Logs:" >&2
docker logs "$NAME" >&2
exit 1
