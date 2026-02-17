#!/bin/bash
echo "=== Arrêt des containers ==="
docker stop $(docker ps -q) 2>/dev/null
docker rm $(docker ps -aq) 2>/dev/null

echo "=== Arrêt ngrok ==="
killall ngrok 2>/dev/null

echo "=== Mise à jour du code ==="
cd ~/pptx-service
git pull 2>/dev/null || echo "Pas un repo git, skip pull"

echo "=== Build Docker ==="
docker build -t pptx-service .

echo "=== Lancement du service ==="
cd ~
bash ~/start.sh

echo "=== Test health ==="
sleep 3
curl -s http://localhost:8000/health
echo ""

echo "=== Lancement ngrok ==="
ngrok http 8000
