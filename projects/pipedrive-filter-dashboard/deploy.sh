#!/bin/bash
set -e

DEPLOY_ID="AKfycbwiVJqCd_KJdigHiyhW974MPjSgaFdOWSlahPkIXhgQOs06vwGJHhT-WZ9jO9FIbJZ7"
DESC="${1:-업데이트}"

cd "$(dirname "$0")"

echo "📤 코드 푸시 중..."
clasp push --force

echo "🚀 배포 중: $DESC"
clasp deploy --deploymentId "$DEPLOY_ID" --description "$DESC"

echo "✅ 완료"
echo "🔗 https://script.google.com/macros/s/$DEPLOY_ID/exec"
