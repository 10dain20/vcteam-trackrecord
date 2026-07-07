#!/bin/bash
# VC투자팀 TR - 원클릭 실행기
# 더블클릭하면 백엔드(5001) + 프론트(8743)를 켜고 브라우저를 엽니다.

cd "$(dirname "$0")" || exit 1
echo "▶ 폴더: $(pwd)"

# 종료 시 두 서버 정리
cleanup() {
  echo "▶ 서버를 종료합니다..."
  kill "$BACK" "$FRONT" 2>/dev/null
  exit 0
}
trap cleanup INT TERM

# 1) 필요 라이브러리 설치 (최초 1회, 이미 있으면 빠르게 넘어감)
echo "▶ 라이브러리 확인/설치 중..."
python3 -m pip install -r requirements.txt >/tmp/tr_pip.log 2>&1

# 2) 백엔드 실행 (포트 5001)
echo "▶ 백엔드 시작 (http://localhost:5001) ..."
python3 app.py >/tmp/tr_backend.log 2>&1 &
BACK=$!

# 3) 프론트 실행 (포트 8743)
echo "▶ 프론트 시작 (http://localhost:8743) ..."
python3 -m http.server 8743 >/tmp/tr_frontend.log 2>&1 &
FRONT=$!

# 4) 백엔드가 응답할 때까지 대기 후 브라우저 열기
echo "▶ 서버 준비 대기 중..."
for i in $(seq 1 20); do
  if curl -s http://localhost:5001/api/health >/dev/null 2>&1; then
    echo "✔ 백엔드 정상 (api/health OK)"
    break
  fi
  sleep 0.5
done

open "http://localhost:8743"
echo ""
echo "=================================================="
echo " 열린 주소:  http://localhost:8743"
echo " 비밀번호 :  Maxibest11^^"
echo "--------------------------------------------------"
echo " 로그:  /tmp/tr_backend.log  /tmp/tr_frontend.log"
echo " 이 창을 닫으면 서버가 종료됩니다. (Ctrl+C)"
echo "=================================================="

# 창을 열어둔 채 서버 유지
wait
