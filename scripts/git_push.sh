#!/bin/bash

# chmod +x scripts/git_push.sh
# 실행 예시(레포 루트에서): ./scripts/git_push.sh "기능 추가: 로그인 페이지 구현"
commit_message="$1"

# 커밋 메시지가 비어있는지 확인
if [ -z "$commit_message" ]; then
  echo "커밋 메시지를 입력해주세요."
  echo "사용법: ./scripts/git_push.sh \"커밋 메시지\""
  exit 1
fi

# 스크립트 위치(scripts/)에서 레포 루트를 계산해 그 위치에서 실행한다 (실행 폴더와 무관)
repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)" || exit 1
cd "$repo_root" || exit 1

git add .
git commit -m "$commit_message"
git push origin main

echo "-----------------------------------"
echo "Git 자동화 완료!"
echo "커밋 메시지: $commit_message"

echo "-----------------------------------"
