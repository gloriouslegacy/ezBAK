# 파일명: git_push.ps1
# 실행 예시: .\git_push.ps1 "기능 추가: 로그인 페이지 구현"

param(
    [Parameter(Mandatory=$true)]
    [string]$commit_message
)

# 1. 모든 변경 사항 스테이징
Write-Host "✅ git add . 실행 중..."
git add .

# 2. 커밋 실행 (오류 발생 시 push 중단)
Write-Host "✅ git commit -m ""$commit_message"" 실행 중..."
try {
    # -ErrorAction Stop: 오류 발생 시 즉시 catch 블록으로 이동
    # Note: 커밋할 변경 사항이 없으면 Git이 오류(Exit Code 1)를 반환합니다.
    git commit -m "$commit_message" -ErrorAction Stop
    
}
catch {
    # Git 커밋 실패 시 실행되는 블록
    Write-Host "-----------------------------------"
    # $_.Exception.Message는 Git이 반환한 상세 오류 메시지입니다.
    # 예: 'Your branch is up to date with 'origin/main'.' 또는 'nothing to commit, working tree clean'
    Write-Host "❌ Git 커밋 실패: 변경 사항이 없거나 다른 문제가 발생했습니다." -ForegroundColor Red
    Write-Host "상세: $($_.Exception.Message.Trim())" -ForegroundColor Yellow
    Write-Host "-----------------------------------"
    
    # Git 커밋 실패 후 스크립트 종료 (push 방지)
    exit 1
}

# 3. 푸시 실행 (커밋이 성공했을 때만 실행됨)
Write-Host "✅ git push origin main 실행 중..."
git push origin main

# 4. 결과 출력
Write-Host "-----------------------------------"
Write-Host "🎉 Git 자동화 완료!" -ForegroundColor Green
Write-Host "커밋 메시지: $commit_message"
Write-Host "-----------------------------------"