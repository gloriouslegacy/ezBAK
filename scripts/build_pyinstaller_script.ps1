param(
    [Parameter(Mandatory=$true)]
    [string]$ScriptName
)

# PyInstaller script
# 실행 예시(레포 루트에서): .\scripts\build_pyinstaller_script.ps1 -ScriptName "ezBAK.py"
# 어느 폴더에서 실행해도 레포 루트 기준으로 동작한다. -ScriptName 은 레포 루트 기준 경로다.

# 스크립트 위치(scripts/)에서 레포 루트를 계산해 그 위치에서 실행한다
$RepoRoot = Split-Path -Parent $PSScriptRoot
Push-Location -Path $RepoRoot
try {

## 1. 기존 빌드 파일 삭제
Write-Host "[1/3] Cleaning old build files..."
# 'build' 폴더가 존재하면 삭제
if (Test-Path -Path ".\build" -PathType Container) {
    Remove-Item -Path ".\build" -Recurse -Force
}
# 'dist' 폴더가 존재하면 삭제
if (Test-Path -Path ".\dist" -PathType Container) {
    Remove-Item -Path ".\dist" -Recurse -Force
}
# '__pycache__' 폴더가 존재하면 삭제 (스크립트 경로에 따라 다를 수 있지만 일반적으로 포함)
if (Test-Path -Path ".\__pycache__" -PathType Container) {
    Remove-Item -Path ".\__pycache__" -Recurse -Force
}
# .spec 파일 삭제 (오류는 무시)
Get-Item -Path ".\*.spec" -ErrorAction SilentlyContinue | Remove-Item -Force

Write-Host "Done!"
Write-Host ""
# -----------------------------------------------------------------------------

# 플래그는 .github/workflows/release.yml 의 빌드 단계와 같다(P6-9, 결정 D3·D4: spec 없음).
# 다른 점은 버전 파일 하나: 로컬은 저장소의 bak_version_info.txt, release.yml 은 태그에서 만든 version_info.txt
pyinstaller --noconfirm --onefile --windowed --icon "./icon/ezbak.ico" --add-data "./icon;./icon" --collect-data customtkinter --hide-console "hide-early" --uac-admin --version-file "bak_version_info.txt" --name "ezBAK" --clean $ScriptName

}
finally {
    Pop-Location
}
