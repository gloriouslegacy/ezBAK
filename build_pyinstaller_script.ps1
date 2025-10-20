<<<<<<< HEAD
param(
    [Parameter(Mandatory=$true)]
    [string]$ScriptName
)

# PyInstaller script
# .\build_script.ps1 -ScriptName "test_app.py"

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

=======
param(
    [Parameter(Mandatory=$true)]
    [string]$ScriptName
)

# PyInstaller script
# .\build_script.ps1 -ScriptName "test_app.py"

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

>>>>>>> 1e32fc836007db295ae960da03d3a840bc47d492
pyinstaller --noconfirm --onefile --windowed --icon "./icon/ezbak.ico" --add-data "./icon/ezbak_title.ico;./icon" --hide-console "hide-early" --uac-admin --version-file "bak_version_info.txt" $ScriptName
