# 1) 대상 확인
$files = Get-ChildItem -Path 'F:\' -Filter '*backup_2025-09-24*.log' | Where-Object { -not $_.PSIsContainer }
$files | Select Name, LastWriteTime

# 2) 읽기전용 해제 + 날짜 -20일로 변경
foreach ($f in $files) {
    if ($f.IsReadOnly) { $f.IsReadOnly = $false }
    [System.IO.File]::SetLastWriteTime($f.FullName, (Get-Date).AddDays(-20))
}

# 3) 결과 확인
Get-ChildItem -Path 'F:\' -Filter '*backup_2025-09-24*.log' | Where-Object { -not $_.PSIsContainer } |
  Select Name, LastWriteTime

