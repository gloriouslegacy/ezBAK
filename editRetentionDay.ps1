# 1) ��� Ȯ��
$files = Get-ChildItem -Path 'F:\' -Filter '*backup_2025-09-24*.log' | Where-Object { -not $_.PSIsContainer }
$files | Select Name, LastWriteTime

# 2) �б����� ���� + ��¥ -20�Ϸ� ����
foreach ($f in $files) {
    if ($f.IsReadOnly) { $f.IsReadOnly = $false }
    [System.IO.File]::SetLastWriteTime($f.FullName, (Get-Date).AddDays(-20))
}

# 3) ��� Ȯ��
Get-ChildItem -Path 'F:\' -Filter '*backup_2025-09-24*.log' | Where-Object { -not $_.PSIsContainer } |
  Select Name, LastWriteTime

