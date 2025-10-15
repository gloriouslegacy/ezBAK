param(
    [Parameter(Mandatory=$true)]
    [string]$ScriptName
)

# PyInstaller script
# .\build_script.ps1 -ScriptName "test_app.py"


pyinstaller --noconfirm --onefile --console --icon "./icon/ezbak.ico" --add-data "./icon/ezbak_title.ico;./icon" --hide-console "hide-early" --uac-admin --version-file "version_info.txt" $ScriptName