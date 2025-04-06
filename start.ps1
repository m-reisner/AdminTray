Clear-Host

$title = "AdminTray Starter"
Write-Host -ForegroundColor Cyan $title
Write-host -ForegroundColor Cyan $("="*$title.Length)

$desktop = "$env:HOMEPATH\Desktop"
if (-not (Test-Path -Path "$desktop\AdminTray.lnk")) {
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut("$desktop\AdminTray.lnk")
    $Shortcut.TargetPath = "pwsh.exe"
    $Shortcut.Arguments = "-nologo -noni -WindowStyle Hidden -executionpolicy bypass -File $PSScriptRoot\start.ps1"
    $Shortcut.IconLocation = "$PSScriptRoot\Torii.ico"
    $Shortcut.WorkingDirectory = "$PSScriptRoot\"
    $Shortcut.Description = "Shortcut to run AdminTray startscript"
    $Shortcut.Save()
    Write-Host -ForegroundColor Yellow "Verknüpfung am Desktop erstellt."
} else {
    Write-Host -ForegroundColor DarkYellow "Verknüpfung am Desktop bereits vorhanden."
}
Write-Host -ForegroundColor Yellow "Starte AdminTray."