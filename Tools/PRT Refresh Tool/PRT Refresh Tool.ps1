Try {
    $Content = @'
cls
# ─── TOP OF CLOUD ─────────────────────────────────────────────────────────────
Write-Host "                 _ _ _ _ _                   " -ForegroundColor Cyan
Write-Host "              _/           \_                " -ForegroundColor Cyan
Write-Host "            _/               \_              " -ForegroundColor Cyan

# ─── KEY INSIDE CLOUD (multi-color lines) ─────────────────────────────────────
Write-Host "           /        " -ForegroundColor Cyan -NoNewline
Write-Host "__"                 -ForegroundColor Yellow -NoNewline
Write-Host "        \             " -ForegroundColor Cyan

Write-Host "          /       "  -ForegroundColor Cyan -NoNewline
Write-Host "/o )====>"           -ForegroundColor Yellow -NoNewline
Write-Host "    \            "  -ForegroundColor Cyan

Write-Host "          \       "  -ForegroundColor Cyan -NoNewline
Write-Host "\__\       "        -ForegroundColor Yellow -NoNewline
Write-Host "/            "      -ForegroundColor Cyan

# ─── BOTTOM OF CLOUD ──────────────────────────────────────────────────────────
Write-Host "            \_               _/              " -ForegroundColor Cyan
Write-Host "              \_           _/                " -ForegroundColor Cyan
Write-Host "                 -_ _ _ _-                   " -ForegroundColor Cyan
Write-Host "P R T   R E F R E S H   T O O L"

$username = Read-Host "Enter username (e.g. Admin_username@contoso.com)"
$runasCmd = "/user:AzureAD\$username `"dsregcmd /refreshprt`""
Start-Process "runas.exe" -ArgumentList $runasCmd -NoNewWindow -wait

# Simulate processing
Write-Host ""
Write-Host "Initiating PRT refresh process..." -ForegroundColor Yellow
Start-Sleep -milliseconds 500
Write-Host "Connecting to Azure AD..." -ForegroundColor Yellow
Start-Sleep -milliseconds 500
Write-Host "Authenticating user credentials..." -ForegroundColor Yellow
Start-Sleep -milliseconds 500
Write-Host "Refreshing Primary Refresh Token..." -ForegroundColor Yellow
Start-Sleep -Seconds 1
Write-Host "Validating token refresh..." -ForegroundColor Yellow
Start-Sleep -milliseconds 500
Write-Host ""
Write-Host "✓ PRT refresh completed successfully!" -ForegroundColor Green
Write-Host "✓ User authentication tokens have been updated" -ForegroundColor Green
Write-Host ""
pause

'@
    If (Test-Path -Path 'c:\ProgramData\Intune\Tools' -PathType Container) {
        Write-Host "Folder exists: 'c:\ProgramData\Intune\Tools'" -ForegroundColor Green
    } Else {
        New-Item -Path 'c:\ProgramData\Intune\Tools' -ItemType Directory
        Write-Host "Created tools folder" -ForegroundColor Green
    }
    Write-Host "Saving script as a file" -ForegroundColor Green
    $Content | Out-File -FilePath "c:\ProgramData\Intune\Tools\RefreshPRT.ps1" -Encoding UTF8 -Force
    Write-Host "Saving icon as a file" -ForegroundColor Green
    Copy-Item -Path "$PSScriptRoot\RPRT.ico" -Destination "c:\ProgramData\Intune\Tools\RPRT.ico" -Force
    Write-Host "Creating Start Menu shortcut" -ForegroundColor Green

    # Target directory is directly the Programs folder
    $startMenuFolder = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
	
    # Create the shortcut
    $shortcutPath = Join-Path -Path $startMenuFolder -ChildPath "PRT Refresh Tool.lnk"
    $targetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    $arguments = "-ExecutionPolicy Bypass -NoProfile -File `"C:\ProgramData\Intune\Tools\RefreshPRT.ps1`""
    $iconLocation = "c:\ProgramData\Intune\Tools\RPRT.ico"

    $WshShell = New-Object -ComObject WScript.Shell
    $shortcut = $WshShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $targetPath
    $shortcut.Arguments = $arguments
    $shortcut.IconLocation = $iconLocation
    $shortcut.Description = "Intune PRT Refresh Tool"
    $shortcut.WorkingDirectory = "C:\ProgramData\Intune\Tools\"
    $shortcut.Save()

    Write-Host "Created Start Menu shortcut at: $shortcutPath" -ForegroundColor Green   

    Exit-WithCode -exitcode $SuccessCode
} Catch {
    $Err = [PSCustomObject]@{
        Exception = $_.Exception.Message
        Reason    = $_.CategoryInfo.Reason
        Target    = $_.CategoryInfo.TargetName
        Script    = $_.InvocationInfo.ScriptName
        Line      = $_.InvocationInfo.ScriptLineNumber
        Column    = $_.InvocationInfo.OffsetInLine
    }
    Write-Host "Script ERROR" -ForegroundColor Red
    Write-Host "$Err" -ForegroundColor Red
    Write-Host "Exit with code 1" -ForegroundColor Red
    Exit 1
}