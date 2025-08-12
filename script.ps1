# Find the latest installed Windows Terminal package
$wtPackage = Get-ChildItem "C:\Program Files\WindowsApps" -Directory -Filter "Microsoft.WindowsTerminal_*" |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $wtPackage) {
    Write-Host "❌ Windows Terminal package not found in WindowsApps. Is it installed from Microsoft Store?"
    exit
}

# Full path to the real Windows Terminal executable (with embedded icon)
$wtExePath = Join-Path $wtPackage.FullName "WindowsTerminal.exe"

# This is what we'll use for the icon in the context menu
$iconPath = "`"$wtExePath`""

# Path to launch executable (App Installer path works for launching)
$appPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"

# --- Background right-click menu ---
New-Item -Path "HKCU:\Software\Classes\Directory\Background\shell\WindowsTerminal" -Force |
    New-ItemProperty -Name "(default)" -Value "Open in Windows Terminal Here" -Force |
    Out-Null

Set-ItemProperty -Path "HKCU:\Software\Classes\Directory\Background\shell\WindowsTerminal" -Name "Icon" -Value $iconPath

New-Item -Path "HKCU:\Software\Classes\Directory\Background\shell\WindowsTerminal\command" -Force |
    New-ItemProperty -Name "(default)" -Value "`"$appPath`" --startingDirectory . " -Force |
    Out-Null

# --- Folder right-click menu ---
New-Item -Path "HKCU:\Software\Classes\Directory\shell\WindowsTerminal" -Force |
    New-ItemProperty -Name "(default)" -Value "Open in Windows Terminal Here" -Force |
    Out-Null

Set-ItemProperty -Path "HKCU:\Software\Classes\Directory\shell\WindowsTerminal" -Name "Icon" -Value $iconPath

New-Item -Path "HKCU:\Software\Classes\Directory\shell\WindowsTerminal\command" -Force |
    New-ItemProperty -Name "(default)" -Value "`"$appPath`" --startingDirectory `"%V`"" -Force |
    Out-Null
