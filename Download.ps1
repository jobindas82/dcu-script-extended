# ============================
# Dell Driver Pack Downloader - Offline Edition
# Features: Progress bars, Sequential downloads, Version tracking
# Author: Jobin Das (jobindas82)
# GitHub: https://github.com/jobindas82
# Email: hello@jobin-d.com
# ============================

#Requires -Version 5.1

# Add: Prevent window from closing immediately
$host.UI.RawUI.WindowTitle = "Dell Driver Pack Downloader"

# Change console colors - set background to black
try {
    $host.UI.RawUI.BackgroundColor = "Black"
    $host.UI.RawUI.ForegroundColor = "White"
    Clear-Host
}
catch {
    # Silently continue if color change fails
}

# Initialize files directory first
$FilesRoot = Join-Path $PSScriptRoot "files"
if (!(Test-Path $FilesRoot)) { 
    New-Item -ItemType Directory -Path $FilesRoot -Force | Out-Null 
}

# Check if config file exists in files directory
$configPath = Join-Path $FilesRoot "config.psd1"
$sampleConfigPath = Join-Path $FilesRoot "config.sample.psd1"

if (!(Test-Path $configPath)) {
    Write-Host ""
    Write-Host "===============================================================================" -ForegroundColor Yellow
    Write-Host "                          CONFIGURATION REQUIRED                               " -ForegroundColor Yellow
    Write-Host "===============================================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Config file not found: config.psd1" -ForegroundColor Red
    Write-Host ""
    
    if (Test-Path $sampleConfigPath) {
        Write-Host "Creating config.psd1 from sample..." -ForegroundColor Cyan
        Copy-Item $sampleConfigPath $configPath
        Write-Host "Config file created successfully!" -ForegroundColor Green
    }
    else {
        Write-Host "Sample config file not found. Creating default config..." -ForegroundColor Cyan
        
        # Create default config
        $defaultConfig = @"
@{
    FilesRoot      = ".\files"
    CatalogUrl     = "https://dl.dell.com/catalog/DriverPackCatalog.cab"
    DriverURL      = "https://dl.dell.com"
    CatalogCABFile = ".\files\DriverPackCatalog.cab"
    CatalogXMLFile = ".\files\catalog\DriverPackCatalog.xml"
    
    DcuUrl         = "https://dl.dell.com/FOLDER13309338M/3/Dell-Command-Update-Application_Y5VJV_WIN64_5.5.0_A00_02.EXE"
    DcuUrl54       = "https://dl.dell.com/FOLDER12208864M/1/Dell-Command-Update-Windows-Universal-Application_R284X_WIN_5.4.1_A00.EXE"
    DcuUrl56       = "https://dl.dell.com/FOLDER13922605M/1/Dell-Command-Update-Application_5CR1Y_WIN64_5.6.0_A00.EXE"
    DcuFileName    = "Dell-Command-Update_55.exe"
    DcuFileName54  = "Dell-Command-Update_54.exe"
    DcuFileName56  = "Dell-Command-Update_56.exe"
    
    DotNetUrl      = "https://download.visualstudio.microsoft.com/download/pr/907765b0-2bf8-494e-93aa-5ef9553c5d68/a9308dc010617e6716c0e6abd53b05ce/windowsdesktop-runtime-8.0.11-win-x64.exe"
    DotNetFileName = "windowsdesktop-runtime-8.0-win-x64.exe"

    TargetModels   = @(
        "Dell Pro 14 Plus PB14250"
        "Dell Pro Max 16 Premium MA16250"
        "Dell Pro Max Slim FCS1250"
        "Latitude 5430"
        "Precision 3660 Tower"
        "Precision 5680"
        "Precision 5690"
        "Latitude 5440"
    )

    TargetOS       = "Windows 11 x64"
}
"@
        Set-Content -Path $configPath -Value $defaultConfig
        Write-Host "Default config file created!" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "===============================================================================" -ForegroundColor Yellow
    Write-Host "                          PLEASE UPDATE CONFIG FILE                            " -ForegroundColor Yellow
    Write-Host "===============================================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Before running this script, please:" -ForegroundColor White
    Write-Host ""
    Write-Host "  1. Open: $configPath" -ForegroundColor Cyan
    Write-Host "  2. Update 'TargetModels' with your Dell computer models" -ForegroundColor Cyan
    Write-Host "  3. Update 'TargetOS' if needed (Windows 10 x64 or Windows 11 x64)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "To find your Dell model name, run this command:" -ForegroundColor White
    Write-Host "  (Get-CimInstance Win32_ComputerSystem).Model" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to open config file in Notepad..." -ForegroundColor Green
    Write-Host ""
    
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Start-Process notepad.exe $configPath
    
    Write-Host "After updating the config file, run this script again." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 0
}

# Load config
$config = Import-PowerShellDataFile -Path $configPath

$FilesRoot = $config.FilesRoot
$CatalogUrl = $config.CatalogUrl
$DriverURL = $config.DriverURL
$CatalogCABFile = $config.CatalogCABFile
$CatalogXMLFile = $config.CatalogXMLFile
$TargetModels = $config.TargetModels
$TargetOS = $config.TargetOS
$DcuUrl = $config.DcuUrl
$DcuUrl54 = $config.DcuUrl54
$DcuUrl56 = $config.DcuUrl56
$DcuFileName = $config.DcuFileName
$DcuFileName54 = $config.DcuFileName54
$DcuFileName56 = $config.DcuFileName56
$DotNetUrl = $config.DotNetUrl
$DotNetFileName = $config.DotNetFileName

# Paths
$LogFile = Join-Path $FilesRoot "Download.log"
$DatabaseFile = Join-Path $FilesRoot "DriverPackDB.xml"


# Initialize
if (!(Test-Path $FilesRoot)) { New-Item -ItemType Directory -Path $FilesRoot -Force | Out-Null }

# ============================
# UI Helper Functions
# ============================

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    
    $colors = @{
        'Info'    = 'Cyan'
        'Success' = 'Green'
        'Warning' = 'Yellow'
        'Error'   = 'Red'
        'Header'  = 'Magenta'
    }
    
    $prefix = switch ($Type) {
        'Info' { '[+]' }
        'Success' { '[OK]' }
        'Warning' { '[!]' }
        'Error' { '[X]' }
        'Header' { '[#]' }
    }
    
    Write-Host "$prefix $Message" -ForegroundColor $colors[$Type]
    Write-Log "$prefix $Message"
}

function Write-Separator {
    param([string]$Char = "=", [int]$Length = 80)
    Write-Host ($Char * $Length) -ForegroundColor DarkGray
}

function Write-Header {
    param([string]$Title)
    Write-Host ""
    Write-Separator
    Write-Host "  $Title" -ForegroundColor Magenta
    Write-Separator
    Write-Host ""
}

function Write-Log {
    param([string]$Message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $LogFile -Value "$timestamp`t$Message" -ErrorAction SilentlyContinue
}

function Show-ProgressBar {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete,
        [int]$Id = 0
    )
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete -Id $Id
}

# ============================
# Database Functions
# ============================

function Initialize-Database {
    if (!(Test-Path $DatabaseFile)) {
        $db = @{
            LastUpdate  = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            DriverPacks = @()
        }
        $db | Export-Clixml -Path $DatabaseFile
        Write-ColorOutput "Database initialized" "Success"
    }
}

function Get-Database {
    if (Test-Path $DatabaseFile) {
        return Import-Clixml -Path $DatabaseFile
    }
    return @{ LastUpdate = $null; DriverPacks = @() }
}

function Save-Database {
    param($Database)
    $Database.LastUpdate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $Database | Export-Clixml -Path $DatabaseFile
}

function Get-DriverPackFromDB {
    param(
        [string]$Model,
        [string]$OS,
        $Database
    )
    return $Database.DriverPacks | Where-Object { $_.Model -eq $Model -and $_.OS -eq $OS } | Select-Object -First 1
}

function Add-DriverPackToDB {
    param(
        [string]$Model,
        [string]$OS,
        [string]$Version,
        [string]$Hash,
        [string]$FilePath,
        $Database
    )
    
    $existing = Get-DriverPackFromDB -Model $Model -OS $OS -Database $Database
    if ($existing) {
        $Database.DriverPacks = $Database.DriverPacks | Where-Object { -not ($_.Model -eq $Model -and $_.OS -eq $OS) }
    }
    
    $Database.DriverPacks += @{
        Model      = $Model
        OS         = $OS
        Version    = $Version
        Hash       = $Hash
        FilePath   = $FilePath
        Downloaded = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }
}

# ============================
# Download Functions
# ============================

function Get-FileHashMD5 {
    param([string]$FilePath)
    if (Test-Path $FilePath) {
        return (Get-FileHash -Path $FilePath -Algorithm MD5).Hash
    }
    return $null
}

function Get-Catalog {
    Write-ColorOutput "Downloading Dell Driver Pack Catalog..." "Header"
    Show-ProgressBar -Activity "Catalog Download" -Status "Downloading catalog..." -PercentComplete 0
    
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "Mozilla/5.0")
        $webClient.DownloadFile($CatalogUrl, $CatalogCABFile)
        
        Show-ProgressBar -Activity "Catalog Download" -Status "Extracting catalog..." -PercentComplete 50
        
        $extractFolder = Split-Path $CatalogXMLFile -Parent
        if (!(Test-Path $extractFolder)) { New-Item -ItemType Directory -Path $extractFolder -Force | Out-Null }
        
        $sevenZip = "C:\Program Files\7-Zip\7z.exe"
        if (!(Test-Path $sevenZip)) {
            $sevenZip = "C:\Program Files (x86)\7-Zip\7z.exe"
        }
        
        if (!(Test-Path $sevenZip)) {
            throw "7-Zip not found. Please install 7-Zip from https://www.7-zip.org/"
        }
        
        & $sevenZip x "$CatalogCABFile" -o"$extractFolder" -y | Out-Null
        
        if (!(Test-Path $CatalogXMLFile)) {
            throw "Catalog XML not found after extraction"
        }
        
        Show-ProgressBar -Activity "Catalog Download" -Status "Complete" -PercentComplete 100
        Write-ColorOutput "Catalog downloaded and extracted successfully" "Success"
    }
    catch {
        Write-ColorOutput "Catalog download failed: $_" "Error"
        throw
    }
    finally {
        Write-Progress -Activity "Catalog Download" -Completed
    }
}

function Start-SequentialDownload {
    param(
        [array]$DownloadJobs,
        $Database
    )
    
    $totalJobs = $DownloadJobs.Count
    if ($totalJobs -eq 0) {
        Write-ColorOutput "No driver packs to download" "Warning"
        return
    }
    
    Write-ColorOutput "Starting download of $totalJobs driver pack(s)..." "Header"
    Write-Host ""
    
    $successCount = 0
    $failedCount = 0
    $currentJob = 0
    
    foreach ($job in $DownloadJobs) {
        $currentJob++
        
        Write-Host ""
        Write-Separator "-"
        Write-Host "  Download $currentJob of $totalJobs" -ForegroundColor Cyan
        Write-Separator "-"
        Write-Host ""
        Write-ColorOutput "Model: $($job.Model)" "Info"
        Write-ColorOutput "OS: $($job.OS)" "Info"
        Write-ColorOutput "Package: $($job.Name)" "Info"
        Write-Host ""
        
        $percent = [math]::Round(($currentJob / $totalJobs) * 100)
        Show-ProgressBar -Activity "Downloading Driver Packs" -Status "Downloading $($job.Model) ($currentJob of $totalJobs)" -PercentComplete $percent
        
        try {
            # Download the file
            $webClient = New-Object System.Net.WebClient
            $webClient.Headers.Add("User-Agent", "Mozilla/5.0")
            
            Write-ColorOutput "Downloading from: $($job.Url)" "Info"
            Write-ColorOutput "Saving to: $($job.Destination)" "Info"
            Write-Host ""
            
            $webClient.DownloadFile($job.Url, $job.Destination)
            
            if (Test-Path $job.Destination) {
                # Verify hash
                Write-ColorOutput "Verifying file integrity..." "Info"
                $hash = (Get-FileHash -Path $job.Destination -Algorithm MD5).Hash
                
                if ($hash -eq $job.Hash) {
                    $fileSize = (Get-Item $job.Destination).Length / 1MB
                    Write-ColorOutput "Download successful! ($([math]::Round($fileSize, 2)) MB)" "Success"
                    Write-ColorOutput "Hash verified: $hash" "Success"
                    
                    # Add to database
                    Add-DriverPackToDB -Model $job.Model -OS $job.OS -Version "Latest" -Hash $job.Hash -FilePath $job.Destination -Database $Database
                    $successCount++
                }
                else {
                    Write-ColorOutput "Hash mismatch!" "Error"
                    Write-ColorOutput "Expected: $($job.Hash)" "Error"
                    Write-ColorOutput "Got: $hash" "Error"
                    Remove-Item $job.Destination -Force -ErrorAction SilentlyContinue
                    $failedCount++
                }
            }
            else {
                Write-ColorOutput "File not created after download" "Error"
                $failedCount++
            }
        }
        catch {
            Write-ColorOutput "Download failed: $($_.Exception.Message)" "Error"
            if (Test-Path $job.Destination) {
                Remove-Item $job.Destination -Force -ErrorAction SilentlyContinue
            }
            $failedCount++
        }
        
        # Small delay between downloads
        Start-Sleep -Milliseconds 500
    }
    
    Write-Progress -Activity "Downloading Driver Packs" -Completed
    
    Write-Host ""
    Write-Separator "="
    Write-Host ""
    Write-Host "  DOWNLOAD RESULTS" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "  Total Downloads : " -NoNewline -ForegroundColor Gray
    Write-Host $totalJobs -ForegroundColor Cyan
    Write-Host "  Successful      : " -NoNewline -ForegroundColor Gray
    Write-Host $successCount -ForegroundColor Green
    Write-Host "  Failed          : " -NoNewline -ForegroundColor Gray
    Write-Host $failedCount -ForegroundColor $(if ($failedCount -gt 0) { "Red" } else { "Gray" })
    Write-Host ""
    Write-Separator "="
    Write-Host ""
}


# ============================
# Prerequisites Download Functions
# ============================

function Download-Prerequisite {
    param(
        [string]$Url,
        [string]$Destination,
        [string]$Name
    )
    
    if (Test-Path $Destination) {
        Write-ColorOutput "$Name already exists, skipping download" "Success"
        return
    }
    
    Write-ColorOutput "Downloading $Name..." "Info"
    Show-ProgressBar -Activity "$Name Download" -Status "Downloading..." -PercentComplete 0
    
    try {
        $destFolder = Split-Path $Destination -Parent
        if (!(Test-Path $destFolder)) { 
            New-Item -ItemType Directory -Path $destFolder -Force | Out-Null 
        }
        
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "Mozilla/5.0")
        
        # Download with progress
        $webClient.DownloadFile($Url, $Destination)
        
        if (Test-Path $Destination) {
            $fileSize = (Get-Item $Destination).Length / 1MB
            Write-ColorOutput "$Name downloaded successfully ($([math]::Round($fileSize, 2)) MB)" "Success"
            Show-ProgressBar -Activity "$Name Download" -Status "Complete" -PercentComplete 100
            Start-Sleep -Milliseconds 500
            Write-Progress -Activity "$Name Download" -Completed
            return
        }
        else {
            throw "File not created after download"
        }
    }
    catch {
        Write-ColorOutput "Failed to download ${Name}: $($_.Exception.Message)" "Error"
        Write-Progress -Activity "$Name Download" -Completed
        if (Test-Path $Destination) {
            Remove-Item $Destination -Force -ErrorAction SilentlyContinue
        }
        return
    }
}

function Download-Prerequisites {
    Write-Header "CHECKING PREREQUISITES"
    
    # Download Dell Command Update 5.5
    $dcuPath = Join-Path $FilesRoot "dcu\$DcuFileName"
    Download-Prerequisite -Url $DcuUrl -Destination $dcuPath -Name "Dell Command Update 5.5"
    
    Write-Host ""
    
    # Download Dell Command Update 5.4
    $dcuPath54 = Join-Path $FilesRoot "dcu\$DcuFileName54"
    Download-Prerequisite -Url $DcuUrl54 -Destination $dcuPath54 -Name "Dell Command Update 5.4"
    
    Write-Host ""
    
    # Download Dell Command Update 5.6
    $dcuPath56 = Join-Path $FilesRoot "dcu\$DcuFileName56"
    Download-Prerequisite -Url $DcuUrl56 -Destination $dcuPath56 -Name "Dell Command Update 5.6"
    
    Write-Host ""
    
    # Download .NET Desktop Runtime
    $dotNetPath = Join-Path $FilesRoot "dotnet\$DotNetFileName"
    Download-Prerequisite -Url $DotNetUrl -Destination $dotNetPath -Name ".NET Desktop Runtime"
    
    Write-Host ""
}

# ============================
# Main Logic
# ============================

function Get-Driver-Pack {
    Clear-Host
    Write-Host ""
    Write-Host "===============================================================================" -ForegroundColor Cyan
    Write-Host "                                                                               " -ForegroundColor Cyan
    Write-Host "           Dell Driver Pack Downloader - Offline Edition                     " -ForegroundColor Cyan
    Write-Host "           Author: Jobin Das (hello@jobin-d.com)                           " -ForegroundColor Cyan
    Write-Host "           GitHub: https://github.com/jobindas82                              " -ForegroundColor Cyan
    Write-Host "                                                                               " -ForegroundColor Cyan
    Write-Host "===============================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    Download-Prerequisites
    
    Initialize-Database
    $database = Get-Database
    
    Write-Header "DOWNLOADING CATALOG"
    Get-Catalog
    
    Write-Header "ANALYZING DRIVER PACKS"
    [XML]$Catalog = Get-Content $CatalogXMLFile
    $DriverPackages = $Catalog.DriverPackManifest.DriverPackage
    
    $downloadJobs = @()
    $skippedCount = 0
    
    Write-ColorOutput "Searching for driver packs matching target models..." "Info"
    Write-Host ""
    
    foreach ($pkg in $DriverPackages) {
        if (-not $pkg.SupportedSystems.Brand.Model) { continue }
        
        $modelNode = $pkg.SupportedSystems.Brand.Model
        
        $models = @()
        if ($modelNode -is [array]) {
            foreach ($m in $modelNode) {
                if ($m.name) { 
                    $models += $m.name.Trim()
                }
            }
        }
        else {
            if ($modelNode.name) {
                $models += $modelNode.name.Trim()
            }
        }
        
        $models = $models | Select-Object -Unique
        
        # Enhanced matching: support both exact match and partial match
        $matchedModel = $null
        $catalogModel = $null
        
        foreach ($model in $models) {
            foreach ($targetModel in $TargetModels) {
                # Try exact match first
                if ($model -ieq $targetModel) {
                    $matchedModel = $targetModel
                    $catalogModel = $model
                    break
                }
                # Try partial match: check if catalog model is contained in target model
                # or if target model is contained in catalog model
                elseif ($targetModel -like "*$model*" -or $model -like "*$targetModel*") {
                    $matchedModel = $targetModel
                    $catalogModel = $model
                    Write-ColorOutput "  Partial match: '$targetModel' matches catalog entry '$model'" "Info"
                    break
                }
            }
            if ($matchedModel) { break }
        }
        
        if (-not $matchedModel) { continue }
        
        Write-Host ""
        Write-Host ">> Model: $matchedModel" -ForegroundColor White
        if ($catalogModel -and $catalogModel -ne $matchedModel) {
            Write-Host "   Catalog Model: $catalogModel" -ForegroundColor Gray
        }
        
        $supportedOSList = $pkg.SupportedOperatingSystems.OperatingSystem
        if (-not $supportedOSList) { 
            Write-ColorOutput "  No OS information found" "Warning"
            continue 
        }
        
        if ($supportedOSList -isnot [array]) {
            $supportedOSList = @($supportedOSList)
        }
        
        foreach ($os in $supportedOSList) {
            $osDisplay = $null
            
            if ($os.Display) {
                if ($os.Display.'#cdata-section') {
                    $osDisplay = $os.Display.'#cdata-section'
                }
                elseif ($os.Display.'#text') {
                    $osDisplay = $os.Display.'#text'
                }
                elseif ($os.Display -is [string]) {
                    $osDisplay = $os.Display
                }
            }
            
            if ([string]::IsNullOrWhiteSpace($osDisplay)) { continue }
            $osDisplay = $osDisplay.Trim()
            
            Write-ColorOutput "   OS: $osDisplay" "Info"
            
            $isTargetOS = $false
            if ($TargetOS -is [array]) {
                foreach ($targetOS in $TargetOS) {
                    if ($osDisplay -ieq $targetOS) {
                        $isTargetOS = $true
                        break
                    }
                }
            }
            else {
                if ($osDisplay -ieq $TargetOS) {
                    $isTargetOS = $true
                }
            }
            
            if (-not $isTargetOS) { 
                Write-ColorOutput "   [SKIP] Not target OS" "Warning"
                continue 
            }
            
            $downloadPath = "$DriverURL/$($pkg.path)"
            
            # Extract the original file extension from the download path
            $originalFileName = [System.IO.Path]::GetFileName($pkg.path)
            $fileExtension = [System.IO.Path]::GetExtension($originalFileName)
            
            $destFolder = "$FilesRoot\$($osDisplay.Replace(' ', '_'))\$($matchedModel.Replace(' ', '_'))"
            if (!(Test-Path $destFolder)) { New-Item -ItemType Directory -Path $destFolder -Force | Out-Null }

            # Use "pack" as base name but keep the original extension
            $fileName = Join-Path $destFolder "pack$fileExtension"
            
            # Get package name before creating metadata
            $pkgName = "Driver Pack"
            if ($pkg.Name.Display.'#cdata-section') {
                $pkgName = $pkg.Name.Display.'#cdata-section'
            }
            elseif ($pkg.Name.Display.'#text') {
                $pkgName = $pkg.Name.Display.'#text'
            }
            elseif ($pkg.Name.Display -is [string]) {
                $pkgName = $pkg.Name.Display
            }
            
            # Get version if available
            $pkgVersion = "N/A"
            if ($pkg.version) {
                $pkgVersion = $pkg.version
            }
            
            # Create a metadata file to help Install.ps1 with model mapping
            $metadataFile = Join-Path $destFolder "model_info.txt"
            $metadataContent = @"
FullModelName=$matchedModel
CatalogModelName=$catalogModel
DriverPackageName=$pkgName
DriverPackageFile=$originalFileName
DriverVersion=$pkgVersion
DownloadDate=$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@
            Set-Content -Path $metadataFile -Value $metadataContent -Force

            $expectedHash = $null
            if ($pkg.hashMD5) {
                $expectedHash = $pkg.hashMD5
                Write-ColorOutput "   Hash (MD5): $expectedHash" "Info"
            }
            else {
                Write-ColorOutput "   Warning: No MD5 hash found for this package, skipping for safety" "Warning"
                continue
            }
            
            $existingPack = Get-DriverPackFromDB -Model $matchedModel -OS $osDisplay -Database $database
            if ($existingPack -and $existingPack.Hash -eq $expectedHash -and (Test-Path $fileName)) {
                Write-ColorOutput "   [CACHED] Already up-to-date (hash matches)" "Success"
                $skippedCount++
                continue
            }
            
            Write-ColorOutput "   [QUEUE] Added to download queue (${originalFileName})" "Info"
            
            $downloadJobs += @{
                Url         = $downloadPath
                Destination = $fileName
                Hash        = $expectedHash
                Model       = $matchedModel
                OS          = $osDisplay
                Name        = $pkgName
            }
        }
        
        Write-Host "-------------------------------------------" -ForegroundColor DarkGray
    }
    
    Write-Host ""
    Write-Separator "-"
    Write-Host ""
    Write-Host "  DOWNLOAD SUMMARY" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "  New/Updated Packs : " -NoNewline -ForegroundColor Gray
    Write-Host $downloadJobs.Count -ForegroundColor Green
    Write-Host "  Up-to-date Packs  : " -NoNewline -ForegroundColor Gray
    Write-Host $skippedCount -ForegroundColor Cyan
    Write-Host ""
    Write-Separator "-"
    Write-Host ""
    
    if ($downloadJobs.Count -gt 0) {
        Start-SequentialDownload -DownloadJobs $downloadJobs -Database $database
        Save-Database -Database $database
    }
    else {
        Write-ColorOutput "No downloads needed. All driver packs are up to date!" "Success"
    }
    
    Write-Host ""
    Write-Separator "="
    Write-Host ""
    Write-Host "  [SUCCESS] Process completed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Separator "="
    Write-Host ""
    
    # Show completion summary
    Write-Host ""
    Write-Host "  NEXT STEPS:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  1. Review the downloaded driver packs in: $FilesRoot" -ForegroundColor White
    Write-Host "  2. Run Install.ps1 to install drivers on target systems" -ForegroundColor White
    Write-Host ""
    Write-Host "  Downloaded files location:" -ForegroundColor Gray
    Write-Host "  - Driver Packs: $FilesRoot\Windows_*_x64\" -ForegroundColor Gray
    Write-Host "  - DCU 5.6:      $FilesRoot\dcu\$DcuFileName56" -ForegroundColor Gray
    Write-Host "  - DCU 5.5:      $FilesRoot\dcu\$DcuFileName" -ForegroundColor Gray
    Write-Host "  - DCU 5.4:      $FilesRoot\dcu\$DcuFileName54" -ForegroundColor Gray
    Write-Host "  - .NET Runtime: $FilesRoot\dotnet\$DotNetFileName" -ForegroundColor Gray
    Write-Host ""
    Write-Separator "-"
    Write-Host ""
    Write-Host "Download process completed. You may now close this window." -ForegroundColor Green
    Write-Host "Press any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Entry point
try {
    Get-Driver-Pack
}
catch {
    $errorMsg = $_.Exception.Message
    Write-ColorOutput "Fatal error: $errorMsg" "Error"
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}