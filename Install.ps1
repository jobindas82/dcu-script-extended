# ============================
# Dell Driver Pack Installer - Enhanced Version
# Features: Auto-detection, Progress bars, Colored output
# Author: Jobin Das (jobindas82)
# GitHub: https://github.com/jobindas82
# Email: jobindas82@gmail.com
# ============================

#Requires -Version 5.1

# Check if running as administrator, if not, restart as admin
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Restarting as Administrator..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Add: Prevent window from closing immediately
$host.UI.RawUI.WindowTitle = "Dell Driver Pack Installer"

# Load config from files directory
$ScriptRoot = $PSScriptRoot
$FilesRoot = Join-Path $ScriptRoot "files"
$configPath = Join-Path $FilesRoot "config.psd1"

if (!(Test-Path $configPath)) {
    Write-Host "Error: Config file not found at $configPath" -ForegroundColor Red
    Write-Host "Please run Download.ps1 first to create the configuration file." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

$config = Import-PowerShellDataFile -Path $configPath

# Don't override FilesRoot from config - use the actual path
$DcuFileName = $config.DcuFileName
$DotNetFileName = $config.DotNetFileName

# Add: Define temp directory on C: drive
$TempRoot = "C:\Temp\DellDriverInstaller"
if (!(Test-Path $TempRoot)) {
    New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null
}

# Move log file to temp directory
$LogFile = Join-Path $TempRoot "Install.log"

# Add: State file to track installation status
$StateFile = Join-Path $TempRoot "install_state.json"

# Initialize
if (!(Test-Path $FilesRoot)) {
    Write-Host "Error: Files directory not found. Please run Download.ps1 first." -ForegroundColor Red
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# ============================
# UI Helper Functions
# ============================

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    
    $colors = @{
        'Info' = 'Cyan'
        'Success' = 'Green'
        'Warning' = 'Yellow'
        'Error' = 'Red'
        'Header' = 'Magenta'
    }
    
    $prefix = switch ($Type) {
        'Info'    { '[+]' }
        'Success' { '[OK]' }
        'Warning' { '[!]' }
        'Error'   { '[X]' }
        'Header'  { '[#]' }
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
        [int]$PercentComplete
    )
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
}

# ============================
# State Management Functions
# ============================

function Set-InstallationState {
    param(
        [string]$State,
        [string]$Step = "",
        [string]$Details = ""
    )
    
    $stateObj = @{
        State = $State
        Step = $Step
        Details = $Details
        Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        ComputerName = $env:COMPUTERNAME
    }
    
    try {
        $stateObj | ConvertTo-Json | Set-Content -Path $StateFile -Force
        Write-Log "Installation state: $State - $Step - $Details"
    }
    catch {
        Write-Log "Failed to save installation state: $_"
    }
}

function Get-InstallationState {
    if (Test-Path $StateFile) {
        try {
            $state = Get-Content $StateFile -Raw | ConvertFrom-Json
            return $state
        }
        catch {
            Write-Log "Failed to read installation state: $_"
            return $null
        }
    }
    return $null
}

function Clear-InstallationState {
    if (Test-Path $StateFile) {
        try {
            Remove-Item $StateFile -Force -ErrorAction SilentlyContinue
            Write-Log "Installation state cleared"
        }
        catch {
            Write-Log "Failed to clear installation state: $_"
        }
    }
}

function Test-PreviousCrash {
    $previousState = Get-InstallationState
    
    if ($previousState -and $previousState.State -eq "IN_PROGRESS") {
        Write-Host ""
        Write-Host "===============================================================================" -ForegroundColor Red
        Write-Host "                          PREVIOUS INSTALLATION FAILED                         " -ForegroundColor Red
        Write-Host "===============================================================================" -ForegroundColor Red
        Write-Host ""
        Write-ColorOutput "A previous installation was interrupted unexpectedly!" "Error"
        Write-Host ""
        Write-Host "  Last Step    : " -NoNewline -ForegroundColor Gray
        Write-Host $previousState.Step -ForegroundColor Yellow
        Write-Host "  Last Action  : " -NoNewline -ForegroundColor Gray
        Write-Host $previousState.Details -ForegroundColor Yellow
        Write-Host "  Time         : " -NoNewline -ForegroundColor Gray
        Write-Host $previousState.Timestamp -ForegroundColor Yellow
        Write-Host ""
        Write-ColorOutput "This may indicate:" "Warning"
        Write-Host "  - System crash (BSOD) during driver installation" -ForegroundColor Yellow
        Write-Host "  - Power failure or forced shutdown" -ForegroundColor Yellow
        Write-Host "  - Incompatible driver causing system instability" -ForegroundColor Yellow
        Write-Host ""
        Write-Separator "-"
        Write-Host ""
        
        Write-ColorOutput "Recommendations:" "Header"
        Write-Host "  1. Boot into Safe Mode if system is unstable" -ForegroundColor Cyan
        Write-Host "  2. Use System Restore to revert changes" -ForegroundColor Cyan
        Write-Host "  3. Check Windows Event Viewer for crash details" -ForegroundColor Cyan
        Write-Host "  4. Review log file: $LogFile" -ForegroundColor Cyan
        Write-Host ""
        Write-Separator "-"
        Write-Host ""
        
        $response = Read-Host "  Do you want to continue anyway? (Y/N)"
        if ($response -ne 'Y' -and $response -ne 'y') {
            Write-ColorOutput "Installation cancelled by user" "Warning"
            Clear-InstallationState
            exit 0
        }
        
        Write-Host ""
        Write-ColorOutput "Continuing installation..." "Info"
        Write-Host ""
        Start-Sleep -Seconds 2
    }
}

# ============================
# System Detection
# ============================

function Get-SystemInfo {
    Write-ColorOutput "Detecting system information..." "Info"
    
    try {
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        
        if (-not $computerSystem -or -not $osInfo) {
            throw "Unable to retrieve system information"
        }
        
        $model = $computerSystem.Model.Trim()
        $osName = if ($osInfo.Caption -match "Windows 11") { "Windows 11 x64" } 
                  elseif ($osInfo.Caption -match "Windows 10") { "Windows 10 x64" } 
                  else { "Unknown" }
        
        Write-ColorOutput "System Model: $model" "Info"
        Write-ColorOutput "Operating System: $osName" "Info"
        
        return @{
            Model = $model
            OS = $osName
            Architecture = $osInfo.OSArchitecture
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-ColorOutput "Failed to detect system information: $errorMsg" "Error"
        throw
    }
}

# ============================
# DCU Functions
# ============================

function Get-DcuPath {
    $paths = @(
        "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe",
        "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
    )
    
    foreach ($path in $paths) {
        if (Test-Path $path) {
            return $path
        }
    }
    return $null
}

function Get-DcuVersion {
    param([string]$DcuPath)
    
    try {
        if (Test-Path $DcuPath) {
            $versionOutput = & $DcuPath /version 2>&1 | Out-String
            if ($versionOutput -match "(\d+\.\d+\.\d+)") {
                return [version]$matches[1]
            }
        }
    }
    catch {
        Write-Log "Failed to get DCU version: $_"
    }
    return $null
}

function Get-InstalledDcuVersion {
    try {
        # Check registry for installed version
        $regPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        foreach ($regPath in $regPaths) {
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | 
                    Where-Object { $_.DisplayName -like "*Dell Command*Update*" }
            
            if ($apps) {
                foreach ($app in $apps) {
                    if ($app.DisplayVersion) {
                        Write-Log "Found DCU version in registry: $($app.DisplayVersion)"
                        return [version]$app.DisplayVersion
                    }
                }
            }
        }
    }
    catch {
        Write-Log "Failed to get installed DCU version from registry: $_"
    }
    return $null
}

function Uninstall-DellCommandUpdate {
    Write-ColorOutput "Uninstalling old Dell Command | Update..." "Info"
    Show-ProgressBar -Activity "DCU Uninstall" -Status "Searching for uninstaller..." -PercentComplete 20
    
    try {
        # Find uninstaller
        $uninstallString = $null
        $regPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        foreach ($regPath in $regPaths) {
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | 
                    Where-Object { $_.DisplayName -like "*Dell Command*Update*" }
            
            if ($apps) {
                $uninstallString = $apps[0].UninstallString
                break
            }
        }
        
        if (-not $uninstallString) {
            Write-ColorOutput "Uninstaller not found in registry" "Warning"
            return $false
        }
        
        Write-ColorOutput "Found uninstaller, removing old version..." "Info"
        Show-ProgressBar -Activity "DCU Uninstall" -Status "Uninstalling..." -PercentComplete 50
        
        # Parse uninstall string
        if ($uninstallString -match '"([^"]+)"') {
            $uninstallerPath = $matches[1]
            $process = Start-Process -FilePath $uninstallerPath -ArgumentList "/S" -Wait -PassThru -NoNewWindow
        }
        else {
            # Try direct execution
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $uninstallString, "/S" -Wait -PassThru -NoNewWindow
        }
        
        # Wait for uninstall to complete
        Start-Sleep -Seconds 3
        
        Show-ProgressBar -Activity "DCU Uninstall" -Status "Complete" -PercentComplete 100
        Start-Sleep -Milliseconds 500
        Write-Progress -Activity "DCU Uninstall" -Completed
        
        if ($process.ExitCode -eq 0 -or -not (Get-DcuPath)) {
            Write-ColorOutput "Old version uninstalled successfully" "Success"
            return $true
        }
        else {
            Write-ColorOutput "Uninstall completed with exit code: $($process.ExitCode)" "Warning"
            return $false
        }
    }
    catch {
        Write-ColorOutput "Uninstall failed: $_" "Error"
        return $false
    }
}

function Install-DotNetRuntime {
    Write-ColorOutput "Checking .NET Desktop Runtime..." "Header"
    Show-ProgressBar -Activity ".NET Runtime Check" -Status "Checking installation..." -PercentComplete 10
    
    # Check if .NET 8.0 is installed
    $dotNetInstalled = $false
    try {
        $dotNetVersions = Get-ChildItem "HKLM:\SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedhost" -ErrorAction SilentlyContinue
        if ($dotNetVersions) {
            foreach ($version in $dotNetVersions.GetValueNames()) {
                if ($version -like "8.0.*") {
                    $dotNetInstalled = $true
                    break
                }
            }
        }
    }
    catch {
        # Registry check failed, assume not installed
    }
    
    if ($dotNetInstalled) {
        Write-ColorOutput ".NET Desktop Runtime 8.0 already installed" "Success"
        Show-ProgressBar -Activity ".NET Runtime Check" -Status "Already installed" -PercentComplete 100
        Start-Sleep -Milliseconds 500
        Write-Progress -Activity ".NET Runtime Check" -Completed
        return
    }
    
    $setupPath = Join-Path $FilesRoot "dotnet\$DotNetFileName"
    if (!(Test-Path $setupPath)) {
        Write-ColorOutput ".NET Runtime installer not found at $setupPath" "Error"
        Write-ColorOutput "Please run Download.ps1 first" "Warning"
        throw ".NET Runtime installer not found"
    }
    
    Write-ColorOutput "Installing .NET Desktop Runtime 8.0..." "Info"
    Show-ProgressBar -Activity ".NET Runtime Installation" -Status "Running installer..." -PercentComplete 50
    
    try {
        $process = Start-Process -FilePath $setupPath -ArgumentList "/install", "/quiet", "/norestart" -Wait -PassThru
        
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            Write-ColorOutput ".NET Desktop Runtime installed successfully" "Success"
            Show-ProgressBar -Activity ".NET Runtime Installation" -Status "Complete" -PercentComplete 100
            Start-Sleep -Milliseconds 500
            Write-Progress -Activity ".NET Runtime Installation" -Completed
            return
        } else {
            throw "Installation failed with exit code: $($process.ExitCode)"
        }
    }
    catch {
        Write-ColorOutput ".NET Runtime installation failed: $_" "Error"
        throw
    }
}

function Install-DellCommandUpdate {
    Write-ColorOutput "Checking for Dell Command | Update..." "Header"
    Show-ProgressBar -Activity "DCU Installation" -Status "Checking installation..." -PercentComplete 10
    
    $dcuPath = Get-DcuPath
    
    # Define minimum required version (adjust as needed)
    $minimumVersion = [version]"5.0.0"
    $needsReinstall = $false
    
    if ($dcuPath) {
        Write-ColorOutput "Dell Command | Update found" "Info"
        
        # Check version
        $installedVersion = Get-InstalledDcuVersion
        
        if ($installedVersion) {
            Write-ColorOutput "Installed version: $installedVersion" "Info"
            Write-Log "DCU installed version: $installedVersion"
            
            if ($installedVersion -lt $minimumVersion) {
                Write-ColorOutput "Installed version is outdated (minimum required: $minimumVersion)" "Warning"
                $needsReinstall = $true
            }
            else {
                Write-ColorOutput "Dell Command | Update version is up to date" "Success"
                Show-ProgressBar -Activity "DCU Installation" -Status "Already installed" -PercentComplete 100
                Start-Sleep -Milliseconds 500
                Write-Progress -Activity "DCU Installation" -Completed
                return $dcuPath
            }
        }
        else {
            Write-ColorOutput "Could not determine DCU version, assuming it's current" "Warning"
            Show-ProgressBar -Activity "DCU Installation" -Status "Already installed" -PercentComplete 100
            Start-Sleep -Milliseconds 500
            Write-Progress -Activity "DCU Installation" -Completed
            return $dcuPath
        }
    }
    
    # Uninstall old version if needed
    if ($needsReinstall) {
        Write-Host ""
        $uninstallSuccess = Uninstall-DellCommandUpdate
        
        if (-not $uninstallSuccess) {
            Write-ColorOutput "Failed to uninstall old version, attempting to install anyway..." "Warning"
        }
        
        Write-Host ""
    }
    
    # Install new version
    $setupPath = Join-Path $FilesRoot "dcu\$DcuFileName"
    if (!(Test-Path $setupPath)) {
        Write-ColorOutput "DCU installer not found at $setupPath" "Error"
        Write-ColorOutput "Please run Download.ps1 first" "Warning"
        throw "DCU installer not found"
    }
    
    Write-ColorOutput "Installing Dell Command | Update..." "Info"
    Show-ProgressBar -Activity "DCU Installation" -Status "Running installer..." -PercentComplete 50
    
    try {
        $process = Start-Process -FilePath $setupPath -ArgumentList "/s" -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Show-ProgressBar -Activity "DCU Installation" -Status "Verifying installation..." -PercentComplete 90
            Start-Sleep -Seconds 2
            
            $dcuPath = Get-DcuPath
            if ($dcuPath) {
                $newVersion = Get-InstalledDcuVersion
                if ($newVersion) {
                    Write-ColorOutput "Dell Command | Update $newVersion installed successfully" "Success"
                }
                else {
                    Write-ColorOutput "Dell Command | Update installed successfully" "Success"
                }
                
                Show-ProgressBar -Activity "DCU Installation" -Status "Complete" -PercentComplete 100
                Start-Sleep -Milliseconds 500
                Write-Progress -Activity "DCU Installation" -Completed
                return $dcuPath
            } else {
                throw "Installation completed but DCU not found"
            }
        } else {
            throw "Installation failed with exit code: $($process.ExitCode)"
        }
    }
    catch {
        Write-ColorOutput "DCU installation failed: $_" "Error"
        throw
    }
}

function Set-DcuConfiguration {
    param([string]$DcuPath)
    
    Write-ColorOutput "Configuring Dell Command | Update..." "Header"
    Show-ProgressBar -Activity "DCU Configuration" -Status "Configuring settings..." -PercentComplete 30
    
    try {
        # Enable Advanced Driver Restore
        & $DcuPath /configure -advancedDriverRestore=enable | Out-Null
        Write-ColorOutput "Advanced Driver Restore enabled" "Success"
        
        Show-ProgressBar -Activity "DCU Configuration" -Status "Configuring driver source..." -PercentComplete 60
        
        # Configure to download driver library from Dell support site
        & $DcuPath /configure -downloadLibrary=enable | Out-Null
        Write-ColorOutput "Driver library download from Dell enabled" "Success"
        
        Show-ProgressBar -Activity "DCU Configuration" -Status "Complete" -PercentComplete 100
        Start-Sleep -Milliseconds 500
        Write-Progress -Activity "DCU Configuration" -Completed
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-ColorOutput "Configuration failed: $errorMsg" "Warning"
    }
}

function Set-OfflineDriverPack {
    param(
        [string]$DcuPath,
        [string]$PackPath
    )
    
    if (!(Test-Path $PackPath)) {
        Write-ColorOutput "Driver pack not found: $PackPath" "Error"
        return
    }
    
    Write-ColorOutput "Configuring offline driver pack..." "Info"
    Show-ProgressBar -Activity "Driver Pack Configuration" -Status "Setting driver library location..." -PercentComplete 50
    
    try {
        & $DcuPath /configure -driverLibraryLocation="$PackPath" | Out-Null
        Write-ColorOutput "Offline driver pack configured: $PackPath" "Success"
        
        Show-ProgressBar -Activity "Driver Pack Configuration" -Status "Complete" -PercentComplete 100
        Start-Sleep -Milliseconds 500
        Write-Progress -Activity "Driver Pack Configuration" -Completed
        return
    }
    catch {
        Write-ColorOutput "Failed to configure driver pack: $_" "Error"
        return
    }
}

function Install-Drivers {
    param(
        [string]$DcuPath,
        [string]$PackPath
    )
    
    Write-ColorOutput "Starting driver installation..." "Header"
    Write-Host ""
    
    # Verify pack path exists
    if (!(Test-Path $PackPath)) {
        Write-ColorOutput "Driver pack not found at: $PackPath" "Error"
        Write-ColorOutput "Attempting online installation via Dell Command Update..." "Warning"
        
        $process = Start-Process -FilePath $DcuPath -ArgumentList "/driverinstall" -Wait -PassThru -NoNewWindow
        Write-Host ""
        Write-ColorOutput "Online installation completed with exit code: $($process.ExitCode)" "Info"
        return $process.ExitCode
    }
    
    # Use temp directory on C: drive for extraction
    $extractPath = Join-Path $TempRoot "extracted"
    
    if (Test-Path $extractPath) {
        Write-ColorOutput "Cleaning up previous extraction..." "Info"
        Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    Write-ColorOutput "Extracting driver pack..." "Info"
    Show-ProgressBar -Activity "Driver Installation" -Status "Extracting driver pack..." -PercentComplete 10
    
    try {
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        
        # Use 7-Zip to extract CAB
        $sevenZip = "C:\Program Files\7-Zip\7z.exe"
        if (!(Test-Path $sevenZip)) {
            $sevenZip = "C:\Program Files (x86)\7-Zip\7z.exe"
        }
        
        if (!(Test-Path $sevenZip)) {
            Write-ColorOutput "7-Zip not found. Please install 7-Zip to extract driver packs." "Error"
            throw "7-Zip not found"
        }
        
        $expandResult = Start-Process $sevenZip -ArgumentList "x `"$PackPath`" -o`"$extractPath`" -y" -Wait -PassThru -NoNewWindow
        
        if ($expandResult.ExitCode -ne 0) {
            throw "Failed to extract CAB file (Exit code: $($expandResult.ExitCode))"
        }
        
        Write-ColorOutput "Driver pack extracted successfully" "Success"
        
        # Look for the actual driver installation executable
        $setupFiles = @(
            (Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse | Where-Object { $_.Name -like "*install*" -or $_.Name -like "*setup*" } | Select-Object -First 1),
            (Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse | Select-Object -First 1)
        ) | Where-Object { $_ -ne $null } | Select-Object -First 1
        
        if ($setupFiles) {
            Write-ColorOutput "Found installer: $($setupFiles.Name)" "Info"
            Show-ProgressBar -Activity "Driver Installation" -Status "Installing drivers..." -PercentComplete 50
            
            $installArgs = @("/s", "/silent", "/quiet", "/norestart")
            $process = Start-Process -FilePath $setupFiles.FullName -ArgumentList $installArgs -Wait -PassThru -NoNewWindow
            
            Show-ProgressBar -Activity "Driver Installation" -Status "Complete" -PercentComplete 100
            Start-Sleep -Milliseconds 500
            Write-Progress -Activity "Driver Installation" -Completed
            
            Write-Host ""
            if ($process.ExitCode -eq 0) {
                Write-ColorOutput "Driver installation completed successfully" "Success"
                Write-ColorOutput "A system restart may be required" "Warning"
            } elseif ($process.ExitCode -eq 3010 -or $process.ExitCode -eq 500) {
                Write-ColorOutput "Driver installation completed - restart required" "Warning"
            } else {
                Write-ColorOutput "Driver installation completed with exit code: $($process.ExitCode)" "Warning"
            }
            
            # Cleanup extraction directory
            if (Test-Path $extractPath) {
                Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            
            return $process.ExitCode
        } else {
            # No installer found, use DCU online mode as fallback
            Write-ColorOutput "No installer found in pack, using Dell Command Update online mode..." "Warning"
            Show-ProgressBar -Activity "Driver Installation" -Status "Installing drivers online..." -PercentComplete 50
            
            $process = Start-Process -FilePath $DcuPath -ArgumentList "/driverinstall" -Wait -PassThru -NoNewWindow
            
            Show-ProgressBar -Activity "Driver Installation" -Status "Complete" -PercentComplete 100
            Start-Sleep -Milliseconds 500
            Write-Progress -Activity "Driver Installation" -Completed
            
            Write-Host ""
            Write-ColorOutput "Driver installation completed with exit code: $($process.ExitCode)" "Info"
            
            # Cleanup extraction directory
            if (Test-Path $extractPath) {
                Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            
            return $process.ExitCode
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-ColorOutput "Driver installation failed: $errorMsg" "Error"
        
        # Cleanup extraction directory
        if (Test-Path $extractPath) {
            Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        # Fallback to DCU online mode
        Write-ColorOutput "Attempting online installation via Dell Command Update..." "Warning"
        
        $process = Start-Process -FilePath $DcuPath -ArgumentList "/driverinstall" -Wait -PassThru -NoNewWindow
        Write-Host ""
        Write-ColorOutput "Fallback installation completed with exit code: $($process.ExitCode)" "Info"
        
        return $process.ExitCode
    }
}

function Test-InternetConnection {
    try {
        $result = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction Stop
        return $result
    }
    catch {
        return $false
    }
}

function Wait-ForInternet {
    Write-ColorOutput "No internet connection detected" "Warning"
    Write-Host ""
    Write-Host "  Waiting for internet connection..." -ForegroundColor Yellow
    Write-Host "  Press Ctrl+C to cancel" -ForegroundColor Gray
    Write-Host ""
    
    $attempt = 0
    while (-not (Test-InternetConnection)) {
        $attempt++
        Write-Host "`r  Checking... (attempt $attempt)" -NoNewline -ForegroundColor Cyan
        Start-Sleep -Seconds 5
    }
    
    Write-Host ""
    Write-Host ""
    Write-ColorOutput "Internet connection established!" "Success"
    Start-Sleep -Seconds 2
}

function Install-DriversOffline {
    param(
        [string]$DcuPath,
        [string]$PackPath
    )
    
    Write-ColorOutput "Attempting offline driver installation..." "Header"
    Write-Host ""
    
    if (!(Test-Path $PackPath)) {
        Write-ColorOutput "Driver pack not found at: $PackPath" "Error"
        return $false
    }
    
    # Set state before starting
    Set-InstallationState -State "IN_PROGRESS" -Step "Offline Driver Installation" -Details "Configuring offline mode"
    
    Write-ColorOutput "Configuring offline driver pack..." "Info"
    Show-ProgressBar -Activity "Driver Installation" -Status "Configuring offline mode..." -PercentComplete 20
    
    try {
        # Copy driver pack to temp location if it's on a read-only drive
        $tempPackPath = Join-Path $TempRoot "driverpack_temp"
        if (!(Test-Path $tempPackPath)) {
            New-Item -ItemType Directory -Path $tempPackPath -Force | Out-Null
        }
        
        $tempPackFile = Join-Path $tempPackPath (Split-Path $PackPath -Leaf)
        
        if ($PackPath -notlike "C:\*") {
            Write-ColorOutput "Copying driver pack to temp location..." "Info"
            Set-InstallationState -State "IN_PROGRESS" -Step "Offline Driver Installation" -Details "Copying driver pack"
            Copy-Item $PackPath $tempPackFile -Force
            $packToUse = $tempPackFile
        } else {
            $packToUse = $PackPath
        }
        
        # Try to configure DCU for offline mode (suppress output)
        $null = & $DcuPath /configure -driverLibraryLocation="$packToUse" 2>&1
        
        Write-ColorOutput "Starting offline driver installation..." "Info"
        Write-ColorOutput "This may take several minutes. Please wait..." "Info"
        Write-Host ""
        Write-ColorOutput "WARNING: System may restart automatically or become unresponsive during installation" "Warning"
        Write-Host ""
        
        Set-InstallationState -State "IN_PROGRESS" -Step "Offline Driver Installation" -Details "Installing drivers via DCU (CRITICAL - DO NOT INTERRUPT)"
        
        Show-ProgressBar -Activity "Driver Installation" -Status "Installing drivers (this may take 10-20 minutes)..." -PercentComplete 30
        
        # Run DCU with output redirection to capture it
        $dcuOutput = & $DcuPath /driverinstall 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
        
        # You can optionally log the full output for debugging
        Write-Log "DCU Output: $dcuOutput"
        Write-Log "DCU Exit Code: $exitCode"
        
        # Cleanup temp files
        if ($packToUse -eq $tempPackFile -and (Test-Path $tempPackFile)) {
            Remove-Item $tempPackFile -Force -ErrorAction SilentlyContinue
        }
        
        Show-ProgressBar -Activity "Driver Installation" -Status "Complete" -PercentComplete 100
        Start-Sleep -Milliseconds 500
        Write-Progress -Activity "Driver Installation" -Completed
        
        Write-Host ""
        
        # Check if installation was successful
        if ($exitCode -eq 0) {
            Set-InstallationState -State "COMPLETED" -Step "Offline Driver Installation" -Details "Success"
            Write-ColorOutput "Offline driver installation completed successfully" "Success"
            Write-ColorOutput "A system restart may be required" "Warning"
            return $true
        }
        elseif ($exitCode -eq 3010 -or $exitCode -eq 500) {
            Set-InstallationState -State "COMPLETED" -Step "Offline Driver Installation" -Details "Success - Restart required"
            Write-ColorOutput "Offline driver installation completed - restart required" "Warning"
            return $true
        }
        else {
            Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Exit code: $exitCode"
            Write-ColorOutput "Offline installation failed with exit code: $exitCode" "Warning"
            return $false
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Exception: $errorMsg"
        Write-ColorOutput "Offline installation failed: $errorMsg" "Error"
        return $false
    }
}

function Install-DriversWithLibrary {
    param([string]$DcuPath)
    
    Write-ColorOutput "Attempting driver installation using Dell driver library..." "Header"
    Write-Host ""
    
    # Check internet connectivity
    if (-not (Test-InternetConnection)) {
        Write-ColorOutput "No internet connection - driver library download requires internet" "Warning"
        return $null
    }
    
    Set-InstallationState -State "IN_PROGRESS" -Step "Online Driver Library" -Details "Configuring Advanced Driver Restore"
    
    Write-ColorOutput "Internet connection available" "Success"
    Write-ColorOutput "Configuring Advanced Driver Restore..." "Info"
    Show-ProgressBar -Activity "Driver Library Download" -Status "Configuring..." -PercentComplete 20
    
    try {
        # Enable Advanced Driver Restore and download from Dell (suppress output)
        $null = & $DcuPath /configure -advancedDriverRestore=enable 2>&1
        $null = & $DcuPath /configure -downloadLibrary=enable 2>&1
        
        Write-ColorOutput "Downloading driver library from Dell..." "Info"
        Write-ColorOutput "This may take several minutes depending on your internet speed..." "Info"
        Write-Host ""
        Write-ColorOutput "WARNING: System may restart automatically or become unresponsive during installation" "Warning"
        Write-Host ""
        
        Set-InstallationState -State "IN_PROGRESS" -Step "Online Driver Library" -Details "Downloading and installing drivers (CRITICAL)"
        
        Show-ProgressBar -Activity "Driver Library Download" -Status "Downloading and installing drivers..." -PercentComplete 30
        
        # Use DCU to download the driver library (suppress output)
        $dcuOutput = & $DcuPath /driverinstall 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
        
        # Log for debugging
        Write-Log "DCU Library Output: $dcuOutput"
        Write-Log "DCU Library Exit Code: $exitCode"
        
        Show-ProgressBar -Activity "Driver Library Download" -Status "Complete" -PercentComplete 100
        Start-Sleep -Milliseconds 500
        Write-Progress -Activity "Driver Library Download" -Completed
        
        Write-Host ""
        
        if ($exitCode -eq 0) {
            Set-InstallationState -State "COMPLETED" -Step "Online Driver Library" -Details "Success"
            Write-ColorOutput "Driver installation with library completed successfully" "Success"
            Write-ColorOutput "A system restart may be required" "Warning"
            return $exitCode
        }
        elseif ($exitCode -eq 3010 -or $exitCode -eq 500) {
            Set-InstallationState -State "COMPLETED" -Step "Online Driver Library" -Details "Success - Restart required"
            Write-ColorOutput "Driver installation with library completed - restart required" "Warning"
            return $exitCode
        }
        else {
            Set-InstallationState -State "FAILED" -Step "Online Driver Library" -Details "Exit code: $exitCode"
            Write-ColorOutput "Driver library installation completed with exit code: $exitCode" "Warning"
            return $exitCode
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Set-InstallationState -State "FAILED" -Step "Online Driver Library" -Details "Exception: $errorMsg"
        Write-ColorOutput "Driver library installation failed: $errorMsg" "Error"
        return $null
    }
}

function Install-DriversOnline {
    param([string]$DcuPath)
    
    Write-ColorOutput "Attempting online driver installation..." "Header"
    Write-Host ""
    
    # Check internet connectivity
    if (-not (Test-InternetConnection)) {
        Wait-ForInternet
    }
    
    Set-InstallationState -State "IN_PROGRESS" -Step "Online Driver Installation" -Details "Starting online installation"
    
    Write-ColorOutput "Internet connection available" "Success"
    Write-ColorOutput "Starting online driver installation..." "Info"
    Write-ColorOutput "This may take several minutes. Please wait..." "Info"
    Write-Host ""
    Write-ColorOutput "WARNING: System may restart automatically or become unresponsive during installation" "Warning"
    Write-Host ""
    
    Show-ProgressBar -Activity "Driver Installation" -Status "Downloading and installing drivers..." -PercentComplete 30
    
    try {
        Set-InstallationState -State "IN_PROGRESS" -Step "Online Driver Installation" -Details "Installing drivers via DCU (CRITICAL)"
        
        # Run DCU with output suppression
        $dcuOutput = & $DcuPath /driverinstall 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
        
        # Log for debugging
        Write-Log "DCU Online Output: $dcuOutput"
        Write-Log "DCU Online Exit Code: $exitCode"
        
        Show-ProgressBar -Activity "Driver Installation" -Status "Complete" -PercentComplete 100
        Start-Sleep -Milliseconds 500
        Write-Progress -Activity "Driver Installation" -Completed
        
        Write-Host ""
        
        if ($exitCode -eq 0) {
            Set-InstallationState -State "COMPLETED" -Step "Online Driver Installation" -Details "Success"
            Write-ColorOutput "Online driver installation completed successfully" "Success"
            Write-ColorOutput "A system restart may be required" "Warning"
        }
        elseif ($exitCode -eq 3010 -or $exitCode -eq 500) {
            Set-InstallationState -State "COMPLETED" -Step "Online Driver Installation" -Details "Success - Restart required"
            Write-ColorOutput "Online driver installation completed - restart required" "Warning"
        }
        else {
            Set-InstallationState -State "FAILED" -Step "Online Driver Installation" -Details "Exit code: $exitCode"
            Write-ColorOutput "Online installation completed with exit code: $exitCode" "Warning"
        }
        
        return $exitCode
    }
    catch {
        $errorMsg = $_.Exception.Message
        Set-InstallationState -State "FAILED" -Step "Online Driver Installation" -Details "Exception: $errorMsg"
        Write-ColorOutput "Online installation failed: $errorMsg" "Error"
        throw
    }
}

# ============================
# Main Logic
# ============================

function Start-Installation {
    Clear-Host
    Write-Host ""
    Write-Host "===============================================================================" -ForegroundColor Cyan
    Write-Host "                                                                               " -ForegroundColor Cyan
    Write-Host "           Dell Driver Pack Installer - Enhanced Edition                      " -ForegroundColor Cyan
    Write-Host "           Author: Jobin Das (jobindas82@gmail.com)                           " -ForegroundColor Cyan
    Write-Host "           GitHub: https://github.com/jobindas82                              " -ForegroundColor Cyan
    Write-Host "                                                                               " -ForegroundColor Cyan
    Write-Host "===============================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Check for previous crash
    Test-PreviousCrash
    
    # Initialize installation state
    Set-InstallationState -State "IN_PROGRESS" -Step "Initialization" -Details "Starting installation"
    
    Write-Header "INSTALLING PREREQUISITES"
    
    Set-InstallationState -State "IN_PROGRESS" -Step "Prerequisites" -Details "Installing .NET Runtime"
    
    # Install .NET Runtime first
    Install-DotNetRuntime
    Write-Host ""
    
    Write-Header "DETECTING SYSTEM"
    
    # Detect system
    $systemInfo = Get-SystemInfo
    
    if (-not $systemInfo) {
        Write-ColorOutput "System detection failed" "Error"
        throw "Unable to detect system information"
    }
    
    Write-Host ""
    
    Write-Separator "-"
    Write-Host ""
    Write-Host "  SYSTEM INFORMATION" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "  Model        : " -NoNewline -ForegroundColor Gray
    Write-Host $systemInfo.Model -ForegroundColor Green
    Write-Host "  OS           : " -NoNewline -ForegroundColor Gray
    Write-Host $systemInfo.OS -ForegroundColor Green
    Write-Host "  Architecture : " -NoNewline -ForegroundColor Gray
    Write-Host $systemInfo.Architecture -ForegroundColor Green
    Write-Host ""
    Write-Separator "-"
    Write-Host ""
    
    # Find driver pack - check for multiple extensions and model name variations
    $osFolder = "$FilesRoot\$($systemInfo.OS.Replace(' ', '_'))"
    
    $hasDriverPack = $false
    $packPath = $null
    $packFolder = $null
    
    # Function to find driver pack with fuzzy model matching
    function Find-DriverPackFolder {
        param(
            [string]$BaseFolder,
            [string]$ModelName
        )
        
        if (!(Test-Path $BaseFolder)) {
            return $null
        }
        
        # Clean model name for folder comparison
        $cleanModel = $ModelName.Replace(' ', '_')
        
        # Try exact match first
        $exactPath = Join-Path $BaseFolder $cleanModel
        if (Test-Path $exactPath) {
            return $exactPath
        }
        
        # Try partial match - find folders that contain parts of the model name
        $modelParts = $ModelName -split '\s+' | Where-Object { $_.Length -gt 2 }
        $folders = Get-ChildItem -Path $BaseFolder -Directory -ErrorAction SilentlyContinue
        
        foreach ($folder in $folders) {
            $folderName = $folder.Name
            
            # Check if folder name matches any significant part of the model name
            $matchScore = 0
            foreach ($part in $modelParts) {
                if ($folderName -like "*$part*") {
                    $matchScore++
                }
            }
            
            # If we have a reasonable match (at least one part matches)
            if ($matchScore -gt 0) {
                Write-ColorOutput "Found potential match: $folderName (score: $matchScore)" "Info"
                
                # Check if this folder has a pack file
                $testPack = Get-ChildItem -Path $folder.FullName -File | 
                            Where-Object { $_.BaseName -eq "pack" -and $_.Extension -match '\.(cab|exe)$' }
                
                if ($testPack) {
                    return $folder.FullName
                }
            }
        }
        
        return $null
    }
    
    Write-ColorOutput "Searching for driver pack..." "Info"
    $packFolder = Find-DriverPackFolder -BaseFolder $osFolder -ModelName $systemInfo.Model
    
    if ($packFolder) {
        Write-ColorOutput "Driver pack folder found: $packFolder" "Success"
        
        # Look for pack files with common extensions
        $packFiles = Get-ChildItem -Path $packFolder -File | 
                     Where-Object { $_.BaseName -eq "pack" -and $_.Extension -match '\.(cab|exe)$' }
        
        if ($packFiles) {
            $packPath = $packFiles[0].FullName
            $hasDriverPack = $true
            Write-ColorOutput "Driver pack found: $packPath" "Success"
            
            # Check for model info metadata
            $metadataFile = Join-Path $packFolder "model_info.txt"
            if (Test-Path $metadataFile) {
                Write-ColorOutput "Model metadata found" "Info"
                $metadata = Get-Content $metadataFile -Raw
                if ($metadata -match "CatalogModelName=(.+)") {
                    Write-Host "  Catalog Model: $($matches[1])" -ForegroundColor Gray
                }
            }
        } else {
            Write-ColorOutput "Pack file not found in folder" "Warning"
        }
    } else {
        Write-ColorOutput "Driver pack not found for this system" "Warning"
        Write-Host ""
        Write-Host "  System Model  : $($systemInfo.Model)" -ForegroundColor Gray
        Write-Host "  Expected in   : $osFolder" -ForegroundColor Gray
        Write-Host ""
        Write-ColorOutput "Will use online installation mode" "Info"
    }
    
    Write-Host ""
    
    Write-Header "INSTALLING DELL COMMAND UPDATE"
    
    # Install DCU
    $dcuPath = Install-DellCommandUpdate
    Write-Host ""
    
    Write-Header "CONFIGURING DELL COMMAND UPDATE"
    
    # Configure DCU
    Set-DcuConfiguration -DcuPath $dcuPath
    Write-Host ""
    
    Write-Header "INSTALLING DRIVERS"
    
    $exitCode = 0
    
    # Try offline installation first if driver pack is available
    if ($hasDriverPack) {
        Write-ColorOutput "Driver pack available - trying offline installation first" "Info"
        Write-Host ""
        
        $offlineSuccess = Install-DriversOffline -DcuPath $dcuPath -PackPath $packPath
        
        if ($offlineSuccess) {
            $exitCode = 0
        }
        else {
            Write-Host ""
            Write-Separator "-"
            Write-Host ""
            Write-ColorOutput "Offline installation failed" "Warning"
            Write-ColorOutput "Trying Advanced Driver Restore with Dell driver library..." "Info"
            Write-Host ""
            Write-Separator "-"
            Write-Host ""
            Start-Sleep -Seconds 2
            
            # Try downloading driver library from Dell
            $libraryExitCode = Install-DriversWithLibrary -DcuPath $dcuPath
            
            if ($libraryExitCode -ne $null) {
                $exitCode = $libraryExitCode
            }
            else {
                Write-Host ""
                Write-Separator "-"
                Write-Host ""
                Write-ColorOutput "Driver library download failed - falling back to online mode" "Warning"
                Write-Host ""
                Write-Separator "-"
                Write-Host ""
                Start-Sleep -Seconds 2
                
                $exitCode = Install-DriversOnline -DcuPath $dcuPath
            }
        }
    }
    else {
        # No driver pack available, try driver library download first
        Write-ColorOutput "No driver pack available" "Warning"
        Write-ColorOutput "Trying Advanced Driver Restore with Dell driver library..." "Info"
        Write-Host ""
        
        $libraryExitCode = Install-DriversWithLibrary -DcuPath $dcuPath
        
        if ($libraryExitCode -ne $null) {
            $exitCode = $libraryExitCode
        }
        else {
            Write-Host ""
            Write-Separator "-"
            Write-Host ""
            Write-ColorOutput "Driver library download failed - falling back to online mode" "Warning"
            Write-Host ""
            Write-Separator "-"
            Write-Host ""
            Start-Sleep -Seconds 2
            
            $exitCode = Install-DriversOnline -DcuPath $dcuPath
        }
    }
    
    Write-Host ""
    Write-Separator "="
    Write-Host ""
    Write-Host "  [SUCCESS] Installation process completed!" -ForegroundColor Green
    Write-Host ""
    Write-Separator "="
    Write-Host ""
    
    # Clear state on successful completion
    Set-InstallationState -State "COMPLETED" -Step "Installation" -Details "All steps completed successfully"
    
    # Prompt for restart if needed
    if ($exitCode -eq 500 -or $exitCode -eq 0 -or $exitCode -eq 3010) {
        Write-Host ""
        $response = Read-Host "  Do you want to restart the computer now? (Y/N)"
        if ($response -eq 'Y' -or $response -eq 'y') {
            Write-ColorOutput "Restarting computer in 5 seconds..." "Info"
            Clear-InstallationState
            Start-Sleep -Seconds 5
            Restart-Computer -Force
        } else {
            Write-Host ""
            Write-ColorOutput "Please restart your computer to complete the installation" "Warning"
            Write-Host ""
            Clear-InstallationState
        }
    } else {
        Clear-InstallationState
    }
}

# Entry point
try {
    Start-Installation
}
catch {
    $errorMsg = $_.Exception.Message
    Write-ColorOutput "Fatal error: $errorMsg" "Error"
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
finally {
    # Cleanup temp directory (but keep logs and state file)
    if (Test-Path $TempRoot) {
        try {
            # Remove extracted files but keep logs
            $extractPath = Join-Path $TempRoot "extracted"
            if (Test-Path $extractPath) {
                Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            
            $tempPackPath = Join-Path $TempRoot "driverpack_temp"
            if (Test-Path $tempPackPath) {
                Remove-Item $tempPackPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Log "Failed to cleanup temp files: $_"
        }
    }
    
    # Add: Final pause before closing
    Write-Host ""
    Write-Host "Script execution completed. Press any key to exit..." -ForegroundColor Green
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}