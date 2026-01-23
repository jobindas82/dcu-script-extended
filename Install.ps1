# ============================
# Dell Driver Pack Installer - Offline Edition
# Features: Auto-detection, Progress bars, Colored output
# Author: Jobin Das (jobindas82)
# GitHub: https://github.com/jobindas82
# Email: hello@jobin-d.com
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

# Change console colors - set background to black
try {
    $host.UI.RawUI.BackgroundColor = "Black"
    $host.UI.RawUI.ForegroundColor = "White"
    Clear-Host
}
catch {
    # Silently continue if color change fails
}

# Load config from files directory
$ScriptRoot = $PSScriptRoot
$FilesRoot = Join-Path $ScriptRoot "files"
$configPath = Join-Path $FilesRoot "config.psd1"

# Check if running from read-only media and warn user
$isReadOnly = $false
try {
    $testFile = Join-Path $ScriptRoot "write_test_$(Get-Random).tmp"
    [System.IO.File]::WriteAllText($testFile, "test")
    Remove-Item $testFile -Force -ErrorAction SilentlyContinue
}
catch {
    $isReadOnly = $true
    Write-Host ""
    Write-Host "===============================================================================" -ForegroundColor Yellow
    Write-Host "                          READ-ONLY MEDIA DETECTED                             " -ForegroundColor Yellow
    Write-Host "===============================================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Running from read-only media (CD/DVD/USB)" -ForegroundColor Cyan
    Write-Host "  All logs and temp files will be stored in: C:\Temp\DellDriverInstaller" -ForegroundColor Cyan
    Write-Host ""
    Start-Sleep -Seconds 2
}

if (!(Test-Path $configPath)) {
    Write-Host "Error: Config file not found at $configPath" -ForegroundColor Red
    Write-Host "Please run Download.ps1 first to create the configuration file." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

$config = Import-PowerShellDataFile -Path $configPath

# Don't override FilesRoot from config - use the actual path (read-only is OK for reading)
$DcuFileName = $config.DcuFileName
$DotNetFileName = $config.DotNetFileName

# ALWAYS use C:\Temp for all write operations
$TempRoot = "C:\Temp\DellDriverInstaller"
if (!(Test-Path $TempRoot)) {
    New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null
}

# All writable files go to temp directory
$LogFile = Join-Path $TempRoot "Install.log"
$StateFile = Join-Path $TempRoot "install_state.json"

# Log that we're using temp directory
Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tScript started from: $ScriptRoot" -ErrorAction SilentlyContinue
Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tRead-only media: $isReadOnly" -ErrorAction SilentlyContinue
Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tTemp directory: $TempRoot" -ErrorAction SilentlyContinue

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

function Select-DcuVersion {
    Write-Host ""
    Write-Host "===============================================================================" -ForegroundColor Cyan
    Write-Host "                         SELECT DCU VERSION                                    " -ForegroundColor Cyan
    Write-Host "===============================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Please select the Dell Command | Update version to install:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    [1] DCU 5.4" -ForegroundColor Cyan
    Write-Host "    [2] DCU 5.5" -ForegroundColor Cyan
    Write-Host "    [3] DCU 5.6 (Recommended)" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Default: DCU 5.6 will be selected automatically in 5 seconds..." -ForegroundColor Gray
    Write-Host ""
    
    $timeout = 5
    $selectedVersion = $null
    
    for ($i = $timeout; $i -gt 0; $i--) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.KeyChar -eq '1') {
                $selectedVersion = "5.4"
                break
            }
            elseif ($key.KeyChar -eq '2') {
                $selectedVersion = "5.5"
                break
            }
            elseif ($key.KeyChar -eq '3') {
                $selectedVersion = "5.6"
                break
            }
        }
        
        Write-Host "`r  Time remaining: $i seconds... " -NoNewline -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }
    
    if ($null -eq $selectedVersion) {
        $selectedVersion = "5.6"
        Write-Host "`r                                    " -NoNewline
        Write-Host "`r  No selection made. Using default: DCU 5.6" -ForegroundColor Green
    }
    else {
        Write-Host "`r                                    " -NoNewline
        Write-Host "`r  Selected: DCU $selectedVersion" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "===============================================================================" -ForegroundColor Cyan
    Write-Host ""
    Start-Sleep -Seconds 1
    
    return $selectedVersion
}

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
        State        = $State
        Step         = $Step
        Details      = $Details
        Timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
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
            Model        = $model
            OS           = $osName
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
        }
        else {
            throw "Installation failed with exit code: $($process.ExitCode)"
        }
    }
    catch {
        Write-ColorOutput ".NET Runtime installation failed: $_" "Error"
        throw
    }
}

function Install-DellCommandUpdate {
    param([string]$PreferredVersion = "5.6")
    
    Write-ColorOutput "Checking for Dell Command | Update..." "Header"
    Show-ProgressBar -Activity "DCU Installation" -Status "Checking installation..." -PercentComplete 10
    
    $dcuPath = Get-DcuPath
    
    # Define minimum required version based on preference
    $minimumVersion = if ($PreferredVersion -eq "5.4") { 
        [version]"5.4.0" 
    }
    elseif ($PreferredVersion -eq "5.5") { 
        [version]"5.5.0" 
    }
    else { 
        [version]"5.6.0" 
    }
    $needsReinstall = $false
    
    if ($dcuPath) {
        Write-ColorOutput "Dell Command | Update found" "Info"
        
        # Check version
        $installedVersion = Get-InstalledDcuVersion
        
        if ($installedVersion) {
            Write-ColorOutput "Installed version: $installedVersion" "Info"
            Write-Log "DCU installed version: $installedVersion"
            
            # Check if installed version matches preferred version
            $installedMajorMinor = "$($installedVersion.Major).$($installedVersion.Minor)"
            
            if ($installedMajorMinor -ne $PreferredVersion) {
                Write-ColorOutput "Installed version ($installedMajorMinor) does not match preferred version ($PreferredVersion)" "Warning"
                $needsReinstall = $true
            }
            elseif ($installedVersion -lt $minimumVersion) {
                Write-ColorOutput "Installed version is outdated (minimum required: $minimumVersion)" "Warning"
                $needsReinstall = $true
            }
            else {
                Write-ColorOutput "Dell Command | Update version $PreferredVersion is installed" "Success"
                Show-ProgressBar -Activity "DCU Installation" -Status "Already installed" -PercentComplete 100
                Start-Sleep -Milliseconds 500
                Write-Progress -Activity "DCU Installation" -Completed
                return $dcuPath
            }
        }
        else {
            Write-ColorOutput "Could not determine DCU version, will reinstall with version $PreferredVersion" "Warning"
            $needsReinstall = $true
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
    
    # Determine which installer to use based on preferred version
    $dcuInstallerName = if ($PreferredVersion -eq "5.4") {
        # Look for DCU 5.4 installer in config or use default name
        if ($config.DcuFileName54) {
            $config.DcuFileName54
        }
        else {
            "Dell-Command-Update_54.exe"
        }
    }
    elseif ($PreferredVersion -eq "5.6") {
        # Use DCU 5.6
        if ($config.DcuFileName56) {
            $config.DcuFileName56
        }
        else {
            "Dell-Command-Update_56.exe"
        }
    }
    else {
        # Use DCU 5.5 (from config)
        $DcuFileName
    }
    
    # Install new version
    $setupPath = Join-Path $FilesRoot "dcu\$dcuInstallerName"
    if (!(Test-Path $setupPath)) {
        Write-ColorOutput "DCU $PreferredVersion installer not found at $setupPath" "Error"
        Write-ColorOutput "Please run Download.ps1 first" "Warning"
        throw "DCU installer not found"
    }
    
    Write-ColorOutput "Installing Dell Command | Update $PreferredVersion..." "Info"
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
            }
            else {
                throw "Installation completed but DCU not found"
            }
        }
        else {
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
        Write-Host ""
        & $DcuPath /configure -advancedDriverRestore=enable
        Write-Host ""
        
        Write-ColorOutput "Advanced Driver Restore enabled" "Success"
        
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
        Write-Host ""
        & $DcuPath /configure -driverLibraryLocation="$PackPath"
        Write-Host ""
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

function Install-DriversOffline {
    param(
        [string]$DcuPath,
        [string]$PackPath
    )
    
    Write-ColorOutput "Attempting offline driver installation..." "Header"
    Write-Host ""
    
    if (!(Test-Path $PackPath)) {
        Write-ColorOutput "Driver pack not found at: $PackPath" "Error"
        throw "Driver pack is required for offline installation"
    }
    
    # Set state before starting
    Set-InstallationState -State "IN_PROGRESS" -Step "Offline Driver Installation" -Details "Configuring offline mode"
    
    Write-ColorOutput "Configuring offline driver pack..." "Info"
    Show-ProgressBar -Activity "Driver Installation" -Status "Configuring offline mode..." -PercentComplete 20
    
    try {
        # Configure DCU for offline mode using the original pack path directly
        Write-Host ""
        & $DcuPath /configure -driverLibraryLocation="$PackPath"
        Write-Host ""
        
        Write-ColorOutput "Starting offline driver installation..." "Info"
        Write-ColorOutput "This may take several minutes. Please wait..." "Info"
        Write-Host ""
        Write-ColorOutput "WARNING: System may restart automatically or become unresponsive during installation" "Warning"
        Write-Host ""
        
        Set-InstallationState -State "IN_PROGRESS" -Step "Offline Driver Installation" -Details "Installing drivers via DCU (CRITICAL - DO NOT INTERRUPT)"
        
        Show-ProgressBar -Activity "Driver Installation" -Status "Installing drivers (this may take 10-20 minutes)..." -PercentComplete 30
        Write-Progress -Activity "Driver Installation" -Completed
        
        Write-Host ""
        Write-Separator "-"
        Write-Host "  DCU DRIVER INSTALLATION OUTPUT" -ForegroundColor Cyan
        Write-Separator "-"
        Write-Host ""
        
        # Run DCU - show all output
        & $DcuPath /driverinstall
        $exitCode = $LASTEXITCODE
        
        Write-Host ""
        Write-Separator "-"
        Write-Host ""
        
        
        Write-Log "DCU Exit Code: $exitCode"
        
        # Comprehensive DCU CLI exit code handling
        switch ($exitCode) {
            # Success codes
            0 {
                Set-InstallationState -State "COMPLETED" -Step "Offline Driver Installation" -Details "Success"
                Write-ColorOutput "Offline driver installation completed successfully" "Success"
                Write-ColorOutput "A system restart may be required" "Warning"
                return $true
            }
            1 {
                Set-InstallationState -State "COMPLETED" -Step "Offline Driver Installation" -Details "Success - Reboot required"
                Write-ColorOutput "Driver installation completed successfully" "Success"
                Write-ColorOutput "A system reboot is required to complete the operation" "Warning"
                return $true
            }
            
            # Generic application errors (2-8)
            2 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Unknown application error"
                Write-ColorOutput "An unknown application error has occurred" "Error"
                return $false
            }
            3 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Not a Dell system"
                Write-ColorOutput "The current system manufacturer is not Dell" "Error"
                Write-ColorOutput "Dell Command | Update can only be run on Dell systems" "Error"
                return $false
            }
            4 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Insufficient privileges"
                Write-ColorOutput "The CLI was not launched with administrative privilege" "Error"
                Write-ColorOutput "Please run this script as Administrator" "Warning"
                return $false
            }
            5 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Reboot pending"
                Write-ColorOutput "A reboot is pending from a previous operation" "Error"
                Write-ColorOutput "Please reboot the system and try again" "Warning"
                return $false
            }
            6 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Another instance running"
                Write-ColorOutput "Another instance of Dell Command | Update is already running" "Error"
                Write-ColorOutput "Close any running instance of DCU UI or CLI and retry" "Warning"
                return $false
            }
            7 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "System model not supported"
                Write-ColorOutput "The application does not support the current system model" "Error"
                Write-ColorOutput "Contact your administrator if the current system model is not supported" "Warning"
                return $false
            }
            8 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "No update filters"
                Write-ColorOutput "No update filters have been applied or configured" "Error"
                return $false
            }
            
            # Scan operation errors (500-503)
            500 {
                Set-InstallationState -State "COMPLETED" -Step "Offline Driver Installation" -Details "No updates found"
                Write-ColorOutput "No updates were found for the system" "Warning"
                Write-ColorOutput "The system is up to date or no updates match the provided filters" "Info"
                return $true
            }
            501 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Error determining updates"
                Write-ColorOutput "An error occurred while determining the available updates" "Error"
                Write-ColorOutput "Please retry the operation" "Warning"
                return $false
            }
            502 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Scan operation canceled"
                Write-ColorOutput "The scan operation was canceled" "Warning"
                return $false
            }
            503 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Download error during scan"
                Write-ColorOutput "An error occurred while downloading a file during the scan operation" "Error"
                Write-ColorOutput "Check your network connection and ensure there is Internet connectivity" "Warning"
                return $false
            }
            
            # Apply updates errors (1000-1002)
            1000 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Error retrieving apply updates result"
                Write-ColorOutput "An error occurred when retrieving the result of the apply updates operation" "Error"
                Write-ColorOutput "Please retry the operation" "Warning"
                return $false
            }
            1001 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Apply updates canceled"
                Write-ColorOutput "The apply updates operation was canceled" "Warning"
                return $false
            }
            1002 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Download error during apply"
                Write-ColorOutput "An error occurred while downloading a file during the apply updates operation" "Error"
                Write-ColorOutput "Check your network connection and ensure there is Internet connectivity" "Warning"
                return $false
            }
            
            # Driver install errors (2000-2007)
            2000 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Error retrieving driver install result"
                Write-ColorOutput "An error occurred when retrieving the result of the Advanced Driver Restore operation" "Error"
                Write-ColorOutput "Please retry the operation" "Warning"
                return $false
            }
            2001 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Advanced Driver Restore failed"
                Write-ColorOutput "The Advanced Driver Restore process failed" "Error"
                Write-ColorOutput "Please retry the operation" "Warning"
                return $false
            }
            2002 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Multiple driver CABs provided"
                Write-ColorOutput "Multiple driver CABs were provided for the Advanced Driver Restore operation" "Error"
                Write-ColorOutput "Ensure that you provide only one driver CAB file" "Warning"
                return $false
            }
            2003 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Invalid driver CAB path"
                Write-ColorOutput "An invalid path for the driver CAB was provided" "Error"
                Write-ColorOutput "Ensure that the file path exists and has valid permissions" "Warning"
                return $false
            }
            2004 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Driver install canceled"
                Write-ColorOutput "The driver install operation was canceled" "Warning"
                return $false
            }
            2005 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Download error during driver install"
                Write-ColorOutput "An error occurred while downloading a file during the driver install operation" "Error"
                Write-ColorOutput "Check your network connection and ensure there is Internet connectivity" "Warning"
                return $false
            }
            2006 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Advanced Driver Restore disabled"
                Write-ColorOutput "The Advanced Driver Restore feature is disabled" "Error"
                Write-ColorOutput "Enable the feature using: /configure -advancedDriverRestore=enable" "Warning"
                return $false
            }
            2007 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Advanced Driver Restore not supported"
                Write-ColorOutput "The Advanced Driver Restore feature is not supported" "Error"
                Write-ColorOutput "Disable FIPS mode on the system" "Warning"
                return $false
            }
            
            # Dell Client Management Service errors (3000-3005)
            3000 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Service not running"
                Write-ColorOutput "The Dell Client Management Service is not running" "Error"
                Write-ColorOutput "Start the Dell Client Management Service in Windows services" "Warning"
                return $false
            }
            3001 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Service not installed"
                Write-ColorOutput "The Dell Client Management Service is not installed" "Error"
                Write-ColorOutput "Download and install the Dell Client Management Service from Dell support site" "Warning"
                return $false
            }
            3002 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Service disabled"
                Write-ColorOutput "The Dell Client Management Service is disabled" "Error"
                Write-ColorOutput "Enable the Dell Client Management Service from Windows services" "Warning"
                return $false
            }
            3003 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Service busy"
                Write-ColorOutput "The Dell Client Management Service is busy" "Error"
                Write-ColorOutput "Wait until the service is available to process new requests" "Warning"
                return $false
            }
            3004 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Service self-update in progress"
                Write-ColorOutput "The Dell Client Management Service has initiated a self-update install" "Error"
                Write-ColorOutput "Wait until the service is available to process new requests" "Warning"
                return $false
            }
            3005 {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Service installing updates"
                Write-ColorOutput "The Dell Client Management Service is installing pending updates" "Error"
                Write-ColorOutput "Wait until the service is available to process new requests" "Warning"
                return $false
            }
            
            # Any other exit code
            default {
                Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Exit code: $exitCode"
                Write-ColorOutput "Offline installation failed with exit code: $exitCode" "Error"
                Write-ColorOutput "Please check the Dell Command | Update documentation for this error code" "Warning"
                return $false
            }
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Set-InstallationState -State "FAILED" -Step "Offline Driver Installation" -Details "Exception: $errorMsg"
        Write-ColorOutput "Offline installation failed: $errorMsg" "Error"
        return $false
    }
}

function Invoke-CustomConfiguration {
    Write-Header "APPLYING CUSTOM CONFIGURATION"
    
    Write-ColorOutput "Running custom configuration tasks..." "Info"
    Write-Host ""
    
    try {
        # Get computer name
        $computerName = $env:COMPUTERNAME
        Write-ColorOutput "Computer Name: $computerName" "Info"
        
        # Extract location prefix (first 2 characters)
        if ($computerName.Length -ge 2) {
            $locationPrefix = $computerName.Substring(0, 2).ToUpper()
            Write-ColorOutput "Location Prefix: $locationPrefix" "Info"
            
            # Map location prefix to timezone
            $timezoneMap = @{
                'CA' = 'Mountain Standard Time'      # Calgary
                'VA' = 'Pacific Standard Time'       # Vancouver
            }
            
            if ($timezoneMap.ContainsKey($locationPrefix)) {
                $targetTimezone = $timezoneMap[$locationPrefix]
                Write-ColorOutput "Target Timezone: $targetTimezone" "Info"
                
                # Get current timezone
                $currentTimezone = (Get-TimeZone).Id
                Write-ColorOutput "Current Timezone: $currentTimezone" "Info"
                
                if ($currentTimezone -ne $targetTimezone) {
                    Write-ColorOutput "Setting timezone to $targetTimezone..." "Info"
                    
                    try {
                        Set-TimeZone -Id $targetTimezone -ErrorAction Stop
                        Write-ColorOutput "Timezone changed successfully to $targetTimezone" "Success"
                        Write-Log "Timezone changed from $currentTimezone to $targetTimezone"
                    }
                    catch {
                        Write-ColorOutput "Failed to set timezone: $($_.Exception.Message)" "Error"
                        Write-Log "Timezone change failed: $($_.Exception.Message)"
                    }
                }
                else {
                    Write-ColorOutput "Timezone already set to $targetTimezone" "Success"
                }
            }
            else {
                Write-ColorOutput "No timezone mapping found for location prefix: $locationPrefix" "Warning"
                Write-ColorOutput "Skipping timezone configuration" "Info"
            }
        }
        else {
            Write-ColorOutput "Computer name too short to extract location prefix" "Warning"
            Write-ColorOutput "Skipping timezone configuration" "Info"
        }
        
        Write-Host ""
        Write-ColorOutput "Custom configuration completed" "Success"
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-ColorOutput "Custom configuration failed: $errorMsg" "Error"
        Write-Log "Custom configuration error: $errorMsg"
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
    Write-Host "           Dell Driver Pack Installer - Offline Edition                       " -ForegroundColor Cyan
    Write-Host "           Author: Jobin Das (hello@jobin-d.com)                           " -ForegroundColor Cyan
    Write-Host "           GitHub: https://github.com/jobindas82                              " -ForegroundColor Cyan
    Write-Host "                                                                               " -ForegroundColor Cyan
    Write-Host "===============================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Check for previous crash
    Test-PreviousCrash
    
    # Select DCU version at the start
    $dcuVersion = Select-DcuVersion
    Write-Log "User selected DCU version: $dcuVersion"
    
    # Initialize installation state
    Set-InstallationState -State "IN_PROGRESS" -Step "Initialization" -Details "Starting installation with DCU $dcuVersion"
    
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
    Write-Host "  DCU Version  : " -NoNewline -ForegroundColor Gray
    Write-Host $dcuVersion -ForegroundColor Green
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
        }
        else {
            Write-ColorOutput "Pack file not found in folder" "Warning"
        }
    }
    
    if (-not $hasDriverPack) {
        Write-ColorOutput "Driver pack not found for this system" "Error"
        Write-Host ""
        Write-Host "  System Model  : $($systemInfo.Model)" -ForegroundColor Gray
        Write-Host "  Expected in   : $osFolder" -ForegroundColor Gray
        Write-Host ""
        Write-ColorOutput "Offline installation requires a driver pack" "Error"
        Write-ColorOutput "Please run Download.ps1 to download the driver pack first" "Warning"
        throw "Driver pack not found"
    }
    
    Write-Host ""
    
    Write-Header "INSTALLING DELL COMMAND UPDATE"
    
    # Install DCU with selected version
    $dcuPath = Install-DellCommandUpdate -PreferredVersion $dcuVersion
    Write-Host ""
    
    Write-Header "CONFIGURING DELL COMMAND UPDATE"
    
    # Configure DCU
    Set-DcuConfiguration -DcuPath $dcuPath
    Write-Host ""
    
    Write-Header "INSTALLING DRIVERS"
    
    # Offline installation only
    Write-ColorOutput "Starting offline driver installation" "Info"
    Write-Host ""
    
    $offlineSuccess = Install-DriversOffline -DcuPath $dcuPath -PackPath $packPath
    
    if (-not $offlineSuccess) {
        Write-ColorOutput "Offline installation failed" "Error"
        throw "Driver installation failed"
    }
    
    Write-Host ""
    
    # Run custom configuration
    Invoke-CustomConfiguration
    
    Write-Host ""
    Write-Separator "="
    Write-Host ""
    Write-Host "  [SUCCESS] Installation process completed!" -ForegroundColor Green
    Write-Host ""
    Write-Separator "="
    Write-Host ""
    
    # Clear state on successful completion
    Set-InstallationState -State "COMPLETED" -Step "Installation" -Details "All steps completed successfully"
    
    # Prompt for restart
    Write-Host ""
    $response = Read-Host "  Do you want to restart the computer now? (Y/N)"
    if ($response -eq 'Y' -or $response -eq 'y') {
        Write-ColorOutput "Restarting computer in 5 seconds..." "Info"
        Clear-InstallationState
        Start-Sleep -Seconds 5
        Restart-Computer -Force
    }
    else {
        Write-Host ""
        Write-ColorOutput "Please restart your computer to complete the installation" "Warning"
        Write-Host ""
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