# ============================
# Dell Driver Pack Configuration - Sample
# Copy this file to "config.psd1" and update with your settings
# ============================

@{
    # Root folder where all downloaded files will be stored
    FilesRoot      = ".\files"
    
    # Dell Driver Pack Catalog URL (do not change unless Dell updates the location)
    CatalogUrl     = "https://dl.dell.com/catalog/DriverPackCatalog.cab"
    DriverURL      = "https://dl.dell.com"
    CatalogCABFile = ".\files\DriverPackCatalog.cab"
    CatalogXMLFile = ".\files\catalog\DriverPackCatalog.xml"
    
    # Dell Command Update - Latest version URL
    # Check for updates at: https://www.dell.com/support/home/en-us/drivers/driversdetails?driverid=Y5VJV
    DcuUrl         = "https://dl.dell.com/FOLDER13309338M/3/Dell-Command-Update-Application_Y5VJV_WIN64_5.5.0_A00_02.EXE"
    DcuFileName    = "Dell-Command-Update-Setup.exe"
    
    # .NET Desktop Runtime 8.0 LTS - Latest version URL
    # Check for updates at: https://dotnet.microsoft.com/download/dotnet/8.0
    DotNetUrl      = "https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop/8.0.22/windowsdesktop-runtime-8.0.22-win-x64.exe"
    DotNetFileName = "windowsdesktop-runtime-8.0-win-x64.exe"

    # ============================
    # TODO: UPDATE THESE SETTINGS
    # ============================
    
    # Target Dell Models - Add your specific Dell computer models here
    # To find your model name, run: (Get-CimInstance Win32_ComputerSystem).Model
    # Examples:
    #   - "Latitude 5440"
    #   - "Precision 3650 Tower"
    #   - "OptiPlex 7090"
    #   - "XPS 13 9310"
    TargetModels   = @(
        "Latitude 5440",
        "Latitude 5430",
        "Precision 3650 Tower"
        # Add more models here as needed
    )

    # Target Operating System - Specify which OS drivers to download
    # Options: "Windows 11 x64" or "Windows 10 x64"
    TargetOS       = "Windows 11 x64"
}

# ============================
# INSTRUCTIONS:
# ============================
# 1. Copy this file to "config.psd1" in the same directory
# 2. Update TargetModels with your Dell computer models
# 3. Update TargetOS if you need Windows 10 instead of Windows 11
# 4. Run Download.ps1 to download driver packs
# 5. Run Install.ps1 on target computers to install drivers
# ============================
