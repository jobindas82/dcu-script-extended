# Dell Driver Pack Downloader & Installer

**Author:** Jobin Das (jobindas82)  
**GitHub:** https://github.com/jobindas82  
**Email:** hello@jobin-d.com

Automated solution for downloading and installing Dell driver packs for multiple computer models.

## Features

- ✅ Automated driver pack downloads from Dell
- ✅ Support for multiple Dell models
- ✅ MD5 hash verification
- ✅ Version tracking and caching
- ✅ Automatic Dell Command Update installation
- ✅ .NET Desktop Runtime 8.0 installation
- ✅ Offline driver installation support
- ✅ Progress bars and colored output
- ✅ Comprehensive logging

## Requirements

- Windows 10/11
- PowerShell 5.1 or later
- Administrator rights (for Install.ps1)

## Quick Start

### 1. Initial Setup

```powershell
# Run Download.ps1 for the first time
.\Download.ps1
```

The script will create a `config.psd1` file and open it in Notepad. Update the configuration with your Dell models.

### 2. Find Your Dell Model Name

```powershell
(Get-CimInstance Win32_ComputerSystem).Model
```

### 3. Update Configuration

Edit `config.psd1`:

```powershell
TargetModels = @(
    "Latitude 5440",
    "Precision 3650 Tower",
    "Your Model Name Here"
)

TargetOS = "Windows 11 x64"  # or "Windows 10 x64"
```

### 4. Download Driver Packs

```powershell
.\Download.ps1
```

### 5. Install Drivers (Run as Administrator)

```powershell
.\Install.ps1
```

## File Structure

```
.
├── Download.ps1           # Main download script
├── Install.ps1            # Driver installation script
├── config.psd1            # Your configuration
├── README.md              # This file
└── files/                 # Downloaded files (auto-created)
    ├── dcu/               # Dell Command Update installer
    ├── dotnet/            # .NET Desktop Runtime installer
    ├── catalog/           # Driver pack catalog
    ├── Download.log       # Download log
    ├── Install.log        # Installation log
    └── Windows_11_x64/    # Driver packs by OS
        ├── Latitude_5440/
        │   └── pack.cab
        └── Precision_3650_Tower/
            └── pack.cab
```

## Configuration Options

### TargetModels
Array of Dell model names to download driver packs for.

**Example:**
```powershell
TargetModels = @(
    "Latitude 5440",
    "Latitude 5430",
    "Precision 5690",
    "OptiPlex 7090"
)
```

### TargetOS
Operating system to download drivers for.

**Options:**
- `"Windows 11 x64"`
- `"Windows 10 x64"`

### DcuUrl / DotNetUrl
URLs for Dell Command Update and .NET Runtime. Update these if newer versions are available.

## Common Issues

### Config File Not Found
Run `Download.ps1` once to create the default config, then update it.

### Driver Pack Not Found During Install
Ensure you've run `Download.ps1` first and the model name matches exactly.

### Hash Mismatch
The downloaded file may be corrupted. Delete the file and run `Download.ps1` again.

## Updating Driver Packs

Simply run `Download.ps1` again. The script will:
- Check for updated driver packs
- Only download new/updated packs
- Skip already downloaded packs (based on MD5 hash)

## Logs

- **Download.log**: All download activity
- **Install.log**: All installation activity

Both logs are in the `files/` directory.

## Advanced Usage

### Offline Installation
1. Copy the entire `files/` folder to a USB drive
2. Run `Install.ps1` from the USB drive on target computers

### Multiple OS Support
Configure multiple target OS versions:
```powershell
TargetOS = @("Windows 11 x64", "Windows 10 x64")
```

## Support

For issues, questions, or contributions:
- GitHub: https://github.com/jobindas82
- Email: hello@jobin-d.com

## License

MIT License

This project is provided as-is for use with Dell computers. Dell, Dell Command Update, and related trademarks are property of Dell Inc.
