# Loxone Config Version Launcher

# Function to get config file path (using user's temp directory instead of system root)
function Get-ConfigFilePath {
    # Use user's temp directory instead of system root to avoid permission issues
    $tempPath = [System.IO.Path]::GetTempPath()
    return Join-Path $tempPath "LoxoneLauncher.config"
}

# Function to read config file
function Read-ConfigFile {
    $configPath = Get-ConfigFilePath
    
    if (Test-Path $configPath) {
        try {
            $jsonConfig = Get-Content $configPath -Raw | ConvertFrom-Json
            # Convert PSCustomObject to hashtable for consistent handling
            $config = @{
                LoxonePath = $jsonConfig.LoxonePath
                CreateDesktopShortcut = $jsonConfig.CreateDesktopShortcut
                LastUpdated = $jsonConfig.LastUpdated
            }
            return $config
        }
        catch {
            Write-Warning "Could not read config file. Using default settings."
            return $null
        }
    }
    
    return $null
}

# Function to write config file
function Write-ConfigFile {
    param($config)
    
    $configPath = Get-ConfigFilePath
    
    # Ensure all required properties exist
    $configToSave = @{
        LoxonePath = if ($config.LoxonePath) { $config.LoxonePath } else { $null }
        CreateDesktopShortcut = if ($config.ContainsKey('CreateDesktopShortcut')) { $config.CreateDesktopShortcut } else { $null }
        LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    try {
        $configToSave | ConvertTo-Json | Set-Content -Path $configPath -ErrorAction Stop
        Write-Host "Configuration saved to $configPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Could not save configuration file: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Config will not be persistent between sessions." -ForegroundColor Yellow
    }
}

# Function to get current script directory
function Get-ScriptDirectory {
    return Split-Path -Parent $PSCommandPath
}

# Function to find Loxone Config executable for icon extraction
function Find-LoxoneConfigExe {
    param($basePath)
    
    $loxoneFolders = Get-ChildItem -Path $basePath -Directory | 
                     Where-Object { $_.Name -like "*config*" -or $_.Name -like "*loxone*" }
    
    foreach ($folder in $loxoneFolders) {
        $possibleExes = @("LoxoneConfig.exe", "Loxone Config.exe", "Config.exe")
        
        foreach ($exe in $possibleExes) {
            $exePath = Join-Path $folder.FullName $exe
            if (Test-Path $exePath) {
                return $exePath
            }
        }
        
        # Search subdirectories
        $foundExe = Get-ChildItem -Path $folder.FullName -Filter "*.exe" -Recurse | 
                    Where-Object { $_.Name -like "*config*" -or $_.Name -like "*loxone*" } | 
                    Select-Object -First 1
        
        if ($foundExe) {
            return $foundExe.FullName
        }
    }
    
    return $null
}

# Function to create desktop shortcut
function New-DesktopShortcut {
    param($config, $versionList)
    
    try {
        $scriptPath = $PSCommandPath
        $scriptDirectory = Get-ScriptDirectory
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $shortcutPath = Join-Path $desktopPath "Loxone Config Launcher.lnk"
        
        # Create WScript.Shell object
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        
        # Set shortcut properties
        $Shortcut.TargetPath = "powershell.exe"
        $Shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
        $Shortcut.WorkingDirectory = $scriptDirectory
        $Shortcut.Description = "Loxone Config Version Launcher"
        
        # Use the latest version's executable for the icon (first item in sorted list)
        $latestVersionExe = $null
        if ($versionList -and $versionList.Count -gt 0) {
            $latestVersionExe = $versionList[0].Executable
        }
        
        if ($latestVersionExe -and (Test-Path $latestVersionExe)) {
            $Shortcut.IconLocation = "$latestVersionExe,0"
            Write-Host "Using Loxone Config icon from latest version: $latestVersionExe" -ForegroundColor Green
        } else {
            # Fallback to finding any Loxone Config executable
            $loxonePath = if ($config.LoxonePath) { $config.LoxonePath } else { "C:\Program Files (x86)\Loxone" }
            $loxoneExe = Find-LoxoneConfigExe -basePath $loxonePath
            
            if ($loxoneExe -and (Test-Path $loxoneExe)) {
                $Shortcut.IconLocation = "$loxoneExe,0"
                Write-Host "Using Loxone Config icon from: $loxoneExe" -ForegroundColor Green
            } else {
                # Fallback to PowerShell icon
                $Shortcut.IconLocation = "powershell.exe,0"
                Write-Host "Using PowerShell icon (Loxone Config icon not found)" -ForegroundColor Yellow
            }
        }
        
        # Save the shortcut
        $Shortcut.Save()
        
        # Clean up COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Shortcut) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "Desktop shortcut created successfully: $shortcutPath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error creating desktop shortcut: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to ask user about desktop shortcut creation
function Ask-CreateShortcut {
    Write-Host ""
    Write-Host "=== Desktop Shortcut Creation ===" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Would you like to create a desktop shortcut for easy access to this launcher?" -ForegroundColor Yellow
    
    do {
        $response = Read-Host "Create desktop shortcut? (y/N)"
        
        if ([string]::IsNullOrWhiteSpace($response) -or $response -eq 'n' -or $response -eq 'N') {
            return $false
        }
        elseif ($response -eq 'y' -or $response -eq 'Y') {
            return $true
        }
        else {
            Write-Host "Please enter 'y' for yes or 'n' for no (or press Enter for no)." -ForegroundColor Red
        }
    } while ($true)
}

# Function to get Loxone installation path with notification and error handling
function Get-LoxonePathWithNotification {
    param($config)

    $defaultPath = "C:\Program Files (x86)\Loxone"

    if ($config.LoxonePath) {
        Write-Host "Found configuration file - checking path: $($config.LoxonePath)" -ForegroundColor Cyan

        if (Test-Path $config.LoxonePath -PathType Container) {
            $resolved  = (Resolve-Path $config.LoxonePath).Path
            $driveRoot = [System.IO.Path]::GetPathRoot($resolved)
            if ($resolved -eq $driveRoot) {
                Write-Host "Configured path is a drive root - not allowed." -ForegroundColor Red
                Write-Host "Invalid path found in configuration. Please specify a new path." -ForegroundColor Yellow
                return Get-ValidLoxonePath
            }
            else {
                Write-Host "Configuration path is valid." -ForegroundColor Green
                return $resolved
            }
        }
        else {
            Write-Host "Configuration path does not exist: $($config.LoxonePath)" -ForegroundColor Red
            Write-Host "Invalid path found in configuration. Please specify a new path." -ForegroundColor Yellow
            return Get-ValidLoxonePath
        }
    }

    if (Test-Path $defaultPath) {
        Write-Host "Using default path: $defaultPath" -ForegroundColor Green
        return $defaultPath
    }

    Write-Host "No valid Loxone installation found." -ForegroundColor Yellow
    return Get-ValidLoxonePath
}

# Ask repeatedly until the user enters a *non-root* directory
function Get-ValidLoxonePath {
    do {
        $newPath = Read-Host "Enter the Loxone installation path"

        # ---------- basic checks ----------
        if (-not (Test-Path $newPath)) {
            Write-Host "Path does not exist: $newPath" -ForegroundColor Red
            continue
        }
        if (-not ((Get-Item $newPath) -is [System.IO.DirectoryInfo])) {
            Write-Host "Path exists but is not a directory: $newPath" -ForegroundColor Red
            continue
        }

        # ---------- extra rule: not a drive root ----------
        $resolved     = (Resolve-Path -LiteralPath $newPath).Path           # full, canonical form
        $driveRoot    = [System.IO.Path]::GetPathRoot($resolved)            # e.g. 'C:\'
        if ($resolved -eq $driveRoot) {
            Write-Host "A drive root ($resolved) cannot be used." -ForegroundColor Red
            continue
        }

        Write-Host "Path confirmed: $resolved" -ForegroundColor Green
        return $resolved
    } while ($true)
}

# Function to change installation path with enhanced validation
function Set-LoxonePath {
    param($config)
    
    Write-Host ""
    Write-Host "=== Change Loxone Installation Path ===" -ForegroundColor Cyan
    Write-Host ""
    
    $currentPath = if ($config.LoxonePath) { $config.LoxonePath } else { "C:\Program Files (x86)\Loxone (default)" }
    Write-Host "Current path: $currentPath" -ForegroundColor Yellow
    Write-Host ""
    
    do {
        # Use the validation function to get a valid path
        $newPath = Get-ValidLoxonePath
        
        # Check for Loxone Config folders in the new path
        try {
            $loxoneFolders = Get-ChildItem -Path $newPath -Directory -ErrorAction Stop | 
                            Where-Object { $_.Name -like "*config*" -or $_.Name -like "*loxone*" }
        }
        catch {
            Write-Host "Error scanning directory: $($_.Exception.Message)" -ForegroundColor Red
            continue
        }
        
        if ($loxoneFolders.Count -eq 0) {
            Write-Host ""
            Write-Host "No Loxone Config versions were found in `"$newPath`"." -ForegroundColor Yellow
            Write-Host "If this directory really is the correct installation folder, type Y to save it anyway." -ForegroundColor Yellow
            Write-Host "Otherwise press Enter and specify another path." -ForegroundColor Yellow
            
            $confirmSave = Read-Host "Save this path? (y/N)"
            
            if ($confirmSave -notmatch '^[yY]$') {
                Write-Host "Path not saved. Please enter a new path." -ForegroundColor Cyan
                continue
            }
        }
        
        # If we reach here, either Loxone folders were found OR user confirmed to save anyway
        break
        
    } while ($true)
    
    # Save the new path
    Write-Host "Path saved: $newPath" -ForegroundColor Green
    $config.LoxonePath = $newPath
    $config.LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-ConfigFile -config $config
    return $newPath
}

# Function to extract version from unins000.dat file
function Get-LoxoneVersion {
    param($folderPath)
    
    $uninsFile = Join-Path $folderPath "unins000.dat"
    
    if (Test-Path $uninsFile) {
        try {
            # Read the file content and search for version pattern
            $content = Get-Content $uninsFile -Raw -ErrorAction SilentlyContinue
            
            # Look for "Loxone Config" followed by version number pattern
            if ($content -match "Loxone Config\s+(\d+\.\d+(?:\.\d+)?(?:\.\d+)?)") {
                return $matches[1]
            }
            
            # Alternative pattern search if first doesn't match
            if ($content -match "(\d+\.\d+(?:\.\d+)?(?:\.\d+)?)") {
                return $matches[1]
            }
        }
        catch {
            Write-Warning "Could not read version from $uninsFile"
        }
    }
    
    return "Unknown"
}

# Function to convert version string to comparable object
function ConvertTo-VersionObject {
    param($versionString)
    
    if ($versionString -eq "Unknown") {
        return [Version]"0.0.0.0"
    }
    
    try {
        # Ensure we have at least 4 parts for proper comparison
        $parts = $versionString.Split('.')
        while ($parts.Count -lt 4) {
            $parts += "0"
        }
        $normalizedVersion = $parts[0..3] -join '.'
        return [Version]$normalizedVersion
    }
    catch {
        return [Version]"0.0.0.0"
    }
}

# Function to find Loxone Config executable
function Find-LoxoneExecutable {
    param($folderPath)
    
    # Common executable names for Loxone Config
    $possibleExes = @("LoxoneConfig.exe", "Loxone Config.exe", "Config.exe")
    
    foreach ($exe in $possibleExes) {
        $exePath = Join-Path $folderPath $exe
        if (Test-Path $exePath) {
            return $exePath
        }
    }
    
    # If not found in root, search subdirectories
    $foundExe = Get-ChildItem -Path $folderPath -Filter "*.exe" -Recurse | 
                Where-Object { $_.Name -like "*config*" -or $_.Name -like "*loxone*" } | 
                Select-Object -First 1
    
    if ($foundExe) {
        return $foundExe.FullName
    }
    
    return $null
}

# Main script execution with error handling
Clear-Host
Write-Host "=== Loxone Config Version Launcher ===" -ForegroundColor Cyan
Write-Host ""

# Read existing config or create new one
$config = Read-ConfigFile
$isFirstRun = $config -eq $null

if ($isFirstRun) {
    # Create new config with default values
    $config = @{
        LoxonePath = $null
        CreateDesktopShortcut = $null
        LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
}

# Get the Loxone installation path with error handling
$basePath = $null
$pathSuccess = $false

try {
    $basePath = Get-LoxonePathWithNotification -config $config
    $pathSuccess = $true
    
    # Update config with the path if it was determined or corrected
    if ($basePath -ne $config.LoxonePath) {
        $config.LoxonePath = $basePath
        Write-ConfigFile -config $config
    }
}
catch {
    Write-Host "Error determining Loxone installation path: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "This usually indicates a problem with the configured path." -ForegroundColor Yellow
}

if (-not $pathSuccess) {
    Write-Host ""
    Write-Host "Would you like to:" -ForegroundColor Cyan
    Write-Host "1. Manually set a new installation path" -ForegroundColor White
    Write-Host "2. Exit the script" -ForegroundColor White
    Write-Host ""
    
    do {
        $choice = Read-Host "Enter your choice (1-2)"
        
        if ($choice -eq '1') {
            try {
                $basePath = Get-ValidLoxonePath
                $config.LoxonePath = $basePath
                Write-ConfigFile -config $config
                Write-Host "Path set successfully: $basePath" -ForegroundColor Green
                break
            }
            catch {
                Write-Host "Error setting path: $($_.Exception.Message)" -ForegroundColor Red
                continue
            }
        }
        elseif ($choice -eq '2') {
            Write-Host "Exiting..." -ForegroundColor Yellow
            exit 0
        }
        else {
            Write-Host "Invalid choice. Please enter 1 or 2." -ForegroundColor Red
        }
    } while ($true)
}


# Ask about desktop shortcut on first run
if ($isFirstRun -and $config.CreateDesktopShortcut -eq $null) {
    $createShortcut = Ask-CreateShortcut
    $config['CreateDesktopShortcut'] = $createShortcut
    
    if ($createShortcut) {
        # Note: versionList might not be available yet during first run
        # This will use the fallback method to find an executable
        New-DesktopShortcut -config $config -versionList $null | Out-Null
    }
    
    # Save the config with the new preferences
    Write-ConfigFile -config $config
}

# Get all Loxone Config folders with error handling
Write-Host ""
Write-Host "Scanning for Loxone Config installations..." -ForegroundColor Yellow

$loxoneFolders = @()
$scanSuccess = $false

try {
    $loxoneFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction Stop | 
                     Where-Object { $_.Name -like "*config*" -or $_.Name -like "*loxone*" }
    $scanSuccess = $true
}
catch {
    Write-Host "Error scanning directory '$basePath': $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "This usually means the path is invalid or inaccessible." -ForegroundColor Yellow
}

if (-not $scanSuccess -or $loxoneFolders.Count -eq 0) {
    if ($scanSuccess) {
        Write-Host "No Loxone Config folders found in $basePath" -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "Would you like to:" -ForegroundColor Cyan
    Write-Host "1. Change the installation path" -ForegroundColor White
    Write-Host "2. Exit the script" -ForegroundColor White
    Write-Host ""
    
    do {
        $choice = Read-Host "Enter your choice (1-2)"
        
        if ($choice -eq '1') {
            $newPath = Set-LoxonePath -config $config
            if ($newPath -ne $config.LoxonePath) {
                Write-Host ""
                Write-Host "Path updated. Please restart the script to use the new path." -ForegroundColor Green
            }
            Write-Host ""
            Write-Host "Press any key to exit..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit 0
        }
        elseif ($choice -eq '2') {
            Write-Host "Exiting..." -ForegroundColor Yellow
            exit 0
        }
        else {
            Write-Host "Invalid choice. Please enter 1 or 2." -ForegroundColor Red
        }
    } while ($true)
}

# Build list of available versions
$versionList = @()

foreach ($folder in $loxoneFolders) {
    $version = Get-LoxoneVersion -folderPath $folder.FullName
    $executable = Find-LoxoneExecutable -folderPath $folder.FullName
    
    if ($executable) {
        $versionInfo = [PSCustomObject]@{
            FolderName = $folder.Name
            Version = $version
            VersionObject = ConvertTo-VersionObject -versionString $version
            Path = $folder.FullName
            Executable = $executable
        }
        $versionList += $versionInfo
    }
}

# Check if any valid installations were found
if ($versionList.Count -eq 0) {
    Write-Host "No valid Loxone Config installations found with executables." -ForegroundColor Red
    Write-Host "This means no Loxone Config .exe files were found in the scanned folders." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Would you like to:" -ForegroundColor Cyan
    Write-Host "1. Change the installation path" -ForegroundColor White
    Write-Host "2. Exit the script" -ForegroundColor White
    Write-Host ""
    
    do {
        $choice = Read-Host "Enter your choice (1-2)"
        
        if ($choice -eq '1') {
            $newPath = Set-LoxonePath -config $config
            if ($newPath -ne $config.LoxonePath) {
                Write-Host ""
                Write-Host "Path updated. Please restart the script to use the new path." -ForegroundColor Green
            }
            Write-Host ""
            Write-Host "Press any key to exit..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit 0
        }
        elseif ($choice -eq '2') {
            Write-Host "Exiting..." -ForegroundColor Yellow
            exit 0
        }
        else {
            Write-Host "Invalid choice. Please enter 1 or 2." -ForegroundColor Red
        }
    } while ($true)
}


# Sort by version in descending order (newest first)
$versionList = $versionList | Sort-Object VersionObject -Descending

# Add numbering after sorting
$counter = 1
foreach ($item in $versionList) {
    $item | Add-Member -MemberType NoteProperty -Name "Number" -Value $counter
    $counter++
}

# Display available versions
Write-Host ""
Write-Host "Available Loxone Config Versions:" -ForegroundColor Green
Write-Host "=================================" -ForegroundColor Green

foreach ($item in $versionList) {
    $displayName = if ($item.Number -eq 1) {
        "Loxone Config $($item.Version) (latest)"
    } else {
        "Loxone Config $($item.Version)"
    }
    Write-Host "$($item.Number). $displayName" -ForegroundColor White
}

Write-Host ""
Write-Host "Press ENTER to launch the latest version, enter a number (1-$($versionList.Count)), 'c' to change path/create shortcut, or 'q' to quit:" -ForegroundColor Cyan

# Get user selection with default to latest version
do {
    $selection = Read-Host "Selection"
    
    # If user just pressed Enter (empty input), select the latest version (1)
    if ([string]::IsNullOrWhiteSpace($selection)) {
        $selectedNumber = 1
        Write-Host "Launching latest version..." -ForegroundColor Green
        break
    }
    
    # Change installation path or create shortcut
    if ($selection -eq 'c' -or $selection -eq 'C') {
        Write-Host ""
        Write-Host "=== Configuration Options ===" -ForegroundColor Cyan
        Write-Host "1. Change installation path"
        Write-Host "2. Create/recreate desktop shortcut"
        Write-Host ""
        
        do {
            $configChoice = Read-Host "Enter your choice (1-2)"
            
            if ($configChoice -eq '1') {
                $newPath = Set-LoxonePath -config $config
                Write-Host ""
                Write-Host "Path changed. Please restart the script to use the new path." -ForegroundColor Yellow
                Write-Host "Press any key to exit..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                exit 0
            }
            elseif ($configChoice -eq '2') {
                Write-Host ""
                $success = New-DesktopShortcut -config $config -versionList $versionList
                if ($success) {
                    # Use hashtable assignment method
                    $config['CreateDesktopShortcut'] = $true
                    Write-ConfigFile -config $config
                }
                Write-Host ""
                Write-Host "Press any key to continue..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                
                # Re-display the version list after shortcut creation
                Write-Host ""
                Write-Host "Available Loxone Config Versions:" -ForegroundColor Green
                Write-Host "=================================" -ForegroundColor Green
                
                foreach ($item in $versionList) {
                    $displayName = if ($item.Number -eq 1) {
                        "Loxone Config $($item.Version) (latest)"
                    } else {
                        "Loxone Config $($item.Version)"
                    }
                    Write-Host "$($item.Number). $displayName" -ForegroundColor White
                }
                break
            }
            else {
                Write-Host "Invalid choice. Please enter 1 or 2." -ForegroundColor Red
            }
        } while ($true)
        
        # Return to main menu
        Write-Host ""
        Write-Host "Press ENTER to launch the latest version, enter a number (1-$($versionList.Count)), 'c' to change path/create shortcut, or 'q' to quit:" -ForegroundColor Cyan
        continue
    }
    
    # Quit
    if ($selection -eq 'q' -or $selection -eq 'Q') {
        Write-Host "Exiting..." -ForegroundColor Yellow
        exit 0
    }
    
    # Number selection
    $selectedNumber = $null
    if ([int]::TryParse($selection, [ref]$selectedNumber)) {
        if ($selectedNumber -ge 1 -and $selectedNumber -le $versionList.Count) {
            break
        }
    }
    
    Write-Host "Invalid selection. Please enter a number between 1 and $($versionList.Count), press ENTER for latest, 'c' for configuration, or 'q' to quit." -ForegroundColor Red
    
} while ($true)

# Launch selected version
$selectedVersion = $versionList[$selectedNumber - 1]

Write-Host ""
Write-Host "Launching Loxone Config $($selectedVersion.Version)..." -ForegroundColor Green

try {
    Start-Process -FilePath $selectedVersion.Executable -WorkingDirectory $selectedVersion.Path -ErrorAction Stop
    Write-Host "Successfully launched Loxone Config!" -ForegroundColor Green
    
    # Brief pause to allow the process to start
    Start-Sleep -Milliseconds 500
    
    # Exit automatically after successful launch
    Write-Host "Exiting launcher..." -ForegroundColor Cyan
    exit 0
}
catch {
    Write-Host "Error launching Loxone Config: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Executable path: $($selectedVersion.Executable)" -ForegroundColor Yellow
    
    # Throw an error to indicate failure
    throw "Failed to launch Loxone Config: $($_.Exception.Message)"
}
