# Loxone Config Version Launcher

A PowerShell script that automatically detects and launches different versions of Loxone Config with an intuitive menu interface.

## 🚀 Features

- **Automatic Version Detection**: Scans your Loxone installation directory and extracts version information from `unins000.dat` files
- **Smart Version Sorting**: Always displays the newest version first for quick access
- **One-Click Latest Launch**: Press Enter to instantly launch the latest version
- **Desktop Shortcut Creation**: Automatically creates a desktop shortcut with the Loxone Config icon
- **Persistent Configuration**: Remembers your custom installation paths and preferences
- **Custom Path Support**: Configure non-standard Loxone installation directories
- **User-Friendly Interface**: Clean, colorized menu with clear navigation options

## 📋 Requirements

- **Windows OS** with PowerShell 5.0 or later
- **Loxone Config** installations in a common directory structure
- **PowerShell Execution Policy** allowing script execution (handled automatically)

## 🛠️First-Time
- The script will automatically detect Loxone Config installations
- You'll be prompted to create a desktop shortcut (optional)
- Configuration preferences are saved automatically

## 🎯 Usage

### Quick Launch
- **Press Enter**: Launch the latest version immediately
- **Enter a number**: Launch a specific version (1-6, etc.)
- **Press 'q'**: Quit the application
  
### Configuration Options
- **Press 'c'**: Access configuration menu
- Change installation path
- Create/recreate desktop shortcut
  
### Example Output

```
=== Loxone Config Version Launcher ===

# Available Loxone Config Versions:

1. Loxone Config 16.0.6.10 (latest)
    
2. Loxone Config 16.0.6.3
    
3. Loxone Config 15.5.3.4
    
4. Loxone Config 15.3.12.13
    
5. Loxone Config 15.3.12.2
    
6. Loxone Config 15.2.10.14
    

Press ENTER to launch the latest version, enter a number (1-6), 'c' to change path/create shortcut, or 'q' to quit:
```

## ⚙️ Configuration
The script automatically creates a configuration file at:

```
%TEMP%\LoxoneLauncher.config
```

### Configuration Options
- **Custom Installation Path**: Set non-standard Loxone installation directories
- **Desktop Shortcut Preference**: Remember shortcut creation choice
- **Automatic Path Detection**: Fallback to default paths when custom paths aren't available

### Default Installation Path
```
C:\Program Files (x86)\Loxone
```


## 🔧 Advanced Features

### Desktop Shortcut
- **Automatic Icon Extraction**: Uses the latest Loxone Config version's icon
- **Proper Execution Context**: Runs with correct working directory and execution policy
- **One-Click Access**: Double-click desktop shortcut to launch the version selector

### Version Detection Logic
The script identifies Loxone Config versions by:
1. Scanning folders matching `*config*` or `*loxone*` patterns
2. Reading version information from `unins000.dat` files
3. Locating executable files (`LoxoneConfig.exe`, `Loxone Config.exe`, `Config.exe`)
4. Sorting versions numerically (newest first)

### Error Handling
- **Path Validation**: Verifies installation paths exist before proceeding
- **Permission Management**: Uses user temp directory for configuration storage
- **Graceful Fallbacks**: Continues operation even if some features fail
- **Clear Error Messages**: Provides helpful feedback for troubleshooting

## 🐛 Troubleshooting

### Common Issues

**"No Loxone Config folders found"**
- Verify your Loxone installation path
- Use the 'c' option to set a custom path
- Ensure folder names contain "config" or "loxone"

**"Permission denied" errors**
- Run PowerShell as Administrator if needed
- Check that the script location is writable
- Verify antivirus isn't blocking the script

**Desktop shortcut not working**
- Ensure the script file hasn't been moved
- Recreate the shortcut using the 'c' option
- Check that PowerShell execution policy allows script execution

### Debug Information
The script provides detailed feedback including:
- Configuration file location
- Detected installation paths
- Version extraction results
- Executable file locations

## 🙏 Acknowledgments

- **Loxone** for creating the Config software
- **PowerShell Community** for excellent documentation and examples

**Made with ❤️ for the Loxone community**
