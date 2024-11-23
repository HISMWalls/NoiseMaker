# Define the source and destination paths
$sourcePath = "\\172.16.2.6\IT Public Drive\"
$destPath = "C:\IT\NoiseMaker\"
$appPath = "T:\Software\_Scripts\NoiseMaker\NoiMkr.ps1"

# Try block for error handling
try {
    # Create a new PSDrive
    New-PSDrive -Name T -PSProvider FileSystem -Root $sourcePath -ErrorAction Stop > $null
    Get-PSDrive -PSProvider FileSystem

    # Ensure the destination directory exists
    if (-not (Test-Path $destPath)) {
        New-Item -Path $destPath -ItemType Directory -Force -ErrorAction Stop > $null
    }

    # Ensure the Public Desktop directory exists
    $PublicDesktopPath = "$env:PUBLIC\Desktop"
    if (-not (Test-Path $PublicDesktopPath)) {
        New-Item -Path $PublicDesktopPath -ItemType Directory -Force -ErrorAction Stop > $null
    }

    # Copy the file to the destination
    Copy-Item -Path $appPath -Destination "$destPath\NoiMkr.ps1" -ErrorAction Stop

    # Define the source file location for the shortcut
    $SourceFilePath = "$destPath\NoiMkr.ps1"

    # Define the shortcut file location and name
    $ShortcutPath = "$PublicDesktopPath\CallManager.lnk"

    # Create a new WScript.Shell object
    $WScriptShell = New-Object -ComObject WScript.Shell

    # Create the shortcut
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)

    # Set the target path
    $Shortcut.TargetPath = $SourceFilePath

    # Save the shortcut
    $Shortcut.Save()
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    # Clean up resources
    if (Get-PSDrive -Name T -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name T -ErrorAction SilentlyContinue
    }
}
