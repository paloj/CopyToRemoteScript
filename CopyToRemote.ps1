# FolderCopyTool.ps1
param(
    [string]$SourcePath,
    [string]$TargetNickname
)

# This script is designed to be run in PowerShell and provides a menu-driven interface for managing folder copy targets.
# It allows users to add, remove, and manage target locations for copying files and folders using Robocopy.
# The script also creates a right-click context menu for easy access to the copy functionality.

# The script uses a CSV file to store target locations and their nicknames, making it easy to manage multiple targets.
# The script also includes error handling to ensure that the user is informed of any issues that arise during execution.

# OPTIONS
$exclusions = @("Thumbs.db", ".DS_Store") # Example exclusions
$EnableLog = $true  # Set to $false to disable robocopy log file creation
$AskEnter = $true # Set to $false to disable the prompt for pressing Enter after copying

# The storage path for the CSV file that contains target locations.
# This is set to the user's AppData folder to ensure it is user-specific and not system-wide.
$storagePath = "$env:APPDATA\FolderCopyTargets"
$targetsFile = "$storagePath\targets.csv"

# Function to initialize the storage directory and CSV file.
# This function checks if the storage directory exists, and if not, creates it.
function Initialize-Storage {
    try {
        # Create storage directory if it doesn't exist
        if (!(Test-Path $storagePath)) {
            New-Item -ItemType Directory -Path $storagePath -Force -ErrorAction Stop | Out-Null
        }
        # If the CSV file is missing or empty, create it with proper CSV header.
        if (!(Test-Path $targetsFile) -or ((Get-Content $targetsFile -Raw).Trim()) -eq "") {
            "nickname;path" | Out-File -Encoding utf8 -FilePath $targetsFile
            Write-Host "Storage initialized. You can now add targets."
        }
    }
    catch {
        Write-Error "Failed to initialize storage: $($_.Exception.Message)"
        exit 1
    }
}

# Function to load the target list from the CSV file.
# It reads the file, processes each line, and returns an array of target objects.
function Get-Targets {
    Initialize-Storage
    try {
        Write-Host "Loading target list from $targetsFile"
        if (!(Test-Path $targetsFile) -or ((Get-Content $targetsFile -Raw).Trim()) -eq "") {
            return @()  # Return an empty array if the file is empty
        }
        $raw = Get-Content $targetsFile -Encoding utf8
        if ($raw.Count -lt 2) {
            return @()  # Only header exists, no target entries
        }
        # Assume header is "nickname;path"
        $targets = @()
        # Process each line except the header
        for ($i = 1; $i -lt $raw.Count; $i++) {
            $line = $raw[$i].Trim()
            if ($line -eq "") { continue }
            $fields = $line.Split(";")
            $targets += [PSCustomObject]@{
                nickname = $fields[0].Trim()
                path     = $fields[1].Trim()
            }
        }
        return $targets
    }
    catch {
        Write-Error "Error loading target list. CSV might be corrupted or improperly formatted. Error: $($_.Exception.Message)"
        return @()
    }
}

# Function to save the target list back to the CSV file.
# It takes an array of target objects and writes them to the file in CSV format.
function Save-Targets {
    param (
        [Parameter(Mandatory = $true)]
        [array]$targets
    )

    try {
        # Build CSV lines manually
        $csvLines = @("nickname;path")
        foreach ($target in $targets) {
            $csvLines += "$($target.nickname);$($target.path)"
        }
        $csvLines | Out-File -Encoding utf8 -FilePath $targetsFile -Force
        Write-Host "Targets saved to $targetsFile"
    }
    catch {
        Write-Error "Failed to save targets: $($_.Exception.Message)"
    }
}

# Function to display the main menu and handle user input.
# This function provides a simple text-based menu for the user to interact with.
function Show-Menu {
    Write-Host "1. Create new target location (cmd)"
    Write-Host "2. Create new target location (Prompt folder selection)"
    Write-Host "3. Remove target location"
    Write-Host "4. Create right-click menu item"
    Write-Host "5. Uninstall right-click menu"
    Write-Host "0. Exit"
    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1' { Add-Target }
        '2' { Add-TargetPrompt }
        '3' { Remove-Target }
        '4' { New-ContextMenu }
        '5' { Remove-ContextMenu }
        '0' { exit }
        default { Write-Host "Invalid choice"; Show-Menu }
    }
}

# Function to add a new target location.
# This function prompts the user for a path and a nickname, validates them, and saves the new target to the CSV file.
function Add-Target {
    # Force the result of Get-Targets into an array
    $targets = @(Get-Targets)

    $path = Read-Host "Enter the path for the target"
    # Validate the path
    if (!(Test-Path $path)) {
        Write-Host "Invalid path. Please enter a valid path."
        return
    }
        
    $nickname = Read-Host "Enter a nickname for the target $path"
    # Check if the nickname already exists
    if ($targets | Where-Object { $_.nickname -eq $nickname }) {
        Write-Host "Nickname already exists. Please choose a different one."
        return
    }   

    # Create a new target object
    $newTarget = [PSCustomObject]@{
        nickname = $nickname
        path     = $path
    }

    # Add the new target to the array
    $targets += $newTarget

    # Save the updated targets
    Save-Targets $targets
    New-ContextMenu
    Write-Host "Target added and context menu updated."
}

# Function to add a new target location using a folder selection dialog.
# This function uses Windows Forms to prompt the user for a folder path and a nickname, validates them, and saves the new target to the CSV file.
# This function is similar to Add-Target but uses a GUI for folder selection.
function Add-TargetPrompt {
    Add-Type -AssemblyName System.Windows.Forms
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialogResult = $folderBrowser.ShowDialog()
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $path = $folderBrowser.SelectedPath
        $nickname = Read-Host "Enter a nickname for the target $path"
        # Validate the path
        if (!(Test-Path $path)) {
            Write-Host "Invalid path. Please enter a valid path."
            return
        }
        # Check if the nickname already exists
        if ($targets | Where-Object { $_.nickname -eq $nickname }) {
            Write-Host "Nickname already exists. Please choose a different one."
            return
        }

        $targets = @(Get-Targets)
        $newTarget = [PSCustomObject]@{
            nickname = $nickname
            path     = $path
        }
        $targets += $newTarget
        Save-Targets $targets
        New-ContextMenu
        Write-Host "Target added and context menu updated."
    }
    else {
        Write-Host "Folder selection cancelled."
    }
    Show-Menu
}

# Function to remove a target location.
# This function displays the list of targets, prompts the user to select one, and removes it from the CSV file.
function Remove-Target {
    $targets = @(Get-Targets)
    if ($targets.Count -eq 0) {
        Write-Host "No targets found."
        Show-Menu
        return
    }

    # Display the list of targets
    for ($i = 0; $i -lt $targets.Count; $i++) {
        Write-Host "$(($i+1)). $($targets[$i].nickname) => $($targets[$i].path)"
    }

    # Prompt the user to select a target to remove
    $choice = Read-Host "Select number to remove"
    if ($choice -match '^\d+$' -and $choice -gt 0 -and $choice -le $targets.Count) {
        # Remove the selected target by index
        $indexToRemove = $choice - 1
        $targets = $targets | Where-Object { $_ -ne $targets[$indexToRemove] }

        # Handle the case where the list becomes empty
        if ($targets.Count -eq 0) {
            Remove-Item -Path $targetsFile -Force -ErrorAction SilentlyContinue
            Write-Host "All targets removed. CSV file deleted."
        }
        else {
            Save-Targets $targets  # Save the updated list to the CSV file
        }

        New-ContextMenu        # Update the context menu
        Write-Host "Target removed and context menu updated."
    }
    else {
        Write-Host "Invalid selection."
    }

    Show-Menu
}

# Function to create a right-click context menu for the script.
# This function creates a registry entry for the context menu and adds sub-commands for each target location.
# It uses the script path to ensure that the correct script is executed when the menu item is clicked.
# The context menu allows users to right-click on a folder and copy it to the specified target locations.
function New-ContextMenu {
    # Use $PSCommandPath to reliably get the current script file path.
    $scriptPath = $PSCommandPath
    $keyPath = "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\shell\CopyToRemote"
    try {
        Remove-Item -Path $keyPath -Recurse -ErrorAction SilentlyContinue
        New-Item -Path $keyPath -Force | Out-Null
        Set-ItemProperty -Path $keyPath -Name "MUIVerb" -Value "Copy to Remote"
        Set-ItemProperty -Path $keyPath -Name "SubCommands" -Value ""

        $subCmdPath = Join-Path $keyPath 'shell'
        $targets = Get-Targets
        foreach ($target in $targets) {
            if (-not [string]::IsNullOrWhiteSpace($target.nickname)) {
                $sub = New-Item -Path (Join-Path $subCmdPath $target.nickname) -Force
                # Remove newlines and escape quotes from script path
                $escapedScriptPath = ($scriptPath -replace '[\r\n]+', '') -replace '"', '""'
                $cmd = "powershell.exe -NoProfile -WindowStyle Normal -ExecutionPolicy Bypass -File `"$escapedScriptPath`" -SourcePath `"%1`" -TargetNickname `"$($target.nickname)`""
                # Ensure the command is a single line (remove any stray newlines)
                $cmd = $cmd -replace "[\r\n]+", " "
                $commandKey = Join-Path $sub 'command'
                # Ensure the command key path uses the Registry:: prefix
                $commandKey = "Registry::" + $commandKey
                # Create the command key and set the default value
                $null = New-Item -Path $commandKey -Force
                New-ItemProperty -Path $commandKey -Name '(default)' -Value $cmd -PropertyType String -Force | Out-Null
            }
        }
        Write-Host "Right-click context menu created."
    }
    catch {
        Write-Error "Failed to create context menu: $($_.Exception.Message)"
    }
    Show-Menu
}

# Function to remove the right-click context menu.
# This function deletes the registry entry for the context menu, effectively removing it from the right-click options.
function Remove-ContextMenu {
    $keyPath = "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\shell\CopyToRemote"
    try {
        if (Test-Path $keyPath) {
            Remove-Item -Path $keyPath -Recurse -Force
            Write-Host "Context menu removed."
        }
        else {
            Write-Host "Context menu not found."
        }
    }
    catch {
        Write-Error "Failed to remove context menu: $($_.Exception.Message)"
    }
    Show-Menu
}

# Function to copy files and folders with versioning.
# This function takes a source path and a target base path, creates a versioned target directory, and uses Robocopy to copy the files.
function Copy-With-Versioning($source, $targetBase) {
    if (!(Test-Path $source)) {
        Write-Error "Source path does not exist: $source"
        exit 1
    }
    if (!(Test-Path $targetBase)) {
        Write-Error "Target base path does not exist: $targetBase"
        exit 1
    }

    $date = Get-Date -Format "yyyy-MM-dd"
    $baseName = Split-Path $source -Leaf
    $i = 1
    do {
        $targetName = "$date $baseName V$i"
        $targetPath = Join-Path $targetBase $targetName
        $i++
    } while (Test-Path $targetPath)

    try {
        New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
    }
    catch {
        Write-Error "Failed to create target directory: $($_.Exception.Message)"
        exit 1
    }

    # Build an array for robocopy parameters
    $robocopyArgs = @(
        "$source",
        "$targetPath",
        "/V",
        "/S",
        "/E",
        "/DCOPY:DA",
        "/COPY:DAT",
        "/Z",
        "/ETA",
        "/TEE",
        "/MT:32",
        "/R:2",
        "/W:10"
    )
    # Add exclusion parameters separately
    $robocopyArgs += ($exclusions | ForEach-Object { "/XF"; $_ })
    $robocopyArgs += ($exclusions | ForEach-Object { "/XD"; $_ })

    # Add optional logging parameter if enabled
    if ($EnableLog) {
        $robocopyArgs += "/LOG+:`"$targetPath\robocopy.log`""
    }

    # Call robocopy using splatting
    robocopy @robocopyArgs

    if ($LASTEXITCODE -ge 8) {
        Write-Error "Robocopy failed with exit code $LASTEXITCODE"
    }
    else {
        Write-Host "Copied successfully to $targetPath"
    }

    # Optionally prompt the user to press Enter before exiting
    if ($AskEnter) {
        Read-Host -Prompt "Press Enter to exit"
    }
    else {
        exit 0
    }
}

# === MAIN ENTRY POINT ===

if ($PSBoundParameters.ContainsKey('SourcePath') -and $PSBoundParameters.ContainsKey('TargetNickname')) {
    $targets = Get-Targets
    $target = $targets | Where-Object { $_.nickname -eq $TargetNickname }
    if ($null -eq $target) {
        Write-Error "Target nickname not found."
        exit 1
    }
    Copy-With-Versioning -source $SourcePath -targetBase $target.path
    exit
}
else {
    Show-Menu
}
# End of script