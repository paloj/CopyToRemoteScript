# FolderCopyTool.ps1
param(
    [string]$SourcePath,
    [string]$TargetNickname
)

$storagePath = "$env:APPDATA\FolderCopyTargets"
$targetsFile = "$storagePath\targets.csv"

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
    } catch {
        Write-Error "Failed to initialize storage: $($_.Exception.Message)"
        exit 1
    }
}

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
    } catch {
        Write-Error "Error loading target list. CSV might be corrupted or improperly formatted. Error: $($_.Exception.Message)"
        return @()
    }
}

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
    } catch {
        Write-Error "Failed to save targets: $($_.Exception.Message)"
    }
}

function Show-Menu {
    Write-Host "1. Create new target location"
    Write-Host "2. Remove target location"
    Write-Host "3. Create right-click menu item"
    Write-Host "4. Uninstall right-click menu"
    Write-Host "0. Exit"
    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1' { Add-Target }
        '2' { Remove-Target }
        '3' { New-ContextMenu }
        '4' { Remove-ContextMenu }
        '0' { exit }
        default { Write-Host "Invalid choice"; Show-Menu }
    }
}

function Add-Target {
    # Force the result of Get-Targets into an array
    $targets = @(Get-Targets)
    $nickname = Read-Host "Enter a nickname for the target"
    $path = Read-Host "Enter the path for the target"

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
        } else {
            Save-Targets $targets  # Save the updated list to the CSV file
        }

        New-ContextMenu        # Update the context menu
        Write-Host "Target removed and context menu updated."
    } else {
        Write-Host "Invalid selection."
    }

    Show-Menu
}

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
                $escapedScriptPath = ($scriptPath -replace '[\r\n]+','') -replace '"','""'
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

function Remove-ContextMenu {
    $keyPath = "Registry::HKEY_CURRENT_USER\Software\Classes\Directory\shell\CopyToRemote"
    try {
        if (Test-Path $keyPath) {
            Remove-Item -Path $keyPath -Recurse -Force
            Write-Host "Context menu removed."
        } else {
            Write-Host "Context menu not found."
        }
    }
    catch {
        Write-Error "Failed to remove context menu: $($_.Exception.Message)"
    }
    Show-Menu
}

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
    } catch {
        Write-Error "Failed to create target directory: $($_.Exception.Message)"
        exit 1
    }

    robocopy "$source" "$targetPath" /E /Z /MT:32 /NDL /NFL /ETA
    if ($LASTEXITCODE -ge 8) {
        Write-Error "Robocopy failed with exit code $LASTEXITCODE"
    } else {
        Write-Host "Copied successfully to $targetPath"
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
} else {
    Show-Menu
}
# End of script