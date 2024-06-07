Add-Type –assemblyName PresentationFramework

function Download-File {
    param (
        [Parameter(Mandatory = $true)]
        [string] $Uri,
        [Parameter(Mandatory = $true)]
        [string] $Output
    )

    $webClient = New-Object System.Net.WebClient
    try {
        $webClient.DownloadFile($Uri, $Output)
        Write-Host "Downloaded file from $Uri to $Output." -ForegroundColor Cyan
    } catch {
        Write-Host "Error downloading file from $Uri to $Output." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    } finally {
        $webClient.Dispose()
    }
}

function Get-GitHubLatestRelease {
    param (
        [Parameter(Mandatory = $true)]
        [string] $repo
    )

    $uri = "https://api.github.com/repos/$repo/releases/latest"
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (compatible; PowerShell/5.0)")

    try {
        $responseJson = $webClient.DownloadString($uri)
        $response = ConvertFrom-Json -InputObject $responseJson
        return $response.tag_name
    } catch {
        Write-Host "Error retrieving latest release from GitHub for repository: $repo" -ForegroundColor Cyan
        return $null
    } finally {
        $webClient.Dispose()
    }
}

function Get-GitHubLatestAssets {
    param ([string] $repo)

    $uri = "https://api.github.com/repos/$repo/releases/latest"
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (compatible; PowerShell/5.0)")

    try {
        $responseJson = $webClient.DownloadString($uri)
        if ($null -eq $responseJson) {
            throw "No response from GitHub. Check the URL and your internet connection."
        }
        
        $response = ConvertFrom-Json -InputObject $responseJson
        return $response.assets
    } catch {
        Write-Host "Error retrieving GitHub assets. Ensure the repo exists and is reachable." -ForegroundColor DarkCyan
        return $null
    } finally {
        $webClient.Dispose()
    }
}

function archipelagoClientLocation {
    $ArchipelagoClientDir = ""

    Write-Host "Getting Archipelago installation location" -ForegroundColor Yellow

    $archipelagoWMI = Get-WmiObject -Class Win32_Product -Filter 'Name like "%Archipelago%"' |
        Select-Object -First 1 -Property InstallLocation

    if ($archipelagoWMI.InstallLocation) {
        $ArchipelagoClientDir = $archipelagoWMI.InstallLocation
    }

    if (-not $ArchipelagoClientDir) {
        $archipelagoReg = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
            ForEach-Object { Get-ItemProperty $_.PsPath } |
            Where-Object { $_.DisplayName -like "Archipelago*" } |
            Select-Object -First 1 -Property InstallLocation

        if ($archipelagoReg.InstallLocation) {
            $ArchipelagoClientDir = $archipelagoReg.InstallLocation
        }
    }

    if (-not $ArchipelagoClientDir) {
        $ArchipelagoClientDir = if (Test-Path $ArchipelagoDefaultDir) {
            $ArchipelagoDefaultDir
        } else {
            Write-Host "`r`nArchipelago client not found. Please enter the installation path:`r`n" -ForegroundColor Red
            Read-Host "Enter the installed location of the Archipelago Client"
        }
    }

    if ([string]::IsNullOrWhiteSpace($ArchipelagoClientDir)) {
        Write-Host "Archipelago client not found. Ensure it is installed or provide a valid path." -ForegroundColor Red
        return
    }

    Write-Host "Archipelago client found at $ArchipelagoClientDir." -ForegroundColor Yellow
    return $ArchipelagoClientDir
}

function Extract-ZipFile {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ZipPath,
        [Parameter(Mandatory = $true)]
        [string] $Destination
    )

    if (-not (Test-Path $ZipPath)) {
        Write-Host "Zip file $ZipPath does not exist." -ForegroundColor Red
        return
    }

    # Create a temporary extraction directory
    $TempExtractDir = Join-Path (Split-Path -Parent $ZipPath) "TempExtract"
    if (-not (Test-Path $TempExtractDir)) {
        New-Item -Path $TempExtractDir -ItemType Directory | Out-Null
    }

    try {
        # Extract the zip file to the temporary directory
        Expand-Archive -Path $ZipPath -DestinationPath $TempExtractDir -Force
        Write-Host "Extracted $ZipPath to $TempExtractDir." -ForegroundColor Cyan

        # Move the contents to the final destination
        Get-ChildItem -Path $TempExtractDir -Recurse | ForEach-Object {
            $TargetPath = Join-Path $Destination -ChildPath $_.FullName.Substring($TempExtractDir.Length + 1)
            if (-not (Test-Path (Split-Path $TargetPath -Parent))) {
                New-Item -Path (Split-Path $TargetPath -Parent) -ItemType Directory | Out-Null
            }
            Move-Item -Path $_.FullName -Destination $TargetPath -Force
        }
        Write-Host "Moved extracted contents to $Destination." -ForegroundColor Cyan
    } catch {
        Write-Host "Error extracting ($ZipPath): $_" -ForegroundColor Red
    } finally {
        Remove-Item -Path $TempExtractDir -Recurse -Force
    }
}

function ConvertFrom-Yaml {
    param (
        [string]$YamlString
    )
    $YamlObject = @{}
    $YamlString -split "`n" | ForEach-Object {
        if ($_ -match "^\s*(?<Key>[^:]+)\s*:\s*(?<Value>.*)\s*$") {
            $YamlObject[$matches.Key.Trim()] = $matches.Value.Trim()
        }
    }
    return $YamlObject
}

function ConvertTo-Yaml {
    param (
        [hashtable]$YamlObject
    )
    $YamlString = ""
    $YamlObject.GetEnumerator() | ForEach-Object {
        $YamlString += "$($_.Key): $($_.Value)`n"
    }
    return $YamlString
}

function Find-LunacidInstallation {
    # Determine if the system is 64-bit or 32-bit
    $arch = if (($env:PROCESSOR_ARCHITECTURE -eq "AMD64") -or ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64")) {
        "64-bit"
    } else {
        "32-bit"
    }

    $SteamKeyPath = if ($arch -eq "64-bit") {
        "HKLM:\SOFTWARE\Wow6432Node\Valve\Steam"
    } else {
        "HKLM:\SOFTWARE\Valve\Steam"
    }

    # Get Steam install path from the registry
    $SteamPath = Get-ItemProperty -Path $SteamKeyPath -Name "InstallPath" | Select-Object -ExpandProperty InstallPath
    Write-Host "Steam install path is $SteamPath." -ForegroundColor Yellow

    # Path to the libraryfolders.vdf file
    $SteamLibraryFoldersPath = Join-Path $SteamPath "steamapps\libraryfolders.vdf"

    # Ensure the file exists
    if (-not (Test-Path $SteamLibraryFoldersPath)) {
        Write-Host "Steam libraryfolders.vdf file not found." -ForegroundColor Red
        return $null
    }

    # Read the content of the file and extract library paths
    $SteamLibraryFoldersRaw = Get-Content $SteamLibraryFoldersPath -Raw
    $libraryPaths = @()  # Array to store Steam library paths
    $lines = $SteamLibraryFoldersRaw -split "`n"

    foreach ($line in $lines) {
        $trimmedLine = $line.Trim()
        if ($trimmedLine -match '^"path"\s+"([^"]+)"') {
            $libraryPath = $matches[1]  # Get the path from the match
            $libraryPaths += $libraryPath
        }
    }

    # Check each library path for Lunacid
    foreach ($path in $libraryPaths) {
        $LunacidCheckPath = Join-Path $path "steamapps/common/Lunacid"
        if (Test-Path $LunacidCheckPath) {
            Write-Host "`r`nLunacid installation directory found at $LunacidCheckPath." -ForegroundColor Yellow
            return $LunacidCheckPath
        }
    }

    Write-Host "`r`nLunacid installation directory not found. Ensure Steam and Lunacid are installed." -ForegroundColor Red
    return $null
}

function Install {
    $ScriptDir = "$PSScriptRoot"
    $TempDir = Join-Path $ScriptDir "temp"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    cd $PSScriptRoot

    if (-not (Test-Path $TempDir)) {
        New-Item -Path $TempDir -ItemType Directory | Out-Null
    }

    $ArchipelagoDir = archipelagoClientLocation

    # Retrieve assets from GitHub
    $BepInExAssets = Get-GitHubLatestAssets -repo "BepInEx/BepInEx"
    $ArchipelagoAssets = Get-GitHubLatestAssets -repo "ArchipelagoMW/Archipelago"
    $LunacidAssets = Get-GitHubLatestAssets -repo "Witchybun/LunacidAPClient"

    if ($null -eq $BepInExAssets -or $null -eq $ArchipelagoAssets -or $null -eq $LunacidAssets) {
        Write-Host "Error: Failed to retrieve one or more assets from GitHub. Exiting." -ForegroundColor Red
        return
    }

    $BepInExFilename = $BepInExAssets | Where-Object { $_.name -like "BepInEx_win*.zip" } | Select-Object -First 1
    $ArchipelagoSetup = $ArchipelagoAssets | Where-Object { $_.name -like "Setup.Archipelago*.exe" } | Select-Object -First 1
    $LunacidWorldFile = $LunacidAssets | Where-Object { $_.name -like "*.apworld" } | Select-Object -First 1
    $LunacidZipFile = $LunacidAssets | Where-Object { $_.name -like "*.zip" } | Select-Object -First 1

    if ($null -eq $BepInExFilename -or $null -eq $ArchipelagoSetup -or $null -eq $LunacidWorldFile -or $null -eq $LunacidZipFile) {
        Write-Host "Error: Failed to assign asset download URLs. Exiting." -ForegroundColor Red
        return
    }

    if ($BepInExFilename) {
        $BepInExZip = Join-Path $TempDir $BepInExFilename.name
        Download-File -Uri $BepInExFilename.browser_download_url -Output $BepInExZip
    }

    if ($ArchipelagoSetup) {
        $ArchipelagoExe = Join-Path $TempDir $ArchipelagoSetup.name
        Download-File -Uri $ArchipelagoSetup.browser_download_url -Output $ArchipelagoExe
    }

    if ($LunacidWorldFile -and $LunacidZipFile) {
        $LunacidApWorld = Join-Path $TempDir $LunacidWorldFile.name
        Download-File -Uri $LunacidWorldFile.browser_download_url -Output $LunacidApWorld

        $LunacidApZip = Join-Path $TempDir $LunacidZipFile.name
        Download-File -Uri $LunacidZipFile.browser_download_url -Output $LunacidApZip
    }

    # Extract files to their respective directories
    $LunacidDir = Find-LunacidInstallation

    if ($LunacidDir) {
        $BepInExDir = Join-Path $LunacidDir "BepInEx"
        $LunacidBepInExPluginsDir = Join-Path $BepInExDir "plugins"

        if (Test-Path $BepInExZip) {
            Extract-ZipFile -ZipPath $BepInExZip -Destination $LunacidDir
        }

        if (Test-Path $LunacidApZip) {
            Extract-ZipFile -ZipPath $LunacidApZip -Destination $LunacidBepInExPluginsDir
        }

        if (Test-Path $LunacidApWorld) {
            $ArchipelagoWorldsDir = Join-Path $ArchipelagoDir "lib/worlds"
            if (-not (Test-Path $ArchipelagoWorldsDir)) {
                New-Item -Path $ArchipelagoWorldsDir -ItemType Directory | Out-Null
            }
            Move-Item -Path $LunacidApWorld -Destination $ArchipelagoWorldsDir -Force
        }
    }

    if ($ArchipelagoDir) {
        $ArchipelagoLauncher = Join-Path $ArchipelagoDir "ArchipelagoLauncher.exe"
        if (Test-Path $ArchipelagoLauncher) {
            Write-Host "`r`nInstallation and Setup complete. Archipelago Client has been started." -ForegroundColor Green
            Write-host "`r`nContinue with the In-Game setup instructions at' https://github.com/Witchybun/LunacidAPClient/blob/main/Documentation/Setup.md#in-game-setup '`r`nTip: You can CTRL + Left-Click the link to open it in your browser or copy it out.`r`n" -ForegroundColor Green
            generateGame
        } else {
            Write-Host "`r`nArchipelagoLauncher.exe not found in Archipelago directory." -ForegroundColor Red
        }
    } else {
        Write-Host "`r`nArchipelago Client directory not specified. Please check the installation process." -ForegroundColor Red
    }
}

function startServer {
    param (
        [string] $ArchipelagoDir = $null
    )

    if (-not $ArchipelagoDir) {
        $ArchipelagoDir = archipelagoClientLocation
    }

    if (-not $ArchipelagoDir) {
        Write-Host "Error: Archipelago directory is not specified. Please ensure Archipelago is installed." -ForegroundColor Red
        return
    }

    $LunacidDir = Find-LunacidInstallation
    if (-not $LunacidDir) {
        Write-Host "Error: Lunacid directory is not specified. Please ensure Lunacid is installed." -ForegroundColor Red
        return
    }

    $PlayersDir = Join-Path $ArchipelagoDir "Players"
    $LunacidYamlPath = Join-Path $PlayersDir "lunacid.yaml"

    if (Test-Path $LunacidYamlPath) {
        $playername = Get-Content $LunacidYamlPath | Where-Object { $_ -match '^name:' } | ForEach-Object {
            $_ -replace '^name:\s*', ''
        }
    } else {
        Write-Host "Lunacid.yaml file not found." -ForegroundColor Red
        return
    }

    Write-Host "`r`n Starting Server `r`n" -ForegroundColor Yellow
    Start-Sleep -Seconds 1

    $ArchipelagoServer = Join-Path $ArchipelagoDir "ArchipelagoServer.exe"
    Start-Process $ArchipelagoServer

    $IPADDRESS = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration).Where{ $_.IPAddress }.foreach{ $_.IPAddress[0] }
    Write-Host "`r`nLaunch the game and use the connection detail below to connect.`r`nIf trouble connection, please check the setup page here: https://github.com/Witchybun/LunacidAPClient/blob/main/Documentation/Setup.md" -ForegroundColor Green
    Write-Host "`r`n Player Slot: $playername `r`n Host: $IPADDRESS `r`n Port: 38281 `r`n`r`n " -ForegroundColor Cyan

    Start-Sleep -Seconds 10
    Read-Host -Prompt "Press any key to start Lunacid"
    $LunacidExe = Join-Path $LunacidDir "Lunacid.exe"
    Start-Process $LunacidExe
    return
}

function generateGame {
    $ArchipelagoDir = archipelagoClientLocation
    $ArchipelagoClientDir = $ArchipelagoDir
    $TemplatesDir = Join-Path $ArchipelagoDir "Players\Templates"
    $PlayersDir = Join-Path $ArchipelagoDir "Players"
    $OutputDir = Join-Path $ArchipelagoDir "output"
    $ArchiveDir = Join-Path $ArchipelagoDir "Archive"

    $dateStampExecuted = $false

    Write-Host "`r`nCreating Archive folder and moving previous generated games and yaml files there with the following format`r`n daymonthyear.old`r`nTo use these files go to '$ArchiveDir', rename the files and remove the daymonthyear.old extension`r`nTip: If you do not see the .old extension, go to 'Folder Options' -> 'View Tab' -> and untick 'Hide extensions for known file types' and once done recheck it." -ForegroundColor Green

    if (-not (Test-Path $ArchiveDir)) {
        New-Item -Path $ArchiveDir -ItemType Directory | Out-Null
    }

    $ArchivePlayersDir = Join-Path $ArchiveDir "Players"
    if (-not (Test-Path $ArchivePlayersDir)) {
        New-Item -Path $ArchivePlayersDir -ItemType Directory | Out-Null
    }

    $ArchiveOutputDir = Join-Path $ArchiveDir "Output"
    if (-not (Test-Path $ArchiveOutputDir)) {
        New-Item -Path $ArchiveOutputDir -ItemType Directory | Out-Null
    }

    $OutputFiles = Get-ChildItem $OutputDir
    foreach ($File in $OutputFiles) {
        if ($File -is [System.IO.FileInfo]) {
            $DateStamp = Get-Date -Format "MM-dd-yy"
            $dateStampExecuted = $true
            $OldFilePath = Join-Path $ArchiveOutputDir ($File.Name + ".$DateStamp.old")
            Move-Item -Path $File.FullName -Destination $OldFilePath -Force
        }
    }

    $PlayersYaml = Join-Path $PlayersDir "lunacid.yaml"
    if (Test-Path $PlayersYaml) {
        $archiveExistingYAML = $(Write-Host "An existing YAML configuration file (lunacid.yaml) exists. Do you want to archive it? (Y/N)" -ForegroundColor Magenta -NoNewLine) + $(Write-Host " - Y = Yes | N = no  Enter (y\n):" -NoNewLine; Read-Host)
        if ($archiveExistingYAML -eq "Y") {
            if (-not $dateStampExecuted) {
                $DateStamp = Get-Date -Format "MM-dd-yy"
                $dateStampExecuted = $true
            }
            $OldYamlPath = Join-Path -Path $ArchivePlayersDir -ChildPath ("lunacid.yaml.$DateStamp.old")
            Move-Item -Path $PlayersYaml -Destination $OldYamlPath -Force
            
            Write-Host "`r`nMoving Lunacid.yaml from $TemplatesDir\lunacid.yaml template to $PlayersDir`r`n" -ForegroundColor Yellow
            Copy-Item -Path $TemplatesDir\lunacid.yaml -Destination $PlayersDir -Force
        }
    }

    Write-Host "`r`nOpening the Players Lunacid Yaml file in Notepad`r`nPlease edit and save the file to your liking and close notepad to continue" -ForegroundColor Magenta
    Start-Sleep -Seconds 3
    $LunacidYamlPath = Join-Path $PlayersDir "lunacid.yaml"
    Start-Process "notepad.exe" -ArgumentList $LunacidYamlPath -Wait

    Start-Sleep -Seconds 3

    Write-Host "`r`nGenerating Seed - Please wait`r`n" -ForegroundColor Yellow
    $ArchipelagoGenerate = Join-Path $ArchipelagoDir "ArchipelagoGenerate.exe"
    Start-Process $ArchipelagoGenerate
    $ArchiveDirMessage = "Enjoy the randomizer and run the script again to generate a new seed. All previous randomizer settings and generations have been moved to $ArchiveDir"
    Start-Sleep -Seconds 3

    Write-Host "Starting server in 10 seconds..." -ForegroundColor Green
    10..1 | ForEach-Object { Write-Host "Countdown: $_" -ForegroundColor DarkGreen; Start-Sleep -Seconds 1 }
    startServer($ArchipelagoDir)
}

function Restore-ArchivedFiles {
    $ArchipelagoDir = archipelagoClientLocation
    $ArchiveDir = Join-Path $ArchipelagoDir "Archive"

    if (-not (Test-Path $ArchiveDir)) {
        Write-Host "Archive directory not found. Please ensure it exists." -ForegroundColor Red
        return
    }

    Write-Host "`r`nExperimental Function to Restore Saves and Seeds from the archive directory. Proceed with caution as save file integrity is not guaranteed." -ForegroundColor DarkYellow

    $RestoreDir = Join-Path $ArchipelagoDir\output "Restore"
    if (-not (Test-Path $RestoreDir)) {
        New-Item -Path $RestoreDir -ItemType Directory | Out-Null
    }

    $ArchivedFiles = Get-ChildItem -Path $ArchiveDir -Recurse -File

    $FileGroups = $ArchivedFiles | Group-Object { $_.BaseName -replace '\.\d{6}\.\w{3,5}$' }

    $counter = 1
    Write-Host "Select a group to restore:" -ForegroundColor Yellow
    foreach ($group in $FileGroups) {
        $groupFiles = $group.Group | Sort-Object LastWriteTime -Descending
        $groupString = "{0}. {1}" -f $counter, $group.Name
        Write-Host $groupString -ForegroundColor DarkYellow
        $groupFiles | ForEach-Object {
            Write-Host "   $_" -ForegroundColor White
        }
        $counter++
    }

    $selection = Read-Host "Enter the number of the group you want to restore"

    if ($selection -lt 1 -or $selection -gt $FileGroups.Count) {
        Write-Host "Invalid selection. Please select a valid group number." -ForegroundColor Red
        return
    }

    $selectedGroup = $FileGroups[$selection - 1]

    $selectedGroup.Group | ForEach-Object {
        $fileName = $_.Name -replace '\.old', ''
        $destinationPath = Join-Path $RestoreDir $fileName
        Copy-Item -Path $_.FullName -Destination $destinationPath -Force
        Write-Host "Restored file: $destinationPath" -ForegroundColor Yellow
    }

    Write-Host "Files restored successfully to $RestoreDir" -ForegroundColor Green
}

function Show-Menu {
    Clear-Host
    
    Write-Host "Welcome to the Lunacid Archipelago Installer" -ForegroundColor Cyan
    Write-Host "-------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Please select an option:" -ForegroundColor Green
    Write-Host "1. Install - Installs Archipelago, Lunacip Ap Randomizer, and BepEnIx" -ForegroundColor Yellow
    Write-Host "2. Generate Seed and Launch Server" -ForegroundColor Yellow
    Write-Host "3. Start Archipelago Server" -ForegroundColor Yellow
    Write-Host ""
    
    Write-Host "By selecting an option, you agree to the terms and conditions. <URL>" -ForegroundColor DarkGray
    Write-Host ""
    
    $choice = Read-Host "Enter the number of your choice (1, 2, or 3)"
    
    switch ($choice) {
        '1' {
            Write-Host "`r`nInstalling Lunacid AP Randomizer, prerequisites, generating game, and starting the Archipelago server." -ForegroundColor Yellow
            install
        }
        '2' {
            Write-Host "`r`nGenerating Game and Launching Server" -ForegroundColor Yellow
            generateGame
        }
        '3' {
            Write-Host "`r`nStarting the Archipelago server." -ForegroundColor Yellow
            startServer
        }
        '4' {
            Write-Host "`r`nProceeding to Restore old saves and Seeds - !!EXPERIMENTAL!! (Close script anytime to stop)." -ForegroundColor DarkYellow
            Restore-ArchivedFiles
        }
        default {
            Write-Host "`r`nInvalid choice. `nPlease select 1, 2, or 3 by typing them into the console and pressing return\enter key." -ForegroundColor Red
            Show-Menu
        }
    }
}

Show-Menu
Write-host "`r`nPress Any Key to close this prompt`r`n" -ForegroundColor Red
cmd /c 'pause'
