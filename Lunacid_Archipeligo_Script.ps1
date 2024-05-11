Add-Type –assemblyName PresentationFramework
#region begin Gather release information from Git
# Retrieve Release Information
function Get-GitHubLatestRelease {
    param (
        [Parameter(Mandatory = $true)]
        [string] $repo
    )

    # URL for the latest GitHub release
    $uri = "https://api.github.com/repos/$repo/releases/latest"
    
    # Initialize WebClient
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (compatible; PowerShell/5.0)")

    try {
        # Download the JSON data
        $responseJson = $webClient.DownloadString($uri)

        # Parse the JSON
        $response = ConvertFrom-Json -InputObject $responseJson
        
        # Return the tag name
        return $response.tag_name
    } catch {
        Write-Host "Error retrieving latest release from GitHub for repository: $repo" -ForegroundColor Cyan
        return $null
    } finally {
        $webClient.Dispose()
    }
}
#endregion
#region begin Github download function
# Get GitHub Assets and Assign Variables
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
#endregion
#region begin Find Archipelago Client Install Location Function
#function to find the archipelagoClientLocation
function archipelagoClientLocation {
# Setup Archipelago Integration
# Check if Archipelago is installed using the default path
$ArchipelagoClientDir = ""

Write-Host "Getting Arhipelago installation location" -ForegroundColor Yellow

# Attempt to locate Archipelago using WMI (for Windows installations)
$archipelagoWMI = Get-WmiObject -Class Win32_Product -Filter 'Name like "%Archipelago%"' |
    Select-Object -First 1 -Property InstallLocation

if ($archipelagoWMI.InstallLocation) {
    $ArchipelagoClientDir = $archipelagoWMI.InstallLocation
}

# If not found with WMI, check the registry
if (-not $ArchipelagoClientDir) {
    $archipelagoReg = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
        ForEach-Object { Get-ItemProperty $_.PsPath } |
        Where-Object { $_.DisplayName -like "Archipelago*" } |
        Select-Object -First 1 -Property InstallLocation

    if ($archipelagoReg.InstallLocation) {
        $ArchipelagoClientDir = $archipelagoReg.InstallLocation
    }
}

# If still not found, ask the user for the location
if (-not $ArchipelagoClientDir) {
    $ArchipelagoClientDir = if (Test-Path $ArchipelagoDefaultDir) {
        $ArchipelagoDefaultDir
    } else {
        Write-Host "`r`nArchipelago client not found. Please enter the installation path:`r`n" -ForegroundColor Red
        Read-Host "Enter the installed location of the Archipelago Client"
    }
}

# If the user enters a blank path, assume Archipelago is not installed
if ([string]::IsNullOrWhiteSpace($ArchipelagoClientDir)) {
    Write-Host "Archipelago client not found. Ensure it is installed or provide a valid path." -ForegroundColor Red
    return
}

Write-Host "Archipelago client found at $ArchipelagoClientDir." -ForegroundColor Yellow
return $ArchipelagoClientDir
}
#endregion
#region begin Download File Function
# Function to download files from a given URL
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
        throw
    } finally {
        $webClient.Dispose()
    }
}
#endregion
#region begin Start Server Function
#Function to start Archipelago server
function startServer {
    param (
        [string] $ArchipelagoDir = $null
    )

    # If $ArchipelagoDir is null, call archipelagoClientLocation to set its value
    if (-not $ArchipelagoDir) {
        $ArchipelagoDir = archipelagoClientLocation
    }

    # Check if Archipelago directory is still null
    if (-not $ArchipelagoDir) {
        Write-Host "Error: Archipelago directory is not specified. Please ensure Archipelago is installed." -ForegroundColor Red
        return
    }

    $PlayersDir = Join-Path $ArchipelagoDir "Players"
    $LunacidYamlPath = Join-Path $PlayersDir "lunacid.yaml"

    # Read the Lunacid.yaml file to get the player name
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
    # Run ArchipelagoServer.exe asynchronously
    Write-Host "`r`n Please open the .zip file located in the directory named 'Ouput' inside the Archipelagos install directory in the File Explorer window that opens with the server`r`n Example File: '$ArchipelagoDir\Output\AP_88054721713267540549.zip' `r`n" -ForegroundColor Yellow
    $ArchipelagoServer = Join-Path $ArchipelagoDir "ArchipelagoServer.exe"
    Start-Process $ArchipelagoServer
    $IPADDRESS = (Get-WmiObject  -Class Win32_NetworkAdapterConfiguration).Where{$_.IPAddress}.foreach{$_.IPAddress[0]}
    Write-Host "`r`nLaunch the game and use the connection detail below to connect.`r`nIf trouble connection, please check the setup page here: https://github.com/Witchybun/LunacidAPClient/blob/main/Documentation/Setup.md" -ForegroundColor Green
    Write-Host "`r`n`r`n Player Slot: $playername `r`n Host: $IPADDRESS `r`n Port: 38281 `r`n " -ForegroundColor Cyan
    return
}
#endregion
#region begin Generate Game Function
#function to generateGame with Archipelago
function generateGame {

    # Define directories
    $ArchipelagoDir = archipelagoClientLocation
    #$ArchipelagoClientDir = $ArchipelagoDir
    $TemplatesDir = Join-Path $ArchipelagoDir "Players\Templates"
    $PlayersDir = Join-Path $ArchipelagoDir "Players"
    $OutputDir = Join-Path $ArchipelagoDir "output"
    $ArchiveDir = Join-Path $ArchipelagoDir "Archive"

    $dateStampExecuted = $false

    Write-Host "`r`nCreating Archive folder and moving previous generated games and yaml files there with the following format`r`n daymonthyear.old`r`nTo use these files go to '$ArchiveDir', rename the files and remove the daymonthyear.old extension`r`n Tip: If you do not see the .old extension, go to 'Folder Options' -> 'View Tab' -> and untick 'Hide extensions for known file types' and once done recheck it." -ForegroundColor Green

    # Create Archive directory if it doesn't exist
    if (-not (Test-Path $ArchiveDir)) {
        New-Item -Path $ArchiveDir -ItemType Directory | Out-Null
    }

    # Create Players directory in Archive if it doesn't exist
    $ArchivePlayersDir = Join-Path $ArchiveDir "Players"
    if (-not (Test-Path $ArchivePlayersDir)) {
        New-Item -Path $ArchivePlayersDir -ItemType Directory | Out-Null
    }

    # Create Output directory in Archive if it doesn't exist
    $ArchiveOutputDir = Join-Path $ArchiveDir "Output"
    if (-not (Test-Path $ArchiveOutputDir)) {
        New-Item -Path $ArchiveOutputDir -ItemType Directory | Out-Null
    }

        # Move files from Output directory to Archive\Output and rename if exists
    $OutputFiles = Get-ChildItem $OutputDir
    foreach ($File in $OutputFiles) {
        if ($File -is [System.IO.FileInfo]) {
            $DateStamp = Get-Date -Format "MM-dd-yy"
            $dateStampExecuted = $true
            $OldFilePath = Join-Path $ArchiveOutputDir ($File.Name + ".$DateStamp.old")
            Move-Item -Path $File.FullName -Destination $OldFilePath -Force
        }
    }

    # Check if existing YAML configuration file exists
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
            
        # Copy Lunacid.yaml from Templates to Players directory
        Write-Host "`r`nMoving Lunacid.yaml from $TemplatesDir\lunacid.yaml template to $PlayersDir`r`n" -ForegroundColor Yellow
        Copy-Item -Path $TemplatesDir\lunacid.yaml -Destination $PlayersDir -Force
        }
    }

    # Open Lunacid.yaml in default text editor
    Write-Host "`r`nOpening the Players Lunacid Yaml file in Notepad`r`nPlease edit and save the file to your liking and close notepad to continue" -ForegroundColor Magenta
    Start-Sleep -Seconds 3
    $LunacidYamlPath = Join-Path $PlayersDir "lunacid.yaml"
    Start-Process "notepad.exe" -ArgumentList $LunacidYamlPath -Wait

    # Wait for 3 seconds
    Start-Sleep -Seconds 3

    # Run ArchipelagoGenerate.exe
    Write-Host "`r`nGenerating Seed - Please wait`r`n" -ForegroundColor Yellow
    $ArchipelagoGenerate = Join-Path $ArchipelagoDir "ArchipelagoGenerate.exe"
    Start-Process $ArchipelagoGenerate
    $ArchiveDirMessage = "Enjoy the randomizer and run the script again to generate a new seed. All previous randomizer settings and generations have been moved to $ArchiveDir"
    Start-Sleep -Seconds 3

        # Wait for 10 seconds with countdown
    Write-Host "Starting server in 10 seconds..." -ForegroundColor Green
    10..1 | ForEach-Object { Write-Host "Countdown: $_" -ForegroundColor DarkGreen; Start-Sleep -Seconds 1 }
    startServer($ArchipelagoDir)
}
#endregion
#region begin Install Lunacid AP Rando Function
#function to Install Lunacid AP Rando
function install {

    #Initialize Directories and Security Protocols
    $ScriptDir = "$PSScriptRoot"
    $TempDir = Join-Path $ScriptDir "temp"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    cd $PSScriptRoot

    #Create Temp Directory
    if (-not (Test-Path $TempDir)) {
        New-Item -Path $TempDir -ItemType Directory | Out-Null
    }

    # Retrieve assets from GitHub
    $BepInExAssets = Get-GitHubLatestAssets -repo "BepInEx/BepInEx"
    $ArchipelagoAssets = Get-GitHubLatestAssets -repo "ArchipelagoMW/Archipelago"
    $LunacidAssets = Get-GitHubLatestAssets -repo "Witchybun/LunacidAPClient"

    # Check if assets retrieval is successful
    if ($null -eq $BepInExAssets -or $null -eq $ArchipelagoAssets -or $null -eq $LunacidAssets) {
        Write-Host "Error: Failed to retrieve one or more assets from GitHub. Exiting." -ForegroundColor Red
     return
    }

    # Assign URLs for BepInEx, Archipelago, and Lunacid
    $BepInExFilename = $BepInExAssets | Where-Object { $_.name -like "BepInEx_win*.zip" } | Select-Object -First 1
    $ArchipelagoSetup = $ArchipelagoAssets | Where-Object { $_.name -like "Setup.Archipelago*.exe" } | Select-Object -First 1
    $LunacidWorldFile = $LunacidAssets | Where-Object { $_.name -like "*.apworld" } | Select-Object -First 1
    $LunacidZipFile = $LunacidAssets | Where-Object { $_.name -like "*.zip" } | Select-Object -First 1

    # Check if variables contain valid objects with download URLs
    if ($null -eq $BepInExFilename -or $null -eq $ArchipelagoSetup -or $null -eq $LunacidWorldFile -or $null -eq $LunacidZipFile) {
        Write-Host "Error: Failed to assign asset download URLs. Exiting." -ForegroundColor Red
        return
    }

    # Download the Files
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

    # Unpack ZIP Files
    # Check if the downloaded files exist
    if (Test-Path $BepInExZip) {
        $BepInExTempDir = Join-Path $TempDir "BepInEx"
        if (-not (Test-Path $BepInExTempDir)) {
            New-Item -Path $BepInExTempDir -ItemType Directory | Out-Null
        }
        Expand-Archive -Path $BepInExZip -DestinationPath $BepInExTempDir -Force
    } else {
        Write-Host "BepInEx ZIP file not found. Check the download process." -ForegroundColor Red
    }

    if (Test-Path $LunacidApZip) {
        $LunacidApTempDir = Join-Path $TempDir "LunacidAPClient"
        if (-not (Test-Path $LunacidApTempDir)) {
            New-Item -Path $LunacidApTempDir -ItemType Directory | Out-Null
        }
        Expand-Archive -Path $LunacidApZip -DestinationPath $LunacidApTempDir -Force
    } else {
         Write-Host "Lunacid AP Client ZIP file not found. Check the download process." -ForegroundColor Red
    }

    $ArchipelagoDir = archipelagoClientLocation
    $ArchipelagoClientDir = $ArchipelagoDir

    # Query Steam Installation Paths
    # Determine if the system is 64-bit or 32-bit
    if (($env:PROCESSOR_ARCHITECTURE -eq "AMD64") -or ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64")) {
        $arch = "64-bit"
    } else {
        $arch = "32-bit"
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
        return
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
    $LunacidDir = ""

    foreach ($path in $libraryPaths) {
        $LunacidCheckPath = Join-Path $path "steamapps/common/Lunacid"
        if (Test-Path $LunacidCheckPath) {
            $LunacidDir = $LunacidCheckPath
            break
        }
    }

    # Verify if the correct installation directory is found
    if (-not $LunacidDir) {
        Write-Host "`r`nLunacid installation directory not found. Ensure Steam and Lunacid are installed." -ForegroundColor Red
        return
    }

    Write-Host "`r`nLunacid installation directory found at $LunacidDir." -ForegroundColor Yellow

    # Install Archipelago Client
    # Function to check if a specific process is running
    function Is-ProcessRunning {
        param (
            [Parameter(Mandatory = $true)]
            [string] $ProcessPattern
        )
        $processes = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like $ProcessPattern }
        return $processes.Count -gt 0
    }

    # Start the Archipelago client installation (setup) process
    if (Test-Path $ArchipelagoExe) {
        Start-Process -FilePath $ArchipelagoExe
    } else {
        Write-Host "`r`nArchipelago setup executable not found. Check the download process." -ForegroundColor Red
        return
    }

    # Check if the installation process has started
    $maxLoops = 30  # Total loops for a 5-minute timeout (300 seconds)
    $interval = 10  # Check every 10 seconds
    $processPattern = "Setup.Archipelago*"  # Wildcard pattern for setup process

    # Loop to wait for the installation process to finish
    while ($maxLoops -gt 0) {
        Start-Sleep -Seconds $interval
    
        # Log running processes for debugging (optional)
        $runningProcesses = Get-Process -ErrorAction SilentlyContinue | Select-Object -Property ProcessName

         # Check if the expected process is still running
        if (-not (Is-ProcessRunning -ProcessPattern $processPattern)) {
            Write-Host "`r`nArchipelago setup process has finished." -ForegroundColor Green
            break
        }
    
        $maxLoops--
    }

    # If the loop count reaches zero, prompt the user to continue
    if ($maxLoops -le 0) {
        Write-Host "`nThe Archipelago setup process is taking longer than expected. Please press any key to continue..." -ForegroundColor Yellow
        $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")  # Wait for any key press
    }

    # Install BepInEx
    if ($LunacidDir) {
        $BepInExDir = Join-Path $LunacidDir "BepInEx"
        if (-not (Test-Path $BepInExDir)) {
            New-Item -Path $BepInExDir -ItemType Directory | Out-Null
        }

        # Copy BepInEx to Lunacid directory
        if (Test-Path $BepInExTempDir) {
            Copy-Item -Path $BepInExTempDir -Destination $LunacidDir -Recurse -Force
        } else {
            Write-Host "BepInEx extraction directory not found. Check the unpacking process." -ForegroundColor Red
        }
    }

    # Create or ensure the Archipelago Worlds directory exists
    $ArchipelagoWorldsDir = Join-Path $ArchipelagoClientDir "lib/worlds"
    if (-not (Test-Path $ArchipelagoWorldsDir)) {
        Write-Host "Creating Archipelago Worlds directory at $ArchipelagoWorldsDir" -ForegroundColor Yellow
        New-Item -Path $ArchipelagoWorldsDir -ItemType Directory | Out-Null
    }

    # Check if the lunacid.apworld file exists
    if (Test-Path $LunacidApWorld) {
        Write-Host "Moving Lunacid.apworld to $ArchipelagoWorldsDir" -ForegroundColor Yellow
        Move-Item -Path $LunacidApWorld -Destination $ArchipelagoWorldsDir -Force
    } else {
        Write-Host "`r`nLunacid.apworld file not found at $LunacidApWorld. Check the download process." -ForegroundColor Red
    }

    # Check if LunacidAPClient extraction directory exists
    if (Test-Path $LunacidApTempDir) {
        Write-Host "Moving LunacidAPClient contents to $LunacidBepInExPluginsDir" -ForegroundColor Yellow
        Copy-Item -Path (Join-Path $LunacidApTempDir "*") -Destination $LunacidBepInExPluginsDir -Recurse -Force
    } else {
        Write-Host "`r`nLunacidAPClient extraction directory not found. Check the unpacking process." -ForegroundColor Red
    }

    # Move LunacidAPClient contents to BepInEx\Plugins
    $LunacidBepInExPluginsDir = Join-Path $BepInExDir "plugins"
    if (-not (Test-Path $LunacidBepInExPluginsDir)) {
        New-Item -Path $LunacidBepInExPluginsDir -ItemType Directory | Out-Null
    }

    if (Test-Path $LunacidApTempDir) {
        Copy-Item -Path (Join-Path $LunacidApTempDir "*") -Destination $LunacidBepInExPluginsDir -Recurse -Force
    } else {
        Write-Host "`r`nLunacidAPClient extraction directory not found. Check the unpacking process." -ForegroundColor Red
    }

    # Run Archipelago Client
    if ($ArchipelagoClientDir) {
        $ArchipelagoLauncher = Join-Path $ArchipelagoClientDir "ArchipelagoLauncher.exe"
        if (Test-Path $ArchipelagoLauncher) {
            Write-Host "`r`nInstallation and Setup complete. Archipelago Client has been started." -ForegroundColor Green
            Write-host "`r`nContinue with the In-Game setup instructions at' https://github.com/Witchybun/LunacidAPClient/blob/main/Documentation/Setup.md#in-game-setup '`r`n Tip: You can CTRL + Left-Click the link to open it in your browser or copy it out.`r`n" -ForegroundColor Green
            generateGame
        } else {
            Write-Host "`r`nArchipelagoLauncher.exe not found in Archipelago directory." -ForegroundColor Red
            }
        } else {
        Write-Host "`r`nArchipelago Client directory not specified. Please check the installation process." -ForegroundColor Red
    }
}
#endregion
#region begin Test Functions
#Test Restore Function
function Restore-ArchivedFiles {
    # Define directories
    $ArchipelagoDir = archipelagoClientLocation
    $ArchiveDir = Join-Path $ArchipelagoDir "Archive"

    # Check if Archive directory exists
    if (-not (Test-Path $ArchiveDir)) {
        Write-Host "Archive directory not found. Please ensure it exists." -ForegroundColor Red
        return
    }

    Write-Host "`r`nExperimental Function to Restore Saves and Seeds from the archive directory. Proceed with caution as save file integrity is not guaranteed." -ForegroundColor Orange

    # Create Restore directory if it doesn't exist
    $RestoreDir = Join-Path $ArchiveDir "Restore"
    if (-not (Test-Path $RestoreDir)) {
        New-Item -Path $RestoreDir -ItemType Directory | Out-Null
    }

    # Get archived files
    $ArchivedFiles = Get-ChildItem -Path $ArchiveDir -Recurse -File

    # Group archived files by their unique string
    $FileGroups = $ArchivedFiles | Group-Object { $_.BaseName -replace '\.\d{6}\.\w{3,5}$' }

    # Display file groups to the user
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

    # Prompt the user to select a group
    $selection = Read-Host "Enter the number of the group you want to restore"

    if ($selection -lt 1 -or $selection -gt $FileGroups.Count) {
        Write-Host "Invalid selection. Please select a valid group number." -ForegroundColor Red
        return
    }

    # Get the selected group
    $selectedGroup = $FileGroups[$selection - 1]

    # Copy files from the selected group to the Restore directory
    $selectedGroup.Group | ForEach-Object {
        $fileName = $_.Name -replace '\.old', ''
        $destinationPath = Join-Path $RestoreDir $fileName
        Copy-Item -Path $_.FullName -Destination $destinationPath -Force
        Write-Host "Restored file: $destinationPath" -ForegroundColor Yellow
    }

    Write-Host "Files restored successfully to $RestoreDir" -ForegroundColor Green
}
#endregion
#region begin Main Menu
function Show-Menu {
    Clear-Host
    
    # Title text
    Write-Host "Welcome to the Lunacid Archipelago Installer" -ForegroundColor Cyan
    Write-Host "-------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    
    # Body text
    Write-Host "Please select an option:" -ForegroundColor Green
    Write-Host "1. Install - Installs Archipelago, Lunacip Ap Randomizer, and BepEnIx"  -ForegroundColor Yellow
    Write-Host "2. Generate Seed and Launch Server" -ForegroundColor Yellow
    Write-Host "3. Start Archipelago Server" -ForegroundColor Yellow
    Write-Host ""
    
    # Footer text
    Write-Host "By selecting an option, you agree to the terms and conditions. <URL>" -ForegroundColor DarkGray
    Write-Host ""
    
    # Prompt user for input
    $choice = Read-Host "Enter the number of your choice (1, 2, or 3)"
    
    switch ($choice) {
        '1' {
            Write-Host "`r`nInstalling Lunacid AP Randomizer, prerequisites, generating game, and starting the Archipelago server." -ForegroundColor Yellow
            Install
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
#endregion
# Call the function to display the menu   startServer
Show-Menu
Write-host "`r`nPress Any Key to close this prompt`r`n" -ForegroundColor Red
cmd /c 'pause'