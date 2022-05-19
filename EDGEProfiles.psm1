<#
 .Synopsis
  Allows for easy backup and restore of Microsoft EDGE (Anaheim) Profiles.
  EDGE MUST BE CLOSED DURING!

 .Description
  Will backup all EDGE "User Data" for the current user. This data contains all the "Profiles" within the browser, and the corresponding registry keys will also be saved alongside the backup.
  Backups are zipped to allow for easy storage on locations like OneDrive.
  Before archiving the backup, all profiles have their Cache emptied.

  Restore will replace the current users EDGE data. The command requires that the user chooses how to handle existing data.

 .Example
   # Backup the current users EDGE Profiles to the _EdgeProfilesBackup folder in the users own OneDrive.
   Backup-EDGEProfiles

 .Example
   # Backup the current users EDGE Profiles to the users own TEMP folder.
   Backup-EDGEProfiles -Destination $env:TEMP

 .Example
   # Restore a previous backup and remove existing user data.
   Restore-EDGEProfiles -ZIPSource EDGE-UserData30July2021-MichaelMardahl.zip -REGSource EDGE-ProfilesRegistry30July2021-MichaelMardahl.reg -ExistingDataAction Remove

 .NOTES
        Author:      Michael Mardahl
        Contact:     @michael_mardahl
        Created:     2021-30-07
        Updated:     2021-31-07
        Version history:
        1.0.0 - (2021-30-07) Script created
        1.0.1 - (2021-31-07) Minor output fixes
        1.0.2 - (2021-01-08) Changed from exit codes to breaks
        1.0.3 - (2021-01-08) Changed from exit codes to breaks
        1.0.4 - (2021-01-08) Default destination validation bug fix (Thanks @byteben)
        1.0.5 - (2022-05-19) Added Support for Edge Dev and Edge SxS Profiles (Edge Dev, Edge Beta, Edge Canary)

#>
#Requires -Version 5

function Backup-EDGEProfiles {
<#
 .Synopsis
  Backup current users Microsoft EDGE (Anaheim) Profiles.

 .Description
  Will backup all EDGE "User Data" for the current user.

 .Parameter Verbose
  Enables extended output

 .Parameter Destination
  (optional)
  Location in which to save the backup ZIP and REG files
  Defaults to the users OneDrive

 .Parameter AddDate
  (optional - $true/$false)
  Applies a date stamp to the filenames.
  Defaults to $true

 .Example
   # Backup the current users EDGE Profiles to the _EdgeProfilesBackup folder in the users own OneDrive.
   Backup-EDGEProfiles

 .Example
   # Backup the current users EDGE Profiles to the users own TEMP folder.
   Backup-EDGEProfiles -Destination $env:TEMP
#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false,
                    HelpMessage="Destination of the EDGE profile backup (Defaults to OneDrive root \_EdgeProfilesBackup)")]
        [string]$Destination = (Join-Path -Path $env:OneDrive -ChildPath "\_EdgeProfilesBackup"),
        [Parameter(Mandatory=$false,
                    HelpMessage="Append the current date to the backup (Defaults to true)")]
        [bool]$AddDate = $true
    )

    #region Execute

    #Verify that the entered destination exists
    if ((-not (Test-Path $Destination) -and ($Destination -eq (Join-Path -Path $env:OneDrive -ChildPath "\_EdgeProfilesBackup")))){
        #Create default destination
        New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    }
    elseif (-not (Test-Path $Destination)){
        Write-Warning "The entered destination path could not be validated ($Destination)"
        break
    }

    #Verify EDGE is closed
    if (Get-Process msedge -ErrorAction SilentlyContinue) {
        Write-Error "EDGE is still running, please close any open EDGE Browsers and try again."
        break
    }

    Write-Output "Starting EDGE profiles backup for $($env:USERNAME) to ($Destination) - DON'T OPEN EDGE! and please wait..."
    Write-Verbose "Destination root   : $Destination"
    Write-Verbose "Append date        : $AddDate"

    #Date name addition check
    if($AddDate) {
        $dateName = (get-date -Format yyyy-mm-dd).ToString() + '-'
    } else {
        $dateName = ""
    }

    #Setting some important variables
    $edgeProfilesPath = (Join-Path -Path $env:LOCALAPPDATA -ChildPath "\Microsoft\Edge")
    $edgeDevProfilesPath = (Join-Path -Path $env:LOCALAPPDATA -ChildPath "\Microsoft\Edge Dev")
    $edgeProfilesRegistry = "HKCU\Software\Microsoft\Edge\PreferenceMACs"
    $edgeDevProfilesRegistry = "HKCU\Software\Microsoft\Edge Dev\PreferenceMACs"
    
    #Export registry key
    $regBackupDestination = = (Join-Path -Path $Destination -ChildPath "\$($dateName)EDGE-ProfilesRegistry-$($env:USERNAME).reg")
    $regBackupDestination_Dev = (Join-Path -Path $Destination -ChildPath "\$($dateName)EDGE_DEV-ProfilesRegistry-$($env:USERNAME).reg")
    Write-Verbose "Exporting Registry backup to $regBackupDestination"
    #Remove any existing destination file, else the export will stall.
    if(($regBackupDestination -ilike "*.reg") -and (Test-Path $regBackupDestination)) {
        Remove-Item $regBackupDestination -Force -ErrorAction SilentlyContinue
    }
    if(($regBackupDestination_Dev -ilike "*.reg") -and (Test-Path $regBackupDestination_Dev)) {
        Remove-Item $regBackupDestination_Dev -Force -ErrorAction SilentlyContinue
    }
    
    $regCMD = Invoke-Command {reg export "$edgeProfilesRegistry" "$regBackupDestination"}
    $regCMD_Dev = Invoke-Command {reg export "$edgeDevProfilesRegistry" "$regBackupDestination_Dev"}

    #Export user data

    #Cleaning cache
    Write-Verbose "Cleaning up cache before export."
    if(Test-Path $edgeProfilesPath){
        $cacheFolders = Get-ChildItem -Path $edgeProfilesPath -r  | Where-Object { $_.PsIsContainer -and $_.Name -eq "Cache" }
        Foreach ($folder in $cacheFolders)
        {
            $rmPath = Join-Path -Path $folder.fullname -ChildPath "\*"
            Write-Verbose "Emptying $rmPath"
            Remove-Item $rmPath -Recurse -Force
        }
        Write-Verbose "Cleanup completed."
    } else {
        Write-Error "EDGE user data folder missing - terminating!"
        break
    }
    
    if(Test-Path $edgeDevProfilesPath){
        $cacheFolders = Get-ChildItem -Path $edgeDevProfilesPath -r  | Where-Object { $_.PsIsContainer -and $_.Name -eq "Cache" }
        Foreach ($folder in $cacheFolders)
        {
            $rmPath = Join-Path -Path $folder.fullname -ChildPath "\*"
            Write-Verbose "Emptying $rmPath"
            Remove-Item $rmPath -Recurse -Force
        }
        Write-Verbose "Cleanup completed."
    } else {
        Write-Error "EDGE DEV user data folder missing - terminating!"
        break
    }

    
    #Creating ZIP Archive
    $zipBackupDestination = (Join-Path -Path $Destination -ChildPath "\EDGE-UserData$($dateName)-$($env:USERNAME).zip")
    $zipBackupDestination_Dev = (Join-Path -Path $Destination -ChildPath "\$($dateName)-EDGE_DEV-UserData-$($env:USERNAME).zip")
    Write-Verbose "Exporting user data backup to $zipBackupDestination"
    #Remove any existing destination file, else the export will fail.
    if(($zipBackupDestination -ilike "*.zip") -and (Test-Path $zipBackupDestination)) {
        Remove-Item $zipBackupDestination -Force -ErrorAction SilentlyContinue
    }
    if(($zipBackupDestination_Dev -ilike "*.zip") -and (Test-Path $zipBackupDestination_Dev)) {
        Remove-Item $zipBackupDestination_Dev -Force -ErrorAction SilentlyContinue
    }
    #Compressing data to backup location
    try {
        Get-ChildItem -Path $edgeProfilesPath | Compress-Archive -DestinationPath $zipBackupDestination -CompressionLevel Fastest
        Write-Output "EDGE Profile export completed to: $Destination"
        Get-ChildItem -Path $edgeDevProfilesPath | Compress-Archive -DestinationPath $zipBackupDestination_Dev -CompressionLevel Fastest
        Write-Output "EDGE DEV Profile export completed to: $Destination"
    } catch {
        #Error out and cleanup
        Write-Error $_
        Remove-Item $zipBackupDestination -Force -ErrorAction SilentlyContinue
        Remove-Item $regBackupDestination_Dev -Force -ErrorAction SilentlyContinue
        Write-Error "EDGE Backup failed, did you forget to keep EDGE / EDGE DEV closed?!"
        break
    }
    #endregion Execute
}

function Restore-EDGEProfiles {
<#
 .Synopsis
  Restore Microsoft EDGE (Anaheim) Profiles to the current users EDGE Browser.

 .Description
  Will restore all EDGE "User Data" for the current user from an archive created by the Backup-EDGEProfiles function.

 .Parameter Verbose
  Enables extended output

 .Parameter ZIPSource
  (Mandatory - file path)
  Location of the User Data backup archive file.

 .Parameter REGSource
  (Mandatory - file path)
  Location of the profile data registry file.

 .Parameter ExistingDataAction
  (Mandatory - Rename/Remove)
  Choose wheather to have the existing User Data removed completely or just renamed. Renaming will add a datestamp to the existing USer Data folder.

 .Example
   # Restore a previous backup and remove existing user data.
   Restore-EDGEProfiles -ZIPSource EDGE-UserData30July2021-MichaelMardahl.zip -REGSource EDGE-ProfilesRegistry30July2021-MichaelMardahl.reg -ExistingDataAction Remove
#>

    #Add the -verbose parameter to commandline to get extra output.
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
                    HelpMessage="Source of the EDGE User Data profile backup archive")]
        [string]$ZIPSource,
        [Parameter(Mandatory=$true,
                    HelpMessage="Source of the EDGE Registry profile backup file")]
        [string]$REGSource,
        [Parameter(Mandatory=$true,
                    HelpMessage="How to handle the existing profiles? Options are Backup or Remove")]
        [ValidateSet('Rename','Remove')]
        [string]$ExistingDataAction
    )

    #region Execute

    #Verify that the entered sources exits and have the right fileextention
    if(-not ((Test-Path $ZIPSource) -or (-not ($ZIPSource -ilike "*.zip")))){
        Write-Error "The entered source file could not be validated ($ZIPSource)"
        break
    }
    if(-not ((Test-Path $REGSource) -or (-not ($REGSource -ilike "*.reg")))){
        Write-Error "The entered source file could not be validated ($REGSource)"
        break
    }

    #Verify EDGE is closed
    if (Get-Process msedge -ErrorAction SilentlyContinue) {
        Write-Error "EDGE is still running, please close any open EDGE Browsers and try again."
        Break
    }

    Write-Output "Starting EDGE profiles restore for $($env:USERNAME) - (DON'T OPEN EDGE!) please wait..."
    Write-Verbose "Source archive   : $ZIPSource"
    Write-Verbose "Source registry  : $REGSource"

    #Define location of EDGE Profile for current user
    $edgeProfilesPath = (Join-Path -Path $env:LOCALAPPDATA -ChildPath "\Microsoft\Edge")

    #Handle existing User Data
    $UserData = (Join-Path -Path $edgeProfilesPath -ChildPath "\User Data")
    if (Test-Path $UserData){
        Write-Verbose "Existing User Data folder found in $edgeProfilesPath"
        if($ExistingDataAction -eq "Rename") {
            $renameFolder = "$($UserData)-$((get-date -Format ddMMMMyyyy-HHmmss).ToString())"
            Write-Verbose "Rename parameter set - Renaming folder to '$renameFolder'"
            Rename-Item $UserData $renameFolder
        }
        else {
            Write-Verbose "Remove parameter set - Deleting existing data."
            Remove-Item $UserData -Recurse -Force
        }
    }

    #Import registry key
    Write-Verbose "Importing Registry backup from $REGSource"
    $regCMD = Invoke-Command {reg import "$REGSource"}

    #Import user data
    #
    Write-Verbose "Decompressing '$ZIPSource' to $edgeProfilesPath"
    try {
        Expand-Archive -Path $ZIPSource -DestinationPath $edgeProfilesPath -Force
        Write-Output "EDGE Profile import completed to: $UserData"
    } catch {
        #Error out and cleanup
        Write-Error $_
        Remove-Item $zipBackupDestination -Force -ErrorAction SilentlyContinue
        Remove-Item $regBackupDestination -Force -ErrorAction SilentlyContinue
        Write-Error "EDGE import failed, did you forget to keep EDGE closed?!"
        break
    }
    #endregion Execute
}

Export-ModuleMember -Function Backup-EDGEProfiles, Restore-EDGEProfiles
