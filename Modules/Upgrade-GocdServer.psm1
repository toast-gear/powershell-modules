<#
.Synopsis
   Module for upgrading the GoCD Server
#>
function Upgrade-GocdServer
{
    [CmdletBinding()]
    Param
    (
        [String]
        [ValidateNotNullOrEmpty()]
        $GocdRootDirectoryFullName,
        [String]
        [ValidateNotNullOrEmpty()]
        $DestinationBackupDirectoryFullName,
        [String]
        [ValidateNotNullOrEmpty()]
        $DeployedArtefactDirectoryFullName,
        [String]
        [ValidateNotNullOrEmpty()]
        $ServiceName,
        [String]
        [ValidateNotNullOrEmpty()]
        $InstallerFullName
    )
}

Write-Host "$(Get-Date -Format s) : Cleaning down the backup location"
$Files = Get-ChildItem -Path $InstallerFullName -ErrorAction Stop
if($files.Count -lt 1)
{
    foreach ($Err in $Error | Select-Object -Unique)
    {
        $Err
    }
    Write-Error "$(Get-Date -Format s) : No files were found in the declared installer path. Exiting script."
    Exit
}
elseif($files.Count -gt 1)
{
    foreach ($Err in $Error | Select-Object -Unique)
    {
        $Err
    }
    Write-Error "$(Get-Date -Format s) : More than 1 installer was found in the installer path. Exiting script."
    Exit
}

Write-Host "$(Get-Date -Format s) : Installer Found - $($File.Name)"
Write-Host "$(Get-Date -Format s) : Stopping GoCD Windows Service"
try
{
    Get-Service -Name $ServiceName -ErrorAction Stop | Stop-Service -ErrorAction Stop -ErrorAction Stop
}
catch
{
    foreach ($Err in $Error | Select-Object -Unique)
    {
        $Err
    }
    Write-Error "$(Get-Date -Format s) : Errors detected when trying to stop the Windows Service."
    Exit
}

Write-Host "$(Get-Date -Format s) : Clearing configured target backup directory"
Get-Item -Path $DestinationBackupDirectoryFullName -Force | Remove-Item -Force -Recurse

Write-Host "$(Get-Date -Format s) : Starting backup process"
Get-ChildItem -Path $DestinationBackupDirectoryFullName | Remove-Item -Force
Write-Host "$(Get-Date -Format s) : Backing up deployed cruise-config.xml"
Get-Item -Path (Join-Path -Path $GocdRootDirectoryFullName -ChildPath 'config\cruise-config.xml') | Copy-Item -Destination $DestinationBackupDirectoryFullName
Write-Host "$(Get-Date -Format s) : Backing up deployed cipher"
Get-Item -Path (Join-Path -Path $GocdRootDirectoryFullName -ChildPath 'config\cipher') | Copy-Item -Destination $DestinationBackupDirectoryFullName
Write-Host "$(Get-Date -Format s) : Backing up deployed db"
robocopy.exe (Join-Path -Path $GocdRootDirectoryFullName -ChildPath 'db') $DestinationBackupDirectoryFullName /MIR /LOG+:$(Join-Path -Path $GocdBackupDirectoryFullName -ChildPath 'log.txt')
Write-Host "$(Get-Date -Format s) : Backing up deployed Artefacts"
robocopy.exe $DeployedArtefactDirectoryFullName $DestinationBackupDirectoryFullName /MIR /LOG+:$(Join-Path -Path $GocdBackupDirectoryFullName -ChildPath 'log.txt')

Write-Host "$(Get-Date -Format s) : Starting GoCD Upgrade"
& $InstallerFullName /S