<#
.Synopsis
   Module for upgrading GoCD
#>
function Upgrade-GoCD
{
    [CmdletBinding()]
    Param
    (
        [String]
        [ValidateNotNullOrEmpty()]
        $GoCDRootDirectoryDullName,
        [String]
        [ValidateNotNullOrEmpty()]
        $GocdBackupDirectoryFullName,
        [String]
        [ValidateNotNullOrEmpty()]
        $GoCDArtefactDirectoryFullName,
        [String]
        [ValidateNotNullOrEmpty()]
        $GoCDServiceName,
        [String]
        [ValidateNotNullOrEmpty()]
        $GoCDInstallerFullName
    )
}

$Files = Get-ChildItem -Path $GoCDInstallerFullName
if($files.Count -ne 1)
{
    foreach ($Err in $Error | Select-Object -Unique)
    {
        $Err
    }
    Write-Error "$(Get-Date -Format s) : More than 1 installer was found in the declared path. Exiting script."
    Exit
}

Write-Verbose "$(Get-Date -Format s) : Installer Found - $($File.Name)"
try
{
    Get-Service -Name $GoCDServiceName | Stop-Service -ErrorAction Stop
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
$Files = Get-ChildItem -Path $GocdBackupDirectoryFullName -Recurse -Force | Where-Object { $_.Attributes -eq 'Archive' }
foreach ($File in $Files)
{
    Remove-Item $File -Force
}
Get-ChildItem -Path $GocdBackupDirectoryFullName | Remove-Item -Force
Write-Host "$(Get-Date -Format s) : Backing up deployed cruise-config.xml"
Get-ChildItem -Path ([System.IO.Path]::Combine($GoCDRootDirectoryDullName, 'config', 'cruise-config.xml')) | Copy-Item $GocdBackupDirectoryFullName
Write-Host "$(Get-Date -Format s) : Backing up deployed cipher"
Get-ChildItem -Path ([System.IO.Path]::Combine($GoCDRootDirectoryDullName, 'config', 'cipher')) | Copy-Item $GocdBackupDirectoryFullName
Write-Host "$(Get-Date -Format s) : Backing up deployed db"
Get-ChildItem -Path ([System.IO.Path]::Combine($GoCDRootDirectoryDullName, 'db')) | Copy-Item $GocdBackupDirectoryFullName -Recurse -Force
Write-Host "$(Get-Date -Format s) : Backing up deployed Artefacts"
Get-ChildItem -Path $GoCDArtefactDirectoryFullName | Copy-Item $GocdBackupDirectoryFullName -Recurse -Force

Write-Host "$(Get-Date -Format s) : Starting GoCD Upgrade"
Write-Verbose "$(Get-Date -Format s) : Option provided:"
Write-Verbose "$(Get-Date -Format s) : /S"
& $GoCDInstallerFullName /S