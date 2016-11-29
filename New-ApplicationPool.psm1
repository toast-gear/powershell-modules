<#
.Synopsis
   This module creates an IIS application pool with teh configured values
.DESCRIPTION
   
.EXAMPLE  
   New-ApplicationPool -UserDomain 'DOMAIN' `
                       -UserName 'Steve' `
                       -UserPassword 'KhalaBindsUs' `
                       -EnableCredentialValidation False `
                       -PoolName 'NameOfPool' `
                       -RuntimeVersion v2.0 `
                       -ManagedPipelineMode classic `
                       -Verbose 4>&1 | Out-File -FilePath 'C:\log.txt'
.INPUTS
   [string] UserDomain 
   [string] UserName 
   [string] UserPassword 
   [string] PoolName 
   [boolean] EnableCredentialValidation 
   [ValidateSet('v2.0', 'v4.0')] RuntimeVersion
   [ValidateSet('integrated', 'classic')] ManagedPipelineMode
   [Int64] MaximumWorkerProcesses
.OUTPUTS
   * Verbose output   
#>

function New-ApplicationPool
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
		[String]
        $UserDomain,
        [Parameter(Mandatory=$true)]
		[String]
        $UserName,
        [Parameter(Mandatory=$true)]
		[String]
        $UserPassword,
        [Parameter(Mandatory=$true)]
        [ValidateSet($true,$false)]
        $EnableCredentialValidation,
        [Parameter(Mandatory=$true)]
        [String]
        $PoolName,
        [Parameter(Mandatory=$true)]
        [ValidateSet('v2.0','v4.0')]
        $RuntimeVersion,
        [Parameter(Mandatory=$true)]
        [ValidateSet('integrated','classic')]
        $ManagedPipelineMode,
        [Parameter(Mandatory=$false)]
        [Int32]
        $MaximumWorkerProcesses
    )

    ##### ##### ##### ##### ##### ##### ##### ##### ##### #####
    ##### ##### ##### VALIDATE USER PROVIDED  ##### ##### #####
    ##### ##### ##### ##### ##### ##### ##### ##### ##### #####

    if($EnableCredentialValidation -eq $true)
    {
        $Assembly1 = [System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement")
        if($Assembly1 -eq $null)
        {
			throw "Assemly failed to import. Check permissions of user running the script."
            EXIT
        }
        try
        {
            $DomainObj = [System.DirectoryServices.AccountManagement.ContextType]::Domain
            $PrincipalContextObject = New-Object System.DirectoryServices.AccountManagement.PrincipalContext -ArgumentList $DomainObj,$UserDomain
            if(($PrincipalContextObject.ValidateCredentials($UserName,$UserPassword)) -eq $false)
            {
                throw "Credentials failed validation. Check supplied values and user status."
				EXIT
            }
        }
        catch
        {
            throw "Error on validating credentials"
            EXIT
        }
    }

    ##### ##### ##### ##### ##### ##### ##### #####
    ##### ##### ##### CREATE POOL ##### ##### #####
    ##### ##### ##### ##### ##### ##### ##### #####

    try
    {
            Import-Module -Name WebAdministration -ErrorAction Stop
    }
    catch
    {
            throw "Error caught when importing the WebAdministration module, check for permissions"
            EXIT
    }
    try
    {
		$PoolFullName = [System.IO.Path]::Combine(IIS:\AppPools, $PoolName)
        if(!(Test-Path $PoolFullName)
        {
            $CurrentPath = Get-Location
            IIS:
            New-Item $PoolFullName
            $Pool = Get-Item $PoolFullName -ErrorAction Stop
            $Pool.processModel.userName = [System.IO.Path]::Combine($UserDomain,$UserName)
            $Pool.processModel.password = $UserPassword
            $Pool.processModel.identityType = 3
            $Pool.managedRuntimeVersion = $RuntimeVersion
            $Pool.managedPipelineMode = $ManagedPipelineMode
            if(!($MaximumWorkerProcesses -eq $null))
            {
                $Pool.processModel.maxProcesses = $MaximumWorkerProcesses
            }
            Set-Item -Path $PoolFullName -Value $Pool
            Set-Location $CurrentPath
        }
        else
        {
            throw "A pool with the supplied name already exists. Each pool must have a unique name"
            EXIT
        }
    }
    catch
    {
        throw "Error on setting application pool configuration"
        EXIT
    }
}