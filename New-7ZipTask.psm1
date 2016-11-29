<#
.Synopsis
   Wrapper for the 7-Zip application
.DESCRIPTION
   CONTAINS DYNAMIC PARAMETERS - USE Get-Help New-7ZipTask -Full to see a complete list of parameters - CONTAINS DYNAMIC PARAMETERS
.EXAMPLE
   ##########
   New-7zipTask -SourceFullName 'C:\InputFolder' `
                -DestinationDirectoryFullName 'C:\OutputFolder' `
                -OutputBaseName 'Output' `
                -OutputArtefact zip `
                -CompressionLevel 5 `
                -EnablePasswordProtection True `
                -Password 'MyPassword'
.EXAMPLE
   ##########
   New-7zipTask -SourceFullName 'C:\InputFolder' `
                -DestinationDirectoryFullName 'C:\OutputFolder' `
                -OutputBaseName 'Output' `
                -OutputArtefact zip `
                -CompressionLevel 5 `
                -EnablePasswordProtection False
   #################
.INPUTS
   [String] $SourceFullName
   [String] $DestinationDirectoryFullName
   [String] $OutputBaseName
   [ValidateSet('zip', '7z')] $OutputArtefact
   [ValidateSet(0,1,3,5,7,9)] $CompressionLevel
   [ValidateSet($true, $false)] $EnablePasswordProtection
   [String] $Password (Dynamic)
.OUTPUTS
   * Verbose stream
.NOTES

#>

function New-7ZipTask
{
    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $SourceFullName,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $DestinationDirectoryFullName,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $OutputBaseName,
        [Parameter(Mandatory=$true)]
        [ValidateSet('zip','7z')]
        $OutputArtefact,
        [Parameter(Mandatory=$true)]
        [ValidateSet(0,1,3,5,7,9)]
        [Int32]
        $CompressionLevel,
        [ValidateSet($true, $false)]
        [Parameter(Mandatory=$true)]
        $EnablePasswordProtection

    )
    DynamicParam
    {
		# You can chain this to support multiple seperate IF clauses which is pretty neat
        if($EnablePasswordProtection -eq $true)
        {
            # Creates parameter object
            $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
            $ParameterAttribute.Mandatory = $true

            # Create a collection to hold the parameter in and add our parameter to it.
            $ParameterAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $ParameterAttributeCollection.Add($ParameterAttribute)

            # Below is how you would add a dynamic attribute validation block. Just pass it a System.String[]. Not useful for this moduel but I am putting it here
            # as I think this could be useful going foward for something. 
            # $ParameterAttributeCollection.Add((New-Object System.Management.Automation.ValidateSetAttribute((Get-ChildItem C:\TheAwesome -File | Select-Object -ExpandProperty Name))))

            # Add validation
            $ParameterAttributeCollection.Add((New-Object System.Management.Automation.ValidateNotNullOrEmptyAttribute))
            
            # Create the runtime constructor and add our collection
            $PasswordParam = New-Object System.Management.Automation.RuntimeDefinedParameter('Password', [string], $ParameterAttributeCollection)
            
            # Create dictionary to hold the runtime object exposing it at to the runspace
            $ParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $ParameterDictionary.Add('Password', $PasswordParam)

        }
        # Return collection
        return $ParameterDictionary
    }

    process
    {
        $StartDateTime = Get-Date -Format s
		
		# Better approach would be to grab install path from registry or something like that
        $X86BinaryFullName = [System.IO.Path]::Combine(${Env:ProgramFiles(x86)}, '7-Zip\7z.exe')
        $X64BinaryFullName = [System.IO.Path]::Combine($Env:ProgramFiles, '7-Zip\7z.exe')      

        switch($OutputArtefact)
        {
            zip
            {
                if($EnablePasswordProtection -eq $true)
                {
                    [Array]$Arguments = "a", "-tzip", $([System.IO.Path]::Combine($DestinationDirectoryFullName, "$OutputBaseName.zip")), "$SourceFullName", "-r", "-mx=$($CompressionLevel)", "-bd", '-mmt=on', '-y', "-P$($PasswordParam.Value)"
                }
                else
                {
                    [Array]$Arguments = "a", "-tzip", $([System.IO.Path]::Combine($DestinationDirectoryFullName, "$OutputBaseName.zip")), "$SourceFullName", "-r", "-mx=$($CompressionLevel)", "-bd", '-mmt=on', '-y'
                }
                break
            }
            7z
            {
                if($EnablePasswordProtection -eq $true)
                {
                    [Array]$Arguments = "a", "-t7z", $([System.IO.Path]::Combine($DestinationDirectoryFullName, "$OutputBaseName.7z")), "$SourceFullName", "-r", "-mx=$($CompressionLevel)", "-bd", '-mmt=on', '-y', "-P$($PasswordParam.Value)"
                }
                else
                {
                    [Array]$Arguments = "a", "-t7z", $([System.IO.Path]::Combine($DestinationDirectoryFullName, "$OutputBaseName.7z")), "$SourceFullName", "-r", "-mx=$($CompressionLevel)", "-bd", '-mmt=on', '-y'
                }
                break
            }
        }

        if (Test-Path $X64BinaryFullName)
        {
            & $X64BinaryFullName $Arguments
        }
        elseif(Test-Path $X86BinaryFullName)
        {         
            & $X86BinaryFullName $Arguments
        }
        else
        {
            throw "Could not find the 7-Zip binary to perform task."
        }
        switch($LASTEXITCODE)
        {
            0
            {
                Write-Verbose -Message "$StartDateTime : 0 (success) report as last exit code"
                Break
            }

            1
            {
                throw "Warning (Non fatal error(s)). For example, one or more files were locked by some other application, so they were not compressed"

                Break
            }

            2
            {
                throw "Fatal error"
                Break
            }

            7
            {
                throw "Command line error"
                Break
            }

            8
            {
                throw "Not enough memory for operation."
                Break
            }
            255
            {
                throw "User stopped the process"
                Break
            }
        }
    }
}