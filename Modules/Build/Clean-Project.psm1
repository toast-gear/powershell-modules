<#
.Synopsis
   Module for getting a list of all projects attached to a specific .NET solution files
.EXAMPLE
   (Get-ChildItem -Path .\* -Recurse -Include *.csproj) | Clean-Project 
.EXAMPLE
   Clean-Project -FullName someProject.csproj
#>
function Clean-Project
{
    [CmdletBinding()]
	param
	(
        [Parameter(ValueFromPipelineByPropertyName,
                   ParameterSetName='Path to project file')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
            if((Test-Path -Path $_ ))
            {
                return $true
            }
            else
            {
                Throw "$_ failed a Test-Path check, ensure the specific project file exists"
            }
        })]
        [String]
		$FullName
	)

    Begin
    {
        Write-Output 'Cleaning supplied projects for a fresh build'
    }
    Process
    {
        $ProjectFile = Get-ChildItem -Path $FullName
        Remove-Item "$($ProjectFile.Directory.FullName)\bin" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
        Remove-Item "$($ProjectFile.Directory.FullName)\obj" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    }
    End
    {
        return
    }
}
