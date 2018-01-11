<#
.Synopsis
   Module for getting a list of all projects attached to a specific .NET solution files
.EXAMPLE
   (Get-ChildItem -Path .\* -Recurse -Include *.sln) | Get-SolutionProjects 
.EXAMPLE
   Get-SolutionProjects -Path status_console.sln
#>
function Get-SolutionProjects
{
    [CmdletBinding()]
	param
	(
        [Parameter(ValueFromPipelineByPropertyName,
                   ParameterSetName='Path to solution file')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
            if((Test-Path -Path $_ ))
            {
                return $true
            }
            else
            {
                Throw "$_ failed a Test-Path check, ensure the specific solution file exists"
            }
        })]
        [String]
		$FullName
	)

    Begin
    {
        # This DLL comes with Visual Studio
        Add-Type -Path (Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath 'Microsoft Visual Studio\2017\Community\MSBuild\15.0\Bin\Microsoft.Build.dll')
        ($SolutionList = [System.Collections.Generic.List[System.Collections.Generic.List[System.Object]]]("")).RemoveAt(0)
    }
    Process
    {
        $SolutionFile = Get-ChildItem -Path $FullName
        $Solution = [Microsoft.Build.Construction.SolutionFile]::Parse($SolutionFile.FullName)
        ($ProjectsList = [System.Collections.Generic.List[System.Object]]("")).RemoveAt(0)
        $Solution.ProjectsInOrder |
        # This filters out web site projects that are part of the solution e.g. a React website
        Where-Object { $_.ProjectType -eq 'KnownToBeMSBuildFormat'} |
            ForEach-Object {
                # This is used to determine if it is a web project (web projects have 2 output directories (a bin folder and a web folder))
                $IsWebProject = (Select-String -Pattern "<UseIISExpress>.+</UseIISExpress>" -Path $_.AbsolutePath) -ne $null
                $Object = New-Object System.Object
                $Object | Add-Member -MemberType NoteProperty -Name AttachedSolutionFullName -Value $SolutionFile.FullName
                $Object | Add-Member -MemberType NoteProperty -Name IsWebProject -Value $IsWebProject
                $Object | Add-Member -MemberType NoteProperty -Name FullName -Value $_.AbsolutePath
                $Object | Add-Member -MemberType NoteProperty -Name DirectoryFullName -Value "$(Split-Path -Path $_.AbsolutePath -Resolve)"
                $Object | Add-Member -MemberType NoteProperty -Name Name -Value $_.ProjectName
                $ProjectsList.Add($Object)
            }
        $SolutionList.Add($ProjectsList)
    }
    End
    {
        return $SolutionList
    }
}
