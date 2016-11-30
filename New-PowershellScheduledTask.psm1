<#
.DESCRIPTION
   Module to create a scheduled task
.EXAMPLE
   New-PowershellScheduledTask -TaskName $TaskName -TaskPath $TaskPath -TargetFileFullName $TargetFileStoreFullName -TriggerTime $Trigger -RunLevel $RunWithHighestPrivilegesHolder -UserName $DomainUserName -Password $Password
.NOTES
   * Assumes the TaskPath provided is valid
#>

function New-PowershellScheduledTask
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $TaskName,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $TaskPath,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $TargetFileFullName,
        [Parameter(Mandatory=$true)]
        [DateTime]
        $HourMinuteTriggerTime,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Highest','LUA')]
        $RunLevel,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $DomainUserName,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Password
    )
    ($SplitHolder = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
    foreach ($item in $TaskPath.Split('\'))
    {
        if(!([System.String]::IsNullOrWhiteSpace($item)))
        {
            $SplitHolder.Add($item)
        }
    }
    $WorkingTaskPath = "\$($SplitHolder -join '\')"
    $Action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-File $TargetFileFullName -NoLogo -Noninteractive -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden"
    $Trigger =  New-ScheduledTaskTrigger -Daily -At $(Get-Date $TriggerTime -Format 'HH:mm' -ErrorAction Stop)
    Register-ScheduledTask -Action $Action `
                           -Trigger $Trigger `
                           -TaskName $TaskName `
                           -TaskPath $WorkingTaskPath `
                           -User $DomainUserName `
                           -Password $Password `
                           -RunLevel $RunLevel
}