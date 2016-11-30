<#
.DESCRIPTION
   Module to create a task scheduler folder if it does not exist
.EXAMPLE
   New-ScheduledTaskFolder -TaskPath '\Tasks'
.NOTES
   * Handles creating folders along the entire declared path without any problems
#>

function New-TaskSchedulerFolder
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $TaskPath
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
    $ScheduleObject = New-Object -ComObject Schedule.Service
    $ScheduleObject.connect()
    $RootFolder = $ScheduleObject.GetFolder('\')
    try
    {
        $null = $ScheduleObject.GetFolder($WorkingTaskPath)
    }
    catch
    {
        $null = $RootFolder.CreateFolder($WorkingTaskPath)
        # The below turns on history logging for the event viewer.
        $MMC_View = 'Microsoft-Windows-TaskScheduler/Operational'
        $EvenLogConfig = New-Object System.Diagnostics.Eventing.Reader.EventLogConfiguration $MMC_View
        $EvenLogConfig.IsEnabled = $true
        $EvenLogConfig.SaveChanges()
    }
}