<#
.DESCRIPTION
   Module to configure thea scheduled tasks settings
.EXAMPLE
   Create-NewScheduledTaskSettings -TaskName $TaskName -TaskPath $TaskPath -Compatibility Win8 -UserName $DomainUserName -Password $Password 
.NOTES
   * Assumes the TaskPath provided is valid
#>

function Set-ScheduledTaskSettings
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
        [ValidateSet('At','V1', 'Vista', 'Win7', 'Win8')]
        $Compatibility,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $DomainUserName,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Password
    )
    $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
                                             -DontStopIfGoingOnBatteries `
                                             -Compatibility $Compatibility
    Set-ScheduledTask -TaskName $TaskName `
                      -Settings $Settings `
                      -TaskPath $TaskPath `
                      -User $DomainUserName `
                      -Password $Password
}