<#
.DESCRIPTION
   Module to grant file system Full Control rights on a NTFS file system to a declared user
.EXAMPLE
   Grant-FullControl -TargetFullName 'C:\Users\person\GIT\test\' -InheritanceFlags ContainerInherit,ObjectInherit -PropagationFlags NoPropagateInherit -DomainUserName 'STUDY-PC\Steve Jobs'
#>
function Grant-FullControl
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [String]
        $TargetFullName,
        [Parameter(Mandatory=$true)]
        [ValidateSet('None','ContainerInherit','ObjectInherit')]
        $InheritanceFlags,
        [Parameter(Mandatory=$true)]
        [ValidateSet('None','InheritOnly','NoPropagateInherit')]
        $PropagationFlags,
        [Parameter(Mandatory=$true)]
        [String]
        $DomainUserName
    )

    try
    {
        $ACL = Get-ACL -Path $TargetFullName -ErrorAction Stop
    }
    catch
    {
        foreach ($Err in $Error | Select-Object -Unique)
        {
            $Err
        }
        return
    }

    $RuleDirection = [System.Security.AccessControl.AccessControlType]'Allow'
    $GroupsOrUser = [System.Security.Principal.NTAccount]$DomainUserName
    $FileSystemRights = [System.Security.AccessControl.FileSystemRights]'FullControl'
    $Inherit = [system.security.accesscontrol.InheritanceFlags]$InheritanceFlags -join ','
    $Propagation = [system.security.accesscontrol.PropagationFlags]$PropagationFlags -join ','

    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($GroupsOrUser, $FileSystemRights, $Inherit, $Propagation, $RuleDirection)
    $ACL.AddAccessRule($AccessRule)

    try
    {
        Set-ACL -Path $AssetFullName -AclObject $ACL -ErrorAction Stop
    }
    catch
    {
        foreach ($Err in $Error | Select-Object -Unique)
        {
            $Err
        }
        return
    }
}