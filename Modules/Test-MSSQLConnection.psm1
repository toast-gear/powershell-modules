<#
.Synopsis
   Module to connect to a MS SQL server to test the connection / credentials
.DESCRIPTION
   CONTAINS DYNAMIC PARAMETERS - USE Get-Help Test-MSSQLConnection -Full to see a complete list of parameters - CONTAINS DYNAMIC PARAMETERS
.EXAMPLE
   Test-MSSQLConnection -MS_SQL_Server . -EnableWindowsAuthentication $False -MS_SQL_Login 'test' -MS_SQL_Password 'test'
.INPUTS
   [string] MS_SQL_Server
   [Boolean] EnableWindowsAuthentication
   [string] MS_SQL_Login (Dynamic)
   [string] MS_SQL_Password (Dynamic)
.NOTES
#>
function Test-MSSQLConnection
{
    [CmdletBinding()]
        Param
        (
            [Parameter(Mandatory=$true)]
            [String]
            $MS_SQL_Server,
            [ValidateSet($true,$false)]
            [Parameter(Mandatory=$true)]
            $EnableWindowsAuthentication
        )
        DynamicParam
        {
            if($EnableWindowsAuthentication -eq $false)
            {
                # Creates parameter object
                $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
                $ParameterAttribute.Mandatory = $true

                # Create a collection to hold the parameter in and add our parameter to it.
                $ParameterAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
                $ParameterAttributeCollection.Add($ParameterAttribute)

                # Create the runtime constructor and add our collection
                $MS_SQL_Login_Param = New-Object System.Management.Automation.RuntimeDefinedParameter('MS_SQL_Login', [string], $ParameterAttributeCollection)
                $MS_SQL_Password_Parameter = New-Object System.Management.Automation.RuntimeDefinedParameter('MS_SQL_Password', [string], $ParameterAttributeCollection)

                # Create dictionary to hold the runtime object exposing it at to the runspace
                $ParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
                $ParameterDictionary.Add('MS_SQL_Login', $MS_SQL_Login_Param)
                $ParameterDictionary.Add('MS_SQL_Password', $MS_SQL_Password_Parameter)
            }

            # Return collection
            return $ParameterDictionary
        }

    process
    {
        # Can't use a try catch with the LoadWithPartialName method. Method does not throw error if the assembly can;t be found. 
        # We must check the variable is not null instead
        $Assembly = [Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.Smo')
        if($Assembly -eq $null)
        {
            throw "Import of Microsoft.SqlServer.Smo failed, you most likely are not operating with admin rights"      
        }
        $ServerObject = New-Object ('Microsoft.SqlServer.Management.Smo.Server') ($MS_SQL_Server)
        $ConnectionObject = $ServerObject.ConnectionContext
        if($EnableWindowsAuthentication -eq $true)
        {
            $ConnectionObject.LoginSecure = $true
        }
        else
        {
            $ConnectionObject.LoginSecure = $false
            $ConnectionObject.Login = $MS_SQL_Login_Param.Value
            $ConnectionObject.Password = $SQL_Password_Param.Value
        }
        try
        {
            $ServerObject.ConnectionContext.Connect()
            $ServerObject.ConnectionContext.Disconnect()
        }
        catch
        {
            $_.Exception
            $_.Exception.InnerException
            $_.Exception.InnerException.InnerException
        }
    }
}