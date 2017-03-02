<#
.Synopsis
   This is a module for starting a MS SQL job via a PowerShell script.
.DESCRIPTION
   WARNING ----- CONTAINS DYNAMIC PARAMETERS - USE Get-Help Invoke-SQLJob -Full to see a complete list of parameters ----- WARNING
.EXAMPLE
   Start-MSSQLAgentJob -MS_SQL_Server '.' -EnableWindowsAuthentication $false -MS_SQL_Login 'My_SQL_UserName' -MS_SQL_Password 'My_SQL_Password' -MS_SQL_AgentJob_ID 'CEF69218-756E-4B63-B5A1-51E561CE9B09' -TimeoutInSeconds 500
.EXAMPLE
   Start-MSSQLAgentJob -MS_SQL_Server '.' -EnableWindowsAuthentication $true -MS_SQL_AgentJob_ID 'CEF69218-756E-4B63-B5A1-51E561CE9B09' -TimeoutInSeconds 500
.INPUTS
   [String]    MS_SQL_Server
   [Boolean]   EnableWindowsAuthentication
   [String]	   MS_SQL_Job_ID
   [Int16]	   TimeoutInSeconds
   [String]    MS_SQL_Login                        (Dynmaic)
   [String]    MS_SQL_Password                     (Dynmaic)
.NOTES
   * Tested on Windows 8.1 against SQL Server 2008 R2. Should also work against anything newer than 2008 R2
   * Possible job statuses contained in link below. At the we only handle the Idle status in the module. May want to look into handling more of them as the module develops.
   https://msdn.microsoft.com/en-us/library/microsoft.sqlserver.management.smo.agent.jobexecutionstatus%28v=sql.120%29.aspx   
#>
function Start-MSSQLAgentJob
{
    [CmdletBinding()]
    Param
    (
        [ValidateScript({ $_ -is [string] })]
        [Parameter(Mandatory=$true)]
        $MS_SQL_Server,
        [ValidateSet($true,$false)]
        [Parameter(Mandatory=$true)]
        $EnableWindowsAuthentication,
        [String]
        [ValidateNotNullOrEmpty()]
        [Parameter(Mandatory=$true)]
        $MS_SQL_AgentJob_ID,
		[Int16]
		[Parameter(Mandatory=$true)]
		$TimeoutInSeconds
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
            $MS_SQL_LoginParam = New-Object System.Management.Automation.RuntimeDefinedParameter('MS_SQL_Login', [string], $ParameterAttributeCollection)
            $MS_SQL_PasswordParam = New-Object System.Management.Automation.RuntimeDefinedParameter('MS_SQL_Password', [string], $ParameterAttributeCollection)
            
            # Create dictionary to hold the runtime object exposing it at to the runspace
            $ParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $ParameterDictionary.Add('MS_SQL_Login', $MS_SQL_LoginParam)
            $ParameterDictionary.Add('MS_SQL_Password', $MS_SQL_PasswordParam)
        }

        # Return collection
        return $ParameterDictionary
    }

    process
    {
		# Load the assembly required for this module to work, this method will load w/e version is avaliable on your system but we are not using anything special so thats fine
		# Due to the use of the LoadWithPartialName method checking to see if the variable is null is the only accurate way of determining if a module was loaded.
		$Assembly1 = [Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.Smo')
		if ($Assembly1 -eq $null)
		{
			throw "Import of Microsoft.SqlServer.Smo failed, you most likely are not operating with admin rights or SQL is not installed on the machine that is trying to import the assembly"
		}

		# Creating the connection object so you can authenticate
		$ServerObject = New-Object ('Microsoft.SqlServer.Management.Smo.Server') ($MS_SQL_Server)
		$ConnectionObject = $ServerObject.ConnectionContext
        $ConnectionObject.DatabaseName = 'msdb'
		if($EnableWindowsAuthentication -eq $false)
		{
			$ConnectionObject.LoginSecure = $false
			$ConnectionObject.Login = $SQL_LoginParam.Value
			$ConnectionObject.Password = $SQL_PasswordParam.Value
		}
		else
		{
			$ConnectionObject.LoginSecure = $true     
		}
		try
		{
			$Job = $serverObject.JobServer.Jobs | Where-Object { $_.JobID -eq $MS_SQL_AgentJob_ID }
			$ConnectionObject.ExecuteNonQuery("EXEC dbo.sp_start_job @job_id = '$($Job.JobID)'")
            $Gandalf = 0
			while ($TimeoutInSeconds -gt $Gandalf)
			{
				$Gandalf++
				Start-Sleep -Seconds 1
				$Job.Refresh()
				# CurrentRunStatus returns a JobExecutionStatus object, could not find any other way of successfully comparing the status.
				if($Job.CurrentRunStatus -eq [Microsoft.SqlServer.Management.Smo.Agent.JobExecutionStatus]::Idle)
				{   
                    <#
                        KEEPING FOR THE MOMENT, IT WAS USED FOR LOGGING WHICH I HAVE PULLED OUT UNTIL i HAVE FIGURED OUT HOW TO USE THE STREAMS CORRECTLY
					    Select the highest number ID to get the latest message. Then put that value into a variable so we can select the message to log
					    Includes a Select First just in case multiple values are in found for some reason. First used instead of last as highest numbers are 
					    searched first

					    $InstanceID = $job.EnumHistory() | Select-Object { $_.InstanceID } | Select -ExpandProperty ' $_.InstanceID ' |  Select -First 1
					    $Result = $job.EnumHistory() | Where-Object { $_.InstanceID -eq $InstanceID}
					    $JobMessage = $Result.Message
                    #>
					break
				}
				if($TimeoutInSeconds -eq $Gandalf)
				{

                    <#
                        KEEPING FOR THE MOMENT, IT WAS USED FOR LOGGING WHICH I HAVE PULLED OUT UNTIL i HAVE FIGURED OUT HOW TO USE THE STREAMS CORRECTLY
					    Select the highest number ID to get the latest message. Then put that value into a variable so we can select the message to log
					    Includes a Select First just in case multiple values are in found for some reason. First used instead of last as highest numbers are 
					    searched first
					
                        $InstanceID = $Job.EnumHistory() | Select-Object { $_.InstanceID } | Select -ExpandProperty ' $_.InstanceID ' |  Select -First 1
					    $Result = $Job.EnumHistory() | Where-Object { $_.InstanceID -eq $InstanceID}
					    $JobMessage = $Result.Message

					    $jobState = $job.CurrentRunStatus
					    $jobStep = $job.CurrentRunStep
					    Write-Verbose -Message "$StartDateTime : Timeout Reached:"
					    Write-Verbose -Message "$StartDateTime : Job Status Read as - $jobState"
    					Write-Verbose -Message "$StartDateTime : Job Current Run Step Read as - $jobStep"
					    Write-Verbose -Message "$StartDateTime : Job Message Read as - $jobMessage"
                    #>
					break
				}
			}
		}
		catch
		{
            $_.Exception
            $_.Exception.InnerException
            $_.Exception.InnerException.InnerException
		}
	}
}