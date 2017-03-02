<#
.Synopsis
   This module runs MS SQL scripts
.DESCRIPTION
   This uses the SMO assemblies to run scripts into a MS SQL DB and provide feedback.
.EXAMPLE
   $folders = 'C:\Scripts\Main','D:\Scripts\Second\'
   $URL = 'http://dashboard.com/api/v1/logs/deploymentlogpost'
   Invoke-ScriptRunnerSMO -SQL_Server 'STUDY-PC\SQLEXPRESS' `
                          -EnableWindowsAuthentication $false `
                          -SQL_Login 'username' `
                          -SQL_Password 'password' `
                          -InitialDatabase 'target-database-name' `
                          -SQL_ScriptsLocations $folders `
                          -EnableTransactions $true `
                          -Enable_API_Logging $true `
                          -URL_To_API $URL `
                          -Company_ID 1 `
                          -Product_ID 1 `
                          -Environment_ID 10 `
                          -EnableVersioning $true `
                          -EnableVersionCheck $true `
                          -EnableEmailNotifications $false `
                          -FromEmailAddress 'name.surname@domain.com' `
                          -ToSuccessEmailAddresses 'name.surname@domain.com'`
                          -ToFailEmailAddresses 'name.surname@domain.com', 'name.surname@domain.com'`
                          -SMTP_Server '10.201.176.10'| Out-File -FilePath 'C:\Output.txt' -Append
.INPUTS
   -SQL_Server [String]
   -EnableWindowsAuthentication [Boolean]
   -SQL_Login [String]
   -SQL_Password [String]
   -InitialDatabase [String]
   -SQL_ScriptsLocations [String]
   -EnableTransactions [Boolean]
   -EnableVersioning [Boolean]
   -EnableVersionCheck [Boolean]
   -Enable_API_Logging [Boolean]
   -URL_To_API [string]
   -Company_ID [int]
   -Product_ID [int]
   -Environment_ID [int]
   -EnableEmailNotifications [Boolean]
   -FromEmailAddress [String]
   -ToSuccessEmailAddresses [array]
   -ToFailEmailAddresses [array]
   -SMTP_Server [String]
.OUTPUTS
   * All output is sent to the verbose stream when the output stream is sued you get blank lines logged from the posts to the API
   * All handled errors that should terminate the module are swsent to the error stream. Exceptions from scripts should not stop the script so they go to the output stream.
.NOTES
   * This has been tested against 2008 R2 and 2012.
   * The declared DB is just the initial DB. This can be changed with USE statements and the module will not know
   * If email notifications are turned on then a report email will be sent at the end of the run detailing success / fail and any recorded errors including why a script failed..
   * Currently the versioning option will try to log into the intial database regardless of if the script uses a USE or something to target another database
#>
function Invoke-ScriptRunnerSMO
{
    [CmdletBinding()]
        Param
        (
            [String]
            [ValidateNotNullOrEmpty()]
            [Parameter(Mandatory=$true)]
            $SQL_Server,
            [ValidateSet($true,$false)]
            [Parameter(Mandatory=$true)]
            $EnableWindowsAuthentication,
            [String]
            [ValidateNotNullOrEmpty()]
            [Parameter(Mandatory=$false)]
            $SQL_Login,
            [String]
            [ValidateNotNullOrEmpty()]
            [Parameter(Mandatory=$false)]
            $SQL_Password,
            [String]
            [ValidateNotNullOrEmpty()]
            [Parameter(Mandatory=$true)]
            $InitialDatabase,
            [String []]
            [Parameter(Mandatory=$true)]
            $SQL_ScriptsLocations,
            [ValidateSet(1,2)]
            [Parameter(Mandatory=$true)]
            $AssetDepth,
            [ValidateSet($true,$false)]
            [Parameter(Mandatory=$true)]
            $EnableTransactions,
            [ValidateSet($true,$false)]
            [Parameter(Mandatory=$true)]
            $EnableVersioning,
            [ValidateSet($true,$false)]
            [Parameter(Mandatory=$true)]
            $EnableVersionCheck,
            [ValidateSet($true,$false)]
            [Parameter(Mandatory=$true)]
            $Enable_API_Logging,
            [String]
            [ValidateNotNullOrEmpty()]
            $URL_To_API,
            [int]
            [Parameter(Mandatory=$false)]
            $Company_ID,
            [int]
            [Parameter(Mandatory=$false)]
            $Product_ID,
            [int]
            [Parameter(Mandatory=$false)]
            $Environment_ID,
            [ValidateSet($true,$false)]
            [Parameter(Mandatory=$true)]
            $EnableEmailNotifications,
            [String]
            [ValidateNotNullOrEmpty()]
            [Parameter(Mandatory=$false)]
            $FromEmailAddress,
            [String []]
            [Parameter(Mandatory=$false)]
            $ToSuccessEmailAddresses,
            [String []]
            [Parameter(Mandatory=$false)]
            $ToFailEmailAddresses,
            [String]
            [ValidateNotNullOrEmpty()]
            [Parameter(Mandatory=$false)]
            $SMTP_Server
        )
    
    # Get a start time for logging
    $StartDateTime = Get-Date -Format s

    ##### ##### ##### ##### ##### ##### #####
    ##### IMPORT API LOGGING MODULE ### #####
    ##### ##### ##### ##### ##### ##### #####

    if($Enable_API_Logging -eq $true)
    {
        try
        {
            Import-Module Invoke-Dashboard_API -ErrorAction Stop | Out-Null
        }
        catch
        {
            Write-Error "$StartDateTime : Failed to import Invoke-Dashboard_API. Check that your PS_Modules environment variable is corract and teh module is present in the variable locations."
            EXIT
        }
        # Create JSON object used for external logging.
        $JSON = @{
            Company_ID = $Company_ID;
            Product_ID = $Product_ID;
            Environment_ID = $Environment_ID;
            EntryString = ''
        }
    }

    ##### ##### ##### ##### #####
    ##### VALIDATE INPUTS ##### #
    ##### ##### ##### ##### #####

    # Really basic handling of switches, this is included as well as the validation script section as we need to check variable values if the switch for the functinality is enabled
    # currently I am not too sure how I would do this other than this. 
    if($EnableEmailNotifications -eq $true)
    {
        if([System.String]::IsNullOrWhiteSpace($FromEmailAddress))
        {
            Write-Error "$StartDateTime : $`fromEmailAddress variable is null or white space, this is not valid when email notifications are enabled"
            $JSON.EntryString = "$StartDateTime : $`fromEmailAddress variable is null or white space, this is not valid when email notifications are enabled"
            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json)
            EXIT
        }
        if([System.String]::IsNullOrWhiteSpace($ToSuccessEmailAddresses))
        {
            Write-Error "$StartDateTime : $`ToSuccessEmailAddresses variable is null or white space, this is not valid when email notifications are enabled"
            $JSON.EntryString = "$StartDateTime : $`ToSuccessEmailAddresses variable is null or white space, this is not valid when email notifications are enabled"
            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json)
            EXIT
        }
        if([System.String]::IsNullOrWhiteSpace($ToFailEmailAddresses))
        {
            Write-Error "$StartDateTime : $`ToFailEmailAddresses variable is null or white space, this is not valid when email notifications are enabled"
            $JSON.EntryString = "$StartDateTime : $`ToFailEmailAddresses variable is null or white space, this is not valid when email notifications are enabled"
            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json)
            EXIT
        }
        if([System.String]::IsNullOrWhiteSpace($SMTP_Server))
        {
            Write-Error "$StartDateTime : $`SmtpServer variable is null or white space, this is not valid when email notifications are enabled"
            $JSON.EntryString = "$StartDateTime : $`SmtpServer variable is null or white space, this is not valid when email notifications are enabled"
            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json)
            EXIT
        }
    }

    if($EnableWindowsAuthentication -eq $false)
    {
        if([System.String]::IsNullOrWhiteSpace($SQL_Login) -or [System.String]::IsNullOrWhiteSpace($SQL_Password))
        {
            Write-Verbose -Message "$StartDateTime : Windows authentication has been disabled. You must provide a SQL login and password to if Windows authentication is disabled"
            if($Enable_API_Logging -eq $true)
            {
               $JSON.EntryString = "$StartDateTime : Windows authentication has been disabled. You must provide a SQL login and password to if Windows authentication is disabled"
               Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json)
            }
            EXIT   
        }
    }
    if($EnableDefaultRestoreLocations -eq $false)
    {
        if([System.String]::IsNullOrWhiteSpace($RestoreDestination))
        {
            Write-Verbose -Message "$StartDateTime : EnableDefaultRestoreLocations is disabled. You must provide a restore path if this is disabled"
            if($Enable_API_Logging -eq $true)
            {
               $JSON.EntryString = "$StartDateTime : EnableDefaultRestoreLocations is disabled. You must provide a restore path if this is disabled"
               Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json)
            }
            EXIT   
        }
    }

    ##### ##### ##### ##### #####
    ##### SCRIPT STARTS HERE ####
    ##### ##### ##### ##### #####

    # Load the assembly required for this module to work, this method will load w/e version is avaliable on your system but we are not using anything special so thats fine
    # Due to the use of the LoadWithPartialName method checking to see if the variable is null is the onyl accurate way of determining if a module was loaded.
    $Assembly1 = [Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.Smo')

    if ($Assembly1 -eq $null)
    {
        Write-Error "$StartDateTime : Import of Microsoft.SqlServer.Smo failed, you most likely are not operating with admin rights or SQL is not installed on the machine that is trying to import the assembly"
        if($Enable_API_Logging -eq $true)
        {
            $JSON.EntryString = "$StartDateTime : Import of Microsoft.SqlServer.Smo failed, you most likely are not operating with admin rights or SQL is not installed on the machine that is trying to import the assembly"
            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
        }
        EXIT
    }

    # Create lists for storing results.
    ($ExceptionList = [System.Collections.Generic.List[System.Object]]("")).RemoveAt(0)
    ($RunningOrder = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
    ($CommittedScripts = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
    ($RolledBackScripts = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
    ($ZeroRowScripts = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
    ($SkippedScripts = [System.Collections.Generic.List[System.Object]]("")).RemoveAt(0)

    # Variable for assessing the path, there for future expansion more than anything, may never get used
    foreach ($Location in $SQL_ScriptsLocations)
    {
        # Used to determine if an error has occured to prevent the last success email running if an error occured
        $Gandalf = $false

        # Create list to store all scripts to be ran in in the order found
        ($Scripts = [System.Collections.Generic.List[System.Object]]("")).RemoveAt(0)

        # The $WinSort variables sort the assets in the same way Windows does e.g. do a Get-ChildItem -Path 'C:\PATH' | Select FullName and compare that to a list sorted via name in Windows. 
        # By default, PowerShell does not sort files in the same way as Windows. The method used for the $WinSort variables sorts the assets like Windows This was the case has been tested
        # on PowerShell 5.0 and below. May not be the case in later editions of PowerShell.
        
        try
        {
            switch ($AssetDepth)
            {
                0
                {
                    $WinSort = Get-ChildItem -Path $Location -Include *.sql | Sort-Object
                    foreach ($SQL_File in $WinSort)
                    {
                        $Scripts.Add($SQL_File)    
                    }
                    break
                }
                1
                {
                    $WinSort = Get-ChildItem -Path $Location | Where-Object { $_.Attributes -eq 'd' } | Sort-Object
                    foreach ($Folder in $WinSort)
                    {
                        $WinSort2 = Get-ChildItem -Path $([System.IO.Path]::Combine($($Folder.FullName), "*")) -Include *.sql | Sort-Object
                        foreach ($SQL_File in $WinSort2)
                        {
                            $Scripts.Add($SQL_File)    
                        }          
            
                    }
                    break
                }
                2 
                {
                    $WinSort = Get-ChildItem -Path $Location | Where-Object { $_.Attributes -eq 'd' } | Sort-Object
                    foreach ($Folder in $WinSort)
                    {
                        $WinSort2 = Get-ChildItem -Path $Folder.FullName | Where-Object { $_.Attributes -eq 'd' } | Sort-Object
                        foreach ($DepthOne in $WinSort2)
                        {
                            $WinSort3 = Get-ChildItem -Path $([System.IO.Path]::Combine($($DepthOne.FullName), "*")) -Include *.sql | Sort-Object
                            foreach ($SQL_File in $WinSort3)
                            {
                                $Scripts.Add($SQL_File)
                            }     
                        }
                    }
                    break
                }
                Default 
                {
                    throw '$AssetDepth provided is out of range, values 1 and 2 are supported only'
                }
            }
        }
        catch
        {
            Write-Verbose -Message "$StartDateTime : Depth provided is outside of the supported range. Only depths between 1 and 2 are supported"
            if($Enable_API_Logging -eq $true)
            {
               $JSON.EntryString = "$StartDateTime : Depth provided is outside of the supported range. Only depths between 1 and 2 are supported"
               Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
            }
            EXIT
        }
        if($Scripts.Count -lt 1)
        {
            Write-Verbose -Message "$StartDateTime : No scripts found in declared location. Check your depth configuration"
            Write-Verbose -Message "$StartDateTime : Declared location - $Location"
            if($Enable_API_Logging -eq $true)
            {
               $JSON.EntryString = "$StartDateTime : No scripts found in declared location. Check your depth configuration`r`n$StartDateTime : Declared location - $Location"
               Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
            }
            if($EnableEmailNotifications -eq $true)
            {
                $ExceptionObj = New-Object System.Object
                $ExceptionObj | Add-Member -type NoteProperty -name ScriptName -Value 'No scripts found in declared location. Check your depth configuration'
                $ExceptionObj | Add-Member -type NoteProperty -name ScriptDirectory -Value 'No scripts found in declared location. Check your depth configuration'
                $ExceptionObj | Add-Member -type NoteProperty -name ExceptionMessage -Value 'No scripts found in declared location. Check your depth configuration'
                $ExceptionObj | Add-Member -type NoteProperty -name InnerExceptionMessge -Value 'No scripts found in declared location. Check your depth configuration'
                $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionMessage -Value 'No scripts found in declared location. Check your depth configuration'
                $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionTrace -Value 'No scripts found in declared location. Check your depth configuration'
                $ExceptionList.Add($ExceptionObj)
                Send-MailMessage -From $FromEmailAddress `
                                 -To $ToFailEmailAddresses `
                                 -Subject "FAIL - Script Runner Report" `
                                 -BodyAsHtml "<h3>General Details</h3>
                                              <b>Success Status </b>FAIL<br />
                                              <b>Start Time Of Run </b>$StartDateTime<br />
                                              <h3>Environment Details</h3>
                                              <b>SQL Server Login </b>$SQL_Login<br />
                                              <b>Target SQL Server </b>$SQL_Server<br />
											  <b>Enabled Transactions</b> $EnableTransactions<br />
                                              <b>Initial Database </b>$InitialDatabase<br />
                                              <b>Path To Scripts Used </b>$Location
                                              <h3>Exception Details</h3>No scripts found in declared location. Check your depth configuration" `
                                 -Priority High `
                                 -SmtpServer $SMTP_Server `
                                 -ErrorAction SilentlyContinue
            }
            continue
        }

        # Create sql object for connection
        $ServerObject = New-Object Microsoft.SqlServer.Management.Smo.Server ($SQL_Server)

        # Set options for connection
        $connectionObject = $serverObject.ConnectionContext
        $connectionObject.DatabaseName = $InitialDatabase
        if($EnableWindowsAuthentication -eq $false)
        {
            $ConnectionObject.LoginSecure = $false
            $ConnectionObject.Login = $SQL_Login
            $ConnectionObject.Password = $SQL_Password
        }
        else
        {
            $ConnectionObject.LoginSecure = $true     
        }
        
        $ServerObject.UserOptions.AnsiNullDefaultOn = $true
        $ServerObject.UserOptions.QuotedIdentifier = $true
        $ServerObject.UserOptions.NoCount = $false

        # Setup running list so we can log this during the loop
        ($RunningOrder = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)

        ###### ###### ###### ###### ###### ###### #####
        ###### ###### TRANSACTION RUNNER ###### #######
        ###### ###### ###### ###### ###### ###### #####
        Write-Verbose -Message "############### START OF SMO RUN ###############"
        if($Enable_API_Logging -eq $true)
        {
            $JSON.EntryString = "############### START OF SMO RUN ###############"
            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
        }

        if($EnableTransactions -eq $true)
        {
            # Loop through scripts first checking we can open a transaction
            foreach ($Asset in $scripts)
            {
                $tracker = $false
                $ScriptBaseName = ($Asset.BaseName).Trim()
                $ScriptName = ($Asset.Name).Trim()
                $ScriptFullName = ($Asset.FullName).Trim()
                #$ScriptDirectory = ($Asset.Directory.FullName).Trim()

                # Done this way as opposed to getting the directory property from the file object to prevent the path being logged on 2 seperate lines. The problem seemed to be 
                # related to the length of the path so I just take last 2 directories the script resides in as that will give me the sprint and ticket number.
                # TO DO - Find out if this is an out-file problem and replace that but for now this works fine
                $Split = $Asset.Directory.FullName.Split('\') | Select-Object -last 2
                $ScriptDirectory = ([System.String]::Join("\",$Split)).Trim()

                $RunningOrder.Add($scriptName)
            
                # Gets the query to be ran
                $query = Get-Content -Path $ScriptFullName -Raw

                # Opens a connection with the SQL instance and sets some global properties
                try
                {
                    $serverObject.ConnectionContext.BeginTransaction()
                }
                catch
                {
                    $Gandalf = $true
                    $e1 = $_.Exception
                    $e1m = $e1.Message
                    $e2 = $e1.InnerException
                    $e2m = $e1.InnerException.Message
                    $e3 = $e1.InnerException.InnerException
                    $e3m = $e1.InnerException.InnerException.Message
                    Write-Verbose -Message "$StartDateTime : Transaction connection failed"
                    Write-Verbose -Message "$StartDateTime : $e3m"
                    Write-Verbose -Message "$StartDateTime : $e2m"
                    Write-Verbose -Message "$StartDateTime : $e1m"
                    Write-Verbose -Message "$StartDateTime : $e3"                
                    Write-Verbose -Message "$StartDateTime : $e2"
                    Write-Verbose -Message "$StartDateTime : $e1"
                    if($Enable_API_Logging -eq $true)
                    {
                        $JSON.EntryString = "$StartDateTime : Transaction connection failed"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                        $JSON.EntryString = "$StartDateTime : $e3m`r`n$StartDateTime : $e2m`r`n$StartDateTime : $e1m`r`n$StartDateTime : $e3`r`n$StartDateTime : $e2`r`n$StartDateTime : $e1"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                    if($EnableEmailNotifications -eq $true)
                    {
                        $ExceptionObj = New-Object System.Object
                        $ExceptionObj | Add-Member -type NoteProperty -name ScriptName -Value $ScriptName
                        $ExceptionObj | Add-Member -type NoteProperty -name ScriptDirectory -Value $ScriptDirectory
                        $ExceptionObj | Add-Member -type NoteProperty -name ExceptionMessage -Value $e1m
                        $ExceptionObj | Add-Member -type NoteProperty -name InnerExceptionMessge -Value $e2m
                        $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionMessage -Value $e3m
                        $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionTrace -Value $e3
                        $ExceptionList.Add($ExceptionObj)
                    }
                    EXIT
                }
                Write-Verbose -Message "$StartDateTime : Script Name: $ScriptName"
                Write-Verbose -Message "$StartDateTime : Script Directory: $ScriptDirectory"
                if($Enable_API_Logging -eq $true)
                {
                    $JSON.EntryString = "$StartDateTime : Script Name: $ScriptName`r`n$StartDateTime : Script Directory: $ScriptDirectory"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
                # Now we know we can open transactions check the script name is not in history table, if it is then it has been ran in before and we want to skip it as a result
                if($EnableVersionCheck -eq $true)
                {
                    # List used for storing all rows in versioning table for comparison with script in loop.
                    ($RowHolder = [System.Collections.Generic.List[System.Object]]("")).RemoveAt(0)
                    $output = $serverObject.ConnectionContext.ExecuteWithResults("SELECT * FROM [$InitialDatabase].[dbo].[_IncrementalCustomScriptLog]")
                    foreach ($t in $output.Tables)
                    {
                        foreach ($r in $t.Rows)
                        {
                            $Row = New-Object System.Object    
                            $Row | Add-Member -type NoteProperty -name ID -Value $r.ID
                            $Row | Add-Member -type NoteProperty -name ScriptFilename -Value $r.ScriptFilename
                            $Row | Add-Member -type NoteProperty -name AppliedDate -Value $r.AppliedDate
                            $RowHolder.Add($Row)
                        }
                    }
                    foreach ($row in $RowHolder)
                    {
                        if($row.ScriptFilename -eq $ScriptName)
                        {
                            $SkippedScripts.Add($scriptName)

                            Write-Verbose -Message "$StartDateTime : Checking script in history table"
                            Write-Verbose -Message "$StartDateTime : Script found in history table, skipping script"
                            if($Enable_API_Logging -eq $true)
                            {
                                $JSON.EntryString = "$StartDateTime : Checking script in history table`r`n$StartDateTime : Script found in history table, skipping script"
                                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                            }
                            # $tracker variable used to break the outer loop. Our intial break is for the $row loop. I could not get it break the main $Asset loop
                            # from here. Tried using named loops and other advanced stuff but it couldn't get it working within a module, only ina  script. This 
                            # method was the only way I could successfully break the outer loop. Would be nice at some point to work out how to do it in 1 line.
                            # Marking as TO-DO
                            $tracker = $true
                            break 
                        }
                    }
                }
                try
                {
                    # Break the main $Asset loop
                    if($tracker -eq $true)
                    {
                        continue
                    }
                    $ScriptStartTime = Get-Date -Format s
                    Write-Verbose -Message "$StartDateTime : Script Start Time: $ScriptStartTime"
                    if($Enable_API_Logging -eq $true)
                    {
                        $JSON.EntryString = "$StartDateTime : Script Start Time: $ScriptStartTime"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                    # Excecute script. There are a few methods to do it, this gives you best best rows affected return value although its not very accurate still
                    $retVal = $serverObject.ConnectionContext.ExecuteNonQuery($query)
                        
                    # Checking for open transactions within the script
                    if($serverObject.ConnectionContext.TransactionDepth -gt 1)
                    {
                        $Gandalf = $true
                        Write-Verbose -Message "$StartDateTime : More than 1 transaction detected. Check script for an open transaction"
                        Write-Verbose -Message "$StartDateTime : Rolling transaction back"
                        if($Enable_API_Logging -eq $true)
                        {
                            $JSON.EntryString = "$StartDateTime : More than 1 transaction detected. Check script for an open transaction`r`n$StartDateTime : Rolling transaction back"
                            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                        }
                    
                        $serverObject.ConnectionContext.RollBackTransaction()
                        Write-Verbose -Message "$StartDateTime : Transaction rolled back"
                        if($Enable_API_Logging -eq $true)
                        {
                            $JSON.EntryString = "$StartDateTime : Transaction rolled back"
                            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                        }

                        # Creating new connection to kill the process that is stuck
                        $serverObjectTwo = New-Object Microsoft.SqlServer.Management.Smo.Server ($SQL_Server)
                        $connectionObject = $serverObjectTwo.ConnectionContext
                        $connectionObject.LoginSecure = $false
                        $connectionObject.Login = $SQL_Login
                        $connectionObject.Password = $SQL_Password
                
                        # Connect, get ProcessID from main server object and the kill the open transaction
                        $serverObjectTwo.ConnectionContext.Connect()
                        $processID = $serverObject.ConnectionContext.ProcessID
                        Write-Verbose -Message "$StartDateTime : Killing process ID $processID"
                        if($Enable_API_Logging -eq $true)
                        {
                            $JSON.EntryString = "$StartDateTime : Killing process ID $processID"
                            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                        }

                        $serverObjectTwo.KillProcess($processID)
                        Write-Verbose -Message "$StartDateTime : ProcessID $processID Killed"
                        if($Enable_API_Logging -eq $true)
                        {
                            $JSON.EntryString = "$StartDateTime : ProcessID $processID Killed"
                            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                        }

                        $serverObjectTwo.ConnectionContext.Disconnect()

                        # Alternatively we could kill all connnections on this new connection, depends. The above method seems to work OK though. Leaving in case we ever want to use it
                        # $serverObject.KillAllProcesses($InitialDatabase)
                                        
                        # Add the script to our rollback scripts to log out at the end
                        $RolledBackScripts.Add($ScriptName)
                        
                        if($EnableEmailNotifications -eq $true)
                        {
                            $ExceptionObj = New-Object System.Object
                            $ExceptionObj | Add-Member -type NoteProperty -name ScriptName -Value $ScriptName
                            $ExceptionObj | Add-Member -type NoteProperty -name ScriptDirectory -Value $ScriptDirectory
                            $ExceptionObj | Add-Member -type NoteProperty -name ExceptionMessage -Value 'More than 1 transaction detected. Check script for an open BEGIN TRANSACTION statement'
                            $ExceptionObj | Add-Member -type NoteProperty -name InnerExceptionMessge -Value 'More than 1 transaction detected. Check script for an open BEGIN TRANSACTION statement'
                            $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionMessage -Value 'More than 1 transaction detected. Check script for an open BEGIN TRANSACTION statement'
                            $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionTrace -Value 'More than 1 transaction detected. Check script for an open BEGIN TRANSACTION statement'
                            $ExceptionList.Add($ExceptionObj)
                        }
                        continue
                    }
            
                    $ScriptEndTime = Get-Date -Format s
                    Write-Verbose -Message "$StartDateTime : Script End Time: $ScriptEndTime"
                    Write-Verbose -Message "$StartDateTime : $retVal rows affected"
                    Write-Verbose -Message "$StartDateTime : Committing transaction"
                    if($Enable_API_Logging -eq $true)
                    {
                        $JSON.EntryString = "$StartDateTime : Script End Time: $ScriptEndTime`r`n$StartDateTime : $retVal rows affected`r`n$StartDateTime : Committing transaction"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }

                    $serverObject.ConnectionContext.CommitTransaction()
                    Write-Verbose -Message "$StartDateTime : Transaction committed"
                    if($Enable_API_Logging -eq $true)
                    {
                        $JSON.EntryString = "$StartDateTime : Transaction committed"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                    if ($EnableVersioning -eq $true)
                    {
                        $serverObject.ConnectionContext.ExecuteNonQuery("INSERT INTO [$InitialDatabase].[dbo].[_IncrementalCustomScriptLog] (ScriptFilename, AppliedDate) VALUES ( '$ScriptName', GETDATE() )") 
                    }
                    $committedScripts.Add($ScriptName)
 	                if($retVal -eq 0)
                    {
 	    	            $ZeroRowScripts.Add($ScriptName)
                    }
                }
                catch
                {
                    $Gandalf = $true
                    $e1 = $_.Exception
                    $e1m = $e1.Message
                    $e2 = $e1.InnerException
                    $e2m = $e1.InnerException.Message
                    $e3 = $e1.InnerException.InnerException
                    $e3m = $e1.InnerException.InnerException.Message
                    Write-Verbose -Message "$StartDateTime : Transaction connection failed"
                    Write-Verbose -Message "$StartDateTime : $e3m"
                    Write-Verbose -Message "$StartDateTime : $e2m"
                    Write-Verbose -Message "$StartDateTime : $e1m"
                    Write-Verbose -Message "$StartDateTime : $e3"
                    Write-Verbose -Message "$StartDateTime : $e2"
                    Write-Verbose -Message "$StartDateTime : $e1"
                    if($Enable_API_Logging -eq $true)
                    {
                        $JSON.EntryString = "$StartDateTime : Transaction connection failed`r`n$StartDateTime : $e3m`r`n$StartDateTime : $e2m`r`n$StartDateTime : $e1m`r`n$StartDateTime : $e3`r`n$StartDateTime : $e2`r`n$StartDateTime : $e1"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                    if($EnableEmailNotifications -eq $true)
                    {
                        $ExceptionObj = New-Object System.Object
                        $ExceptionObj | Add-Member -type NoteProperty -name ScriptName -Value $ScriptName
                        $ExceptionObj | Add-Member -type NoteProperty -name ScriptDirectory -Value $ScriptDirectory
                        $ExceptionObj | Add-Member -type NoteProperty -name ExceptionMessage -Value $e1m
                        $ExceptionObj | Add-Member -type NoteProperty -name InnerExceptionMessge -Value $e2m
                        $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionMessage -Value $e3m
                        $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionTrace -Value $e3
                        $ExceptionList.Add($ExceptionObj)
                    }
                    # Roll back anything committed
                    $serverObject.ConnectionContext.RollBackTransaction()    
                    $RolledBackScripts.Add($ScriptName)
                    Write-Verbose -Message "$StartDateTime : Transaction rolled back"
                    if($Enable_API_Logging -eq $true)
                    {
                        $JSON.EntryString = "$StartDateTime : Transaction rolled back"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                }
            }
            # Get a stop time for logging
            $StopDateTime = Get-Date -Format s

            # Log and email report for run
            Write-Verbose -Message "$StartDateTime : ########################## END OF SMO RUN : REPORT OF RUN BELOW ##########################"
            Write-Verbose -Message "$StartDateTime : Run Start Time : $StartDateTime"
            Write-Verbose -Message "$StartDateTime : Run Stop Time : $StopDateTime"
            Write-Verbose -Message "$StartDateTime : Committed Scripts:"
            if($committedScripts.Count -lt 1)
            {
                Write-Verbose -Message "$StartDateTime : None"
            }
            else
            {
                foreach ($CommittedScript in $committedScripts)
                {
                    Write-Verbose -Message "$StartDateTime : $CommittedScript"
                }
            }
            Write-Verbose -Message "$StartDateTime : Rolled Back Scripts:"
            if($RolledBackScripts.Count -lt 1)
            {

                Write-Verbose -Message "$StartDateTime : None"
            }
            else
            {
                foreach ($RolledBackScript in $RolledBackScripts)
                {
                    Write-Verbose -Message "$StartDateTime : $RolledBackScript"
                }
            }
            Write-Verbose -Message "$StartDateTime : Zero Rows Affected Scripts:"
            if($ZeroRowScripts.Count -lt 1)
            {

                Write-Verbose -Message "$StartDateTime : None"
            }
            else
            {
                foreach ($ZeroRowScript in $ZeroRowScripts)
                {
                    Write-Verbose -Message "$StartDateTime : $ZeroRowScript"
                }
            }
            Write-Verbose -Message "$StartDateTime : Skipped scripts:"
            if($SkippedScripts.Count -lt 1)
            {
                Write-Verbose -Message "$StartDateTime : None"
            }
            else
            {
                foreach ($SkippedScript in $SkippedScripts)
                {
                    Write-Verbose -Message "$StartDateTime : $SkippedScript"
                }
            }
            Write-Verbose -Message "$StartDateTime : Running Order:"
            foreach ($RunningScript in $RunningOrder)
            {
                Write-Verbose -Message "$StartDateTime : $RunningScript"
            }
            if($Enable_API_Logging -eq $true)
            {
                $JSON.EntryString = "$StartDateTime : ###############  END OF SMO RUN : REPORT OF RUN BELOW ###############`r`n$StartDateTime : Run Start Time : $StartDateTime`r`n$StartDateTime : Run Stop Time : $StopDateTime`r`n$StartDateTime : Committed Scripts:"
                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                if($committedScripts.Count -lt 1)
                {
                    $JSON.EntryString = "$StartDateTime : $StartDateTime : None"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
                else
                {
                    foreach ($CommittedScript in $committedScripts)
                    {
                        $JSON.EntryString = "$StartDateTime : $CommittedScript"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                }
                $JSON.EntryString = "$StartDateTime : Rolled Back Scripts:"
                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                if($RolledBackScripts.Count -lt 1)
                {
                    $JSON.EntryString = "$StartDateTime : None"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
                else
                {
                    foreach ($RolledBackScript in $RolledBackScripts)
                    {
                        $JSON.EntryString = "$StartDateTime : $RolledBackScript"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                }
                $JSON.EntryString = "$StartDateTime : Zero Rows Affected Scripts:"
                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                if($ZeroRowScripts.Count -lt 1)
                {
                    $JSON.EntryString = "$StartDateTime : None"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
                else
                {
                    foreach ($ZeroRowScript in $ZeroRowScripts)
                    {
                        $JSON.EntryString = "$StartDateTime : $ZeroRowScript"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                }
                $JSON.EntryString = "$StartDateTime : Skipped scripts:"
                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                if($SkippedScripts.Count -lt 1)
                {
                    $JSON.EntryString = "$StartDateTime : None"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
                else
                {
                    foreach ($SkippedScript in $SkippedScripts)
                    {
                        $JSON.EntryString = "$StartDateTime : $SkippedScript"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                }
                $JSON.EntryString = "$StartDateTime : Running Order:"
                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                foreach ($RunningScript in $RunningOrder)
                {
                    $JSON.EntryString = "$StartDateTime : $RunningScript"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
            }
            if($EnableEmailNotifications -eq $true)
            {
                # If Gandalf is set to true if something went wrong, true = none shall pass
                if($Gandalf -eq $false)
                {
                    ($HTML_Report = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
                    $HTML_Report.Add("<b>Committed Scripts:</b><br />")
                    foreach ($CommittedScript in $committedScripts)
                    {
                        $HTML_Report.Add("$CommittedScript<br />")
                    }
                    $HTML_Report.Add("<b>Rolled Back Scripts:</b><br />")
                    foreach ($RolledBackScript in $RolledBackScripts)
                    {
                        $HTML_Report.Add("$RolledBackScript<br />")
                    }
                    $HTML_Report.Add("<b>Zero Rows Affected Scripts:</b><br />")
                    foreach ($ZeroRowScript in $ZeroRowScripts)
                    {
                        $HTML_Report.Add("$ZeroRowScript<br />")
                    }
                    $HTML_Report.Add("<b>Skipped Scripts:</b><br />")
                    foreach ($SkippedScript in $SkippedScripts)
                    {
                        $HTML_Report.Add("$SkippedScript<br />")
                    }
                    $HTML_Report.Add("<b>Running Order:</b><br />")
                    foreach ($RunningScript in $RunningOrder)
                    {
                        $HTML_Report.Add("$RunningScript<br />")
                    }
                    Send-MailMessage -From $FromEmailAddress `
                                     -To $ToSuccessEmailAddresses `
                                     -Subject "SUCCESS - Complete Runner Report" `
                                     -BodyAsHtml "<h3>General Details</h3>
                                                  <b>Success Status </b>SUCCESS<br />
                                                  <b>Start Time Of Run </b>$StartDateTime<br />
                                                  <b>End Time Of Run </b>$StopDateTime<br />
                                                  <h3>Environment Details</h3>
                                                  <b>SQL Server Login </b>$SQL_Login<br />
                                                  <b>Target SQL Server </b>$SQL_Server<br />
												  <b>Enabled Transactions</b> $EnableTransactions<br />
                                                  <b>Initial Database </b>$InitialDatabase<br />
                                                  <b>Path To Scripts </b>$Location
                                                  <h3>Run Details</h3>
                                                  $HTML_Report" `
                                     -Priority Low `
                                     -SmtpServer $SMTP_Server `
                                     -ErrorAction SilentlyContinue
                }
                else
                {
                    ($HTML_Report = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
                    $HTML_Report.Add("<b>Committed Scripts:</b><br />")
                    foreach ($CommittedScript in $committedScripts)
                    {
                        $HTML_Report.Add("$CommittedScript<br />")
                    }
                    $HTML_Report.Add("<b>Rolled Back Scripts:</b><br />")
                    foreach ($RolledBackScript in $RolledBackScripts)
                    {
                        $HTML_Report.Add("$RolledBackScript<br />")
                    }
                    $HTML_Report.Add("<b>Zero Rows Affected Scripts:</b><br />")
                    foreach ($ZeroRowScript in $ZeroRowScripts)
                    {
                        $HTML_Report.Add("$ZeroRowScript<br />")
                    }
                    $HTML_Report.Add("<b>Skipped Scripts:</b><br />")
                    foreach ($SkippedScript in $SkippedScripts)
                    {
                        $HTML_Report.Add("$SkippedScript<br />")
                    }
                    $HTML_Report.Add("<b>Running Order:</b><br />")
                    foreach ($RunningScript in $RunningOrder)
                    {
                        $HTML_Report.Add("$RunningScript<br />")
                    }
                    $HTML_Report.Add("<h3>Exception Details</h3>")
                    foreach ($ExceptionObj in $ExceptionList)
                    {
                        $HTML_Report.Add("<b>Script Name </b>$($ExceptionObj.ScriptName)<br />")
                        $HTML_Report.Add("<b>Script Directory </b>$($ExceptionObj.ScriptDirectory)<br />")
                        $HTML_Report.Add("<b>Inner Inner Exception Message </b>$($ExceptionObj.InnerInnerExceptionMessage)<br />")
                        $HTML_Report.Add("<b>Inner Inner Exception Trace </b><br />$($ExceptionObj.InnerInnerExceptionTrace)<br />")
                        $HTML_Report.Add("<b>Inner Exception Message </b>$($ExceptionObj.InnerExceptionMessge)<br />")
                        $HTML_Report.Add("<b>Exception Message </b>$($ExceptionObj.ExceptionMessage)<br /><br />")
                    }
                    Send-MailMessage -From $FromEmailAddress `
                                     -To $ToFailEmailAddresses `
                                     -Subject "FAIL - Complete Runner Report" `
                                     -BodyAsHtml "<h3>General Details</h3>
                                                  <b>Success Status </b>FAIL<br />
                                                  <b>Start Time Of Run </b>$StartDateTime<br />
                                                  <b>End Time Of Run </b>$StopDateTime<br />
                                                  <h3>Environment Details</h3>
                                                  <b>SQL Server Login </b>$SQL_Login<br />
                                                  <b>Target SQL Server </b>$SQL_Server<br />
												  <b>Enabled Transactions</b> $EnableTransactions<br />
                                                  <b>Initial Database </b>$InitialDatabase<br />
                                                  <b>Path To Scripts</b> $Location
                                                  <h3>Run Details</h3>
                                                  $HTML_Report" `
                                     -Priority High `
                                     -SmtpServer $SMTP_Server `
                                     -ErrorAction SilentlyContinue
                }
            }
        }
        ###### ###### ###### ###### ###### ###### ##### ###
        ###### ###### TRANSACTIONLESS RUNNER ###### #######
        ###### ###### ###### ###### ###### ###### ##### ###
        else
        {
            try
            {
                $serverObject.ConnectionContext.Connect()
            }
            catch
            {
                $Gandalf = $true
                $e1 = $_.Exception
                $e1m = $e1.Message
                $e2 = $e1.InnerException
                $e2m = $e1.InnerException.Message
                $e3 = $e1.InnerException.InnerException
                $e3m = $e1.InnerException.InnerException.Message
                Write-Verbose -Message "$StartDateTime : Connection failed"
                Write-Verbose -Message "$StartDateTime : $e3m"
                Write-Verbose -Message "$StartDateTime : $e2m"
                Write-Verbose -Message "$StartDateTime : $e1m"
                Write-Verbose -Message "$StartDateTime : $e3"                
                Write-Verbose -Message "$StartDateTime : $e2"
                Write-Verbose -Message "$StartDateTime : $e1"
                if($Enable_API_Logging -eq $true)
                {
                    $JSON.EntryString = "$StartDateTime : Connection failed`r`n$StartDateTime : $e3m`r`n$StartDateTime : $e2m`r`n$StartDateTime : $e1m`r`n$StartDateTime : $e3`r`n$StartDateTime : $e2`r`n$StartDateTime : $e1"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
                if($EnableEmailNotifications -eq $true)
                {
                    $ExceptionObj = New-Object System.Object
                    $ExceptionObj | Add-Member -type NoteProperty -name ScriptName -Value $ScriptName
                    $ExceptionObj | Add-Member -type NoteProperty -name ScriptDirectory -Value $ScriptDirectory
                    $ExceptionObj | Add-Member -type NoteProperty -name ExceptionMessage -Value $e1m
                    $ExceptionObj | Add-Member -type NoteProperty -name InnerExceptionMessge -Value $e2m
                    $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionMessage -Value $e3m
                    $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionTrace -Value $e3
                    $ExceptionList.Add($ExceptionObj)
                }
                EXIT
            }

            # Loop through scripts first checking we can open a transaction
            foreach ($Asset in $scripts)
            {
                $tracker = $false
                $ScriptBaseName = ($Asset.BaseName).Trim()
                $ScriptName = ($Asset.Name).Trim()
                $ScriptFullName = ($Asset.FullName).Trim()
                #$ScriptDirectory = ($Asset.Directory.FullName).Trim()

                # Done this way as opposed to getting the directory property to prevent the path being logged on 2 lines. Problem seemed to be related to the length of the 
                # line so I just take the directory is resides within to get around this. Not ideal, would be best to find out if this is an out-file problem adn replace that
                # but for now this works fin
                $Split = $Asset.Directory.FullName.Split('\') | Select-Object -last 1
                $ScriptDirectory = ([System.String]::Join("\",$Split)).Trim()

                $RunningOrder.Add($scriptName)

                # Gets the query to be ran
                $query = Get-Content -Path $ScriptFullName -Raw

                Write-Verbose -Message "$StartDateTime : Script Name: $ScriptName"
                Write-Verbose -Message "$StartDateTime : Script Directory: $ScriptDirectory"
                # Now we know we can open transactions check the script name is not in history table, if it is then it has been ran in before and we want to skip it as a result
                if($EnableVersionCheck -eq $true)
                {
                    ($RowHolder = [System.Collections.Generic.List[System.Object]]("")).RemoveAt(0)
                    $output = $serverObject.ConnectionContext.ExecuteWithResults("SELECT * FROM [$InitialDatabase].[dbo].[_IncrementalCustomScriptLog]")
                    foreach ($t in $output.Tables)
                    {
                        foreach ($r in $t.Rows)
                        {
                            $Row = New-Object System.Object    
                            $Row | Add-Member -type NoteProperty -name ID -Value $r.ID
                            $Row | Add-Member -type NoteProperty -name ScriptFilename -Value $r.ScriptFilename
                            $Row | Add-Member -type NoteProperty -name AppliedDate -Value $r.AppliedDate
                            $RowHolder.Add($Row)
                        }
                    }
                    foreach ($row in $RowHolder)
                    {
                        if($row.ScriptFilename -eq $ScriptName)
                        {
                            $SkippedScripts.Add($scriptName)
                            Write-Verbose -Message "$StartDateTime : Checking script in history table"
                            Write-Verbose -Message "$StartDateTime : Script found in history table, skipping script"
                            if($Enable_API_Logging -eq $true)
                            {
                                $JSON.EntryString = "$StartDateTime : Checking script in history table`r`nScript found in history table, skipping script"
                                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                            }
                            $tracker = $true
                            # Breaks $row loop, could not get it breka the main script loop in a single line. Tried using named loops and other advanced stuff but it couldn't get it working within a module, this was the only way.
                            break 
                        }
                    }
                }

                # The script has been determined to be needed or the versioning option is off. Run in the SQL.
                try
                {
                    if($tracker -eq $true)
                    {
                        continue
                    }
                    $ScriptStartTime = Get-Date -Format s
                    Write-Verbose -Message "$StartDateTime : Script Start Time: $ScriptStartTime"
                    if($Enable_API_Logging -eq $true)
                    {
                        $JSON.EntryString = "$StartDateTime : Script Start Time: $ScriptStartTime"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                    # Excecute script. There are a few methods to do it, this gives you best best rows affected return value although its not very accurate still
                    $retVal = $serverObject.ConnectionContext.ExecuteNonQuery($query)
                    
                    # Checking for open transactions within the script as the ExcecuteNonQuery method will finish when teh script has finished running.
                    # If the script has not performed a self commit or roll bakc we need to warn someone of this.
                    if($serverObject.ConnectionContext.TransactionDepth -gt 0)
                    {
                        $Gandalf = $true
                        Write-Verbose -Message "$StartDateTime : Open transaction detected after deploying script. Check script for an open transaction with no open COMMIT"
                        Write-Verbose -Message "$StartDateTime : Rolling transaction back"
                        if($Enable_API_Logging -eq $true)
                        {
                            $JSON.EntryString = "$StartDateTime : Open transaction detected after deploying script. Check script for an open transaction with no open COMMIT`r`n$StartDateTime : Rolling transaction back"
                            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                        }
                        $serverObject.ConnectionContext.RollBackTransaction()
                        Write-Verbose -Message "$StartDateTime : Transaction rolled back"
                        if($Enable_API_Logging -eq $true)
                        {
                            $JSON.EntryString = "$StartDateTime : Transaction rolled back"
                            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                        }

                        # Creating new connection to kill the process that is stuck
                        $serverObjectTwo = New-Object Microsoft.SqlServer.Management.Smo.Server ($SQL_Server)
                        $connectionObject = $serverObjectTwo.ConnectionContext
                        $connectionObject.LoginSecure = $false
                        $connectionObject.Login = $SQL_Login
                        $connectionObject.Password = $SQL_Password
                
                        # Connect, get ProcessID from main server object and the kill the open transaction
                        $serverObjectTwo.ConnectionContext.Connect()
                        $processID = $serverObject.ConnectionContext.ProcessID
                        Write-Verbose -Message "$StartDateTime : Killing process ID $processID"
                        if($Enable_API_Logging -eq $true)
                        {
                            $JSON.EntryString = "$StartDateTime : Killing process ID $processID"
                            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                        }

                        $serverObjectTwo.KillProcess($processID)
                        Write-Verbose -Message "$StartDateTime : ProcessID $processID Killed"
                        if($Enable_API_Logging -eq $true)
                        {
                            $JSON.EntryString = "$StartDateTime : ProcessID $processID Killed"
                            Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                        }

                        $serverObjectTwo.ConnectionContext.Disconnect()

                        # Alternatively we could kill all connnections on this new connection, depends. The above method seems to work OK though. Leaving in case we ever want to use it
                        # $serverObject.KillAllProcesses($InitialDatabase)
                                        
                        # Add the script to our rollback scripts to log out at the end
                        $RolledBackScripts.Add($ScriptName)
                        if($EnableEmailNotifications -eq $true)
                        {
                            $ExceptionObj = New-Object System.Object
                            $ExceptionObj | Add-Member -type NoteProperty -name ScriptName -Value $ScriptName
                            $ExceptionObj | Add-Member -type NoteProperty -name ScriptDirectory -Value $ScriptDirectory
                            $ExceptionObj | Add-Member -type NoteProperty -name ExceptionMessage -Value 'More than 1 transaction detected. Check script for an open BEGIN TRANSACTION statement'
                            $ExceptionObj | Add-Member -type NoteProperty -name InnerExceptionMessge -Value 'More than 1 transaction detected. Check script for an open BEGIN TRANSACTION statement'
                            $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionMessage -Value 'More than 1 transaction detected. Check script for an open BEGIN TRANSACTION statement'
                            $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionTrace -Value 'More than 1 transaction detected. Check script for an open BEGIN TRANSACTION statement'
                            $ExceptionList.Add($ExceptionObj)
                        }              
                        continue
                    }
            
                    $ScriptEndTime = Get-Date -Format s
                    Write-Verbose -Message "$StartDateTime : Script End Time: $ScriptEndTime"
                    Write-Verbose -Message "$StartDateTime : $retVal rows affected"
                    if($Enable_API_Logging -eq $true)
                    {
                        $JSON.EntryString = "$StartDateTime : Script End Time: $ScriptEndTime`r`n$StartDateTime : $retVal rows affected"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                    if ($EnableVersioning -eq $true)
                    {
                        $serverObject.ConnectionContext.ExecuteNonQuery("INSERT INTO [$InitialDatabase].[dbo].[_IncrementalCustomScriptLog] (ScriptFilename, AppliedDate) VALUES ( '$ScriptName', GETDATE() )") 
                    }
                    $committedScripts.Add($ScriptName)
 	                if($retVal -eq 0)
                    {
 	    	            $ZeroRowScripts.Add($ScriptName)
                    }
                }
                catch
                {
                    $Gandalf = $true
                    $e1 = $_.Exception
                    $e1m = $e1.Message
                    $e2 = $e1.InnerException
                    $e2m = $e1.InnerException.Message
                    $e3 = $e1.InnerException.InnerException
                    $e3m = $e1.InnerException.InnerException.Message
                    Write-Verbose -Message "$StartDateTime : Query failed"
                    Write-Verbose -Message "$StartDateTime : $e3m"
                    Write-Verbose -Message "$StartDateTime : $e2m"
                    Write-Verbose -Message "$StartDateTime : $e1m"
                    Write-Verbose -Message "$StartDateTime : $e3"
                    Write-Verbose -Message "$StartDateTime : $e2"
                    Write-Verbose -Message "$StartDateTime : $e1"
                    if($Enable_API_Logging -eq $true)
                    {
                        $JSON.EntryString = "$StartDateTime : Query failed`r`n$StartDateTime : $e3m`r`n$StartDateTime : $e2m`r`n$StartDateTime : $e1m`r`n$StartDateTime : $e3`r`n$StartDateTime : $e2`r`n$StartDateTime : $e1"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                    if($EnableEmailNotifications -eq $true)
                    {                    
                        $ExceptionObj = New-Object System.Object
                        $ExceptionObj | Add-Member -type NoteProperty -name ScriptName -Value $ScriptName
                        $ExceptionObj | Add-Member -type NoteProperty -name ScriptDirectory -Value $ScriptDirectory
                        $ExceptionObj | Add-Member -type NoteProperty -name ExceptionMessage -Value $em
                        $ExceptionObj | Add-Member -type NoteProperty -name InnerExceptionMessge -Value $e2m
                        $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionMessage -Value $e3m
                        $ExceptionObj | Add-Member -type NoteProperty -name InnerInnerExceptionTrace -Value $e3
                        $ExceptionList.Add($ExceptionObj)
                    }
                    # Roll back anything committed
                    $serverObject.ConnectionContext.RollBackTransaction()    
                    $RolledBackScripts.Add($ScriptName)
                    Write-Verbose -Message "$StartDateTime : Transaction rolled back"
                    if($Enable_API_Logging -eq $true)
                    {
                        $JSON.EntryString = "$StartDateTime : Transaction rolled back"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                }
            }
            # Get a stop time for logging
            $StopDateTime = Get-Date -Format s

            # Log and email report for run
            Write-Verbose -Message "$StartDateTime : ########################## END OF SMO RUN : REPORT OF RUN BELOW ##########################"
            Write-Verbose -Message "$StartDateTime : Run Start Time : $StartDateTime"
            Write-Verbose -Message "$StartDateTime : Run Stop Time : $StopDateTime"
            Write-Verbose -Message "$StartDateTime : Committed Scripts:"
            if($committedScripts.Count -lt 1)
            {
                Write-Verbose -Message "$StartDateTime : None"
            }
            else
            {
                foreach ($CommittedScript in $committedScripts)
                {
                    Write-Verbose -Message "$StartDateTime : $CommittedScript"
                }
            }
            Write-Verbose -Message "$StartDateTime : Rolled Back Scripts:"
            if($RolledBackScripts.Count -lt 1)
            {

                Write-Verbose -Message "$StartDateTime : None"
            }
            else
            {
                foreach ($RolledBackScript in $RolledBackScripts)
                {
                    Write-Verbose -Message "$StartDateTime : $RolledBackScript"
                }
            }
            Write-Verbose -Message "$StartDateTime : Zero Rows Affected Scripts:"
            if($ZeroRowScripts.Count -lt 1)
            {

                Write-Verbose -Message "$StartDateTime : None"
            }
            else
            {
                foreach ($ZeroRowScript in $ZeroRowScripts)
                {
                    Write-Verbose -Message "$StartDateTime : $ZeroRowScript"
                }
            }
            Write-Verbose -Message "$StartDateTime : Skipped scripts:"
            if($SkippedScripts.Count -lt 1)
            {
 
                Write-Verbose -Message "$StartDateTime : None"
            }
            else
            {

                foreach ($SkippedScript in $SkippedScripts)
                {
                    Write-Verbose -Message "$StartDateTime : $SkippedScript"
                }
            }
            Write-Verbose -Message "$StartDateTime : Running Order:"
            foreach ($RunningScript in $RunningOrder)
            {
                Write-Verbose -Message "$StartDateTime : $RunningScript"
            }
            if($Enable_API_Logging -eq $true)
            {
                $JSON.EntryString = "$StartDateTime : ########################## END OF SMO RUN : REPORT OF RUN BELOW ##########################`r`n$StartDateTime : Run Start Time : $StartDateTime`r`n$StartDateTime : Run Stop Time : $StopDateTime`r`n$StartDateTime : Committed Scripts:"
                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                if($committedScripts.Count -lt 1)
                {
                    $JSON.EntryString = "$StartDateTime : None"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
                else
                {
                    foreach ($CommittedScript in $committedScripts)
                    {
                        $JSON.EntryString = "$StartDateTime : $CommittedScript"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                }
                $JSON.EntryString = "$StartDateTime : Rolled Back Scripts:"
                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                if($RolledBackScripts.Count -lt 1)
                {
                    $JSON.EntryString = "$StartDateTime : None"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
                else
                {
                    foreach ($RolledBackScript in $RolledBackScripts)
                    {
                        $JSON.EntryString = "$StartDateTime : $RolledBackScripts"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                }
                $JSON.EntryString = "$StartDateTime : Zero Rows Affected Scripts:"
                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                if($ZeroRowScripts.Count -lt 1)
                {
                    $JSON.EntryString = "$StartDateTime : None"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
                else
                {
                    foreach ($ZeroRowScript in $ZeroRowScripts)
                    {
                        $JSON.EntryString = "$StartDateTime : $ZeroRowScripts"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                }
                $JSON.EntryString = "$StartDateTime : Skipped scripts:"
                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                if($SkippedScripts.Count -lt 1)
                {
                    $JSON.EntryString = "$StartDateTime : None"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
                else
                {
                    foreach ($SkippedScript in $SkippedScripts)
                    {
                        $JSON.EntryString = "$StartDateTime : $SkippedScripts"
                        Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                    }
                }
                $JSON.EntryString = "$StartDateTime : Running Order:"
                Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                foreach ($RunningScript in $RunningOrder)
                {
                    $JSON.EntryString = "$StartDateTime : $RunningOrder"
                    Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json) | Out-Null
                }
            }
            if($EnableEmailNotifications -eq $true)
            {
                # If Gandalf is set to true if something went wrong, true = none shall pass
                if($Gandalf -eq $false)
                {
                    ($HTML_Report = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
                    $HTML_Report.Add("<b>Rolled Back Scripts:</b><br />")
                    foreach ($RolledBackScript in $RolledBackScripts)
                    {
                        $HTML_Report.Add("$RolledBackScript<br />")
                    }
                    $HTML_Report.Add("<b>Zero Rows Affected Scripts:</b><br />")
                    foreach ($ZeroRowScript in $ZeroRowScripts)
                    {
                        $HTML_Report.Add("$ZeroRowScript<br />")
                    }
                    $HTML_Report.Add("<b>Skipped Scripts:</b><br />")
                    foreach ($SkippedScript in $SkippedScripts)
                    {
                        $HTML_Report.Add("$SkippedScript<br />")
                    }
                    $HTML_Report.Add("<b>Running Order:</b><br />")
                    foreach ($RunningScript in $RunningOrder)
                    {
                        $HTML_Report.Add("$RunningScript<br />")
                    }
                    Send-MailMessage -From $FromEmailAddress `
                                     -To $ToSuccessEmailAddresses `
                                     -Subject "SUCCESS - Complete Runner Report" `
                                     -BodyAsHtml "<h3>General Details</h3>
                                                  <b>Success Status </b>SUCCESS<br />
                                                  <b>Start Time Of Run </b>$StartDateTime<br />
                                                  <b>End Time Of Run </b>$StopDateTime<br />
                                                  <h3>Environment Details</h3>
                                                  <b>SQL Server Login </b>$SQL_Login<br />
                                                  <b>Target SQL Server </b>$SQL_Server<br />
												  <b>Enabled Transactions</b> $EnableTransactions<br />
                                                  <b>Initial Database </b>$InitialDatabase<br />
                                                  <b>Path To Scripts </b>$Location
                                                  <h3>Run Details</h3>
                                                  $HTML_Report" `
                                     -Priority Low `
                                     -SmtpServer $SMTP_Server `
                                     -ErrorAction SilentlyContinue
                }
                else
                {
                    ($HTML_Report = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
                    $HTML_Report.Add("<b>Rolled Back Scripts:</b><br />")
                    foreach ($RolledBackScript in $RolledBackScripts)
                    {
                        $HTML_Report.Add("$RolledBackScript<br />")
                    }
                    $HTML_Report.Add("<b>Zero Rows Affected Scripts:</b><br />")
                    foreach ($ZeroRowScript in $ZeroRowScripts)
                    {
                        $HTML_Report.Add("$ZeroRowScript<br />")
                    }
                    $HTML_Report.Add("<b>Skipped Scripts:</b><br />")
                    foreach ($SkippedScript in $SkippedScripts)
                    {
                        $HTML_Report.Add("$SkippedScript<br />")
                    }
                    $HTML_Report.Add("<b>Running Order:</b><br />")
                    foreach ($RunningScript in $RunningOrder)
                    {
                        $HTML_Report.Add("$RunningScript<br />")
                    }
                    $HTML_Report.Add("<h3>Exception Details</h3>")
                    foreach ($ExceptionObj in $ExceptionList)
                    {
                        $HTML_Report.Add("<b>Script Name </b>$($ExceptionObj.ScriptName)<br />")
                        $HTML_Report.Add("<b>Script Directory </b>$($ExceptionObj.ScriptDirectory)<br />")
                        $HTML_Report.Add("<b>Inner Inner Exception Message </b>$($ExceptionObj.InnerInnerExceptionMessage)<br />")
                        $HTML_Report.Add("<b>Inner Inner Exception Trace </b><br />$($ExceptionObj.InnerInnerExceptionTrace)<br />")
                        $HTML_Report.Add("<b>Inner Exception Message </b>$($ExceptionObj.InnerExceptionMessge)<br />")
                        $HTML_Report.Add("<b>Exception Message </b>$($ExceptionObj.ExceptionMessage)<br /><br />")
                    }
                    Send-MailMessage -From $FromEmailAddress `
                                     -To $ToFailEmailAddresses `
                                     -Subject "FAIL - Complete Runner Report" `
                                     -BodyAsHtml "<h3>General Details</h3>
                                                  <b>Success Status </b>FAIL<br />
                                                  <b>Start Time Of Run </b>$StartDateTime<br />
                                                  <b>End Time Of Run </b>$StopDateTime<br />
                                                  <h3>Environment Details</h3>
                                                  <b>SQL Server Login </b>$SQL_Login<br />
                                                  <b>Target SQL Server </b>$SQL_Server<br />
												  <b>Enabled Transactions</b> $EnableTransactions<br />
                                                  <b>Initial Database </b>$InitialDatabase<br />
                                                  <b>Path To Scripts </b>$Location
                                                  <h3>Run Details</h3>
                                                  $HTML_Report" `
                                     -Priority High `
                                     -SmtpServer $SMTP_Server `
                                     -ErrorAction SilentlyContinue
                }
            }
        }
    }
}