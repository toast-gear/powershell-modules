<#
.DESCRIPTION
   This function restores all bak files from the declared location. It will overwrite databases with the same name.
.EXAMPLE
   $URL = 'http://dashboard.domain.com/api/v1/logs/deploymentlogpost'
   Restore-Baks -SQL_Server 'BOAR\SQLEXPRESS' `
                -EnableWindowsAuthentication $false `
                -SQL_Login 'username' `
                -SQL_Password 'password' `
                -DatabaseNameRegexValues '-.' `
                -BakFilesLocation $files `
                -EnableDefaultRestoreLocations $false `
                -RestoreDestination $restore `
                -Enable_API_Logging $true `
                -URL_To_API $URL `
                -Company_ID 1 `
                -Product_ID 1 `
                -Environment_ID 10 `
                -EnableVersioning $true `
                -EnableVersionCheck $true `
                -EnableEmailNotifications $false `
                -FromEmailAddress 'firstname.surname@domain.com' `
                -ToSuccessEmailAddresses 'firstname.surname@domain.com'`
                -ToFailEmailAddresses 'firstname.surname@domain.com', 'firstname.surname@domain.com.uk'`
                -SMTP_Server "10.201.176.165" | Out-File -FilePath "F:\testtest\test.txt" -Append

   This is how you log this module, all output added by the author was added to the verbose stream. You need ot redirect that output to the normal stdout stream to then log it. At this point you can treat it like normal stdout
   output. You use the out-file append flag to append rather than overwrite. 
    
.INPUTS
    -SQL_Server [String]
    -EnableWindowsAuthentication [Boolean]
    -SQL_Login [String]
    -SQL_Password [String]
    -DatabaseNameRegexValues [string]
    -BakFilesLocation [String]
    -EnableDefaultRestoreLocations [Boolean]
    -RestoreDestination [String]
    -Enable_API_Logging [Boolean]
    -URL_To_API [string]
    -Company_ID [int]
    -Product_ID [int]
    -Environment_ID [int]
    -EnableEmailNotifications [Boolean]
    -FromEmailAddress [String]
    -ToSuccessEmailAddresses [String]
    -ToFailEmailAddresses [String]
    -SMTP_Server [String]
.OUTPUTS
    * All output is sent to the verbose stream when the output stream is sued you get blank lines logged from the posts to the API
.NOTES
    * mTested on SQL 2012, 2008 R2 and with PowerShell 4.0
    * Currently requires a '.' in the regex input is required or the file will have a unexpacted name, will validate some inputs without a '.' however. For example '-' will pass the test. Not sure why atm, most inputs are
    validated however.
#>

function Restore-SQL_Bak
{
    [cmdletbinding()]
        Param
        (
            [String]
            [Parameter(Mandatory=$true)]
            $SQL_Server,
            [ValidateSet($true,$false)]
            [Parameter(Mandatory=$true)]
            $EnableWindowsAuthentication,
            [String]
            [Parameter(Mandatory=$false)]
            $SQL_Login,
            [String]
            [Parameter(Mandatory=$false)]
            $SQL_Password,
            [ValidateScript({ $_ -like '*.*' })]
            [Parameter(Mandatory=$true)]
            $DatabaseNameRegexValues,
            [String]
            $BakFilesLocation,
            [ValidateSet($true,$false)]
            [Parameter(Mandatory=$true)]
            $EnableDefaultRestoreLocations,
            [String]
            [Parameter(Mandatory=$false)]
            $RestoreDestination,
            [ValidateSet($true,$false)]
            [Parameter(Mandatory=$true)]
            $Enable_API_Logging,
            [ValidateScript({
                if(![system.string]::IsNullOrWhiteSpace($_.split('/')[-1]))
                {
                    return $true
                }
                else
                {
                    return $false
                }
            })]
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
            [Parameter(Mandatory=$false)]
            $FromEmailAddress,
            [String]
            [Parameter(Mandatory=$false)]
            $ToSuccessEmailAddresses,
            [String]
            [Parameter(Mandatory=$false)]
            $ToFailEmailAddresses,
            [String]
            [Parameter(Mandatory=$false)]
            $SMTP_Server
        )

    $Error.Clear()
    $StartDateTime = Get-Date -Format s

    if($Enable_API_Logging -eq $true)
    {
        try
        {
            Import-Module Invoke-Dashboard_API -ErrorAction Stop
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
    
    # Can't use a try catch with the LoadWithPartialName method. Method does not throw error, must check variable to see if it is null
    $Assembly1 = [Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.Smo')
    if($Assembly1 -eq $null)
    {
        Write-Verbose -Message "$StartDateTime : Import of Microsoft.SqlServer.Smo failed, you most likely are not operating with admin rights"
        if($Enable_API_Logging -eq $true)
        {
           $JSON.EntryString = "$StartDateTime : Import of Microsoft.SqlServer.Smo failed, you most likely are not operating with admin rights"
           Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json)
        }
        EXIT        
    }
    $Assembly2 = [Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SmoExtended')
    if($Assembly2 -eq $null)
    {
        Write-Verbose -Message "$StartDateTime : Import of Microsoft.SqlServer.SmoExtended failed, you most likely are not operating with admin rights"
        if($Enable_API_Logging -eq $true)
        {
           $JSON.EntryString = "$StartDateTime : Import of Microsoft.SqlServer.SmoExtended failed, you most likely are not operating with admin rights"
           Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json)
        }
        EXIT        
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

    ##### ##### ##### ##### ##### ##### ##### ##### ##### 
    ##### ##### GLOBAL VARIABLES / FUNCTIONS  ##### ##### 
    ##### ##### ##### ##### ##### ##### ##### ##### ##### 

    # Used for Logging for Emails
    ($RunningOrder = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
    ($FailRestores = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
    ($SuccessRestores = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
    
    $ServerObject = New-Object ('Microsoft.SqlServer.Management.Smo.Server') ($SQL_Server)
    $ConnectionObject = $ServerObject.ConnectionContext
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

    ##### ##### ##### ##### #####
    ##### SCRIPT STARTS HERE ####
    ##### ##### ##### ##### #####

    # Check to see how we are setting the restore locations.
    if($EnableDefaultRestoreLocations -eq $true)
    {
        $ConnectionObject.Connect()
        # This is needed as there is a bug with SQL. If the default locations are never change the DefaultFile and DefaultLog properties
        # never get set. As a result we have to look in teh MaterDB Path.
        $DataPath = $ServerObject.Settings.DefaultFile
        $LogPath = $ServerObject.Settings.DefaultLog
        if($DataPath.Length -eq 0)
        {
            $DataPath = $ServerObject.Information.MasterDBPath
        }
        if($LogPath.Length -eq 0)
        {
            $LogPath = $ServerObject.Information.MasterDBLogPath
        }
        $ConnectionObject.Disconnect()
    }

    foreach ($Location in $BakFilesLocation)
    {
        # $Gandalf variable used to determine if an error has occured so we know to email out if there is an error. True = none shall pass (no errors logged)
        $Gandalf = $true
        $ReleaseUnit = Get-ChildItem -Path ([System.IO.Path]::Combine($Location,"*")) -Recurse -Include *.bak
        if($ReleaseUnit.Count -eq 0)
        {
            continue
        }
        # Loop through the supplied location for baks and restore each
        foreach($file in $ReleaseUnit)
        {
            $BakFullName = $file.FullName
            $BakName = $file.Name
            $BakBaseName = $file.BaseName
            $BakDirectory = $file.DirectoryName

            # Add to our running  order list so we can log these details
            $RunningOrder.Add($BakName)

            # Kill connections to the declared target database, commented out in case we care about ti alter
            # $ServerObject.KillAllProcesses($targetDatabase)

            # Create the backup object and provide the bak file
            $backupDevice = New-Object ('Microsoft.SqlServer.Management.Smo.BackupDeviceItem') ($BakFullName, 'File')
    
            # Create the new restore object, get databasename from file and attach backup device to restore object and setting options
            $DB_NameHolder = $BakFullName.Split($DatabaseNameRegexValues)
            $RestoreObject = New-Object('Microsoft.SqlServer.Management.Smo.Restore')
            $RestoreObject.Database = $DB_NameHolder[-2]
            $RestoreObject.Devices.Add($backupDevice)
            $RestoreObject.ReplaceDatabase = $true
            $RestoreObject.NoRecovery = $false
            
            # Restore path construction
            $DB_PhysicalFileNameHolder = $DB_NameHolder[-2]
            if($EnableDefaultRestoreLocations -eq $true)
            {
                $MDF_FullName = [System.IO.Path]::Combine($DataPath,"$DB_PhysicalFileNameHolder`_Data.mdf")
                $LDF_FullName =[System.IO.Path]::Combine($DataPath,"$DB_PhysicalFileNameHolder`_Log.mdf")
            }
            else
            {
                $MDF_FullName = [System.IO.Path]::Combine($RestoreDestination,"$DB_PhysicalFileNameHolder`_Data.mdf")
                $LDF_FullName =[System.IO.Path]::Combine($RestoreDestination,"$DB_PhysicalFileNameHolder`_Log.mdf")
            }
            <#
            # Parse the bak file into a variable from the serverObject we attached it to to extract properties of the file into an object to process
            # $fileList = $restoreObject.ReadFileList($serverObject)
         
            # $dbfile = "$restoreDestination$dbPhysicalFileNameHolder`_Data.mdf"
            # $logfile = "$restoreDestination$dbPhysicalFileNameHolder`_Log.ldf"       
            
            This gives the file a logical and physical filename. We could just append this to the physicalfilename property
            but it is best to check the type flag to make sure we are dealing with a mdf before we name it. Possible flags are@

            L = LogFile , D = Database file , F = FullText Catalog

            Currently the script does not account for the F value

            Changing logical file names can be done on the fly, they are referenced in restore jobs so it should match w/e the 
            restore job expects. The physical path is not as simple to change, As a result I will just use what it present in the 
            file

            #>

            # Loop through the object to extract the logical file name and set the physical file name based on the declared mdf or ldf destinations
            $FileList = $RestoreObject.ReadFileList($ServerObject)
            foreach ($File in $FileList)
            {
                $RelocateFileObject = New-Object('Microsoft.SqlServer.Management.Smo.RelocateFile')
                $RelocateFileObject.LogicalFileName = $File.LogicalName
                if ($File.Type -eq 'D')
                {
                    $RelocateFileObject.PhysicalFileName = $MDF_FullName
                }
                else
                {
                    $RelocateFileObject.PhysicalFileName = $LDF_FullName
                }
                $RestoreObject.RelocateFiles.Add($RelocateFileObject)
            }

            $Holder = $DB_NameHolder[-2]
            $RestoreStartDateTime = Get-Date -Format s
            # It is at this point that 0 and 1 are written into the log. I assume they are return/exit codes. Currently not sure how to get rid of them / stop them, would be nice but very low priority.
            Write-Verbose -Message "$StartDateTime : Restore started at $RestoreStartDateTime"
            Write-Verbose -Message "$StartDateTime : Bak file path - $DB_NameHolder"
            Write-Verbose -Message "$StartDateTime : Database name - $Holder"
            Write-Verbose -Message "$StartDateTime : mdf restore location - $MDF_FullName"
            Write-Verbose -Message "$StartDateTime : ldf restore location - $LDF_FullName"
            if($Enable_API_Logging -eq $true)
            {
               $JSON.EntryString = "$StartDateTime : Restore started at $RestoreStartDateTime`r`n$StartDateTime : Bak file path - $DB_NameHolder`r`n$StartDateTime : Database name - $Holder`r`nmdf restore location - $MDF_FullName`r`nldf restore location - $LDF_FullName"
               Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json)
            }

            # Restore the database and log to success list if nothing is caught
            try
            {
                $RestoreObject.SqlRestore($ServerObject)
                $SuccessRestores.Add($BakName)
                $RestoreStopDateTime = Get-Date -Format s
                Write-Verbose -Message "$StartDateTime : Restore Success!"
                Write-Verbose -Message "$StartDateTime : Restore ended at $RestoreStopDateTime"
            if($Enable_API_Logging -eq $true)
            {
               $JSON.EntryString = "$StartDateTime : Restore Success!`r`n$StartDateTime : Restore ended at $RestoreStopDateTime"
               Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json)
            }
            }
            catch
            {   
                $FailRestores.Add($BakName)
                $e1 = $_.Exception
                $e1m = $e1.Message
                $e2 = $e1.InnerException
                $e2m = $e1.InnerException.Message
                $e3 = $e1.InnerException.InnerException
                $e3m = $e1.InnerException.InnerException.Message
                $RestoreStopDateTime = Get-Date -Format s
                Write-Verbose -Message "$StartDateTime : Restore failed!"
                Write-Verbose -Message "$StartDateTime : Restore ended at $RestoreStopDateTime"
                Write-Verbose -Message "$StartDateTime : $e1m"
                Write-Verbose -Message "$StartDateTime : $e2m"
                Write-Verbose -Message "$StartDateTime : $e3m"
                Write-Verbose -Message "$StartDateTime : $e1"
                Write-Verbose -Message "$StartDateTime : $e2"
                Write-Verbose -Message "$StartDateTime : $e3"
                if($Enable_API_Logging -eq $true)
                {
                   $JSON.EntryString = "$StartDateTime : Restore failed!`r`n$StartDateTime : Restore ended at $RestoreStopDateTime`r`n$StartDateTime : $e1m`r`n$StartDateTime : $e2m`r`n$StartDateTime : $e3m`r`n$StartDateTime : $e1`r`n$StartDateTime : $e2`r`n$StartDateTime : $e3"
                   Invoke-Dashboard_API -URL_To_API $URL_To_API -JSON_Object $($JSON | ConvertTo-Json)
                }
                if($EmailNotificationsOption -eq $true)
                {
                    ($HtmlErrorLog = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
                    $HtmlErrorLog.Add("<b>Exception Messages</b><br />")
                    $HtmlErrorLog.Add("<b>Inner Exception Inner Exception </b>$e3m<br />")
                    $HtmlErrorLog.Add("<b>Inner Exception </b>$e2m<br />")
                    $HtmlErrorLog.Add("<b>Exception </b>$e1m<br /><br />")
                    $HtmlErrorLog.Add("<b>Full Exception Tracebacks </b><br />")
                    $HtmlErrorLog.Add("<b>Inner Exception Inner Exception </b>$e3<br />")
                    $HtmlErrorLog.Add("<b>Inner Exception </b>$e2<br />")
                    $HtmlErrorLog.Add("<b>Exception </b>$e1")
                    Send-MailMessage -From $FromEmailAddress `
                                        -To $ToFailEmailAddresses `
                                        -Subject "FAIL - Single Restore Runner Report" `
                                        -BodyAsHtml "<h3>Run Details</h3>
                                                    <b>Success Status </b>FAIL<br />
                                                    <b>Start Time Of Restore </b>$StartDateTime<br />
                                                    <<b>Block Of Failure </b> `$restoreObject.SqlRestore(`$serverObject)<br />
                                                    <h3>Environment Details</h3>
                                                    <b>Target SQL Server </b>$SQL_Server<br />
                                                    <b>SQL Login </b>$SqlLogin<br />
                                                    <h3>Restore Details</h3>
                                                    <b>Bak File Path </b>$DB_NameHolder<br />
                                                    <b>Database Name </b>$Holder<br />
                                                    <b>Target MDF Restore Location </b>$MDF_FullName<br />
                                                    <b>Target LDF Restore Location </b>$LDF_FullName<br />
                                                    <h3>Exception Details</h3>
                                                    $HtmlErrorLog" `
                                        -Priority High `
                                        -SmtpServer $SMTP_Server `
                                        -ErrorAction SilentlyContinue
                }
                $Gandalf = $false
            }
        } # End of forloop for files (effectively the folder)
        if($EmailNotificationsOption -eq $true)
        {
            ($HTML_Report = [System.Collections.Generic.List[System.String]]("")).RemoveAt(0)
            $HTML_Report.Add("<b>Successful Restores</b><br />")
            foreach ($Success in $SuccessRestores)
            {
                $HTML_Report.Add("$Success<br />")
            }
            $HTML_Report.Add("<b>Failed Restores</b><br />")
            foreach ($Fail in $FailRestores)
            {
                $HTML_Report.Add("$Fail<br />")
            }
            $HTML_Report.Add("<b>Running Order</b><br />")
            foreach ($Running in $RunningOrder)
            {
                $HTML_Report.Add("$Running<br />")
            }
            if($Gandalf -eq $false)
            {
                Send-MailMessage -From $FromEmailAddress `
                                 -To $ToFailEmailAddresses `
                                 -Subject "FAIL - Complete Restore Runner Report" `
                                 -BodyAsHtml "<h3>General Details</h3>
                                              <b>Success Status </b>FAIL<br />
                                              <b>Bak Count </b>$($RunningOrder.Count)<br />
                                              <b>Start Time Of Run </b>$StartDateTime<br />
                                              <h3>Environment Details</h3>
                                              <b>Target SQL Server </b>$SQL_Server<br />
                                              <b>SQL Login </b>$SqlLogin<br />
                                              <h3>Run Details</h3>
                                              $HTML_Report" `
                                 -Priority High `
                                 -SmtpServer $SMTP_Server
            }
            else
            {
                Send-MailMessage -From $FromEmailAddress `
                                 -To $ToSuccessEmailAddresses `
                                 -Subject "SUCCESS - Complete Restore Runner Report" `
                                 -BodyAsHtml "<h3>General Details</h3>
                                              <b>Success Status </b>SUCCESS<br />
                                              <b>Bak Count </b>$($RunningOrder.Count)<br />
                                              <b>Start Time Of Run </b>$StartDateTime<br />
                                              <h3>Environment Details</h3>
                                              <b>Target SQL Server </b>$SQL_Server<br />
                                              <b>SQL Login </b>$SqlLogin<br />
                                              <h3>Run Details</h3>
                                              $HTML_Report" `
                                 -Priority High `
                                 -SmtpServer $SMTP_Server
            }
        }
    }
}