<#
 Remove-CalendarItemsCSV.ps1

 Removes Calendar Items from a CSV from the calendar of multiple users
 Gets data from APIData
 
 Needs EwsManagedApi 2.2 to be installed on the system where this script runs
 Change the path of de log directory the csv file and the password.txt file to the correct path

 By Erwin Rook.
#>

Function LogProgress{
	param($Context,$Status,$Message)
	#
	# Context: Task that is being executed in between square brackets. Like: [CreateNewUser]
	# Status: One of the following values: [INFO],[ERROR]
	# Message: Explain what is being executed or what the error is.
	#
	If ($Context.length -lt 20){1..(20-$Context.length) | %{$Context += " " }}
	[string]$ExportString = "$($Context)`t$($Status)`t$($Message)"
	Add-Content -Path $LogFile -Value $ExportString
		
}

#Set Global Variables
$Password = Get-Content "C:\ScheduledTasks\AfasLeaveData\password.txt" | ConvertTo-SecureString
$Credential = New-Object System.Management.Automation.PsCredential("serviceaccount",$password)
$EndpointUri = "https://url:7843/api/v1/getCanceledLeaveData"
$ScriptPath = "C:\ScheduledTasks\AfasLeaveData\"
$CSVFileName = $ScriptPath + "CanceledLeaveDataExchange" + (Get-Date -Format "ddMMyyyy") + ".csv"
$ProcessedPath = "C:\ScheduledTasks\AfasLeaveData\Processed\"
$LogPath = "C:\ScheduledTasks\AfasLeaveData\Logs\"
$Impersonate = $true
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Constants
$LogFile = $LogPath + "RemoveCalendarItemFromAfas-" + (Get-Date -Format "ddMMyyyy-HHmm") + ".log"

# Log Settings
$Context = "RemoveCalendarItemFromAfas"

# Start Logging
$Today = Get-Date
LogProgress $Context "[START]" "*** Start Logging: $($Today) ***"

$Context = "Load Exchange CmdLets"

# Load PSSessions
If((Get-PSSession | where{$_.ConfigurationName -match "Microsoft.Exchange"}) -eq $null){
   
   Try {
   $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mailserver/PowerShell/ -Authentication Kerberos -Credential $Credential -ErrorAction Stop
   Import-PSSession $session 
   LogProgress $Context "[SUCCESS]" "Exchange CmdLets loaded successfully"
   }
   Catch {
   LogProgress $Context "[ERROR]" "Error loading Exchange CmdLets"
   }
}

$Context = "Test Website"

# First we create the request.
$proxy = [System.Net.WebProxy]::new("http://url:8080")
$HTTP_Request = [System.Net.WebRequest]::Create($EndpointUri)
$HTTP_Request.Credentials = $Credential
$HTTP_Request.Proxy = $proxy 

# We then get a response from the site.
$HTTP_Response = $HTTP_Request.GetResponse()

# We then get the HTTP code as an integer.
$HTTP_Status = [int]$HTTP_Response.StatusCode

If ($HTTP_Status -eq 200) {

     LogProgress $Context "[SUCCESS]" "Site is OK!"
}
Else {
   
   LogProgress $Context "[ERROR]" "The Site may be down, please check!"
   Exit
}

# Finally, we clean up the http request by closing it.
If ($HTTP_Response -eq $null) { } 
Else { $HTTP_Response.Close() 

}

#Check and create CSV file
$Context = "Create  CSV file"
$Files = Get-ChildItem -Path $ScriptPath -Filter "CanceledLeaveDataExchange*"

If ($Files -eq $nul) {
    LogProgress $Context "[SUCCESS]" " CSV dos not exist in $ScriptPath"
    }

Else {
foreach ($file in $files) {
    Move-Item -Path $File -Destination $ProcessedPath
    LogProgress $Context "[WARNING]" " $File already exists in $ScriptPath file moved to $ProcessedPath"
    }
}

$APIData = Invoke-RestMethod -Uri $EndpointUri -Credential $Credential -OutFile $CSVFileName -Proxy "http://url:8080"

$RequiredFields=@{
	"StartDate" = "Start Date";
	"StartTime" = "Start Time";
	"EndDate" = "End Date";
	"EndTime" = "End Time"
}
 
# Import CSV File
# Log Settings
$Context = "Import CSV"
try
{
	$CSVFile = Import-Csv -Path $CSVFileName;
    LogProgress $Context "[SUCCESS]" " CSV file imported successfully"
}
catch { 

LogProgress $Context "[ERROR]" " CSV file not found"

}

if (!$CSVFile)
{
    LogProgress $Context "[ERROR]" " CSV header line not found, using predefined header: StartDate;StartTime;EndDate;EndTime"	
	$CSVFile = Import-Csv -Path $CSVFileName -header StartDate,StartTime,EndDate,EndTime
}

# Check file has required fields
foreach ($Key in $RequiredFields.Keys)
{
	if (!$CSVFile[0].$Key)
	{
		# Missing required field
		LogProgress $Context "[ERROR]" " Import file is missing required field: $Key"
	}
}
 
# Check EWS Managed API available
$Context = "Check EWS Managed API available"
$EWSManagedApiPath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

 if (!(Get-Item -Path $EWSManagedApiPath -ErrorAction SilentlyContinue))
 {
     LogProgress $Context "[ERROR]" " EWS Managed API could not be found at $($EWSManagedApiPath)."
 }
 
# Load EWS Managed API
 [void][Reflection.Assembly]::LoadFile($EWSManagedApiPath);
 
# Create Service Object. 

$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

# Set credentials if specified, or use logged on user.

$service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Credential)

# Get User information   
foreach ($Item in $CSVFile) {

$Context = "Get EmailAddress of ITCode"
try
{
                $MailboxUser = $Item.ITCode
                $EmailAddress = Get-mailbox -Identity $MailboxUser 
                $EmailAddress = $EmailAddress.PrimarySmtpAddress
				$Subject = "MyPlace Leave Booking (automatically created)"
}
catch
{
                LogProgress $Context "[ERROR]" " Could not find email of $Item.ITCode"
}


# Use autodiscover
$Context = "Autodiscover"
try
	{
		$service.AutodiscoverUrl($EmailAddress)
        LogProgress $Context "[SUCCESS]" " Performing autodiscover for $EmailAddress"
	}
	catch
	{
		LogProgress $Context "[ERROR]" " Autodiscover for $EmailAddress not successfull"
	}

 
# Bind to the calendar folder
$Context = "Open user calendar" 
if ($Impersonate)
{
	$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
}
try {
	$CalendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
    LogProgress $Context "[SUCCESS]" " Calendar for user $MailboxUser opened successfully"
} catch {
	LogProgress $Context "[ERROR]" " Cannot open calendar for user $MailboxUser"
}

# Parse the CSV file and add the appointments
  # Create the appointment and set the fields
  $Context = "Find calendar item"
  
  try
	{
        $StartDate=[DateTime]($Item."StartDate" + " " + $Item."StartTime")
        $EndDate=[DateTime]($Item."EndDate" + " " + $Item."EndTime")
	    $CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,10)    
        $fiItems = $service.FindAppointments($CalendarFolder.Id,$CalendarView)
        $FiItems = $fiItems | Where {$_.Subject -eq $Subject} | Where {$_.Start -eq $StartDate} | Where {$_.End -eq $EndDate}
        if($fiItems){
            $NoError=$true
            $type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
            $type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.Item" -as "Type")
            $ItemColl = [Activator]::CreateInstance($type)
            foreach($Items in $fiItems.Items){
                $ItemColl.Add($Items)
            } 
        LogProgress $Context "[SUCCESS]" " Calendar Item found. Required fields set successfully" 
            }
    ELSE {
        $NoError=$false
        LogProgress $Context "[ERROR]" " Calendar Item not found" }
    }
	catch
	{
		# If we fail to set any of the required fields, we will not write the appointment
		$NoError=$false
        LogProgress $Context "[ERROR]" " Failed to set the Required fields"
	}
	if ($NoError)
	{
		$fiItems.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
        LogProgress $Context "[SUCCESS]" " successfully deleted item $($fiItems.Subject) $StartDate $EndDate"
            }
	ELSE
	{
		# Failed to set a required field
		LogProgress $Context "[ERROR]" " Failed to delete appointment: $($fiItems.Subject)"
	    }

$fiItems = $null
$MailboxUser = $null
$EmailAddress = $null 
$CalendarFolder = $null
$CalendarView = $null
$StartDate = $null
$EndDate = $null
$Item = $null
$NoError = $null
}

$Context = "Processed"
Move-Item -Path $CSVFileName -Destination $ProcessedPath
LogProgress $Context "[SUCCESS]" " All items processed successfully file $CSVFileName moved to $ProcessedPath"

$Today = Get-Date
$Context = "ImportCalendarFromAfas"
LogProgress $Context "[STOP]" "*** End Logging: $($Today) ***"