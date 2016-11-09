<#
.SYNOPSIS
ConvertGUIF.ps1 - Generates the ImmutableID attribute from ObjectGUIDs to help with manually hardmatching objects with AAD Connect.

.NOTES

Version 1.0, 20th July 2016
Revision History
---------------------------------------------------------------------
1.0 	- Initial release

Author/Copyright:    Mike Parker - All rights reserved
Email/Blog/Twitter:  mike@mikeparker365.co.uk | www.mikeparker365.co.uk | @MikeParker365

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

.DESCRIPTION
Script to convert a list of users' ObjectGUIDs to the correct ImmutableID value to enable hardmatching with Office 365 Accounts.
	
.PARAMETER CSV
Specify the URL of the CSV file containing the list of users.

.LINK
http://www.mikeparker365.co.uk

.EXAMPLE
PS C:\Mike\Powershell\Office365> .\ConvertGUID.ps1 -CSV C:\Mike\Users.csv

This will import the csv and run the conversion.
#>

[CmdletBinding()]
param (

	[Parameter( Mandatory=$false )]
	[string]$csv

	)

############################################################################
# Functions Start 
############################################################################

#Retrieves the path the script has been run from
function Get-ScriptPath
{ Split-Path $myInvocation.ScriptName
}

#This function is used to write the log file
Function Write-Logfile()
{
 param( $logentry )
$timestamp = Get-Date -DisplayHint Time
"$timestamp $logentry" | Out-File $logfile -Append
Write-Host $logentry
}

#This function enables you to locate files using the file explorer
function Get-FileName($initialDirectory) { 
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
	Out-Null

	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "All files (*.*)| *.*"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
} #end function Get-FileName

function ShowError  ($msg){Write-Host "`n";Write-Host -ForegroundColor Red $msg;   LogErrorToFile  $msg }
function ShowSuccess($msg){Write-Host "`n";Write-Host -ForegroundColor Green  $msg; Write-Logfile   ($msg)}
function ShowProgress($msg){Write-Host "`n";Write-Host -ForegroundColor Cyan  $msg; Write-Logfile   ($msg)}
function ShowInfo($msg){Write-Host "`n";Write-Host -ForegroundColor Yellow  $msg; Write-Logfile   ($msg)}
function Write-Logfile   ($msg){$msg |Out-File -Append -FilePath $logFile -ErrorAction:SilentlyContinue;}
function LogSuccessToFile   ($msg){"Success: $msg" |Out-File -Append -FilePath $logFile -ErrorAction:SilentlyContinue;}
function LogErrorToFile   ($msg){"Error: $msg" |Out-File -Append -FilePath $logFile -ErrorAction:SilentlyContinue;}

function guidtobase64
{
    param($str);
    $g = new-object -TypeName System.Guid -ArgumentList $str;
    $b64 = [System.Convert]::ToBase64String($g.ToByteArray());
    return $b64;
}
############################################################################
# Functions end 
############################################################################


############################################################################
# Variables Start 
############################################################################

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$logfile = "$myDir\ThroughputSummary.log"

$start = Get-Date

$scriptVersion = "0.1"

$output = "$mydir\GUID-to-Immutable.csv"

############################################################################
# Variables End
############################################################################

############################################################################
# Script start   
############################################################################

Write-Logfile "Script started at $start";
Write-Logfile "Running script version $scriptVersion"
# Setup output file

$error.clear()


$headerstring = "Display Name,UserPrincipalName,365UPN,ObjectGuid,ImmutableID"

try{
	ShowInfo "Creating CSV File: $output"
	Out-File -FilePath $output -InputObject $headerstring -Encoding UTF8 -append
}
catch{
	ShowError "There was an error creating the log file: $output"
	ShowError $error
	ShowError "An error has required the script be stopped. Please review the log for further details."
	Exit
}
finally{
	if(!$error){
		ShowSuccess "Successfully created CSV Log File: $output"
		ShowInfo ""
	}
	else{
		ShowError "There was an error creating the CSV Log File: $output"
		ShowError "See the log for further details."
	}
}

if(!$csv){
	# Import CSV with user accounts
	Write-Host "Please select file containing user accounts..."
	$csv = Get-FileName

}

$Users = Import-Csv $csv 
$TotalUsers = $Users.Count
$processedCount = 1
# Process each user
Foreach($User in $Users){
	Write-Progress -Activity "Converting GUIDs to ImmutableIDs. Please wait..." -PercentComplete ($processedCount / $TotalUsers * 100)
	$guid = $User.ObjectGuid
	$ImmutableID = guidtobase64($guid)
	$365UPN = $User.SamAccountName + "@brakesgroup.onmicrosoft.com"
	If(($User.GivenName) -and ($User.Surname)){
		$DisplayName = $User.GivenName + " " + $User.Surname
	}
	Else{
		$DisplayName = $User.Name
	}
	$datastring = ("$DisplayName," + $User.UserPrincipalName + ",$365UPN,$guid,$ImmutableID")

	Out-File -FilePath $output -InputObject $datastring -Encoding UTF8 -append
	$processedCount ++
} 



Write-Logfile "------------Processing Ended---------------------"
$end = Get-Date;
Write-Logfile "Script ended at $end";
$diff = New-TimeSpan -Start $start -End $end
Write-Logfile "Time taken $($diff.Hours)h : $($diff.Minutes)m : $($diff.Seconds)s ";

############################################################################
# Script end   
############################################################################
