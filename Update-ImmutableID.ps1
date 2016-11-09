<#
.SYNOPSIS
Update-ImmutableID.ps1 - Updates a list of users ImmutableID's from a csv file.

.NOTES

Version 1.0, 20th July 2016
Revision History
---------------------------------------------------------------------
1.0 	- Initial release

Author/Copyright:    Mike Parker - All rights reserved
Email/Blog/Twitter:  mike@mikeparker365.co.uk | www.mikeparker365.co.uk | @MikeParker365

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

.DESCRIPTION
Script to update a list of users ImmutableID's from a csv file.
	
.PARAMETER CSV
Specify the URL of the CSV file containing the list of users.

.LINK
http://www.mikeparker365.co.uk

.EXAMPLE
PS C:\Mike\Powershell\Office365> .\Update-ImmutableID.ps1 -CSV C:\Mike\Users.csv

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

$logfile = "$myDir\Update-ImmutableID.log"

$start = Get-Date

$scriptVersion = "0.1"

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

if(!$csv){
	# Import CSV with user accounts
	Write-Host "Please select file containing user accounts..."
	$csv = Get-FileName

}

$Users = Import-Csv $csv 

# Process each user
Foreach($User in $Users){

	$UPN = $User.UserPrincipalName
	$ID = $User.ImmutableID

	$CheckUser = Get-MsolUser -UserPrincipalName $UPN -EA:SilentlyContinue | Select ImmutableID

	If($checkuser){
		$oldImmutableID = $CheckUser.ImmutableID
		ShowInfo "Updating user $upn."
		ShowInfo "Current ImmutableID is: '$oldImmutableID' "
		ShowInfo "Setting ImmutableID to: '$ID'"

		try{
			$error.clear()
			Set-MsolUser -UserPrincipalName $UPN -ImmutableId $ID
		}
		catch{
			ShowError "There was an error updating the immutable ID for user $UPN."
			ShowError $error
		}
		finally{
			If(!$error){
				ShowSuccess "Successfully set immutableID for user $UPN to $ID"
			}
			Else{
				ShowError "User $UPN was not updated."
				ShowError $error
			}
		}
	}
	Else{
		ShowError "The user $upn couldn't be found in Office 365..."
	}
} 



Write-Logfile "------------Processing Ended---------------------"
$end = Get-Date;
Write-Logfile "Script ended at $end";
$diff = New-TimeSpan -Start $start -End $end
Write-Logfile "Time taken $($diff.Hours)h : $($diff.Minutes)m : $($diff.Seconds)s ";

############################################################################
# Script end   
############################################################################
