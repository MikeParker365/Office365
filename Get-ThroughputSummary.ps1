<#
.SYNOPSIS
LicenseAndHold.ps1 - Licenses/Delicenses Office 365 users and enables Litigation hold for new mailboxes.

.NOTES

Version 1.0, 6th January, 2016
Revision History
---------------------------------------------------------------------
1.0 	- Initial release

1.1     - Made setting the litigation hold an optional switch

Author/Copyright:    Mike Parker - All rights reserved
Email/Blog/Twitter:  mike@mikeparker365.co.uk | www.mikeparker365.co.uk | @MikeParker365

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

.DESCRIPTION
Script to either license and enable litigation hold for a batch of users or remove licenses for a batch of users.
The script will prompt the user for a CSV file via File Explorer. This file needs just one column, "UPN".
Use scenario is for when migrating to Office 365 and want to migrate batches of users just for litigation hold purposes, this process will be useful for managing the process.
Designed for use with any Office 365 license SKU that creates a mailbox.
	
.PARAMETER BatchName
Specifies the batch which you would like throughput information for.

.LINK
http://www.mikeparker365.co.uk

.EXAMPLE
PS C:\Mike\Powershell\Office365> .\Get-ThroughputSummary.ps1 -BatchName Test1

This will generate a report on the throughput for the batch named Test1.

#>

[CmdletBinding()]
param (

	[Parameter( Mandatory=$false )]
	[string]$BatchName

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

function New-ExcelReport ($object,$counter)
{
	#Create Excel Worbook
	if ($counter -eq 1)
	{
		$script:ExcelWorkSheets = $script:ExcelWorkbook.WorkSheets
		$script:CurrentWorkSheet = $script:ExcelWorkSheets.Item($counter)
		$script:CurrentWorkSheet.Name = $object[0].AccountSku 
	}
	else
	{
		$script:ExcelWorkSheets.Add()
		$script:CurrentWorkSheet = $script:ExcelWorkSheets.Item(1)
		$script:CurrentWorkSheet.Name = $object[0].AccountSku 
	}
	
	#Get Attributes
	$TempLicenseAttributes = $object | Get-Member -MemberType NoteProperty | Select -Expand Name
	[System.Collections.ArrayList]$LicenseAttributes = $TempLicenseAttributes
	
	$LicenseAttributes.Remove("UserPrincipalName")
	$LicenseAttributes.Remove("DisplayName")
	$LicenseAttributes.Remove("AccountSku")

	$NumRows = $object.Count - 1
	$NumCols = ($object | measure).Count + 1

	$CurrentRow = 1
	$CurrentColumn = 1

	#Create Headers
	[string]$script:CurrentWorkSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = "UserPrincipalName"
	$CurrentColumn = 2
	[string]$script:CurrentWorkSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = "DisplayName"
	$CurrentColumn = 3
	[string]$script:CurrentWorkSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = "AccountSku"
	$CurrentColumn = 4
	foreach ($AttributeName in $LicenseAttributes)
	{
		if ($AttributeName -ne "Anchor")
		{
			[string]$script:CurrentWorkSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $AttributeName
			$CurrentColumn++
		}
	}

	#Add Content
	$CurrentRow = 2
	foreach ($item in $object)
	{
		$CurrentColumn = 1
		[string]$script:CurrentWorkSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $item.UserPrincipalName
		$CurrentColumn++
		[string]$script:CurrentWorkSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $item.DisplayName
		$CurrentColumn++
		[string]$script:CurrentWorkSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $item.AccountSku
		$CurrentColumn++

		foreach ($AttributeName in $LicenseAttributes)
		{
			[string]$script:CurrentWorkSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $item | Select -ExpandProperty $AttributeName
			$CurrentColumn++
		}

		$CurrentRow++
	}
}

Function GenerateReport($BatchName){
	$MoveRequestBatchName = "MigrationService:" + $BatchName
	$Mailboxes = Get-MoveRequest -BatchName $MoveRequestBatchName  
	$MailboxCount = $Mailboxes.Count
	$TotalDataMigratedMB = 0
	$TotalDataMigratedGB = 0
	$TotalSyncDuration = 0
	$processedCount = 1
	$MigrationEndpoint = $Batch.SourceEndpoint.Identity
	$MaxConMigs = $Batch.SourceEndpoint.MaxConcurrentMigrations
	$MaxIncMigs = $Batch.SourceEndpoint.MaxConcurrentIncrementalSyncs


	ForEach($Mailbox in $Mailboxes){
		$error.clear()
		Write-Progress -Activity "Generating Throughput Reports. Please wait..." -PercentComplete ($processedCount / $MailboxCount * 100)
		
		try{
			$error.clear()
			$Alias = $Mailbox.alias
			$DisplayName = $Mailbox.DisplayName.Replace(",","")

			$MailboxStats = Get-MoveRequestStatistics $Mailbox.Identity 
			
			$MigratedDataMB = [math]::Round(($MailboxStats.BytesTransferred.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2);
			$MigratedDataGB = [math]::Round(($MailboxStats.BytesTransferred.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2);

			If($MailboxStats.InitialSeedingCompletedTimestamp){
				$InitialSyncDuration = New-TimeSpan -Start $MailboxStats.StartTimestamp -End $MailboxStats.InitialSeedingCompletedTimestamp
			}
			Else{
				ShowError "The user $DisplayName has not completed and will not be included in the report"
			}

			$avgMBperMin = [math]::Round(($MigratedDataMB/$InitialSyncDuration.TotalMinutes),2)

			# Add this user's info to the totals
			$TotalDataMigratedMB = $TotalDataMigratedMB + $MigratedDataMB
			$TotalDataMigratedGB = $TotalDataMigratedGB + $MigratedDataGB
			$TotalSyncDuration = $TotalSyncDuration + $InitialSyncDuration
			

			$datastring = "$DisplayName,$alias,$InitialSyncDuration,$MigratedDataMB,$MigratedDataGB,$avgMBperMin"
		}
		catch{
			ShowError "There was an error processing user $DisplayName"
		}
		finally{
			if(!$error){
				Out-File -FilePath $UserOutput -InputObject $datastring -Encoding UTF8 -append
			}
			else{
				ShowError "The user $DisplayName was not processed successfully."
				ShowError $error
				ShowError "This user's details will not be included in the report."
			}
		}

		$processedCount ++

	}

	$avgThroughputMBperMin = $TotalDataMigratedMB/$totalSyncDuration.TotalMinutes
	$avgThroughputGBperHr = $TotalDataMigratedGB/$totalSyncDuration.TotalHours
	$avgDurationPerMbx = [math]::Round(($TotalSyncDuration.TotalMinutes / $MailboxCount),2)

	# Calculate estimated mailboxes per 12 hours
	$est12Hrs = ((60/([int]$avgDurationPerMbx))*12)*([int32]($MaxIncMigs.ToString()))

	$datastring = "$BatchName,$MigrationEndpoint,$MailboxCount,$MaxConMigs,$MaxIncMigs,$($totalSyncDuration.Hours)h : $($totalSyncDuration.Minutes)m,$TotalDataMigratedMB,$TotalDataMigratedGB,$avgThroughputMBperMin,$avgThroughputGBperHr,$avgDurationPerMbx,$est12Hrs"
	
	Out-File -FilePath $OverviewOutput -InputObject $datastring -Encoding UTF8 -append
}
############################################################################
# Functions end 
############################################################################


############################################################################
# Variables Start 
############################################################################

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$tenantID = ""

$Licenses = Get-MsolAccountSku
Foreach ($Li in $Licenses){
	$tenantID = $Li.AccountName
}
$LogFolder = "$myDir\Logs\$tenantID"

if( (Test-Path ($LogFolder)) -eq $false)
{
	New-Item $LogFolder -ItemType Directory
	Write-Host "Directory $LogFolder created successfully.";
}


$logfile = "$LogFolder\ThroughputSummary.log"

$start = Get-Date

$scriptVersion = "0.1"

$UserHeaderString = "Display Name,Alias,Initial Sync Duration,Data Moved (MB),Data Moved (GB),Throughput (MB/m)"
$OverviewHeaderString = "Batch Name,Migration Endpoint,Mailboxes in Batch,Max Concurrent Migrations,Max Concurrent Incremental Syncs,Total Initial Sync Duration,Total Data Transferred (MB),Total Data Transferred (GB),Average Throughput (MB/m),Average Throughput (GB/h),Average Processing Time per Mbx (mins),Est. Mailboxes per 12 Hours"

$UserOutput = "$LogFolder\MailboxSummary.csv"
$OverviewOutput = "$LogFolder\OverviewSummary.csv"

############################################################################
# Variables End
############################################################################

############################################################################
# Script start   
############################################################################

Write-Logfile "Script started at $start";
Write-Logfile "Running script version $scriptVersion"

$error.clear()

# Create the output files
# User Summary
try{
	ShowInfo "Creating CSV File: $UserOutput"
	Out-File -FilePath $UserOutput -InputObject $UserHeaderString -Encoding UTF8 -append
}
catch{
	ShowError "There was an error creating the log file: $UserOutput"
	ShowError $error
	ShowError "An error has required the script be stopped. Please review the log for further details."
	Exit
}
finally{
	if(!$error){
		ShowSuccess "Successfully created CSV Log File: $UserOutput"
		ShowInfo ""
	}
	else{
		ShowError "There was an error creating the CSV Log File: $UserOutput"
		ShowError "See the log for further details."
	}
}

# Overview Summary
try{
	ShowInfo "Creating CSV File: $OverviewOutput"
	Out-File -FilePath $OverviewOutput -InputObject $OverviewHeaderString -Encoding UTF8 -append
}
catch{
	ShowError "There was an error creating the log file: $OverviewOutput"
	ShowError $error
	ShowError "An error has required the script be stopped. Please review the log for further details."
	Exit
}
finally{
	if(!$error){
		ShowSuccess "Successfully created CSV Log File: $OverviewOutput"
		ShowInfo ""
	}
	else{
		ShowError "There was an error creating the CSV Log File: $OverviewOutput"
		ShowError "See the log for further details."
	}
}

# Check the batch name is correct

If(!$BatchName){
	
	$Batches = Get-MigrationBatch
	Foreach($Batch in $Batches){
	
		$BatchName = $Batch.Identity
		GenerateReport($BatchName)

	}
}
Else{
	$Batch = Get-MigrationBatch -Identity $BatchName -EA:SilentlyContinue

	If(!$Batch){

		ShowError "The batch name $BatchName does not exist"
	}
	Else{

		GenerateReport($BatchName)
	
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