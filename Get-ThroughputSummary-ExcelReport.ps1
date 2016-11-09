<#
.SYNOPSIS
Get-ThroughputSummary.ps1 - 

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

Function GenerateReport($BatchName){

	$error.Clear()
	$overviewValues = @()
	#$script:BatchRow

	$overviewValues += $BatchName
	$overviewValues += $Batch.SourceEndpoint.Identity

	$MoveRequestBatchName = "MigrationService:" + $BatchName
	$Mailboxes = Get-MoveRequest -BatchName $MoveRequestBatchName  
	If($Mailboxes.Count -eq $Null){
		$MailboxCount = 1
	}
	Else{
		$MailboxCount = $Mailboxes.Count
	}
	

	$overviewValues += $mailboxCount

	$TotalDataMigratedMB = 0
	$TotalDataMigratedGB = 0
	$TotalSyncDuration = 0
	$processedCount = 1

	$overviewValues += $Batch.SourceEndpoint.MaxConcurrentMigrations
	$MaxIncMigs = $Batch.SourceEndpoint.MaxConcurrentIncrementalSyncs.Value
	
	$overviewValues += $MaxIncMigs
	

	ForEach($Mailbox in $Mailboxes){
		$error.clear()
		Write-Progress -Activity "Generating Throughput Reports. Please wait..." -PercentComplete ($processedCount / $MailboxCount * 100)
		$UserValues = @()
		try{
			$error.clear()
			$DisplayName = $Mailbox.DisplayName
			$UserValues += $DisplayName
			$UserValues += $Mailbox.alias
			$UserValues += $Mailbox.BatchName.Replace("MigrationService:","")
			$MailboxStats = Get-MoveRequestStatistics $Mailbox.Identity 
			
			$MigratedDataMB = [math]::Round(($MailboxStats.BytesTransferred.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2);
			$MigratedDataGB = [math]::Round(($MailboxStats.BytesTransferred.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2);

			If($MailboxStats.InitialSeedingCompletedTimestamp){
				$InitialSyncDuration = New-TimeSpan -Start $MailboxStats.StartTimestamp -End $MailboxStats.InitialSeedingCompletedTimestamp
			}
			Else{
				ShowError "The user $DisplayName has not completed and will not be included in the report"
			}
			
			$UserValues += $InitialSyncDuration
			$UserValues += $MigratedDataMB
			$UserValues += $MigratedDataGB
			
				$UserValues += [math]::Round(($MigratedDataMB/$InitialSyncDuration.TotalMinutes),2)

			# Add this user's info to the totals
			$TotalDataMigratedMB = $TotalDataMigratedMB + $MigratedDataMB
			$TotalDataMigratedGB = $TotalDataMigratedGB + $MigratedDataGB
			$TotalSyncDuration = $TotalSyncDuration + $InitialSyncDuration
			
		}
		catch{
			ShowError "There was an error processing user $DisplayName"
		}
		finally{
			if(!$error){

					$CurrentColumn = 1
					$CurrentRow = $script:UserRow

				foreach ($i in $UserValues)
				{
					[string]$ReportWorksheet2.Cells.Item($CurrentRow,$CurrentColumn).value() = $i
					$CurrentColumn++	
				}
			}
			else{
				ShowError "The user $DisplayName was not processed successfully."
				ShowError $error
				ShowError "This user's details will not be included in the report."
			}
		}

		$processedCount ++
		$script:UserRow ++
	}

	$overviewValues += "$($totalSyncDuration.Hours)h : $($totalSyncDuration.Minutes)m"
	$overviewValues += $totalDataMigratedMB
	$overviewValues += $TotalDataMigratedGB
	$overviewValues += $TotalDataMigratedMB/$totalSyncDuration.TotalMinutes	
	$overviewValues += $TotalDataMigratedGB/$totalSyncDuration.TotalHours
	$avgDurationPerMbx = [math]::Round(($TotalSyncDuration.TotalMinutes / $MailboxCount),2)
	$overviewValues += $avgDurationPerMbx

	# Calculate estimated mailboxes per 12 hours
	$overviewValues += ((60/([int]$avgDurationPerMbx))*12)*($MaxIncMigs)

		$CurrentColumn = 1
		$CurrentRow = $script:BatchRow

	foreach ($i in $overviewValues)
	{
		[string]$ReportWorksheet1.Cells.Item($CurrentRow,$CurrentColumn).value() = $i
		$CurrentColumn++	
	}

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

$UserHeader = @("Display Name","Alias","Batch Name","Initial Sync Duration","Data Moved (MB)","Data Moved (GB)","Throughput (MB/m)")
$OverviewHeader = @("Batch Name","Migration Endpoint","Mailboxes in Batch","Max Concurrent Migrations","Max Concurrent Incremental Syncs","Total Initial Sync Duration","Total Data Transferred (MB)","Total Data Transferred (GB)","Average Throughput (MB/m)","Average Throughput (GB/h)","Average Processing Time per Mbx (mins)","Est. Mailboxes per 12 Hours")

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
# Create a new blank document to work with and make the Excel application visible.

$Report = New-Object -ComObject "Excel.Application"
$Report.Visible = $true
$Report.SheetsInNewWorkbook = 2

$ReportWorkbook = $Report.Workbooks.Add()
$ReportWorksheet1 = $ReportWorkbook.Sheets.Item(1)
$ReportWorksheet1.Name = "Batch Summary"
$ReportWorksheet1.Activate()

# Create Overview Headers
$CurrentRow = 1
$CurrentColumn = 1

Foreach($i in $OverviewHeader){

	[string]$ReportWorksheet1.Cells.Item($CurrentRow,$CurrentColumn).value() = $i
	$ReportWorksheet1.Cells.Item($CurrentRow,$CurrentColumn).Font.Bold=$True
	$ReportWorksheet1.Cells.Item($CurrentRow,$CurrentColumn).Font.Size=12
	$CurrentColumn ++

}

$ReportWorksheet2 = $ReportWorkbook.Sheets.Item(2)
$ReportWorksheet2.Name = "Mailbox Summary"
$ReportWorksheet2.Activate()

# Create User Headers
$CurrentRow = 1
$CurrentColumn = 1

Foreach($i in $UserHeader){

	[string]$ReportWorksheet2.Cells.Item($CurrentRow,$CurrentColumn).value() = $i
	$ReportWorksheet2.Cells.Item($CurrentRow,$CurrentColumn).Font.Bold=$True
	$ReportWorksheet2.Cells.Item($CurrentRow,$CurrentColumn).Font.Size=12

	$CurrentColumn ++

}

$script:BatchRow=2
$script:UserRow = 2

If(!$BatchName){
	$Batches = Get-MigrationBatch
	Foreach($Batch in $Batches){
	
		$BatchName = $Batch.Identity
		GenerateReport($BatchName)
		$script:BatchRow ++

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