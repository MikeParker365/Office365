<#
.SYNOPSIS
LicenseAndHold.ps1 - Licenses/Delicenses Office 365 users and enables Litigation hold for new mailboxes.

.NOTES

Version 1.0, 6th January, 2016
Revision History
---------------------------------------------------------------------
1.0 	- Initial release
1.1     - Made setting the litigation hold an optional switch
1.2		- Bug Fixes

Author/Copyright:    Mike Parker - All rights reserved
Email/Blog/Twitter:  mike@mikeparker365.co.uk | www.mikeparker365.co.uk | @MikeParker365

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

.DESCRIPTION
Script to either license and enable litigation hold for a batch of users or remove licenses for a batch of users.
The script will prompt the user for a CSV file via File Explorer. This file needs just one column, "UPN".
Use scenario is for when migrating to Office 365 and want to migrate batches of users just for litigation hold purposes, this process will be useful for managing the process.
Designed for use with any Office 365 license SKU that creates a mailbox.
	
.PARAMETER LicenseSKU
Specifies the license SKU to enable for the users in the batch.

.PARAMETER Enable
Enables the user license and enables litigation hold for the newly created mailbox.

.PARAMETER Hold
Enables Litigation Hold on the new mailbox.

.PARAMETER Disable
Removes the specified license from the Office 365 User.

.LINK
http://www.mikeparker365.co.uk

.EXAMPLE
PS C:\Mike\Powershell\Office365> .\LicenseAndHold.ps1 -LicenseSku ENTERPRISEPACK -Enable

This will enable all of the users in the CSV file with an E3 license and enable litigation hold.

.EXAMPLE
PS C:\Mike\Powershell\Office365> .\LicenseAndHold.ps1 -LicenseSku ENTERPRISEPACK -Disable

This will remove an E3 license from all of the users in the CSV file.
#>

[CmdletBinding()]
param (

	[Parameter( Mandatory=$true )]
	[string]$LicenseSku,
	[Parameter( Mandatory=$false )]
	[switch]$Exchange

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
function ShowError ($msg){Write-Host "`n";Write-Host -ForegroundColor Red $msg; LogErrorToFile $msg }
function ShowSuccess($msg){Write-Host "`n";Write-Host -ForegroundColor Green $msg; LogToFile($msg)}
function ShowProgress($msg){Write-Host "`n";Write-Host -ForegroundColor Cyan $msg; LogToFile ($msg)}
function ShowInfo($msg){Write-Host "`n";Write-Host -ForegroundColor Yellow $msg; LogToFile ($msg)}
function LogSuccessToFile ($msg){"Success: $msg" |Out-File -Append -FilePath $logFile -ErrorAction:SilentlyContinue;}
function LogErrorToFile ($msg){"Error: $msg" |Out-File -Append -FilePath $logFile -ErrorAction:SilentlyContinue;}
function LogToFile ($msg){$msg |Out-File -Append -FilePath $logFile -ErrorAction:SilentlyContinue;}


function SetupOutputFiles ($msg) {
	$error.clear()

 	try{ 
		$headerstring = "Username,Old License,New License,Success/Failure"
		ShowInfo "Creating CSV Log File: $logCSV"
		Out-File -FilePath $logCSV -InputObject $headerstring -Encoding UTF8 -append
	}
	catch{
		ShowError "There was an error creating the log file: $logCSV"
		ShowError "An error has required the script be stopped. Please review the log for further details."
		Exit
		$global:break = $true
	}
	finally{
		if(!$error){
			ShowSuccess "Successfully created CSV Log File: $logCSV"
		}
		else{
			ShowError "There was an error creating the CSV Log File: $logCSV"
			ShowError "See the log for further details."
		}
	}
}


############################################################################
# Functions end 
############################################################################


############################################################################
# Variables Start 
############################################################################

$scriptVersion = "1.2"

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$logfile = "$myDir\LicenseAndHold.log"

$start = Get-Date

$tenantID = ""

$Licenses = Get-MsolAccountSku
Foreach ($Li in $Licenses){
	$tenantID = $Li.AccountName
}

$newLicenseSKU = $tenantID + ":" + $LicenseSKU

$disabledPlans = ('PROJECTWORKMANAGEMENT','SWAY','INTUNE_O365','YAMMER_ENTERPRISE','RMS_S_ENTERPRISE','OFFICESUBSCRIPTION','MCOSTANDARD','SHAREPOINTWAC','EXCHANGE_S_ENTERPRISE')
$exchangeUser = ('PROJECTWORKMANAGEMENT','SWAY','INTUNE_O365','YAMMER_ENTERPRISE','RMS_S_ENTERPRISE','OFFICESUBSCRIPTION','MCOSTANDARD','SHAREPOINTWAC')

$defaultOptions = New-MsolLicenseOptions -AccountSkuId $NewLicenseSKU -DisabledPlans $disabledPlans;
$exchangeOptions = New-MsolLicenseOptions -AccountSkuId $NewLicenseSKU -DisabledPlans $exchangeUser;

$dateForOutput = Get-Date -Format ddMMyyyy
$logCSV = $myDir + "\Set-365License_" + $dateForOutput + ".csv";

############################################################################
# Variables End
############################################################################

############################################################################
# Script start   
############################################################################

Write-Logfile "Script started at $start";
Write-Logfile "Running script version $scriptVersion"

SetupOutputFiles 

Write-Host "Please open the file with the user information.."

$UserList = Get-FileName

$itemCount = Import-Csv $UserList
if($itemCount.Count -eq $null){
	$itemCount = 1
}
else{
	$itemCount = $itemCount.count
}
$processedCount = 1

if($Exchange){

	$defaultOptions = $exchangeOptions
}

import-CSV $UserList | ForEach-Object{
	$error.Clear()

	Write-Progress -Activity "Processing.." -Status "User $processedCount of $itemCount" -PercentComplete ($processedCount / $itemCount * 100)

	$UPN = $_.Username
	$User = get-msoluser -UserPrincipalName $UPN -EA:SilentlyContinue

	<#Testing
	Write-Host $User
	#End of Testing #>

	if ( $User -eq $null ) {

		ShowError "The user $UPN could not be found in the Office 365 Tenant..."
		$datastring = $UPN + ",,," + "Failure - User not found." 

	} # End of IF USER NOT FOUND
	else {

		if( $User.IsLicensed -eq $false){
			Try{
				ShowInfo "The user $UPN is not yet licensed..."
				ShowInfo "Setting the user $UPN with default license option..."

				if($user.UsageLocation -eq $null -or $user.UsageLocation -eq ""){
					ShowProgress "Setting $UPN location for $UPN to United Kingdom (GB)"
					Set-MsolUser -userPrincipalName $UPN -UsageLocation GB
				}
				ShowProgress "Assigning default licensing for user $UPN"

				ShowProgress "The following licenses will be disabled $disabledPlans"

				Set-MsolUserLicense –User $upn -AddLicenses $newLicenseSKU -LicenseOptions $defaultOptions;
			}

			Catch 				{
				#deal with any errors
				ShowError "ERROR - There was an error setting the license for $UPN."
				ShowError "ERROR details - $error"
			}
			Finally{ 
				if(!$error){
					ShowSuccess "Successfully assigned license for user $UPN"

					$datastring = $UPN + ",,SHAREPOINTENTERPRISE," + "Success" 

				}
				else{
					ShowError "There was an error setting the licenses for user $UPN. Review the log."
					ShowError "The latest error logged was - $error"
					$datastring = $UPN + ",None," + $newLicenseSKU + "," + "Failure" 

				}
			}
		} # End of IF USER NOT LICENSED
		else {
			$User = get-msoluser -UserPrincipalName $UPN -EA:SilentlyContinue

			ShowInfo "The user $UPN is already licensed..."
			ShowInfo "Checking existing license..."

			$currentLicense = $User.Licenses.ServiceStatus
			$existingLO = @()
			
			Foreach($i in $currentLicense){ 
				if ($i.ProvisioningStatus -eq "Success") {
					
					$existingLO += $i.ServicePlan.ServiceName 
				} 
			}

			$exchangeEnabled = $false

			foreach ($lo in $currentLicense) {


				if (($lo.ServicePlan.ServiceName -eq "EXCHANGE_S_ENTERPRISE") -and ($lo.ProvisioningStatus -eq 'Success')){
					#LIVE 

					$exchangeEnabled = $true

				}
			}
			<#Testing
			Write-Host "Sharepoint found = $sharepointFound " 
			#EndTesting#>
			if ( $exchangeEnabled -eq $true ){

				try{
					$error.Clear()
					ShowProgress "Assigning Exchange licensing for user $UPN"


					ShowProgress "The following licenses will be disabled $exchangeUser"

					Set-MsolUserLicense –User $upn -LicenseOptions $exchangeOptions;
				}
				Catch 				{
					ShowError "There was an error setting the new license for $UPN..."
					ShowError $error
				}
				Finally{
					if(!$error){
						ShowSuccess "Successfully added new license for $UPN!"

						$newLO = @("SHAREPOINTENTERPRISE","EXCHANGE_S_ENTERPRISE")

						$datastring = $UPN + "," + $existingLO + "," + $newLO + "," + "Success" 

					}
					else{
						ShowError "There was an error setting the new license for $UPN..."
						ShowError $error
						$datastring = $UPN + "," + $existingLO + "," + $newLO + "," + "Failure" 
					}
				} # End of TRY CATCH FINALLY 
			}

			else{
				try{
					$error.Clear()
					ShowProgress "Assigning Default licensing for user $UPN"


					ShowProgress "The following licenses will be disabled $disabledPlans"

					Set-MsolUserLicense –User $upn -LicenseOptions $defaultOptions;
				}
				Catch 				{
					ShowError "There was an error setting the new license for $UPN..."
					ShowError $error
				}
				Finally{
					if(!$error){
						ShowSuccess "Successfully added new license for $UPN!"

						$newLO = @("SHAREPOINTENTERPRISE")

						$datastring = $UPN + "," + $existingLO + "," + $newLO + "," + "Success" 

					}
					else{
						ShowError "There was an error setting the new license for $UPN..."
						ShowError $error
						$datastring = $UPN + "," + $existingLO + "," + $newLO + "," + "Failure" 
					}
				} # End of TRY CATCH FINALLY 
			}
		}
	}

	Out-File -FilePath $logCSV -InputObject $datastring -Encoding UTF8 -append
	$processedCount++
} # End of FOREACH USER IN CSV



Write-Logfile "------------Processing Ended---------------------"
$end = Get-Date;
Write-Logfile "Script ended at $end";
$diff = New-TimeSpan -Start $start -End $end
Write-Logfile "Time taken $($diff.Hours)h : $($diff.Minutes)m : $($diff.Seconds)s ";

############################################################################
# Script end   
############################################################################