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
	[switch]$Enable,

	[Parameter( Mandatory=$false )]
	[switch]$Hold,

	[Parameter( Mandatory=$false )]
	[switch]$Disable

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

############################################################################
# Variables End
############################################################################

############################################################################
# Script start   
############################################################################

Write-Logfile "Script started at $start";
Write-Logfile "Running script version $scriptVersion"

Write-Host "Please open the file with the user information.."

$UserList = Get-FileName
$tenantID = ""

$Licenses = Get-MsolAccountSku
Foreach ($Li in $Licenses){
	$tenantID = $Li.AccountName
}

$newLicenseSKU = $tenantID + ":" + $LicenseSKU

$newLicense = New-MsolLicenseOptions -AccountSkuID $newLicenseSKU -DisabledPlans $null

$itemCount = Import-Csv $UserList
$itemCount = $itemCount.count
$processedCount = 1

if($Enable){

	Write-LogFile "Option to enable users selected..."
	#foreach ($user in @(import-CSV $UserList))
	import-CSV $UserList | ForEach-Object {
		$error.Clear()

		Write-Progress -Activity "Processing.." -Status "User $processedCount of $itemCount" -PercentComplete ($processedCount / $itemCount * 100)
		$UPN = $_.UPN
		$User = get-msoluser -UserPrincipalName $UPN -EA:SilentlyContinue

		#Check the user exists
		if ( $User -eq $null ) {

			Write-Logfile "The user $UPN could not be found in the Office 365 Tenant..."
			$datastring = $UPN + ",,," + "Failure - User not found." 
			$licenseSet = $false

		} # End of IF USER NOT FOUND
		else {

			if( $User.IsLicensed -eq $true){

				$User = get-msoluser -UserPrincipalName $UPN -EA:SilentlyContinue

				Write-Logfile "The user $UPN is already licensed..."

				$currentLicense = $User.Licenses
				Write-Logfile "User has an existing license:"
				Write-Logfile $currentLicense.AccountSkuID

				if((($currentLicense.AccountSkuID).ToLower()) -eq $newLicenseSku) {
					Try{
						$error.clear()
						Write-Logfile "Updating the user license..."
						Set-MsolUserLicense –User $upn -LicenseOptions $newLicense
					 
					}
					Catch 
					{
						#deal with any errors
						Write-Logfile "ERROR - There was an error setting the license for $UPN."
						Write-Logfile "ERROR details - $error"
					}
					Finally{ 
						if(!$error){
							Write-Logfile "Successfully assigned license for user $UPN"
							$datastring = $UPN + ",None," + $newLicenseSKU + "," + "Success" 
							$licenseSet = $true
						}
						else{
							Write-Logfile "There was an error setting the licenses for user $UPN. Review the log."
							Write-Logfile "The latest error logged was - $error"
							$datastring = $UPN + ",None," + $newLicenseSKU + "," + "Failure" 

						}
					}

				}
				Else{
					Write-Logfile "The licenses are not compatible. Please check and re-run."
				}
				Write-Logfile "Moving to next user..."


			} # End of IF USER NOT LICENSED
			else {
				Try{
					Write-Logfile "The user $UPN is not yet licensed..."
					Write-Logfile "Setting the user $UPN with license $LicenseSKU..."

					if($user.UsageLocation -eq $null -or $user.UsageLocation -eq ""){
						Write-Logfile "Setting location for $UPN to United Kingdom (GB)"
						Set-MsolUser -userPrincipalName $UPN -UsageLocation GB
					}

					Set-MsolUserLicense –User $upn -AddLicenses $newLicenseSKU 
				}
				Catch 
				{
					#deal with any errors
					Write-Logfile "ERROR - There was an error setting the license for $UPN."
					Write-Logfile "ERROR details - $error"
				}
				Finally{ 
					if(!$error){
						Write-Logfile "Successfully assigned license for user $UPN"
						$datastring = $UPN + ",None," + $newLicenseSKU + "," + "Success" 
						$licenseSet = $true
					}
					else{
						Write-Logfile "There was an error setting the licenses for user $UPN. Review the log."
						Write-Logfile "The latest error logged was - $error"
						$datastring = $UPN + ",None," + $newLicenseSKU + "," + "Failure" 

					}
				}
			}

			If($licenseSet -and $Hold){
				Write-Logfile "Waiting for mailbox to be provisioned..."
				do {
					sleep -seconds 1
					$mailboxExists = Get-Mailbox $UPN -ErrorAction SilentlyContinue |fw IsValid
					write-host "." -nonewline
				} while (!$mailboxExists)

				if ($mailboxExists) {
					$error.clear()

					Try {
						Set-Mailbox $UPN -LitigationHoldEnabled $true

					}
					Catch{
						Write-Logfile "There was an error setting litigation hold for user $UPN"
						Write-Logfile $error

					}
					Finally{
						If(!$error){
							Write-Logfile "Successfully set litigation hold for user $UPN"
						}
						else{
							Write-Logfile "There was an error setting litigation hold for user $UPN"
							Write-Logfile $error
						}
					}
				}
			}
		}
		$processedCount++
	} # End of foreach user
} # End of Enable Switch

if($Disable){
	import-CSV $UserList | ForEach-Object{
		$error.Clear()

		Write-Progress -Activity "Processing.." -Status "User $processedCount of $itemCount" -PercentComplete ($processedCount / $itemCount * 100)
		$UPN = $_.UPN
		$User = get-msoluser -UserPrincipalName $UPN -EA:SilentlyContinue

		Write-Logfile "Option to disable users selected..."

		#Check the user exists
		if ( $User -eq $null ) {

			Write-Logfile "The user $UPN could not be found in the Office 365 Tenant..."
			$datastring = $UPN + ",,," + "Failure - User not found." 
			$licenseSet = $false

		} # End of IF USER NOT FOUND

		else{

			$currentLicense = $User.Licenses

			foreach ($license in $currentLicense) {

				if ($license.AccountSkuId.ToLower() -like $newLicenseSKU.ToLower()) { # 

					$licenseFound = $true
				}
			}

			#Remove the license if it matches the user input
			if($licensefound -eq $true){

				Try{ 
					$error.Clear()

					Write-Logfile "The user $UPN currently has license..."
					Write-Logfile "Removing existing license..."
					Set-MsolUserLicense –User $upn –RemoveLicenses $newlicenseSKU

				}
				Catch 
				{
					Write-Logfile "There was an error removing the old License..."
					Write-Logfile $error
				}
				Finally{
					if(!$error){
						Write-Logfile "Successfully removed old license for $UPN"
					}
					else{
						Write-Logfile "There was an error removing the old license..."
						Write-Logfile $error

					}
				}
			} # End of license found
			else{
				Write-LogFile "User license did not match the provided SKU"
			}

		} # End of Else (The user exists)
		$processedCount++

	}# End of foreach user
} #End of disable switch

Write-Logfile "------------Processing Ended---------------------"
$end = Get-Date;
Write-Logfile "Script ended at $end";
$diff = New-TimeSpan -Start $start -End $end
Write-Logfile "Time taken $($diff.Hours)h : $($diff.Minutes)m : $($diff.Seconds)s ";

############################################################################
# Script end   
############################################################################