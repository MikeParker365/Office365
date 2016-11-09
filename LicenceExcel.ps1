[CmdletBinding()] 
param (
	[parameter(ParameterSetName='p0',Position=0,Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Connected')][switch]$Connected
)




$Sku = @{
	"DESKLESSPACK" = "Office 365 (Plan K1)"
	"DESKLESSWOFFPACK" = "Office 365 (Plan K2)"
	"LITEPACK" = "Office 365 (Plan P1)"
	"EXCHANGESTANDARD" = "Office 365 Exchange Online Only"
	"STANDARDPACK" = "Office 365 (Plan E1)"
	"STANDARDWOFFPACK" = "Office 365 (Plan E2)"
	"ENTERPRISEPACK" = "Office 365 (Plan E3)"
	"ENTERPRISEPACKLRG" = "Office 365 (Plan E3)"
	"ENTERPRISEWITHSCAL" = "Office 365 (Plan E4)"
	"STANDARDPACK_STUDENT" = "Office 365 (Plan A1) for Students"
	"STANDARDWOFFPACKPACK_STUDENT" = "Office 365 (Plan A2) for Students"
	"ENTERPRISEPACK_STUDENT" = "Office 365 (Plan A3) for Students"
	"ENTERPRISEWITHSCAL_STUDENT" = "Office 365 (Plan A4) for Students"
	"STANDARDPACK_FACULTY" = "Office 365 (Plan A1) for Faculty"
	"STANDARDWOFFPACKPACK_FACULTY" = "Office 365 (Plan A2) for Faculty"
	"ENTERPRISEPACK_FACULTY" = "Office 365 (Plan A3) for Faculty"
	"ENTERPRISEWITHSCAL_FACULTY" = "Office 365 (Plan A4) for Faculty"
	"ENTERPRISEPACK_B_PILOT" = "Office 365 (Enterprise Preview)"
	"STANDARD_B_PILOT" = "Office 365 (Small Business Preview)"
	"EXCHANGEENTERPRISE" = "Office 365 (Exchange Plan 2)"
	"SHAREPOINTENTERPRISE" = "Office 365 (SharePoint Plan 2)"
	"EMS" = "Enterprise Mobility Suite"
	"POWERAPPS_INDIVIDUAL_USER" = "Power BI"
	"VISIOCLIENT" = "Visio Pro for Office 365"
	"POWER_BI_STANDARD" = "Power BI (free)"
	"PROJECTCLIENT" = "Project Pro for Office 365"
}

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

function GenerateVariable ($Total)
{
	for ($i = 1; $i -le $Total; $i++)
	{ 
		$Name = $null
		$Name = "MSOLSku$i"

		Set-Variable -Name $Name -Value @() -Scope Script
	}
}

Function GenerateObject ($LicenseNumber,$header,$Value)
{
	$hash = [ordered]@{}
	$Name = "MSOLSku$LicenseNumber"

	for ($i = 0; $i -lt $header.count; $i++)
	{ 
		$hash.Add($header[$i],$Value[$i])
	}

	$var = Get-Variable -Name $Name -ValueOnly -Scope Script

	$object = new-object psobject -Property $hash
	$var += $object
	Set-Variable -Name $Name -Value $var -Scope Script
	$null
}



#Main
If ($Connected)
{
}
else
{
	Try
	{
		
		Write-Progress -Activity "Connecting to Office 365" -Status "Connecting...";
		$Credentials = Get-Credential
		Connect-MsolService -Credential $Credentials -ErrorAction Stop
		Write-Progress -Activity "Connecting to Office 365" -Completed;
	}
	catch
	{
		Write-Error $Error[0].Exception.Message
	}
}

Write-Progress -Activity "Gathering Tenant information" -Status "Get MSOL Licenses";
$licensetype = Get-MsolAccountSku | Where {$_.ConsumedUnits -ge 1} -ErrorAction Stop
GenerateVariable -Total $licensetype.Count

Write-Progress -Activity "Gathering Tenant information" -Status "Get MSOL users UPN, Licenses, DisplayName";
$MSOLUSERS =  Get-MsolUser -All  | Where-Object {$_.IsLicensed -eq $true} | Select-Object UserPrincipalName,DisplayName,Licenses

$CounterLicense = 0
foreach ($license in $licensetype)
{	
	$CounterLicense ++

	$headers = @("UserPrincipalName","DisplayName","AccountSku")
							  
	#Write-Progress -Activity "Processing each Sku"
	# Build and write the Header for the CSV file
	$headerstring = "UserPrincipalName,DisplayName,AccountSku"
	
	foreach ($row in $($license.ServiceStatus)) 
	{
		# Build header string
		switch -wildcard ($($row.ServicePlan.servicename))
		{
			"EXCHANGE_*" { $thisLicence = "Exchange Online" }
			"EXCHANGE_ANALYTICS*" {$thisLicence = "Delve Analytics"}
			"MCOSTANDARD*" { $thisLicence = "Skype for Business" }
			"MCOMEETADV*" {$thisLicence = "Skype for Business PSTN Conferencing"}
			"MCOEV*" {$thisLicence = "Skype for Business PSTN Conferencing"}
			"LYN*" { $thisLicence = "Skype for Business Cloud PBX" }
			"OFF*" { $thisLicence = "Office Profesional Plus" }
			"SHA*" { $thisLicence = "Sharepoint Online" }
			"*WAC*" { $thisLicence = "Office Web Apps" }
			"WAC*" { $thisLicence = "Office Web Apps" }
			"RMS_S_ENTERPRISE*" { $thisLicence = "Azure Rights Management" }
			"RMS_S_PREMIMUM*" { $thisLicence = "Azure Rights Management Premium" }
			"Intune*" { $thisLicence = "Intune" }
			"AAD*" { $thisLicence = "Azure Active Directory Premium" }
			"MFA*" { $thisLicence = "Mulit-Factor Auth Azure AD" }
			"ATP_ENTERPRISE*" {$thisLicence = "Advanced Threat Protection"}
			"BI_AZURE_P2*" {$thisLicence = "Power BI Pro"}
			default { $thisLicence = $row.ServicePlan.servicename }
		}
		
		$headerstring = ($headerstring + "," + $thisLicence)
		$headers += $thisLicence
	}

	$null
	
	#Out-File -FilePath $File -InputObject $headerstring -Encoding UTF8 -append
	#   
	$counter = 0
	foreach ($user in $MSOLUSERS)
	{
		$counter ++
		Write-Progress -Activity "Processing users Licenses" -Status "Processing $($license.AccountSkuID)" -CurrentOperation "$($counter) / $($MSOLUSERS.Count)" -PercentComplete $(($counter/$($MSOLUSERS.Count))*100)
	#
		$values = @()  
		foreach ($objLicence in $User.Licenses)
		{
			If ($objLicence.AccountSkuid.toString() -eq $license.AccountSkuId)
			{
				$values += $user.UserPrincipalName
				$values += $user.DisplayName
				
				$var = $null
				$var = $Sku.Item($objLicence.AccountSku.SkuPartNumber)
				If ($var -eq $null)
				{
					 $values += $objLicence.AccountSku.SkuPartNumber
				}
				else
				{
					$values += $var
				}        
				
				#write-host ("Processing " + $user.UserPrincipalName)
				$datastring = $user.userprincipalname + "," + $user.DisplayName + "," + $Sku.Item($objLicence.AccountSku.SkuPartNumber)
				foreach ($row in $($objLicence.servicestatus))
				{
					# Build data string
					$datastring = ($datastring + "," + $($row.provisioningstatus))
					$Values += $row.ProvisioningStatus
				}
				
				GenerateObject -LicenseNumber $CounterLicense -header $headers -Value $Values
				$null

				#Out-File -FilePath $File -InputObject $datastring -Encoding UTF8 -append
			}
		}
	}
	#
	Write-Progress -Activity "Processing users Licenses" -Status "Processing $($license.AccountSkuID)" -CurrentOperation "$($counter) / $($MSOLUSERS.Count)" -Completed
}

$script:ExcelApplication = New-Object -ComObject "Excel.Application"
$script:ExcelApplication.Visible = $true

# Create a new blank document to work with and make the Excel application visible.
$script:ExcelWorkbooks = $ExcelApplication.Workbooks
$script:ExcelWorkbook = $script:ExcelWorkbooks.Add()

#$results = @{}
$variables = Get-Variable -Scope Script | Where-Object {$_.Name -like "MSOLSku*"}
$sheetCounter = 0
foreach ($variable in $variables)
{
	$sheetCounter ++

	$value = $null
	$value = Get-Variable $variable.Name -ValueOnly
	New-ExcelReport -object $value -counter $sheetCounter
	#Write-Host "$($value[0].AccountSku)"
	#$results.Add($value[0].AccountSku,$value)
}

