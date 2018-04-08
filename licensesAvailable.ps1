# Mark Rossi
# 12/2/2016
# This script shows the available licenses for Office 365 Admin

# Connects to umusa.net office 365 server -------------------------------------------
$userCredential = Get-Credential
Connect-MsolService -Credential $userCredential
# ---------------------------------------------------------------------------------

#Parse and calculate licenses-----------------------------------------------------
$AccountSku = Get-MsolAccountSku
$count = 0
foreach ($L in $AccountSku) 
{
	$ActiveLic = $L.ActiveUnits
	$ConsumedLic = $L.ConsumedUnits
	$AvailLic = $ActiveLic - $ConsumedLic
	switch ($count){
		0 {echo "`nThere are $AvailLic available E3 License: Enter 'ENTERPRISEPACK'"}
		1 {echo "`nThere are $AvailLic available K1 License: Enter 'DESKLESSPACK'"}
		2 {echo "`nThere are $AvailLic available Exchange License: Enter 'EXCHANGEDESKLESS'"}
		3 {echo "`nThere are $AvailLic available Azure License: Enter 'RIGHTSMANAGEMENT'"}
		4 {echo "`nThere are $AvailLic available SKYPE License: Enter 'MCOIMP'"}
		5 {echo "`nThere are $AvailLic available E1 License: Enter 'STANDARDPACK'"}
		default {continue}
	}
	$count++
}

Read-Host "Continue?"