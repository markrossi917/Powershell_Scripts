# Mark Rossi
# 12/3/2016
# Adds new user to office 365 while adding their dist groups and info

# Connects to umusa.net office 365 server -------------------------------------------
$userCredential = Get-Credential
Connect-MsolService -Credential $userCredential
# ---------------------------------------------------------------------------------

#ENTERPRISEPACK = E3
#DESKLESSPACK = K1
#EXCHANGEDESKLESS = EXCHANGE KIOSK
#RIGHTSMANAGEMENT = AZURE INFO PROTECTION
#MCOIMP = SKYPE FOR BUSINESS
#STANDARDPACK = E1

$exchangeServer = ""

While ($exit -ne 'y')
{
	#----------------------------------------------------------------------------------
	Do
	{	
		$fName = Read-Host "`nEnter First Name"
		$next = Read-Host "Press Enter To Continue or Enter r to re-enter First Name" 
	}
	While ($next -ne '')
	#----------------------------------------------------------------------------------
	Do
	{
		$lName = Read-Host -Prompt "`nEnter Last Name"
		$next = Read-Host "Press Enter To Continue or Enter r to re-enter Last Name"
	}
	While ($next -ne '')
	#----------------------------------------------------------------------------------
	echo "`n`n`n		Select Department"
	echo "1)EXECUTIVE				7)ACCOUNTING"
	echo "2)BILLING				8)HUMAN RESOURCES"
	echo "3)BUSINESS DEVELOPMENT			9)N/A"
	echo "4)MEDICAL PRACTICES			10)LEGAL"
	echo "5)CLINICAL INTEGRATION			11)OFFICE"
	echo "6)CONTRACTING				12)PROVIDER"

	Do
	{
		$depNum = Read-Host "`nPlease select number and press enter for Department"
		
			switch($depNum)
		{
			1{$depment = "EXECUTIVE" ; $next = ''; continue}
			2{$depment = "BILLING" ; $next = ''; continue}
			3{$depment = "BUSINESS DEVELOPMENT" ; $next = ''; continue}
			4{$depment = "MEDICAL PRACTICES" ; $next = ''; continue}
			5{$depment = "CLINICAL INTEGRATION" ; $next = ''; continue}
			6{$depment = "CONTRACTING" ; $next = ''; continue}
			7{$depment = "ACCOUNTING" ; $next = ''; continue}
			8{$depment = "HUMAN RESOURCES" ; $next = ''; continue}
			9{$depment = "N/A" ; $next = ''; continue}
			10{$depment = "LEGAL" ; $next = ''; continue}
			11{$depment = "OFFICE" ; $next = ''; continue}
			12{$depment = "PROVIDER" ; $next = ''; continue}

			Default{Write-Host "`nPlease Enter Valid Number"; $next = "loop"}
		}
		if ($next -ne "loop")
		{
		$next = Read-Host "`nYou have selected $depment : Press Enter to Continue or retry{enter} to go back"
		}
	}
	While ($next -ne '')
	#----------------------------------------------------------------------------------
	echo "`n`n`n			Select Title"
	echo "1)Account Payable				10)IT Specialist"
	echo "2)Accounts Receivables				11)MA"
	echo "3)Business Development				12)Nurse"
	echo "4)Care Coordinator				13)Office Manager"
	echo "5)Credentialing Specialist			14)OR Tech"
	echo "6)Executive Assistant				15)Payment Postings"
	echo "7)Front Desk					16)PROVIDER"
	echo "8)Human Resources				17)Surgical Tech"
	echo "9)Intern					18)Web Designer"

	Do
	{	
		$titleNum = Read-Host "`nPlease select number and press enter for Title"
		
			switch($titleNum)
		{
			1{$t = "Account Payable" ; $next = ''; continue}
			2{$t = "Accounts Receivables" ; $next = ''; continue}
			3{$t = "Business Development" ; $next = ''; continue}
			4{$t = "Care Coordinator" ; $next = ''; continue}
			5{$t = "Credentialing Specialist" ; $next = ''; continue}
			6{$t = "Executive Assistant" ; $next = ''; continue}
			7{$t = "Front Desk" ; $next = ''; continue}
			8{$t = "Human Resources" ; $next = ''; continue}
			9{$t = "Intern" ; $next = ''; continue}
			10{$t = "IT Specialist" ; $next = ''; continue}
			11{$t = "MA" ; $next = ''; continue}
			12{$t = "Nurse" ; $next = ''; continue}
			13{$t = "Office Manager" ; $next = ''; continue}
			14{$t = "OR Tech" ; $next = ''; continue}
			15{$t = "Payment Postings" ; $next = ''; continue}
			16{$t = "PROVIDER" ; $next = ''; continue}
			17{$t = "Surgical Tech" ; $next = ''; continue}
			18{$t = "Web Designer" ; $next = ''; continue}
			Default{Write-Host "`nPlease Enter Valid Number"; $next = "loop"}
		}
		if ($next -ne "loop")
		{
		$next = Read-Host "`nYou have selected $t : Press Enter to Continue or retry{enter} to go back"
		}
	}
	While ($next -ne '')
	#----------------------------------------------------------------------------------
	Do
	{
		$displayName = Read-Host -Prompt "`nEnter Display Name"
		$next = Read-Host "Press Enter To Continue or Enter r to re-enter Display Name"
	}
	While ($next -ne '')
	#----------------------------------------------------------------------------------
	Do
	{
		$emailAddress = Read-Host -Prompt "`nEnter Email Address"
		$next = Read-Host "Press Enter To Continue or Enter r to re-enter Email Address"
		
		$emails = Get-MsolUser
		foreach ($L in $emails) 
		{
			$liveemail = $L.UserPrincipalName
			if ($emailAddress -eq $liveemail)
			{
				echo "Email Address already exist"
				$next = "loop"
			}
		}
	}
	While ($next -ne '')

	#----------------------------------------------------------------------------------
	Do
	{
		$pword = Read-Host -Prompt "`nEnter Password"
		$next = Read-Host "Press Enter to Continue: or any other Key to re-enter password"
	}
	While ($next -ne '')

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
	#add license---------------------------------------------------------------------
	Do
	{
		$licenseAssignment = Read-Host -Prompt "`nEnter License as Above" 
		
		switch($licenseAssignment)
		{
			"EXCHANGEDESKLESS" {$next = ''; continue}
			"RIGHTSMANAGEMENT" {$next = ''; continue}
			"MCOIMP" {$next = ''; continue}
			"ENTERPRISEPACK" {$next = ''; continue}
			"DESKLESSPACK" {$next = ''; continue}
			"STANDARDPACK" {$next = ''; continue}
			default {Write-Host "Please Re-enter the correct license"; $next = "loop"}	
		}
	}
	While ($next -ne '')
	#----------------------------------------------------------------------------------

	Read-Host "Do you want to Create the Email account?"

	#Add User to Office 365 with info--------------------------------------------------
	try
	{
	New-MsolUser -DisplayName $displayName -FirstName $fName -LastName $lName -UserPrincipalName $emailAddress -UsageLocation US -LicenseAssignment UMUSA1:$licenseAssignment -Password $pword -StreetAddress "161 Becks Woods Drive" -State "DE" -PostalCode "19701" -PhoneNumber "302-266-9166" -Fax "302-451-5614" -Office "Beckswoods" -Department $depment -Country "US" -City "Bear" -Title $t
	}
	catch
	{
	Read-Host "You don goof'ed"
	}
	#----------------------------------------------------------------------------------

	Start-Sleep -s 5

	#Connect to Exchange Server--------------------------------------------------------
	if($exchangeServer -eq ""){
	$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $userCredential -Authentication Basic -AllowRedirection
	Import-PSSession $ExchangeSession
	}

	#Adds user to Dist groups and adds company-----------------------------------------
	while ($next -eq '')
	{
		Start-Sleep -s 60
		try
		{
			switch($t)
			{
				"Accounts Receivables" {Add-DistributionGroupMember -Identity "UM - Accounts Receivables" -Member $emailAddress -ErrorAction Stop
				Add-DistributionGroupMember -Identity "United Users" -Member $emailAddress -ErrorAction Stop
				Set-User $emailAddress -Company "United Medical LLC"
				echo "User added to United Users Dist group and company set to United Medical LLC"
				$next = "break" ; continue}
				
				"Coding" {Add-DistributionGroupMember -Identity "UM - Coding" -Member $emailAddress -ErrorAction Stop
				Add-DistributionGroupMember -Identity "United Users" -Member $emailAddress -ErrorAction Stop
				Set-User $emailAddress -Company "United Medical LLC"
				echo "User added to United Users Dist group and company set to United Medical LLC"
				$next = "break" ; continue}
				
				"Human Resources" {Add-DistributionGroupMember -Identity "UM - Human Resources" -Member $emailAddress -ErrorAction Stop
				Add-DistributionGroupMember -Identity "United Users" -Member $emailAddress -ErrorAction Stop
				Set-User $emailAddress -Company "United Medical LLC"
				echo "User added to United Users Dist group and company set to United Medical LLC"
				$next = "break" ; continue}
				
				"Business Development" {Add-DistributionGroupMember -Identity "UM-Business Development" -Member $emailAddress -ErrorAction Stop
				Add-DistributionGroupMember -Identity "United Users" -Member $emailAddress -ErrorAction Stop
				Set-User $emailAddress -Company "United Medical LLC"
				echo "User added to United Users Dist group and company set to United Medical LLC"
				$next = "break" ; continue}
				
				"Care Coordinator" {Add-DistributionGroupMember -Identity "UM-Care Coordinators" -Member $emailAddress -ErrorAction Stop
				Add-DistributionGroupMember -Identity "United Users" -Member $emailAddress -ErrorAction Stop
				Set-User $emailAddress -Company "United Medical LLC"
				echo "User added to United Users Dist group and company set to United Medical LLC"
				$next = "break" ; continue}
				
				"Claims" {Add-DistributionGroupMember -Identity "UM-Claims Department" -Member $emailAddress -ErrorAction Stop
				Add-DistributionGroupMember -Identity "United Users" -Member $emailAddress -ErrorAction Stop
				Set-User $emailAddress -Company "United Medical LLC"
				echo "User added to United Users Dist group and company set to United Medical LLC"
				$next = "break" ; continue}
				
				default{
				Add-DistributionGroupMember -Identity "United Users" -Member $emailAddress -ErrorAction Stop
				echo "User added to United Users Dist group and company set to United Medical LLC"
				$next = "break" ; continue
				}
			}
		}
		catch
		{
			echo "User has not yet synced with exchange server: Sleep for 1 minute"
			$next = ''
		}
	}
	$exchangeServer = "yessir"
	$exit = Read-Host -Prompt "Press Enter to add another user or type y to exit"
}
#----------------------------------------------------------------------------------
