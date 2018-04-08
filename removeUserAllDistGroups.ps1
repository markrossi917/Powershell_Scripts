# Mark Rossi
# 12/2/2016
# This Program removes email address from all existing Dist Groups

try
{
	# Connects to umusa.net exchange server -------------------------------------------
	$userCredential = Get-Credential
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
	Import-PSSession $Session
	# ---------------------------------------------------------------------------------

	# Enter email address to be deleted
	$emailAddress = Read-Host -Prompt 'Input the email address you want to remove'

	# Puts all umusa dist groups into a list
	$groups = Get-DistributionGroup

	Read-Host -Prompt 'Press Enter to Continue...'

	# Goes through each dist group one at a time
	# removes inputed-contact address if it exitst in any dist group
	foreach ( $dg in $groups )
		{
			try
			{
				Remove-DistributionGroupMember $dg.Name -Member $emailAddress -Confirm:$False -ErrorAction Stop
				Write-Host -Prompt "$emailAddress removed from $dg.Name"
			}
			catch
			{
				Write-Host -Prompt "$emailAddress is not in $dg.Name"
			}
		}
	Read-Host 'Press any key to continue'
}
# If failed connection to the exchange server display this error message
catch
{
	Write-Host "You did not successfully connect to the exchange server"
	Read-Host -Prompt 'Press Any Key to Continue...'
}

# Disconnect from the Exchange Server
Remove-PSSession $Session