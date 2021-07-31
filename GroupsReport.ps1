<#	
	.NOTES
	===========================================================================
	 Created on:   	7/31/2021 3:02 PM
	 Created by:   	MoshMan125
	===========================================================================
	.DESCRIPTION
		Exports a report containing distribution groups, their memebers including name and email address, and the group owners.
#>
function MSConnect
{
	if ((Get-Module -Name ExchangeOnlineManagement) -eq $null)
	{
		try
		{
			if ((Get-Module -Name PowerShellGet | Select-Object -ExpandProperty version) -lt '2.2.5')
			{
				Install-Module PowershellGet -Force
			}
			
			Install-Module -Name ExchangeOnlineManagement -Force
			Connect-ExchangeOnline
		}
		catch
		{
			(New-Object -COM WScript.Shell).PopUp("Failed to install MicrosoftTeams module, please install manually and try again.", 0, "Error", 48)
			exit
		}
	}
	else
	{
		Connect-ExchangeOnline
	}
	
}

# connect to exchange
MSConnect

# report on distribution groups
foreach ($group in (Get-DistributionGroup))
{
	$groupMembers = Get-DistributionGroupMember -Identity $group.PrimarySMTPAddress
	New-Object -TypeName PSObject -Property @{
		Group			  = $group.Name
		GroupEmailAddress = $group.PrimarySMTPAddress
		Member		      = ([string]$groupMembers.Name -replace " ",",")
		MemberAddress	  = ([string]$groupMembers.PrimarySMTPAddress -replace " ", ",")
		ManagedBy		  = $group.ManagedBy
		Type			  = 'DistributionGroup'
	} | Export-Csv -Path $env:USERPROFILE\Desktop\MSGroups.csv -NoTypeInformation -Append
}

# report on O365 groups
foreach ($group in (Get-UnifiedGroup))
{
	$groupMembers = Get-UnifiedGroupLinks -LinkType Members -Identity $group.PrimarySMTPAddress
	New-Object -TypeName PSObject -Property @{
		Group			  = $group
		GroupEmailAddress = $group.PrimarySMTPAddress
		Member		      = ([string]$groupMembers.Name -replace " ", ",")
		MemberAddress	  = ([string]$groupMembers.PrimarySMTPAddress -replace " ", ",")
		ManagedBy		  = $group.ManagedBy
		Type			  = $group.GroupType			
	} | Export-Csv -Path $env:USERPROFILE\Desktop\MSGroups.csv -NoTypeInformation -Append
}