# Open PowerShell as Admin
#Install the module
Install-Module -Name ExchangeOnlineManagement

#Bypass script execution policy
Set-ExecutionPolicy -executionpolicy bypass

#Import the module
Import-Module ExchangeOnlineManagement

#Connect to Exchange Online
Connect-ExchangeOnline

#Check Storage of Mailbox
Get-MailboxStatistics -Identity "mailbox@domain.com" | Select-Object DisplayName, TotalItemSize, ItemCount

#Verify that auto-expanding archiving is enabled for organisation. True indicates its enabled
Get-OrganizationConfig | FL AutoExpandingArchiveEnabled

#Verify that auto-expanding archiving is enabled for a mailbox. True indicates its enabled
Get-Mailbox "mailbox@domain.com" | FL AutoExpandingArchiveEnabled

#Enable auto-expanding archiving is enabled for mailbox
Enable-Mailbox "mailbox@domain.com" -AutoExpandingArchive