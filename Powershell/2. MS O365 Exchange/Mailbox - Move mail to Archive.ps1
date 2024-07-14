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

#Run Managed Folder Assistant to reduce mailbox size by archiving - Dont use quotes for Identity 
Start-ManagedFolderAssistant -Identity "mailbox@domain.com" 
