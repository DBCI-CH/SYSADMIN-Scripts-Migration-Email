
#requires -runasadministrator

#defines system box windows assembly type without input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

set-executionpolicy remotesigned

$usercredential=Get-Credential -message "please confirm your Office365 Admin account (if different)" ("$env:username" + "@fia.com")
#New Mailbox name (UPN,Email) 
$PrimaryMailboxEmail=[Microsoft.VisualBasic.Interaction]::InputBox("please input your Mailbox Account email","New user or shared mailbox email","Accountingfia.com")
$DelegateRights =[Microsoft.VisualBasic.Interaction]::InputBox("Please enter an Access Rights", "Validate Editor, PublishingEditor or Reviewer", "Reviewer")
$Delegate =[Microsoft.VisualBasic.Interaction]::InputBox("please input your Delegate  Mailbox Account email","New user or shared mailbox email","JDoe@fia.com")


$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
import-pssession $Session

Add-MailboxFolderPermission -Identity $PrimaryMailboxEmail':\Calendar' -User $delegate -AccessRights $DelegateRights 
add-MailboxFolderPermission -Identity $PrimaryMailboxEmail':\Calendrier' -User $delegate -AccessRights $DelegateRights 