#requires -runasadministrator
#defines system box windows assembly type without input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

set-executionpolicy remotesigned

$usercredential=Get-Credential -message "please confirm your Office365 Admin account (if different)" ("$env:username" + "@fia.com")
#New Mailbox name (UPN,Email) 
$PrimaryMailboxEmail=[Microsoft.VisualBasic.Interaction]::InputBox("please input your Mailbox Account to delegate email","New user or shared mailbox email","Accountingfia.com")
$DelegateEmail =[Microsoft.VisualBasic.Interaction]::InputBox("please input your user email that need delegation to mailbx","New user or shared mailbox email","JDoe@fia.com")


$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
import-pssession $Session


get-mailbox $PrimaryMailboxEmail | set-mailbox -ForwardingSmtpAddress $null
